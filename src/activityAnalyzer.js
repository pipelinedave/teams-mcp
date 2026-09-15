import fs from 'fs';
import path from 'path';
import { activityClient } from './activityClient.js';

/**
 * Layer-2-Analyse für den Teams Activity-Feed.
 *
 * Klassifiziert die von activityClient extrahierten Roh-Items in fünf Kategorien:
 *   - Meeting           (Termin/Abstimmung/Besprechung)
 *   - Task              (Aufgabe/Anfrage/To-do)
 *   - Entscheidung      (Entscheidung/Zusage/Freigabe)
 *   - Risiko            (Risiko/Blocker/Probleme/Fristrisiko)
 *   - Sonstiges         (Rest, z.B. Reaktionen/Kanal-Events/Rauschen)
 *
 * Erzeugt daraus einen täglichen Zusammenfassungs-Bericht (Markdown) mit den
 * Top-3-Einträgen pro Kategorie, priorisiert nach Relevanz.
 *
 * Design:
 *   - Pure, exportierte Funktionen (classify, buildReportText, writeReport) für
 *     Browser-freie Unit-Tests ohne I/O-Nebenwirkungen.
 *   - Injizierbarer Output-Pfad (Reports-Verzeichnis) für Testbarkeit.
 *   - Logging über stderr (verschmutzt den Stdio-JSON-RPC-Kanal nicht).
 */

export const CATEGORIES = ['Meeting', 'Task', 'Entscheidung', 'Risiko', 'Sonstiges'];

export const PRIORITY = {
  HOCH: 'hoch',
  MITTEL: 'mittel',
  NIEDRIG: 'niedrig'
};

// Relevanz-Punkte je Kategorie-Treffer (hoch = dringlicher/aktionpflichtig).
// Reihenfolge der Einträge bestimmt die Präzedenz bei Mehrfach-Matches: frühere
// Kategorien gewinnen, wenn ein Item mehrere Kategorien trifft.
const CLASSIFIER = [
  {
    category: 'Risiko',
    priority: PRIORITY.HOCH,
    keywords: ['risiko', 'blocker', 'blokkiert', 'blockiert', 'funktioniert nicht', 'fehler', 'defekt', 'probleme', 'problem', 'eskalation', 'kritisch', 'dringend', 'deadline', 'frist', 'verzögert', 'verschoben', 'nicht möglich', 'fail', 'error', 'stoppt', 'hängt', 'hängt sich']
  },
  {
    // Entscheidung/Zusage/Freigabe steht vor Task, da eine getroffene Entscheidung
    // präziser und handlungsrelevanter ist als eine reine Aufgaben-Anfrage.
    category: 'Entscheidung',
    priority: PRIORITY.MITTEL,
    keywords: ['entscheidung', 'entschieden', 'zusage', 'freigabe', 'freigegeben', 'approve', 'approved', 'bestätigt', 'bestätigen', 'ok so', 'abgenommen', 'go ahead', 'wir machen', 'wir gehen', 'plan steht', 'so machen wir', 'abnahme', 'erledigt']
  },
  {
    category: 'Task',
    priority: PRIORITY.HOCH,
    keywords: ['kannst du', 'könntest du', 'bitte', 'aufgabe', 'anfrage', 'to-do', 'todo', 'to do', 'erledigen', 'prüfen', 'anschauen', 'review', 'bearbeiten', 'schick mir', 'schicken', 'remember', 'dranbleiben', 'nachfassen', 'erinnerung', 'bis freitag', 'bis montag', 'bis morgen', 'bitte um', 'kann mal', 'wäre gut', 'should', 'must', 'pls', 'please', 'deliver', 'umsetzen', 'implementieren']
  },
  {
    category: 'Meeting',
    priority: PRIORITY.MITTEL,
    keywords: ['meeting', 'termin', 'besprechung', 'abstimmung', 'call', 'videokonferenz', 'sync', 'daily', 'weekly', 'kickoff', 'standup', 'retro', 'demo', 'workshop', 'telefonat', 'verabredung', 'einladen', 'join'] 
  }
];

/**
 * Klassifiziert einen einzelnen Activity-Eintrag (Text) in eine Kategorie mit
 * Priorität und konkreten Treffern. Reiner Text-basierter Klassifikator.
 *
 * @param {string} text - Roher Activity-Text
 * @param {object} [opts]
 * @param {boolean} [opts.lowercase] - Text vorab lowercase anwenden (Standard true)
 * @returns {{category: string, priority: string, matched: string[]}}
 */
export function classify(text, { lowercase = true } = {}) {
  if (!text || typeof text !== 'string') {
    return { category: 'Sonstiges', priority: PRIORITY.NIEDRIG, matched: [] };
  }
  const haystack = lowercase ? text.toLowerCase() : text;

  for (const rule of CLASSIFIER) {
    const hit = [];
    for (const kw of rule.keywords) {
      if (haystack.includes(kw)) hit.push(kw);
    }
    if (hit.length > 0) {
      return {
        category: rule.category,
        priority: rule.priority,
        matched: hit
      };
    }
  }

  return { category: 'Sonstiges', priority: PRIORITY.NIEDRIG, matched: [] };
}

/**
 * Relevanz-Score für die Top-N-Priorisierung innerhalb einer Kategorie.
 * Höhere Punkte = relevanter (Priorität + Autor-Bekanntheit + Zeitstempel-Frische).
 *
 * @param {object} item - Activity-Item {text, author, timestamp, ...}
 * @param {object} classification - Ergebnis von classify()
 * @returns {number}
 */
export function relevanceScore(item, classification) {
  let score = 0;
  score += classification.priority === PRIORITY.HOCH ? 100 : classification.priority === PRIORITY.MITTEL ? 50 : 10;

  // Autor vorhanden -> Kontext nutzbar (vs. anonymes Kanal-Rauschen)
  if (item.author) score += 20;

  // Frischer Zeitstempel (innerhalb der letzten 24h) -> höhere Relevanz
  if (item.timestamp) {
    const ts = new Date(item.timestamp).getTime();
    if (!isNaN(ts)) {
      const ageMs = Date.now() - ts;
      if (ageMs >= 0 && ageMs < 24 * 60 * 60 * 1000) score += 30;
      else if (ageMs >= 0 && ageMs < 72 * 60 * 60 * 1000) score += 10;
    }
  }

  // Mehr Keywords-Treffer -> relevanter
  score += (classification.matched?.length || 0) * 5;
  return score;
}

/**
 * Klassifiziert eine Liste von Activity-Items und gruppiert sie nach Kategorie,
 * je Kategorie absteigend nach Relevanz sortiert.
 *
 * @param {Array<{text: string, author?: string, timestamp?: string}>} items
 * @returns {Object<string, Array<{item, score, category, priority, matched}>>}
 */
export function analyzeItems(items = []) {
  const groups = Object.fromEntries(CATEGORIES.map((c) => [c, []]));

  for (const item of items) {
    if (!item || !item.text) continue;
    const cls = classify(item.text);
    groups[cls.category].push({
      item,
      score: relevanceScore(item, cls),
      ...cls
    });
  }

  // Je Kategorie absteigend nach Relevanz sortieren
  for (const cat of Object.keys(groups)) {
    groups[cat].sort((a, b) => b.score - a.score);
  }
  return groups;
}

/**
 * Erzeugt den Markdown-Berichtstext aus klassifizierten Gruppen.
 * Priorisiert die Top-N Einträge je Kategorie.
 *
 * @param {object} groups - Ergebnis von analyzeItems()
 * @param {object} [opts]
 * @param {number} [opts.topN] - Top-N Einträge je Kategorie (Standard 3)
 * @param {string} [opts.date] - Berichtsdatum (YYYY-MM-DD, Standard heute)
 * @param {string} [opts.tenant] - Tenant-Name für den Report
 * @returns {string} Markdown
 */
export function buildReportText(groups, { topN = 3, date, tenant = 'adesso' } = {}) {
  const reportDate = date || new Date().toISOString().slice(0, 10);
  const lines = [];

  lines.push(`# Teams Activity-Zusammenfassung — ${reportDate}`);
  lines.push('');
  lines.push(`> Automatisch erzeugt aus dem Teams Activity-Feed (tenant: ${tenant}).`);
  lines.push(`> Kategorien priorisiert; je Kategorie die Top-${topN} relevantesten Einträge.`);
  lines.push('');

  // Gesamtzähler
  const totals = {};
  for (const cat of CATEGORIES) {
    totals[cat] = groups[cat]?.length || 0;
  }
  const total = Object.values(totals).reduce((a, b) => a + b, 0);
  lines.push(`**Gesamt:** ${total} Einträge`);
  lines.push('');
  lines.push(`| Kategorie | Anzahl |`);
  lines.push(`|-----------|--------|`);
  for (const cat of CATEGORIES) {
    lines.push(`| ${cat} | ${totals[cat]} |`);
  }
  lines.push('');

  let anyContent = false;
  for (const cat of CATEGORIES) {
    const entries = groups[cat] || [];
    lines.push(`## ${cat} (${entries.length})`);
    const top = entries.slice(0, topN);

    if (top.length === 0) {
      lines.push('_Keine Einträge._');
      lines.push('');
      continue;
    }

    for (const e of top) {
      anyContent = true;
      const author = e.item.author ? ` — *${e.item.author}*` : '';
      const ts = e.item.timestamp ? ` | \`${e.item.timestamp}\`` : '';
      const prioBadge = `\`${e.priority}\``;
      lines.push(`- **[${e.category}] ${prioBadge}** ${e.item.text}${author}${ts}`);
    }

    // Falls mehr als topN vorhanden, Rest als Zähler-Block anfügen
    if (entries.length > topN) {
      lines.push(`  - _… und ${entries.length - topN} weitere Einträge in dieser Kategorie_`);
    }
    lines.push('');
  }

  if (!anyContent && total === 0) {
    lines.push('_In diesem Zeitraum wurden keine Aktivitäten erfasst._');
    lines.push('');
  }

  return lines.join('\n');
}

/**
 * Schreibt den Bericht als Markdown-Datei nach reports/activity-summary-YYYY-MM-DD.md.
 *
 * @param {string} markdown - Berichtstext (aus buildReportText)
 * @param {object} [opts]
 * @param {string} [opts.date] - Berichtsdatum (YYYY-MM-DD, Standard heute)
 * @param {string} [opts.dir] - Basis-Verzeichnis (Standard: CWD des Aufrufers)
 * @returns {{filePath: string, date: string}}
 */
export function writeReport(markdown, { date, dir = process.cwd() } = {}) {
  const reportDate = date || new Date().toISOString().slice(0, 10);
  const reportsDir = path.join(dir, 'reports');
  fs.mkdirSync(reportsDir, { recursive: true });
  const filePath = path.join(reportsDir, `activity-summary-${reportDate}.md`);
  fs.writeFileSync(filePath, markdown, 'utf8');
  return { filePath, date: reportDate };
}

/**
 * Vollständiger Analyse- & Berichts-Lauf (komponiert aus den puren Funktionen).
 *
 * @param {Array<object>} items - Activity-Items (aus activityClient)
 * @param {object} [opts]
 * @param {number} [opts.topN]
 * @param {string} [opts.date]
 * @param {string} [opts.tenant]
 * @param {string} [opts.dir]
 * @param {object} [opts.logger]
 * @returns {{summary, groups, report: string, filePath: string, date: string, counts: object}}
 */
export function runActivityReport(items = [], opts = {}) {
  const logger = opts.logger || {
    info: (m, meta) => console.error(`[analyzer] INFO  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    warn: (m, meta) => console.error(`[analyzer] WARN  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    error: (m, meta) => console.error(`[analyzer] ERROR ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`)
  };

  const groups = analyzeItems(items);
  const report = buildReportText(groups, opts);
  const { filePath, date } = writeReport(report, opts);

  const counts = {};
  for (const cat of CATEGORIES) counts[cat] = groups[cat]?.length || 0;

  logger.info(`Bericht erstellt`, { filePath, date, total: items.length, counts });
  return { summary: report, groups, report, filePath, date, counts };
}

/**
 * Vollständiger End-to-End-Lauf: extrahiert den Activity-Tab (Layer 1) über
 * activityClient, klassifiziert die Items (Layer 2) und erzeugt den Bericht.
 *
 * Dies ist der primäre Einstiegspunkt für den Scheduler (activityReportScheduler).
 * Liefert Roh-Items, Gruppierung, Counts und den Markdown-Report — OHNE den Report
 * bereits zu senden (Sendung übernimmt sendProactiveReport im Scheduler).
 *
 * @param {string} tenant - Tenant-Name/Realm (z.B. "adesso")
 * @param {object} [opts]
 * @param {number} [opts.maxItems] - Maximale Anzahl extrahierter Items (Standard 50)
 * @param {number} [opts.topN] - Top-N je Kategorie im Bericht (Standard 3)
 * @param {string} [opts.date] - Berichtsdatum YYYY-MM-DD (Standard heute)
 * @param {string} [opts.dir] - Basis-Verzeichnis für reports/ (Standard CWD)
 * @param {object} [opts.logger] - Injizierbarer Logger
 * @returns {Promise<{groups, counts, items, report: string, filePath: string, date: string, tenant: string}>}
 */
export async function analyzeActivityTab(tenant = '', opts = {}) {
  const logger = opts.logger || {
    info: (m, meta) => console.error(`[analyzer] INFO  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    warn: (m, meta) => console.error(`[analyzer] WARN  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    error: (m, meta) => console.error(`[analyzer] ERROR ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`)
  };

  // Layer 1: Rohdaten aus dem Activity-Tab extrahieren
  const activity = await activityClient.getActivity(tenant, { maxItems: opts.maxItems || 50 });

  // Layer 2: klassifizieren + gruppieren
  const groups = analyzeItems(activity.items);

  const dates = { ...opts };
  if (!dates.date) dates.date = activity.extractedAt ? activity.extractedAt.slice(0, 10) : undefined;
  const report = buildReportText(groups, dates);

  const { filePath, date } = writeReport(report, opts);

  const counts = {};
  for (const cat of CATEGORIES) counts[cat] = groups[cat]?.length || 0;

  logger.info(`Analyse Activity-Tab abgeschlossen`, {
    tenant: activity.tenant,
    total: activity.count,
    filePath,
    counts
  });

  return {
    tenant: activity.tenant,
    items: activity.items,
    groups,
    counts,
    report,
    filePath,
    date
  };
}
