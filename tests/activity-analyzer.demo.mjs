#!/usr/bin/env node
/**
 * Activity-Analyzer — Demo- & Smoke-Test (Layer 2).
 *
 * Klassifiziert Beispiel-Items aus dem Activity-Feed in Kategorien und erzeugt einen
 * täglichen Markdown-Zusammenfassungs-Bericht unter reports/activity-summary-YYYY-MM-DD.md
 * (Top-3 je Kategorie, priorisiert nach Relevanz).
 *
 * Verwendung:
 *   node tests/activity-analyzer.demo.mjs                 # Mock-Beispieldaten -> reports/
 *   node tests/activity-analyzer.demo.mjs --out /tmp/x    # Berichts-Verzeichnis übersteuern
 */
import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs';
import { analyzeItems, buildReportText, writeReport, runActivityReport, CATEGORIES } from '../src/activityAnalyzer.js';

const outDir = (() => {
  const idx = process.argv.indexOf('--out');
  return idx !== -1 && process.argv[idx + 1] ? process.argv[idx + 1] : process.cwd();
})();

// Realistische Beispiel-Items (Struktur wie aus activityClient.getActivity)
const sampleItems = [
  { text: 'Yannick Bülter erwähnt @David Hallmann: Kannst du die Spec für Wildau nochmal prüfen?', author: 'Yannick Bülter', timestamp: '2026-09-16T09:12:00Z' },
  { text: 'Lukas Behncke antwortete auf Ihre Nachricht: Bitte bis Freitag liefern', author: 'Lukas Behncke', timestamp: '2026-09-16T08:45:00Z' },
  { text: 'Mathis Künzel hat auf Ihren Beitrag reagiert 👍', author: 'Mathis Künzel' },
  { text: 'Theys Schiller: Blocker - Deployment auf dem Testsystem funktioniert nicht', author: 'Theys Schiller', timestamp: '2026-09-16T07:30:00Z' },
  { text: 'Marc hat die Freigabe zur Abnahme erteilt', author: 'Marc', timestamp: '2026-09-16T06:15:00Z' },
  { text: 'Einladung zum Meeting: Kickoff Projekt Wildau am Freitag 10:00', author: 'Amir Rocker', timestamp: '2026-09-16T05:00:00Z' },
  { text: 'Neuer Kanal "Ankündigungen" wurde in Allgemein erstellt' },
  { text: 'Copilot: Erinnerung an die Woche steht an - bitte deadl ine prüfen' },
  { text: 'Alex: Entscheidung steht, wir gehen mit Variante B' , author: 'Alex', timestamp: '2026-09-16T04:00:00Z' }
];

console.log('=== Teams Activity-Analyzer (Layer 2) — Demo ===\n');

// 1) Klassifikation je Item
console.log('--- 1) Einzel-Klassifikation ---');
for (const it of sampleItems) {
  const { category, matched } = (await import('../src/activityAnalyzer.js')).classify(it.text);
  console.log(`[${category.padEnd(11)}] ${it.text} ${matched.length ? '(→ ' + matched.join(',') + ')' : ''}`);
}

// 2) Gruppierung + Scores
console.log('\n--- 2) Gruppen je Kategorie (mit Relevanz-Score) ---');
const groups = analyzeItems(sampleItems);
for (const cat of CATEGORIES) {
  const entries = groups[cat] || [];
  console.log(`\n## ${cat} (${entries.length})`);
  for (const e of entries.slice(0, 3)) {
    console.log(`   [${e.priority}] score=${e.score} | ${e.item.text}${e.item.author ? ' — ' + e.item.author : ''}`);
  }
}

// 3) Markdown-Bericht erzeugen + als Datei speichern
console.log('\n--- 3) Markdown-Bericht (Top-3 je Kategorie) ---');
const date = '2026-09-16';
const md = buildReportText(groups, { topN: 3, date, tenant: 'adesso' });
console.log(md);

const { filePath } = writeReport(md, { date, dir: outDir });
console.log(`\n📄 Bericht gespeichert: ${filePath}`);

// Kurz-Smoke: komponierter Lauf runActivityReport
console.log('\n--- 4) komponierter Lauf runActivityReport ---');
const out = runActivityReport(sampleItems, { date, dir: outDir, tenant: 'adesso' });
console.log(JSON.stringify({ filePath: out.filePath, date: out.date, counts: out.counts, exists: fs.existsSync(out.filePath) }, null, 2));
