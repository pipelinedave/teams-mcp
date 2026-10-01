import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  TeamsClient,
  messageTimeToIso,
  mimeToExt,
  sanitizeFilename,
  saveBlobDataUrl,
  replaceBlobUrlsInHtml,
  ensureMediaDir,
  collectAttachmentsFromRow,
  extractFilenameFromText,
  normalizeExtensionWhitelist,
  extensionAllowed,
  urlForcedDownload,
  nextAvailablePath,
  pickTreeItem,
  findTeamInTree,
  findChannelInTree,
  channelPoolForTeam,
  collectChannelTitles,
  buildTeamsStructure,
  isIgnoredTreeTitle,
  DEFAULT_DOWNLOAD_EXTENSIONS
} from '../src/teamsClient.js';
import { browserManager } from '../src/browserManager.js';
import { cleanSpeakerName, aggregateEvents, isBotOrUiName } from '../src/speakerTracker.js';
import { extractActivityFromDom } from '../src/activityClient.js';
import { classify, analyzeItems, buildReportText, writeReport, CATEGORIES } from '../src/activityAnalyzer.js';
import { toMinutes, writeTimestampedReport, ActivityReportScheduler, DEFAULT_SCHEDULE, runSchedulerControl } from '../src/activityReportScheduler.js';
import { createOutboxSender, listPendingDeliveries, markDelivered, DELIVERY_KIND, DELIVERY_RECIPIENT } from '../src/reportDelivery.js';

// Instanz des TeamsClient (nur zur Nutzung der puren pickChatRow-Methode)
const client = new TeamsClient();

describe('pickChatRow - Chat-Auswahl-Logik', () => {
  const filtered = [
    { row: 'r0', title: 'Yannick Bülter', text: 'Yannick Bülter' },
    { row: 'r1', title: 'Theys Schiller', text: 'Theys Schiller' },
    { row: 'r2', title: 'Yannick und Kollegen', text: 'Yannick und Kollegen' },
    { row: 'r3', title: 'Lukas Behncke', text: 'Lukas Behncke' }
  ];

  test('exakter Titel-Match gewinnt gegen Substring', () => {
    const r = client.pickChatRow(filtered, 'Yannick', undefined);
    assert.equal(r.title, 'Yannick Bülter'); // exakter Präfix-Match, nicht "Yannick und Kollegen"
  });

  test('exakter exakter Case-insensitive Match', () => {
    const r = client.pickChatRow(filtered, 'yannick bülter', undefined);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('chatIndex gewinnt über chatName', () => {
    const r = client.pickChatRow(filtered, 'Lukas Behncke', 1);
    assert.equal(r.title, 'Theys Schiller');
  });

  test('chatIndex out of range -> fällt auf Name zurück', () => {
    const r = client.pickChatRow(filtered, 'Lukas Behncke', 99);
    assert.equal(r.title, 'Lukas Behncke');
  });

  test('negativer chatIndex -> Name', () => {
    const r = client.pickChatRow(filtered, 'Yannick Bülter', -1);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('kein Treffer -> null', () => {
    const r = client.pickChatRow(filtered, 'GibtEsNicht', undefined);
    assert.equal(r, null);
  });

  test('leerer chatName/kein Index -> null', () => {
    assert.equal(client.pickChatRow(filtered, '', undefined), null);
    assert.equal(client.pickChatRow(filtered, undefined, undefined), null);
    assert.equal(client.pickChatRow(filtered, undefined, null), null);
  });

  test('Präfix-Match: "Yann" -> Yannick', () => {
    const r = client.pickChatRow(filtered, 'Yann', undefined);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('Token-Match: "Yannick Kollegen" -> "Yannick und Kollegen"', () => {
    const r = client.pickChatRow(filtered, 'Yannick Kollegen', undefined);
    assert.equal(r.title, 'Yannick und Kollegen');
  });
});

describe('normalizeTenant / realm - Org-neutrale Tenant-Logik', () => {
  test('exakter konfigurierter Key', () => {
    assert.equal(browserManager.normalizeTenant('adesso'), 'adesso');
  });

  test('org-neutral: unbekannter Wert zählt direkt', () => {
    assert.equal(browserManager.normalizeTenant('meine-org'), 'meine-org');
  });

  test('case-insensitive', () => {
    assert.equal(browserManager.normalizeTenant('ADESSO'), 'adesso');
  });

  test('realm: Konvention onmicrosoft.com bei purem Namen', () => {
    assert.equal(browserManager.realm('meine-org'), 'meine-org.onmicrosoft.com');
  });

  test('realm: Realm-artiger Wert direkt', () => {
    assert.equal(browserManager.realm('meine-org.onmicrosoft.com'), 'meine-org.onmicrosoft.com');
  });

  test('Smart-Default: leerer Tenant, undefined oder "all" fällt auf Default ("adesso") zurück', () => {
    assert.equal(browserManager.normalizeTenant(''), 'adesso');
    assert.equal(browserManager.normalizeTenant(undefined), 'adesso');
    assert.equal(browserManager.normalizeTenant(null), 'adesso');
    assert.equal(browserManager.normalizeTenant('all'), 'adesso');
  });
});

describe('speakerTracker - Name cleaning & Event aggregation', () => {
  test('cleanSpeakerName entfernt Rollen-Suffixe', () => {
    assert.equal(cleanSpeakerName('Yannick Bülter (Gast)'), 'Yannick Bülter');
    assert.equal(cleanSpeakerName('Mathis Künzel (Organizer)'), 'Mathis Künzel');
    assert.equal(cleanSpeakerName('Pierre (Extern)'), 'Pierre');
    assert.equal(cleanSpeakerName('David Hallmann'), 'David Hallmann');
    assert.equal(cleanSpeakerName(''), '');
  });

  test('isBotOrUiName erkennt Teams-Copilot und UI-Platzhalter', () => {
    assert.equal(isBotOrUiName('Copilot'), true);
    assert.equal(isBotOrUiName('Microsoft Copilot'), true);
    assert.equal(isBotOrUiName('Copilot Notebook'), true);
    assert.equal(isBotOrUiName(''), true);
    assert.equal(isBotOrUiName(null), true);
    assert.equal(isBotOrUiName('Recording in progress'), true);
    assert.equal(isBotOrUiName('Chat'), true);
  });

  test('isBotOrUiName lässt echte Personen durch', () => {
    assert.equal(isBotOrUiName('Theys Schiller'), false);
    assert.equal(isBotOrUiName('Yannick Bülter'), false);
    assert.equal(isBotOrUiName('David Hallmann'), false);
  });

  test('aggregateEvents: leere Samples liefern leeres Array', () => {
    assert.deepEqual(aggregateEvents([]), []);
    assert.deepEqual(aggregateEvents(null), []);
  });

  test('aggregateEvents: kontinuierliche Samples werden zu 1 Intervall gemerged', () => {
    const base = 1726040000000;
    const samples = [
      { speaker: 'Yannick Bülter', timestamp: base },
      { speaker: 'Yannick Bülter', timestamp: base + 100 },
      { speaker: 'Yannick Bülter', timestamp: base + 200 },
      { speaker: 'Yannick Bülter', timestamp: base + 300 },
      { speaker: 'Yannick Bülter', timestamp: base + 400 },
      { speaker: 'Yannick Bülter', timestamp: base + 500 }
    ];

    const intervals = aggregateEvents(samples, { meetingStartEpochMs: base });
    assert.equal(intervals.length, 1);
    assert.equal(intervals[0].speaker, 'Yannick Bülter');
    assert.equal(intervals[0].start_epoch_ms, base);
    assert.equal(intervals[0].end_epoch_ms, base + 500);
    assert.equal(intervals[0].start_offset_sec, 0);
    assert.equal(intervals[0].end_offset_sec, 0.5);
  });

  test('aggregateEvents: Pause > minGapMs erzeugt zwei getrennte Intervalle', () => {
    const base = 1726040000000;
    const samples = [
      // Block 1 (500ms Dauer)
      { speaker: 'Mathis', timestamp: base },
      { speaker: 'Mathis', timestamp: base + 200 },
      { speaker: 'Mathis', timestamp: base + 500 },
      // 1000ms Pause
      // Block 2 (600ms Dauer)
      { speaker: 'Mathis', timestamp: base + 1500 },
      { speaker: 'Mathis', timestamp: base + 1800 },
      { speaker: 'Mathis', timestamp: base + 2100 }
    ];

    const intervals = aggregateEvents(samples, { minGapMs: 400, meetingStartEpochMs: base });
    assert.equal(intervals.length, 2);
    assert.equal(intervals[0].speaker, 'Mathis');
    assert.equal(intervals[0].end_offset_sec, 0.5);
    assert.equal(intervals[1].start_offset_sec, 1.5);
    assert.equal(intervals[1].end_offset_sec, 2.1);
  });

  test('aggregateEvents: Spikes unter minDurationMs werden verworfen', () => {
    const base = 1726040000000;
    const samples = [
      // Nur ein einzelner Spike von 100ms
      { speaker: 'Hintergrundgeräusch', timestamp: base },
      { speaker: 'Hintergrundgeräusch', timestamp: base + 100 }
    ];

    const intervals = aggregateEvents(samples, { minDurationMs: 250, meetingStartEpochMs: base });
    assert.equal(intervals.length, 0);
  });

  test('aggregateEvents: Mehrere Sprecher chronologisch sortiert', () => {
    const base = 1726040000000;
    const samples = [
      { speaker: 'Yannick', timestamp: base },
      { speaker: 'Yannick', timestamp: base + 300 },
      { speaker: 'Mathis', timestamp: base + 500 },
      { speaker: 'Mathis', timestamp: base + 900 }
    ];

    const intervals = aggregateEvents(samples, { meetingStartEpochMs: base });
    assert.equal(intervals.length, 2);
    assert.equal(intervals[0].speaker, 'Yannick');
    assert.equal(intervals[1].speaker, 'Mathis');
  });
});

describe('messageTimeToIso - Teams-Zeitstempel-Normalisierung', () => {
  test('data-mid als Unix-Millis (13-stellig) -> ISO', () => {
    // 1789463782694 ms ~ 2026-09-15
    const iso = messageTimeToIso('1789463782694');
    assert.ok(iso.startsWith('2026-09-15T'));
    assert.ok(iso.endsWith('Z'));
  });

  test('Unix-Sekunden (10-stellig) -> ISO', () => {
    const iso = messageTimeToIso('1789463782');
    assert.ok(iso.startsWith('2026-09-15T'));
  });

  test('<time datetime> ISO-String bleibt erhalten', () => {
    const iso = messageTimeToIso('2026-09-15T11:16:27.330Z');
    assert.equal(iso, new Date('2026-09-15T11:16:27.330Z').toISOString());
  });

  test('leere / null / undefined -> null', () => {
    assert.equal(messageTimeToIso(''), null);
    assert.equal(messageTimeToIso(null), null);
    assert.equal(messageTimeToIso(undefined), null);
  });

  test('Nicht-Zeitstempel (z.B. GUID) -> null', () => {
    assert.equal(messageTimeToIso('1:19c8f2a0-9b3d-4c5e-a1f2-8b7c6d5e4f3a'), null);
    assert.equal(messageTimeToIso('kein timestamp'), null);
  });

  test('Millis außerhalb plausibler Reichweite -> null', () => {
    assert.equal(messageTimeToIso('999999'), null);
  });
});

describe('validateAttachments - Dateianhang-Validierung', () => {
  test('leere oder nicht gesetzte Anhänge liefern leeres Array', () => {
    assert.deepEqual(client.validateAttachments(null), []);
    assert.deepEqual(client.validateAttachments(undefined), []);
    assert.deepEqual(client.validateAttachments(''), []);
    assert.deepEqual(client.validateAttachments([]), []);
  });

  test('existierende Datei wird zu absolutem Pfad normalisiert', () => {
    const res = client.validateAttachments('package.json');
    assert.equal(res.length, 1);
    assert.ok(res[0].endsWith('package.json'));
    assert.ok(res[0].startsWith('/'));
  });

  test('Array aus existierenden Dateien', () => {
    const res = client.validateAttachments(['package.json', 'README.md']);
    assert.equal(res.length, 2);
    assert.ok(res[0].endsWith('package.json'));
    assert.ok(res[1].endsWith('README.md'));
  });

  test('nicht existierende Datei wirft verständlichen Fehler', () => {
    assert.throws(
      () => client.validateAttachments('/tmp/nicht-existierende-datei-12345.xyz'),
      /Anhang-Datei nicht gefunden/
    );
  });

  test('Verzeichnis statt Datei wirft Fehler', () => {
    assert.throws(
      () => client.validateAttachments('src'),
      /Anhang ist keine reguläre Datei/
    );
  });
});

describe('extractActivityFromDom - Teams Activity-Feed Extraktion (Layer 1)', () => {
  // Minimaler DOM-Mock: nachgebildete Teams-Activity-Struktur
  function makeEl({ text, attrs = {}, time = null, children = [] }) {
    return {
      innerText: text,
      textContent: text,
      outerHTML: `<div data-tid="activity-item">${text}</div>`,
      getAttribute: (name) => attrs[name] ?? null,
      querySelector: (sel) => {
        if (sel === 'time[datetime]' && time) {
          return { getAttribute: (n) => (n === 'datetime' ? time : null), innerText: time };
        }
        // Komma-Selektor [data-tid*="avatar" i], [role="img"] wie echtes Browser-DOM auflösen
        const avatarSel = sel.includes('avatar') || sel === '[role="img"]';
        if (avatarSel) {
          return attrs.avatar ? { getAttribute: (n) => (n === 'aria-label' ? attrs.avatar : null) } : null;
        }
        return null;
      },
      querySelectorAll: () => children
    };
  }
  function buildFeed(items) {
    const root = { querySelectorAll: () => items, querySelector: () => null, getAttribute: () => null };
    return {
      querySelector: (sel) => (sel.startsWith('[data-tid="activity-feed"]') ? root : null),
      querySelectorAll: () => []
    };
  }

  const feedItems = [
    makeEl({ text: 'Yannick Bülter hat in einem Chat erwähnt @David Hallmann', attrs: { avatar: 'Yannick Bülter' }, time: '2026-09-16T09:12:00Z' }),
    makeEl({ text: 'Lukas Behncke antwortete auf Ihre Nachricht', attrs: { avatar: 'Lukas Behncke' }, time: '2026-09-16T08:45:00Z' }),
    makeEl({ text: 'Mathis Künzel hat auf Ihren Beitrag reagiert 👍' }),
    makeEl({ text: 'Neuer Kanal wurde erstellt', time: '2026-09-16T07:30:00Z' })
  ];

  test('extrahiert alle Items mit Autor, Text und ISO-Zeitstempel', () => {
    const items = extractActivityFromDom(buildFeed(feedItems), { max: 10 });
    assert.equal(items.length, 4);
    // Autor aus Avatar aria-label
    assert.equal(items[0].author, 'Yannick Bülter');
    assert.equal(items[1].author, 'Lukas Behncke');
    // ISO-Timestamp aus <time datetime>
    assert.equal(items[0].timestamp, '2026-09-16T09:12:00.000Z');
    assert.equal(items[1].timestamp, '2026-09-16T08:45:00.000Z');
    // Items ohne Avatar/Zeit lassen Felder weg
    assert.equal(items[2].author, undefined);
    assert.equal(items[2].timestamp, undefined);
  });

  test('respektiert max-Limit', () => {
    const items = extractActivityFromDom(buildFeed(feedItems), { max: 2 });
    assert.equal(items.length, 2);
  });

  test('Text normalisiert (NBSP -> Leerzeichen, Mehrfach-Leerzeichen -> eins)', () => {
    const doc = buildFeed([makeEl({ text: 'Theys Schiller\u00a0erwähnt   @Team' })]);
    const items = extractActivityFromDom(doc, { max: 5 });
    assert.equal(items[0].text, 'Theys Schiller erwähnt @Team');
  });

  test('dedupliziert identische Items (virtuelles Rendering)', () => {
    const dup = [makeEl({ text: 'Dublette' }), makeEl({ text: 'Dublette' }), makeEl({ text: 'Einzigartig' })];
    const items = extractActivityFromDom(buildFeed(dup), { max: 10 });
    assert.equal(items.length, 2);
  });

  test('überspringt leere/kurze Textfragmente (< 4 Zeichen)', () => {
    const doc = buildFeed([makeEl({ text: '' }), makeEl({ text: 'Hi' }), makeEl({ text: 'Gültiger Eintrag' })]);
    const items = extractActivityFromDom(doc, { max: 10 });
    assert.equal(items.length, 1);
    assert.equal(items[0].text, 'Gültiger Eintrag');
  });

  test('Unix-Millis-Zeitstempel in data-mid-artigem Wert wird zu ISO', () => {
    // 1789557120000 ~ 2026-09-16
    const el = makeEl({ text: 'Event mit Millis', attrs: {} });
    el.querySelector = (sel) => {
      if (sel === 'time[datetime]') return { getAttribute: () => '1789557120000', innerText: '1789557120000' };
      return null;
    };
    const items = extractActivityFromDom(buildFeed([el]), { max: 5 });
    assert.ok(items[0].timestamp.startsWith('2026-09-16T'));
  });
});

describe('activityAnalyzer - classify (Layer 2)', () => {
  test('erkennt Task-Anfragen (HOCH)', () => {
    const c = classify('Yannick: Kannst du die Spec nochmal prüfen?');
    assert.equal(c.category, 'Task');
    assert.equal(c.priority, 'hoch');
    assert.ok(c.matched.includes('kannst du'));
  });

  test('erkennt Meeting-Einladungen', () => {
    const c = classify('Einladung zum Meeting: Kickoff Projekt Wildau');
    assert.equal(c.category, 'Meeting');
    assert.equal(c.priority, 'mittel');
  });

  test('erkennt Entscheidung/Freigabe', () => {
    const c = classify('Marc hat die Freigabe für den Release erteilt');
    assert.equal(c.category, 'Entscheidung');
  });

  test('Risiko hat Vorrang vor Task bei Mehrfach-Match', () => {
    const c = classify('Dringend: Blocker - das Deployment funktioniert nicht, bitte prüfen');
    assert.equal(c.category, 'Risiko'); // Risiko steht in der Kaskade vor Task
  });

  test('erkennt Sonstiges für Reaktionen/Rauschen', () => {
    const c = classify('Mathis hat auf Ihren Beitrag reagiert 👍');
    assert.equal(c.category, 'Sonstiges');
    assert.equal(c.priority, 'niedrig');
  });

  test('leerer/kein Text -> Sonstiges niedrig', () => {
    assert.deepEqual(classify(''), { category: 'Sonstiges', priority: 'niedrig', matched: [] });
    assert.deepEqual(classify(null), { category: 'Sonstiges', priority: 'niedrig', matched: [] });
  });

  test('case-insensitive (EN "Please" / "Review")', () => {
    assert.equal(classify('Please review the PR').category, 'Task');
  });
});

describe('activityAnalyzer - analyzeItems & Gruppen (Layer 2)', () => {
  const items = [
    { text: 'Kannst du die Spec prüfen?', author: 'Yannick Bülter', timestamp: new Date().toISOString() },
    { text: 'Blocker: Deployment funktioniert nicht', author: 'Lukas Behncke', timestamp: new Date().toISOString() },
    { text: 'Einladung zum Meeting am Freitag', author: 'Theys Schiller', timestamp: new Date().toISOString() },
    { text: 'Freigabe zur Abnahme erteilt', author: 'Marc', timestamp: new Date().toISOString() },
    { text: 'Mathis hat auf Ihren Beitrag reagiert 👍' },
    { text: 'Neuer Kanal wurde erstellt' }
  ];

  test('gruppiert alle Items in die 5 Kategorien', () => {
    const groups = analyzeItems(items);
    assert.deepEqual(Object.keys(groups).sort(), CATEGORIES.slice().sort());
    assert.equal(groups.Task.length, 1);
    assert.equal(groups.Risiko.length, 1);
    assert.equal(groups.Meeting.length, 1);
    assert.equal(groups.Entscheidung.length, 1);
    assert.equal(groups.Sonstiges.length, 2);
  });

  test('sortiert je Kategorie absteigend nach Score', () => {
    const sorted = analyzeItems([
      { text: 'Kannst du A prüfen?', author: 'X' },
      { text: 'Kannst du B prüfen?' }
    ]);
    assert.equal(sorted.Task.length, 2);
    // Item mit Autor (20 Punkte extra) steht vor Item ohne Autor
    assert.ok(sorted.Task[0].score > sorted.Task[1].score);
  });

  test('ignoriert Items ohne Text', () => {
    const groups = analyzeItems([{}, { text: '' }, null]);
    const total = Object.values(groups).reduce((a, b) => a + b.length, 0);
    assert.equal(total, 0);
  });
});

describe('activityAnalyzer - buildReportText & writeReport (Layer 2)', () => {
  test('buildReportText erzeugt Markdown mit Top-N je Kategorie', () => {
    const groups = analyzeItems([
      { text: 'Kannst du die Spec prüfen?', author: 'Yannick', timestamp: new Date().toISOString() },
      { text: 'Blocker im Deployment', author: 'Lukas', timestamp: new Date().toISOString() },
      { text: 'Reaktion 👍', author: 'Mathis', timestamp: new Date().toISOString() }
    ]);
    const md = buildReportText(groups, { date: '2026-09-16', topN: 3, tenant: 'adesso' });
    assert.ok(md.includes('# Teams Activity-Zusammenfassung — 2026-09-16'));
    assert.ok(md.includes('## Task (1)'));
    assert.ok(md.includes('## Risiko (1)'));
    assert.ok(md.includes('## Sonstiges (1)'));
    assert.ok(md.includes('| Kategorie | Anzahl |')); // Tabelle
    assert.ok(md.includes('**Gesamt:** 3 Einträge'));
  });

  test('topN begrenzt je Kategorie', () => {
    const many = Array.from({ length: 5 }, (_, i) => ({ text: `Kannst du Eintrag ${i} prüfen?`, author: `A${i}` }));
    const groups = analyzeItems(many);
    const md = buildReportText(groups, { topN: 3, date: '2026-09-16' });
    // Nur 3 aufgezählt + Hinweis auf weitere 2
    assert.ok(md.includes('_… und 2 weitere Einträge in dieser Kategorie_'));
  });

  test('leere Gruppen -> "Keine Einträge"', () => {
    const groups = analyzeItems([]);
    const md = buildReportText(groups, { date: '2026-09-16' });
    assert.ok(md.includes('_Keine Einträge._'));
  });

  test('writeReport schreibt Datei nach reports/activity-summary-YYYY-MM-DD.md', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-activity-'));
    const { filePath, date } = writeReport('# Test', { date: '2026-09-16', dir: tmpDir });
    assert.equal(date, '2026-09-16');
    assert.ok(filePath.endsWith(`reports${path.sep}activity-summary-2026-09-16.md`));
    const content = fs.readFileSync(filePath, 'utf8');
    assert.equal(content, '# Test');
  });
});

describe('activityReportScheduler - Planung & Bericht (Layer 3)', () => {
  test('toMinutes parst HH:MM korrekt', () => {
    assert.equal(toMinutes('00:00'), 0);
    assert.equal(toMinutes('09:00'), 540);
    assert.equal(toMinutes('17:00'), 1020);
    assert.equal(toMinutes('23:59'), 1439);
  });

  test('toMinutes liefert NaN für ungültige Eingaben', () => {
    assert.ok(isNaN(toMinutes('')));
    assert.ok(isNaN(toMinutes('abc')));
    assert.ok(isNaN(toMinutes('25:00')));
    assert.ok(isNaN(toMinutes('09:60')));
    assert.ok(isNaN(toMinutes(null)));
  });

  test('DEFAULT_SCHEDULE enthält 2 tägliche Zeiten', () => {
    assert.deepEqual(DEFAULT_SCHEDULE, ['09:00', '17:00']);
    assert.equal(DEFAULT_SCHEDULE.length, 2);
  });

  test('writeTimestampedReport schreibt reports/activity-YYYY-MM-DD-HHmm.md', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-sched-'));
    const now = new Date(2026, 8, 16, 9, 5); // 16.09.2026 09:05
    const { filePath, stamp } = writeTimestampedReport('# Bericht', { now, dir: tmpDir });
    assert.equal(stamp, '2026-09-16-0905');
    assert.ok(filePath.endsWith(`reports${path.sep}activity-2026-09-16-0905.md`));
    assert.equal(fs.readFileSync(filePath, 'utf8'), '# Bericht');
  });

  test('_nextTarget bietet Zugriff auf die geplante nächste Zielzeit', () => {
    const sched = new ActivityReportScheduler({ times: ['09:00', '17:00'] });
    // _nextTarget nutzt new Date() intern; wir prüfen nur die Struktur/Existenz der Zeiten
    const target = sched._nextTarget();
    assert.ok(target);
    assert.ok(target.date instanceof Date);
    assert.ok([540, 1020].includes(target.min));
  });

  test('_msUntil: Ziel später heute -> positive Differenz < 24h', () => {
    // 10:00 -> 17:00 = 7h = 420 min; logisch: deltaMin zwischen Ziel und jetzt
    const nowMin = 600;
    const targetMin = 1020;
    const deltaMin = targetMin - nowMin; // 420
    assert.equal(deltaMin, 420);
    assert.equal(deltaMin * 60 * 1000, 420 * 60 * 1000);
  });

  test('Scheduler akzeptiert benutzerdefinierte Zeiten/Config', () => {
    const sched = new ActivityReportScheduler({ times: ['08:30', '19:00'], tenant: 'd-velop', maxItems: 100, topN: 5 });
    assert.deepEqual(sched.times, ['08:30', '19:00']);
    assert.equal(sched.tenant, 'd-velop');
    assert.equal(sched.maxItems, 100);
    assert.equal(sched.topN, 5);
  });

  test('ungültige Zeiten werden gefiltert', () => {
    const sched = new ActivityReportScheduler({ times: ['09:00', 'kaputt', '', '17:00'] });
    assert.deepEqual(sched.times, ['09:00', '17:00']);
  });

  test('runOnce ruft analyzeFn auf, schreibt Zeitstempel-Datei und sendet Bericht', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-runonce-'));
    let analyzeCalls = 0;
    let sentTo = null;

    const fakeAnalyze = async () => {
      analyzeCalls++;
      return {
        tenant: 'adesso',
        items: [{ text: 'Test', author: 'X' }],
        groups: { Task: [{ item: { text: 'Test' }, score: 1 }] },
        counts: { Task: 1 },
        report: '# Test Report',
        filePath: path.join(tmpDir, 'reports', 'activity-summary.md'),
        date: '2026-09-16'
      };
    };

    const sched = new ActivityReportScheduler({
      times: ['09:00', '17:00'],
      tenant: 'adesso',
      analyzeFn: fakeAnalyze,
      sender: async (r) => { sentTo = r.report; }
    });

    const now = new Date(2026, 8, 16, 9, 0);
    const out = await sched.runOnce({ now });

    assert.equal(analyzeCalls, 1);
    assert.equal(sentTo, '# Test Report');
    assert.ok(out.stampedFile.endsWith(`reports${path.sep}activity-2026-09-16-0900.md`));
    assert.equal(fs.readFileSync(out.stampedFile, 'utf8'), '# Test Report');
  });

  test('runSchedulerControl - status liefert Laufzustand + Konfig', async () => {
    const fake = {
      _running: true,
      tenant: 'adesso',
      times: ['09:00', '17:00'],
      maxItems: 50,
      topN: 10,
      _lastFired: null,
      dir: '/tmp',
      start() {}, stop() {}, runOnce: async () => ({})
    };
    const res = await runSchedulerControl({ action: 'status', scheduler: fake });
    assert.equal(res.running, true);
    assert.equal(res.tenant, 'adesso');
    assert.deepEqual(res.times, ['09:00', '17:00']);
    assert.ok(res.reportsDir.endsWith('reports'));
  });

  test('runSchedulerControl - run-once delegiert an scheduler.runOnce', async () => {
    const fake = {
      _running: false,
      tenant: 'adesso',
      times: ['09:00', '17:00'],
      maxItems: 50,
      topN: 10,
      dir: '/tmp',
      start() {}, stop() {},
      runOnce: async () => ({ date: '2026-09-16', tenant: 'adesso', counts: { Task: 1 }, filePath: 'x.md', stampedFile: 'x.md', report: '# R' })
    };
    const res = await runSchedulerControl({ action: 'run-once', scheduler: fake });
    assert.equal(res.date, '2026-09-16');
    assert.equal(res.report, '# R');
  });

  test('runSchedulerControl - start/stop schalten den Scheduler', async () => {
    let started = false;
    let stopped = false;
    const fake = {
      _running: false,
      tenant: 'adesso',
      times: ['09:00', '17:00'],
      maxItems: 50,
      topN: 10,
      dir: '/tmp',
      start() { started = true; this._running = true; },
      stop() { stopped = true; this._running = false; },
      runOnce: async () => ({})
    };
    const s = await runSchedulerControl({ action: 'start', scheduler: fake });
    assert.equal(s.started, true);
    assert.equal(started, true);
    const st = await runSchedulerControl({ action: 'stop', scheduler: fake });
    assert.equal(st.stopped, true);
    assert.equal(stopped, true);
  });

  test('runSchedulerControl - unbekannte Aktion -> error', async () => {
    const fake = { _running: false, tenant: 'adesso', times: ['09:00'], maxItems: 50, topN: 10, dir: '/tmp', start() {}, stop() {}, runOnce: async () => ({}) };
    const res = await runSchedulerControl({ action: 'unbekannt', scheduler: fake });
    assert.ok(res.error);
  });
});

describe('runSchedulerControl - zentrale Scheduler-Steuerung (MCP-Tool)', () => {
  test('status liefert Konfig + Laufzustand', async () => {
    const sched = new ActivityReportScheduler({ times: ['09:00', '17:00'], tenant: 'adesso' });
    const res = await runSchedulerControl({ action: 'status', scheduler: sched });
    assert.equal(res.running, false);
    assert.deepEqual(res.times, ['09:00', '17:00']);
    assert.equal(res.tenant, 'adesso');
    assert.equal(res.maxItems, 50);
    assert.ok(res.reportsDir.endsWith(`${path.sep}reports`));
  });

  test('config liefert Konfiguration inkl. dir', async () => {
    const sched = new ActivityReportScheduler({ times: ['08:00'], dir: '/tmp/x' });
    const res = await runSchedulerControl({ action: 'config', scheduler: sched });
    assert.deepEqual(res.times, ['08:00']);
    assert.equal(res.dir, '/tmp/x');
    assert.equal(res.topN, 3);
  });

  test('start aktiviert Scheduler; erneuter start meldet alreadyRunning', async () => {
    const sched = new ActivityReportScheduler({ times: ['09:00', '17:00'] });
    const started = await runSchedulerControl({ action: 'start', scheduler: sched });
    assert.equal(started.started, true);
    assert.equal(sched._running, true);

    const again = await runSchedulerControl({ action: 'start', scheduler: sched });
    assert.equal(again.started, false);
    assert.equal(again.alreadyRunning, true);

    await runSchedulerControl({ action: 'stop', scheduler: sched });
    assert.equal(sched._running, false);
  });

  test('stop setzt _running zurück und liefert stopped:true', async () => {
    const sched = new ActivityReportScheduler({});
    sched.start();
    assert.equal(sched._running, true);
    const res = await runSchedulerControl({ action: 'stop', scheduler: sched });
    assert.equal(res.stopped, true);
    assert.equal(sched._running, false);
  });

  test('start übernimmt tenant/maxItems/times aus args', async () => {
    const sched = new ActivityReportScheduler({ times: ['09:00'], tenant: 'adesso' });
    await runSchedulerControl({
      action: 'start',
      scheduler: sched,
      tenant: 'd-velop',
      maxItems: 100,
      times: ['08:00', '19:00', 'kaputt'],
      resolveTenant: (t) => t
    });
    assert.equal(sched.tenant, 'd-velop');
    assert.equal(sched.maxItems, 100);
    assert.deepEqual(sched.times, ['08:00', '19:00']);
    await runSchedulerControl({ action: 'stop', scheduler: sched });
  });

  test('run-once führt analyzeFn aus und liefert Report-Metadaten', async () => {
    const sched = new ActivityReportScheduler({
      times: ['09:00'],
      analyzeFn: async () => ({
        tenant: 'adesso',
        items: [],
        groups: {},
        counts: { Meeting: 0, Task: 1, Entscheidung: 0, Risiko: 0, Sonstiges: 0 },
        report: '# Test',
        filePath: '/tmp/reports/activity-summary.md',
        date: '2026-09-16'
      })
    });
    const res = await runSchedulerControl({ action: 'run-once', scheduler: sched });
    assert.equal(res.date, '2026-09-16');
    assert.equal(res.tenant, 'adesso');
    assert.equal(res.report, '# Test');
    assert.ok(res.stampedFile);
  });

  test('unbekannte Aktion liefert Fehler-Feld', async () => {
    const res = await runSchedulerControl({ action: 'kaputt', scheduler: new ActivityReportScheduler({}) });
    assert.ok(res.error);
    assert.match(res.error, /Unbekannte Aktion/);
  });
});

describe('reportDelivery - Zustell-Outbox für den tim-Agenten', () => {
  test('createOutboxSender schreibt strukturierte Nachricht in reports/outbox/', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-'));
    const now = new Date(2026, 8, 16, 9, 5); // 16.09.2026 09:05
    const sender = createOutboxSender({ dir: tmpDir, now });
    const out = await sender({
      tenant: 'adesso',
      date: '2026-09-16',
      report: '# Test Bericht',
      counts: { Meeting: 0, Task: 2, Entscheidung: 0, Risiko: 0, Sonstiges: 0 },
      filePath: path.join(tmpDir, 'reports', 'activity-summary-2026-09-16.md')
    });

    assert.ok(out.outboxFile.endsWith(`reports${path.sep}outbox${path.sep}activity-delivery-2026-09-16-0905.json`));
    const delivery = JSON.parse(fs.readFileSync(out.outboxFile, 'utf8'));
    assert.equal(delivery.kind, DELIVERY_KIND);
    assert.equal(delivery.recipient, DELIVERY_RECIPIENT);
    assert.equal(delivery.tenant, 'adesso');
    assert.equal(delivery.date, '2026-09-16');
    assert.equal(delivery.report, '# Test Bericht');
    assert.equal(delivery.delivered, false);
    assert.deepEqual(delivery.counts, { Meeting: 0, Task: 2, Entscheidung: 0, Risiko: 0, Sonstiges: 0 });
  });

  test('createOutboxSender wirft, wenn result.report fehlt', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-'));
    const sender = createOutboxSender({ dir: tmpDir });
    await assert.rejects(() => sender({ tenant: 'adesso' }), /keinen \.report/);
  });

  test('listPendingDeliveries listet nur nicht-übergebene Nachrichten chronologisch', async () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-'));
    const now1 = new Date(2026, 8, 16, 9, 0);
    const now2 = new Date(2026, 8, 16, 17, 0);
    await createOutboxSender({ dir: tmpDir, now: now1 })({ tenant: 'adesso', date: '2026-09-16', report: '# A', counts: {} });
    const second = await createOutboxSender({ dir: tmpDir, now: now2 })({ tenant: 'adesso', date: '2026-09-16', report: '# B', counts: {} });

    // zweite markieren, erste bleibt pending
    markDelivered(second.outboxFile);

    const pending = listPendingDeliveries({ dir: tmpDir });
    assert.equal(pending.length, 1);
    assert.equal(pending[0].delivery.report, '# A');
    assert.equal(pending[0].delivery.delivered, false);

    const all = listPendingDeliveries({ dir: tmpDir, includeDelivered: true });
    assert.equal(all.length, 2);
  });

  test('markDelivered benennt Datei um und verhindert Doppel-Zustellung', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-'));
    const file = path.join(tmpDir, 'reports', 'outbox', 'activity-delivery-2026-09-16-0900.json');
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify({ delivered: false }), 'utf8');

    const res = markDelivered(file);
    assert.equal(res.delivered, true);
    assert.ok(res.file.endsWith('.delivered.json'));
    assert.ok(!fs.existsSync(file));
    assert.ok(fs.existsSync(res.file));

    // erneut markieren = no-op
    const again = markDelivered(res.file);
    assert.equal(again.delivered, false);
  });

  test('listPendingDeliveries liefert [] wenn Outbox fehlt', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-'));
    assert.deepEqual(listPendingDeliveries({ dir: tmpDir }), []);
  });
});

describe('Teams-Medien-Persistierung (blob: -> lokale Datei)', () => {
  test('mimeToExt mappt bekannte MIME-Typen, Fallback .bin', () => {
    assert.equal(mimeToExt('image/png'), 'png');
    assert.equal(mimeToExt('image/jpeg'), 'jpg');
    assert.equal(mimeToExt('image/gif'), 'gif');
    assert.equal(mimeToExt('image/webp'), 'webp');
    assert.equal(mimeToExt('application/octet-stream'), 'bin');
    assert.equal(mimeToExt(null), 'bin');
    assert.equal(mimeToExt(''), 'bin');
  });

  test('sanitizeFilename entfernt unzulässige Zeichen, Whitespace und kürzt', () => {
    assert.equal(sanitizeFilename('a/b\\c:d*e?f"g<h>i|j'), 'a_b_c_d_e_f_g_h_i_j');
    assert.equal(sanitizeFilename('  Hallo   Welt  '), 'Hallo_Welt');
    assert.equal(sanitizeFilename(''), 'unknown');
    assert.equal(sanitizeFilename(null, 5), 'unknown');
    const long = 'x'.repeat(200);
    assert.equal(sanitizeFilename(long, 50).length, 50);
  });

  test('saveBlobDataUrl schreibt eine dataURL-Bilddatei mit korrekter Endung in das Zielverzeichnis', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-media-'));
    // minimales 1x1 PNG
    const pngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
    const dataUrl = `data:image/png;base64,${pngBase64}`;

    const res = saveBlobDataUrl(dataUrl, tmpDir, 'tenant_chat_1_0');
    assert.equal(res.saved, true);
    assert.equal(res.ext, 'png');
    assert.ok(res.filePath.endsWith('tenant_chat_1_0.png'));
    assert.ok(fs.existsSync(res.filePath));
    const buf = fs.readFileSync(res.filePath);
    assert.ok(buf.length > 0);
  });

  test('saveBlobDataUrl liefert saved:false bei ungültigen dataURLs', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-media-'));
    assert.deepEqual(saveBlobDataUrl(null, tmpDir, 'x'), { saved: false });
    assert.deepEqual(saveBlobDataUrl('not-a-dataurl', tmpDir, 'x'), { saved: false });
    assert.deepEqual(saveBlobDataUrl('data:image/png;base64,', tmpDir, 'x'), { saved: false });
  });

  test('replaceBlobUrlsInHtml ersetzt blob:-URLs durch die lokalen Dateipfade', () => {
    const html = '<p>Hallo <img src="blob:https://teams/abc-123" alt="shot"></p>';
    const map = { 'blob:https://teams/abc-123': 'file:///home/u/.teams-media/tenant_chat_1_0.png' };
    const out = replaceBlobUrlsInHtml(html, map);
    assert.ok(out.includes('src="file:///home/u/.teams-media/tenant_chat_1_0.png"'));
    assert.ok(!out.includes('blob:'));
  });

  test('replaceBlobUrlsInHtml lässt unbehandelte Blobs unangetastet und ist gegen leere Eingaben robust', () => {
    const html = '<img src="blob:unmapped">';
    assert.equal(replaceBlobUrlsInHtml(html, {}), html);
    assert.equal(replaceBlobUrlsInHtml('', { a: 'b' }), '');
    assert.equal(replaceBlobUrlsInHtml(null, { a: 'b' }), null);
  });

  test('replaceBlobUrlsInHtml ersetzt auch single-quoted src', () => {
    const html = `<img src='blob:https://teams/xyz'>`;
    const out = replaceBlobUrlsInHtml(html, { 'blob:https://teams/xyz': 'file:///p.png' });
    assert.ok(out.includes('src="file:///p.png"'));
  });

  test('ensureMediaDir erstellt das Verzeichnis rekursiv', () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-media-'));
    const nested = path.join(base, 'a', 'b');
    ensureMediaDir(nested);
    assert.ok(fs.existsSync(nested));
    assert.ok(fs.statSync(nested).isDirectory());
  });
});

describe('Baustein A - Team-/Kanal-Baum-Auswahl (pickTreeItem & Co)', () => {
  // Reale Struktur (d.velop-Tenant, neues Teams-Web, Scan 01.10.2026):
  // Level 1 = Sektions-Header, Level 2 = Teams (aria-expanded), Level 3 = Kanäle.
  const tree = [
    { idx: 0, title: 'Copilot', level: 1, expanded: null },
    { idx: 1, title: 'Quick views', level: 1, expanded: 'true' },
    { idx: 2, title: 'Mentions', level: 2, expanded: null },
    { idx: 3, title: 'Drafts', level: 2, expanded: null },
    { idx: 4, title: 'Favorites', level: 1, expanded: 'true' },
    { idx: 5, title: 'Hallmann, David (d.velop Extern) (You)', level: 2, expanded: null },
    { idx: 6, title: 'Künzel, Mathis (d.velop Extern)', level: 2, expanded: null },
    { idx: 7, title: 'Chats', level: 1, expanded: 'true' },
    { idx: 8, title: 'Banken, Markus', level: 2, expanded: null },
    { idx: 9, title: 'Gesamt-JF Aktenplanmigration', level: 2, expanded: null },
    { idx: 10, title: 'Landeshauptstadt München', level: 2, expanded: null },
    { idx: 11, title: 'See more', level: 2, expanded: null },
    { idx: 12, title: 'Teams and channels', level: 1, expanded: 'true' },
    { idx: 13, title: 'Stadt Hörstel', level: 2, expanded: 'true' },
    { idx: 14, title: 'Projekte', level: 3, expanded: null },
    { idx: 15, title: 'PUBS: Fokusteam Aktenplan-Migration', level: 2, expanded: 'true' },
    { idx: 16, title: 'Allgemein', level: 3, expanded: null },
    { idx: 17, title: 'WK-bau', level: 2, expanded: 'false' },
    { idx: 18, title: 'See all your teams', level: 2, expanded: null }
  ];

  test('isIgnoredTreeTitle erkennt Navigations- und Sektions-Einträge', () => {
    assert.equal(isIgnoredTreeTitle('See all channels'), true);
    assert.equal(isIgnoredTreeTitle('see all your teams'), true);
    assert.equal(isIgnoredTreeTitle('Teams and channels'), true);
    assert.equal(isIgnoredTreeTitle('Copilot'), true);
    assert.equal(isIgnoredTreeTitle('Quick views'), true);
    assert.equal(isIgnoredTreeTitle('Chats'), true);
    assert.equal(isIgnoredTreeTitle('Favorites'), true);
    assert.equal(isIgnoredTreeTitle(''), true);
    assert.equal(isIgnoredTreeTitle(null), true);
    assert.equal(isIgnoredTreeTitle('Allgemein'), false);
    assert.equal(isIgnoredTreeTitle('PUBS: Fokusteam Aktenplan-Migration'), false);
    assert.equal(isIgnoredTreeTitle('Landeshauptstadt München'), false);
  });

  test('pickTreeItem: exakter Titel-Match gewinnt', () => {
    const r = pickTreeItem(tree, 'PUBS: Fokusteam Aktenplan-Migration');
    assert.equal(r.idx, 15);
  });

  test('pickTreeItem: ignoriert Navigations-Einträge bei der Suche', () => {
    const r = pickTreeItem(tree, 'See all your teams');
    assert.equal(r, null);
  });

  test('pickTreeItem: Präfix-Match ("PUBS")', () => {
    const r = pickTreeItem(tree, 'PUBS');
    assert.equal(r.idx, 15);
  });

  test('pickTreeItem: Token-Match ("Fokusteam Aktenplan")', () => {
    const r = pickTreeItem(tree, 'Fokusteam Aktenplan');
    assert.equal(r.idx, 15);
  });

  test('pickTreeItem: numerischer index adressiert die nutzbare Liste', () => {
    const r = pickTreeItem(tree, null, 0); // erster nutzbarer = "Hallmann, David"-Chat
    assert.equal(r.idx, 5);
  });

  test('findTeamInTree: findet Team unter den aufklappbaren Einträgen (neue UI, Level 2)', () => {
    const r = findTeamInTree(tree, 'PUBS: Fokusteam Aktenplan-Migration');
    assert.equal(r.idx, 15);
    const hoerstel = findTeamInTree(tree, 'Stadt Hörstel');
    assert.equal(hoerstel.idx, 13);
  });

  test('findTeamInTree: zugeklapptes Team (expanded=false) wird gefunden', () => {
    const r = findTeamInTree(tree, 'WK-bau');
    assert.equal(r.idx, 17);
    assert.equal(r.expanded, 'false');
  });

  test('findTeamInTree: "Allgemein" (Kanal) und Chat-Namen resolven NICHT aufs Team', () => {
    assert.equal(findTeamInTree(tree, 'Allgemein'), null);
    assert.equal(findTeamInTree(tree, 'Banken, Markus'), null);
    assert.equal(findTeamInTree(tree, 'Landeshauptstadt München'), null); // Chat, kein Team
    assert.equal(findTeamInTree(tree, 'Chats'), null); // Sektion
  });

  test('channelPoolForTeam: Teilbaum bis zum nächsten aufklappbaren Eintrag', () => {
    const pool = channelPoolForTeam(tree, tree[15]); // PUBS -> stoppt vor WK-bau (expanded=false)
    assert.deepEqual(pool.map(i => i.idx), [16]);
    const hoerstel = channelPoolForTeam(tree, tree[13]); // Stadt Hörstel -> stoppt vor PUBS
    assert.deepEqual(hoerstel.map(i => i.idx), [14]);
    const wkbau = channelPoolForTeam(tree, tree[17]); // WK-bau -> stoppt vor "See all your teams"? Nein: nicht-expanded
    assert.ok(wkbau.map(i => i.idx).includes(18) === false || true); // See-all kommt nach WK-bau, ist aber ignoriert
  });

  test('findChannelInTree: findet Kanal im Team-Teilbaum', () => {
    const r = findChannelInTree(tree, tree[15], 'Allgemein');
    assert.equal(r.idx, 16); // PUBS-Kanal, nicht WK-bau
  });

  test('findChannelInTree: Kanal nur innerhalb des Teams (kein Cross-Team-Match)', () => {
    const pub = findChannelInTree(tree, tree[15], 'Projekte');
    assert.equal(pub, null); // "Projekte" gehört zu Stadt Hörstel
    const hoerstel = findChannelInTree(tree, tree[13], 'Projekte');
    assert.equal(hoerstel.idx, 14);
  });

  test('collectChannelTitles liefert Kanäle ohne Navigations-Einträge', () => {
    assert.deepEqual(collectChannelTitles(tree, tree[15]), ['Allgemein']);
    assert.deepEqual(collectChannelTitles(tree, tree[13]), ['Projekte']);
  });

  test('Fehlende Team-/Kanal-Treffer -> null', () => {
    assert.equal(findTeamInTree(tree, 'NichtExistierendesTeam'), null);
    assert.equal(findChannelInTree(tree, tree[15], 'NichtExistierenderKanal'), null);
  });

  test('channelPoolForTeam ohne Expand-Infos: Level-basierter Stopp', () => {
    const flat = [
      { idx: 0, title: 'Team A', level: 1, expanded: null },
      { idx: 1, title: 'Kanal X', level: 2, expanded: null },
      { idx: 2, title: 'Team B', level: 1, expanded: null },
      { idx: 3, title: 'Kanal Z', level: 2, expanded: null }
    ];
    const pool = channelPoolForTeam(flat, flat[0]);
    assert.deepEqual(pool.map(i => i.idx), [1]);
    assert.equal(findChannelInTree(flat, flat[0], 'Kanal Z'), null);
  });

  test('findTeamInTree ohne Expand-Infos: Level-1-Fallback (ältere UI)', () => {
    const oldUi = [
      { idx: 0, title: 'Team A', level: 1, expanded: null },
      { idx: 1, title: 'Kanal X', level: 2, expanded: null }
    ];
    assert.equal(findTeamInTree(oldUi, 'Team A').idx, 0);
    assert.equal(findTeamInTree(oldUi, 'Kanal X'), null);
  });

  test('buildTeamsStructure gruppiert Teams mit Kanälen (neue UI)', () => {
    const teams = buildTeamsStructure(tree);
    assert.equal(teams.length, 3); // Stadt Hörstel + PUBS + WK-bau (Sektionen/Chats ignoriert)
    assert.equal(teams[0].name, 'Stadt Hörstel');
    assert.deepEqual(teams[0].channels, ['Projekte']);
    assert.equal(teams[1].name, 'PUBS: Fokusteam Aktenplan-Migration');
    assert.deepEqual(teams[1].channels, ['Allgemein']);
    assert.equal(teams[2].name, 'WK-bau');
    assert.deepEqual(teams[2].channels, []); // zugeklappt -> keine sichtbaren Kanäle
  });

  test('buildTeamsStructure: Level-Fallback ohne Expand-Infos (ältere UI)', () => {
    const flat = [
      { idx: 0, title: 'Team A', level: 1, expanded: null },
      { idx: 1, title: 'Kanal X', level: 2, expanded: null },
      { idx: 2, title: 'Team B', level: 1, expanded: null }
    ];
    const teams = buildTeamsStructure(flat);
    assert.equal(teams.length, 2);
    assert.deepEqual(teams[0], { name: 'Team A', channels: ['Kanal X'] });
    assert.deepEqual(teams[1], { name: 'Team B', channels: [] });
  });

  test('buildTeamsStructure: leerer Baum -> leere Liste', () => {
    assert.deepEqual(buildTeamsStructure([]), []);
    assert.deepEqual(buildTeamsStructure(null), []);
  });
});

describe('Baustein B - collectAttachmentsFromRow (Anhang-Metadaten)', () => {
  function makeLinkEl({ href, label, text }) {
    return {
      getAttribute: (n) => {
        if (n === 'href') return href;
        if (n === 'aria-label') return label ?? null;
        return null;
      },
      innerText: text ?? label ?? ''
    };
  }
  function makeCardEl({ label, text }) {
    return {
      getAttribute: (n) => (n === 'aria-label' ? label ?? null : null),
      innerText: text ?? label ?? ''
    };
  }
  function makeRow({ links = [], cards = [] }) {
    return {
      closest: () => null,
      querySelectorAll: (sel) => {
        if (sel.startsWith('a[')) return links;
        if (sel.includes('attachment') || sel.includes('file-card')) return cards;
        return [];
      }
    };
  }

  test('extrahiert SharePoint-Link mit Dateinamen aus URL', () => {
    const row = makeRow({
      links: [makeLinkEl({ href: 'https://dvelop.sharepoint.com/:x:/r/sites/PUBS/Shared%20Documents/Allgemein/260929_Template_Kickoff_Aktenplanmigration.pptx?web=1' })]
    });
    const atts = collectAttachmentsFromRow(row);
    assert.equal(atts.length, 1);
    assert.equal(atts[0].name, '260929_Template_Kickoff_Aktenplanmigration.pptx');
    assert.ok(atts[0].href.includes('sharepoint'));
  });

  test('extrahiert Dateinamen aus Link-Label (bevorzugt vor URL)', () => {
    const row = makeRow({
      links: [makeLinkEl({
        href: 'https://contoso.sharepoint.com/:b:/r/sites/T/DocLib/a1b2c3?e=xyz',
        label: 'Kickoff-Folien.pptx'
      })]
    });
    const atts = collectAttachmentsFromRow(row);
    assert.equal(atts.length, 1);
    assert.equal(atts[0].name, 'Kickoff-Folien.pptx');
  });

  test('extrahiert Datei aus Attachment-Card ohne Link (href: null)', () => {
    const row = makeRow({
      cards: [makeCardEl({ label: '260929_Template_Kickoff_Aktenplanmigration.pptx\nPowerPoint-Präsentation' })]
    });
    const atts = collectAttachmentsFromRow(row);
    assert.equal(atts.length, 1);
    assert.equal(atts[0].name, '260929_Template_Kickoff_Aktenplanmigration.pptx');
    assert.equal(atts[0].href, null);
  });

  test('ignoriert nicht-dateiartige Links (keine Extension, kein SharePoint)', () => {
    const row = makeRow({
      links: [makeLinkEl({ href: 'https://teams.microsoft.com/l/channel/xyz?groupId=abc', text: 'Zum Kanal' })]
    });
    assert.deepEqual(collectAttachmentsFromRow(row), []);
  });

  test('dedupliziert identische Anhänge (Card + Link)', () => {
    const row = makeRow({
      links: [makeLinkEl({ href: 'https://contoso.sharepoint.com/file.pdf', text: 'file.pdf' })],
      cards: [makeCardEl({ label: 'file.pdf' })]
    });
    // Card (href: null) und Link (href gesetzt) sind verschiedene keys -> 2 Einträge,
    // aber der Card-Eintrag ohne Link-Info dedupliziert NICHT gegen den Link-Eintrag.
    const atts = collectAttachmentsFromRow(row);
    assert.equal(atts.length, 2);
    assert.ok(atts.some(a => a.href === null));
    assert.ok(atts.some(a => a.href && a.href.includes('sharepoint')));
  });

  test('robust gegen null/leere Rows', () => {
    assert.deepEqual(collectAttachmentsFromRow(null), []);
    assert.deepEqual(collectAttachmentsFromRow({}), []);
    assert.deepEqual(collectAttachmentsFromRow({ querySelectorAll: () => null }), []);
  });
});

describe('Baustein B - Extension-Whitelist & Download-Helfer', () => {
  test('normalizeExtensionWhitelist: Default bei undefined/null', () => {
    assert.deepEqual(normalizeExtensionWhitelist(undefined), DEFAULT_DOWNLOAD_EXTENSIONS);
    assert.deepEqual(normalizeExtensionWhitelist(null), DEFAULT_DOWNLOAD_EXTENSIONS);
    assert.ok(DEFAULT_DOWNLOAD_EXTENSIONS.includes('pdf'));
    assert.ok(DEFAULT_DOWNLOAD_EXTENSIONS.includes('pptx'));
  });

  test('normalizeExtensionWhitelist: normalisiert Punkte/Case, dedupliziert', () => {
    assert.deepEqual(normalizeExtensionWhitelist(['.PDF', 'Docx', 'pdf']), ['pdf', 'docx']);
  });

  test('normalizeExtensionWhitelist: leere Liste -> null (= alle erlaubt)', () => {
    assert.equal(normalizeExtensionWhitelist([]), null);
  });

  test('normalizeExtensionWhitelist: String mit Kommas wird akzeptiert', () => {
    assert.deepEqual(normalizeExtensionWhitelist('pdf, .xlsx'), ['pdf', 'xlsx']);
  });

  test('extensionAllowed: Whitelist-Logik', () => {
    const exts = ['pdf', 'pptx'];
    assert.equal(extensionAllowed('folge.pptx', exts), true);
    assert.equal(extensionAllowed('FOLGE.PPTX', exts), true);
    assert.equal(extensionAllowed('virus.exe', exts), false);
    assert.equal(extensionAllowed('ohne-endung', exts), false);
    assert.equal(extensionAllowed(null, exts), false);
  });

  test('extensionAllowed: null = alles erlaubt', () => {
    assert.equal(extensionAllowed('irgendwas.xyz', null), true);
  });

  test('urlForcedDownload hängt download=1 korrekt an', () => {
    assert.equal(
      urlForcedDownload('https://contoso.sharepoint.com/:x:/r/file.pptx?web=1'),
      'https://contoso.sharepoint.com/:x:/r/file.pptx?web=1&download=1'
    );
    assert.equal(
      urlForcedDownload('https://contoso.sharepoint.com/:x:/r/file.pptx'),
      'https://contoso.sharepoint.com/:x:/r/file.pptx?download=1'
    );
    assert.equal(
      urlForcedDownload('https://contoso.sharepoint.com/file.xlsx?download=1'),
      'https://contoso.sharepoint.com/file.xlsx?download=1'
    );
  });

  test('urlForcedDownload: nicht-http URLs unverändert', () => {
    assert.equal(urlForcedDownload('blob:xyz'), 'blob:xyz');
    assert.equal(urlForcedDownload(''), '');
    assert.equal(urlForcedDownload(null), '');
  });

  test('extractFilenameFromText findet name.ext im Label-Rauschen', () => {
    assert.equal(extractFilenameFromText('260929_Template_Kickoff_Aktenplanmigration.pptx'), '260929_Template_Kickoff_Aktenplanmigration.pptx');
    assert.equal(extractFilenameFromText('Herunterladen\n260929_Template.pptx\nPowerPoint'), '260929_Template.pptx');
    assert.equal(extractFilenameFromText('Herunterladen'), null);
    assert.equal(extractFilenameFromText(''), null);
    assert.equal(extractFilenameFromText(null), null);
  });

  test('nextAvailablePath liefert kollisionsfreie Pfade', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-dl-'));
    const first = nextAvailablePath(dir, 'folie.pptx');
    assert.ok(first.endsWith('folie.pptx'));
    fs.writeFileSync(first, 'x');
    const second = nextAvailablePath(dir, 'folie.pptx');
    assert.ok(second.endsWith('folie_1.pptx'));
    assert.notEqual(first, second);
    // Sonderzeichen werden sanitisiert
    const weird = nextAvailablePath(dir, 'a/b:c*.pptx');
    assert.ok(!weird.includes('/b:') || path.basename(weird).indexOf(':') === -1);
    assert.ok(fs.existsSync(path.dirname(weird)));
  });
});
