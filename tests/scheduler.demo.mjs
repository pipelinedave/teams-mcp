#!/usr/bin/env node
/**
 * Activity-Report-Scheduler — Demo- & Smoke-Test (Layer 3).
 *
 * Demonstriert die komplette Scheduler-Pipeline ohne echten Teams-Browser:
 * analyzeFn (simuliert) -> zeitgestempelte Markdown-Datei (reports/activity-YYYY-MM-DD-HHmm.md)
 * -> proaktiver Versand (sendProactiveReport).
 *
 * Verwendung:
 *   node tests/scheduler.demo.mjs                    # simulierter Lauf, Ausgabe an stderr
 *   node tests/scheduler.demo.mjs --out /tmp/x      # Berichts-Verzeichnis übersteuern
 */
import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs';
import { ActivityReportScheduler, DEFAULT_SCHEDULE, toMinutes } from '../src/activityReportScheduler.js';

const outDir = (() => {
  const idx = process.argv.indexOf('--out');
  return idx !== -1 && process.argv[idx + 1] ? process.argv[idx + 1] : process.cwd();
})();

console.log('=== Teams Activity-Report-Scheduler (Layer 3) — Demo ===\n');

// Geplanter Tagesrhythmus
console.log('--- 0) Tagesplan ---');
for (const t of DEFAULT_SCHEDULE) {
  console.log(`   ${t} Uhr (Minuten seit Mitternacht: ${toMinutes(t)})`);
}

// Simulierte Analyze-Funktion (ersetzt echten Teams-Browser-Zugriff fürs Demo)
const fakeAnalyze = async (tenant, opts) => {
  const counts = { Meeting: 1, Task: 3, Entscheidung: 2, Risiko: 1, Sonstiges: 2 };
  const report = [
    `# Teams Activity-Zusammenfassung — 2026-09-16`,
    ``,
    `> Automatisch erzeugt (tenant: ${tenant}). Simulierte Demo-Daten.`,
    ``,
    `**Gesamt:** 9 Einträge`,
    ``,
    `## Task (3)`,
    `- **[Task] \`hoch\`** Yannick: Kannst du die Spec für Wildau prüfen? — *Yannick Bülter*`,
    `- **[Task] \`hoch\`** Lukas: Bitte bis Freitag liefern — *Lukas Behncke*`,
    `- **[Task] \`hoch\`** Copilot: Erinnerung an die Woche steht an`,
    ``,
    `## Risiko (1)`,
    `- **[Risiko] \`hoch\`** Theys: Blocker - Deployment funktioniert nicht — *Theys Schiller*`,
    ``,
    `## Entscheidung (2)`,
    `- **[Entscheidung] \`mittel\`** Marc hat die Freigabe zur Abnahme erteilt — *Marc*`,
    ``,
    `## Meeting (1)`,
    `- **[Meeting] \`mittel\`** Kickoff Projekt Wildau am Freitag — *Amir Rocker*`,
    ``
  ].join('\n');
  return {
    tenant,
    items: [],
    groups: {},
    counts,
    report,
    filePath: path.join(outDir, 'reports', 'activity-summary.md'),
    date: '2026-09-16'
  };
};

// Benutzerdefinierter Sender: sammelt den Report (statt an Konsole auszugeben)
let captured = null;
const customSender = async (result) => { captured = result.report; };

console.log('\n--- 1) Scheduler-Lauf (simuliert) ---');
const sched = new ActivityReportScheduler({
  times: DEFAULT_SCHEDULE,
  tenant: 'adesso',
  dir: outDir,
  analyzeFn: fakeAnalyze,
  sender: customSender
});

const now = new Date(2026, 8, 16, 9, 0); // 09:00 Uhr
const out = await sched.runOnce({ now });

console.log(`\nZeitgestempelte Datei: ${out.stampedFile}`);
console.log(`Existiert: ${fs.existsSync(out.stampedFile)}`);
console.log(`Proaktiver Report an Nutzer gesendet: ${captured ? 'JA (' + captured.length + ' Zeichen)' : 'NEIN'}`);

console.log('\n--- 2) Dateiinhalt (reports/activity-2026-09-16-0900.md) ---');
console.log(fs.readFileSync(out.stampedFile, 'utf8'));

// Kurze Demo der Start/Stop-Mechanik (2 Sekunden Laufzeit, ohne echten Treffer zu erzwingen)
console.log('\n--- 3) Start/Stop-Mechanik ---');
const liveSched = new ActivityReportScheduler({
  times: ['09:00', '17:00'],
  analyzeFn: fakeAnalyze,
  sender: customSender
});
liveSched.start();
setTimeout(() => {
  liveSched.stop();
  console.log('Scheduler nach 1.5s sauber gestoppt (kein Lauf ausgelöst, da keine Zielzeit erreicht).');
}, 1500);
