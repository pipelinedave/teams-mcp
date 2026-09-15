#!/usr/bin/env node
/**
 * Off → Outbox → tim-Agent — Demo- & Smoke-Test der Zustellebene.
 *
 * Demonstriert die komplette Zustell-Kette des proaktiven Activity-Berichts ohne
 * echten Teams-Browser:
 *   createOutboxSender (Bericht -> reports/outbox/activity-delivery-*.json)
 *   -> listPendingDeliveries (tim-Agent holt ab)
 *   -> markDelivered (als übergeben markieren, kein Doppel-Versand)
 *
 * Das ist die Zustell-Strecke, die der tim-Agent nutzt, um den Bericht an den Nutzer
 * zu präsentieren — nachdem der Scheduler ihn um 09:00/17:00 erzeugt hat.
 *
 * Verwendung:
 *   node tests/report-delivery.demo.mjs                     # temp-Verzeichnis
 *   node tests/report-delivery.demo.mjs --out /tmp/x        # Berichts-Verzeichnis
 */
import os from 'node:os';
import path from 'node:path';
import fs from 'node:fs';
import { createOutboxSender, listPendingDeliveries, markDelivered, DELIVERY_KIND, DELIVERY_RECIPIENT } from '../src/reportDelivery.js';

const outDir = (() => {
  const idx = process.argv.indexOf('--out');
  return idx !== -1 && process.argv[idx + 1] ? process.argv[idx + 1] : fs.mkdtempSync(path.join(os.tmpdir(), 'teams-delivery-demo-'));
})();

console.log('=== Teams Activity-Bericht Zustellung (Outbox → tim-Agent) — Demo ===\n');

// 1) Sender erzeugt eine strukturierte Zustell-Nachricht aus einem Bericht-Ergebnis
console.log('--- 1) Scheduler-Sender legt Outbox-Nachricht an ---');
const sender = createOutboxSender({ dir: outDir, now: new Date(2026, 8, 16, 9, 0) });
const { outboxFile } = await sender({
  tenant: 'adesso',
  date: '2026-09-16',
  report: [
    '# Teams Activity-Zusammenfassung — 2026-09-16',
    '',
    '**Gesamt:** 3 Einträge',
    '',
    '## Task (2)',
    '- **[Task] `hoch`** Yannick: Spec für Wildau prüfen — *Yannick Bülter*',
    '',
    '## Risiko (1)',
    '- **[Risiko] `hoch`** Blocker - Deployment funktioniert nicht — *Theys Schiller*',
    ''
  ].join('\n'),
  counts: { Meeting: 0, Task: 2, Entscheidung: 0, Risiko: 1, Sonstiges: 0 },
  filePath: path.join(outDir, 'reports', 'activity-summary-2026-09-16.md')
});
console.log(`Outbox-Datei: ${outboxFile}`);
console.log(`Existiert: ${fs.existsSync(outboxFile)}`);
console.log(`Kind: ${DELIVERY_KIND} | Recipient: ${DELIVERY_RECIPIENT}`);

// 2) Der tim-Agent ruft die Pending-Zustellungen ab
console.log('\n--- 2) tim-Agent ruft pending Zustellungen ab ---');
const pending = listPendingDeliveries({ dir: outDir });
console.log(`Pending: ${pending.length}`);
for (const { file, delivery } of pending) {
  console.log(`  - ${file.replace(outDir, '.')}`);
  console.log(`      date=${delivery.date} tenant=${delivery.tenant} delivered=${delivery.delivered}`);
}

// 3) Nach Präsentation an den Nutzer als übergeben markieren (kein Doppel-Versand)
console.log('\n--- 3) Zustellung als übergeben markieren ---');
const first = pending[0];
const marked = markDelivered(first.file);
console.log(`Delivered: ${marked.delivered} -> ${marked.file.replace(outDir, '.')}`);

console.log('\n--- 4) Nach Markierung: keine pending Nachrichten mehr ---');
console.log(`Pending danach: ${listPendingDeliveries({ dir: outDir }).length}`);
console.log(`Inkl. delivered: ${listPendingDeliveries({ dir: outDir, includeDelivered: true }).length}`);
