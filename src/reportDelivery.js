import fs from 'fs';
import path from 'path';

/**
 * Zustell-Ebene für den proaktiven Activity-Bericht (Sender → Outbox → tim-Agent).
 *
 * Der MCP-Server-Prozess (teams-mcp) kann den tim-Agenten (opencode-Subagent) nicht
 * direkt aufrufen — er läuft als separater Prozess. Stattdessen wird der Bericht als
 * STRUKTURIERTE Nachricht in eine deterministische Outbox-Datei geschrieben
 * (reports/outbox/activity-delivery-*.json). Der tim-Agent holt diese Zustellungen ab
 * (listPendingDeliveries), stellt den Bericht an den Nutzer (David) zu und markiert sie
 * als übergeben (markDelivered). So entsteht eine robuste, testbare Zustell-Kette ohne
 * Prozesskopplung.
 *
 * Struktur einer Outbox-Nachricht:
 * {
 *   "schemaVersion": 1,
 *   "kind": "teams-activity-report",   // Fester Typ für den tim-Agenten
 *   "recipient": "tim",                // Ziel-Agent
 *   "createdAt": "<ISO>",              // Erzeugungszeit
 *   "tenant": "adesso",
 *   "date": "2026-09-16",
 *   "report": "<Markdown>",            // Der zu präsentierende Bericht
 *   "counts": { Meeting: 0, Task: 3, ... },
 *   "sourceFile": "<pfad zur .md-Datei>",
 *   "delivered": false
 * }
 */

export const DELIVERY_KIND = 'teams-activity-report';
export const DELIVERY_RECIPIENT = 'tim';
export const SCHEMA_VERSION = 1;

// Status-Datei-Ende für "übergeben" (renamed, damit keine Doppel-Zustellung erfolgt)
const DELIVERED_SUFFIX = '.delivered.json';

function outboxDir(dir) {
  return path.join(dir, 'reports', 'outbox');
}

function pad(n) {
  return String(n).padStart(2, '0');
}

/**
 * Erzeugt einen injizierbaren Sender (passend zu sendProactiveReport), der den Bericht
 * als strukturierte Nachricht in die Outbox schreibt. Mehrere Läufe am selben Tag
 * überschreiben keinen bestehenden Eintrag (Zeitstempel im Dateinamen).
 *
 * @param {object} [opts]
 * @param {string} [opts.dir] - Basis-Verzeichnis (Standard CWD)
 * @param {Date}   [opts.now] - Zeitpunkt (Standard jetzt)
 * @returns {async (result) => {outboxFile: string}}
 */
export function createOutboxSender({ dir = process.cwd(), now = new Date() } = {}) {
  return async function outboxSender(result) {
    if (!result || !result.report) {
      throw new Error('Outbox-Sender: Ergebnis hat keinen .report');
    }
    const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
    const outbox = outboxDir(dir);
    fs.mkdirSync(outbox, { recursive: true });

    const createdAt = now.toISOString();
    const filePath = path.join(outbox, `activity-delivery-${stamp}.json`);

    const delivery = {
      schemaVersion: SCHEMA_VERSION,
      kind: DELIVERY_KIND,
      recipient: DELIVERY_RECIPIENT,
      createdAt,
      tenant: result.tenant || null,
      date: result.date || createdAt.slice(0, 10),
      report: result.report,
      counts: result.counts || {},
      sourceFile: result.filePath || null,
      delivered: false
    };

    fs.writeFileSync(filePath, JSON.stringify(delivery, null, 2), 'utf8');
    return { outboxFile: filePath };
  };
}

/**
 * Listet noch NICHT zugestellte (pending) Outbox-Nachrichten auf — für den tim-Agenten,
 * damit er weiss, welcher Bericht an den Nutzer präsentiert werden soll.
 *
 * @param {object} [opts]
 * @param {string} [opts.dir] - Basis-Verzeichnis (Standard CWD)
 * @param {boolean} [opts.includeDelivered] - auch bereits zugestellte Dateien? (Standard false)
 * @returns {Array<{file: string, delivery: object}>} chronologisch aufsteigend
 */
export function listPendingDeliveries({ dir = process.cwd(), includeDelivered = false } = {}) {
  const outbox = outboxDir(dir);
  if (!fs.existsSync(outbox)) return [];

  const entries = fs.readdirSync(outbox)
    .filter((f) => f.endsWith('.json'))
    .sort();

  const result = [];
  for (const f of entries) {
    const full = path.join(outbox, f);
    const isDelivered = f.endsWith(DELIVERED_SUFFIX);
    if (isDelivered && !includeDelivered) continue;

    try {
      const delivery = JSON.parse(fs.readFileSync(full, 'utf8'));
      result.push({ file: full, delivery });
    } catch (e) {
      // kaputte Zustell-Datei überspringen, loggen auf stderr
      console.error(`[delivery] WARN  Outbox-Datei nicht lesbar: ${full} (${e.message})`);
    }
  }
  return result;
}

/**
 * Markiert eine Outbox-Nachricht als "übergeben" (an den Nutzer zugestellt). Renamed
 * die Datei, damit listPendingDeliveries sie künftig überspringt — kein Doppel-Versand.
 *
 * @param {string} file - voller Pfad einer Outbox-Datei (aus listPendingDeliveries)
 * @returns {{delivered: boolean, file: string}}
 */
export function markDelivered(file) {
  if (!file.endsWith('.json') || file.endsWith(DELIVERED_SUFFIX)) {
    return { delivered: false, file };
  }
  const deliveredFile = file.replace(/\.json$/, DELIVERED_SUFFIX);
  fs.renameSync(file, deliveredFile);
  return { delivered: true, file: deliveredFile };
}
