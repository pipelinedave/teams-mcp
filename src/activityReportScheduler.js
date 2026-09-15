import fs from 'fs';
import path from 'path';
import { analyzeActivityTab, CATEGORIES } from './activityAnalyzer.js';

/**
 * Layer-3-Planer: geplanter proaktiver Activity-Bericht (Cron-Job).
 *
 * Führt 2x täglich (Standard 09:00 + 17:00) analyzeActivityTab() aus, speichert das
 * Ergebnis zusätzlich als zeitgestempelte Markdown-Datei nach
 * reports/activity-YYYY-MM-DD-HHmm.md und sendet den Bericht per sendProactiveReport()
 * an den Nutzer.
 *
 * Design:
 *   - Kein externes Cron-Paket (Repro hat ohnehin keins) — schlanker interner
 *     Timer-basierter Scheduler, der zu konfigurierbaren HH:MM-Zeiten feuert.
 *   - sendProactiveReport ist injizierbar (testbar); Default sendet an die Konsole
 *     (Terminal) und hält den Bezug zur bestehenden Design-Entscheidung (Konsole).
 *   - Fehlerbehandlung: ein fehlgeschlagener Lauf crasht den Scheduler nicht
 *     (try/catch + Logging), sondern wird protokolliert — kein stilles Scheitern.
 *   - Logging über stderr (verschmutzt den Stdio-JSON-RPC-Kanal des MCP-Servers nicht).
 */

// Standard-Zeitfenster (lokal Europe/Berlin): morgens + nachmittags
export const DEFAULT_SCHEDULE = ['09:00', '17:00'];

// Standard-Logger (stderr)
function defaultLogger() {
  return {
    info: (m, meta) => console.error(`[scheduler] INFO  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    warn: (m, meta) => console.error(`[scheduler] WARN  ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`),
    error: (m, meta) => console.error(`[scheduler] ERROR ${m}${meta ? ' ' + JSON.stringify(meta) : ''}`)
  };
}

/**
 * Formatiert HH:MM in Minuten seit Mitternacht (rein numerischer Vergleich).
 * @param {string} hhmm - z.B. "09:00"
 * @returns {number} Minuten seit Mitternacht
 */
export function toMinutes(hhmm) {
  if (!hhmm || typeof hhmm !== 'string') return NaN;
  const [h, m] = hhmm.split(':').map(Number);
  if (!Number.isInteger(h) || !Number.isInteger(m) || h < 0 || h > 23 || m < 0 || m > 59) return NaN;
  return h * 60 + m;
}

/**
 * Schreibt eine zeitgestempelte Berichtsdatei reports/activity-YYYY-MM-DD-HHmm.md.
 * @param {string} markdown
 * @param {object} [opts]
 * @param {Date} [opts.now] - Zeitpunkt (Standard jetzt)
 * @param {string} [opts.dir] - Basis-Verzeichnis
 * @returns {{filePath: string, stamp: string}}
 */
export function writeTimestampedReport(markdown, { now = new Date(), dir = process.cwd() } = {}) {
  const pad = (n) => String(n).padStart(2, '0');
  const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
  const reportsDir = path.join(dir, 'reports');
  fs.mkdirSync(reportsDir, { recursive: true });
  const filePath = path.join(reportsDir, `activity-${stamp}.md`);
  fs.writeFileSync(filePath, markdown, 'utf8');
  return { filePath, stamp };
}

/**
 * Proaktiver Berichts-Versand an den Nutzer.
 *
 * Default-Ausgabe: Konsole (Terminal). Da dieser Scheduler im MCP-Server-Prozess läuft
 * und stdout für den Stdio-JSON-RPC reserviert ist, wird die Konsole über stderr
 * angesprochen bzw. über einen injizierbaren Sender. Der Sender ist die einzige
 * Schnittstelle zur "Verschickung" — so kann über die aufrufende Ebene (z.B. opencode
 * / Agent) auch eine Mail/Teams/Notiz stattfinden, ohne Scheduler-Änderung.
 *
 * @param {object} result - Ergebnis von analyzeActivityTab()
 * @param {object} [opts]
 * @param {function} [opts.sender] - async (result) => void; Standard: Konsole
 * @param {object} [opts.logger]
 * @returns {Promise<{sentTo: string}>}
 */
export async function sendProactiveReport(result, { sender = null, logger = null } = {}) {
  const log = logger || defaultLogger();

  // Konsole-Sender als Default (stderr, da stdout MCP-reserviert ist)
  const defaultSender = async (r) => {
    const text = [
      '\n================= TEAMS ACTIVITY-BERICHT =================',
      r.report,
      '===========================================================\n'
    ].join('\n');
    // Mut: Auf stderr, damit der MCP-Stdio-Kanal (stdout) sauber bleibt.
    process.stderr.write(text);
  };

  const send = sender || defaultSender;
  try {
    await send(result);
    log.info(`Proaktiver Bericht gesendet`, { tenant: result.tenant, date: result.date });
    return { sentTo: sender ? 'custom' : 'console(stderr)' };
  } catch (err) {
    log.error(`Senden des proaktiven Berichts fehlgeschlagen`, { error: err.message });
    throw err;
  }
}

/**
 * Aktivitäts-Report-Scheduler (Cron-artig, 2x täglich).
 *
 * @example
 * const sched = new ActivityReportScheduler({ times: ['09:00','17:00'] });
 * sched.start();
 * sched.stop();
 */
export class ActivityReportScheduler {
  /**
   * @param {object} [opts]
   * @param {string[]} [opts.times] - HH:MM-Zeiten (Standard DEFAULT_SCHEDULE)
   * @param {string} [opts.tenant] - Teams-Tenant (Standard "adesso")
   * @param {number} [opts.maxItems]
   * @param {number} [opts.topN]
   * @param {string} [opts.dir] - Basis für reports/
   * @param {object} [opts.logger]
   * @param {function} [opts.sender] - injizierbare Send-Funktion für sendProactiveReport
   * @param {function} [opts.analyzeFn] - injizierbare Analyze-Funktion (Standard analyzeActivityTab) für Tests
   */
  constructor(opts = {}) {
    this.times = (opts.times || DEFAULT_SCHEDULE).filter((t) => !isNaN(toMinutes(t)));
    this.tenant = opts.tenant || 'adesso';
    this.maxItems = opts.maxItems || 50;
    this.topN = opts.topN ?? 3;
    this.dir = opts.dir || process.cwd();
    this.logger = opts.logger || null;
    this.sender = opts.sender || null;
    this.analyzeFn = opts.analyzeFn || analyzeActivityTab;

    this._timer = null;
    this._running = false;
    this._lastFired = null; // "YYYY-MM-DD HH:MM" zur Entprellung (verhindert Doppel-Feuern)
  }

  /**
   * Nächste Fälligkeit (in ms seit jetzt) für eine HH:MM-Zeit.
   * Berücksichtigt, dass die Zeit ggf. erst morgen wieder ansteht.
   */
  _msUntil(targetMin) {
    const now = new Date();
    const nowMin = now.getHours() * 60 + now.getMinutes();
    let deltaMin = targetMin - nowMin;
    // Wenn die Zielzeit heute bereits vorbei ist, schiebe auf morgen
    if (deltaMin <= 0) deltaMin += 24 * 60;
    return deltaMin * 60 * 1000;
  }

  /**
   * Gibt die nächste geplante Zielzeit (Date) zurück, sortiert nach zeitlicher Nähe.
   */
  _nextTarget() {
    const now = new Date();
    const candidates = this.times
      .map((t) => {
        const min = toMinutes(t);
        const d = new Date(now);
        const today = now.getHours() * 60 + now.getMinutes();
        if (min <= today) d.setDate(d.getDate() + 1); // morgen
        d.setHours(Math.floor(min / 60), min % 60, 0, 0);
        return { min, date: d };
      })
      .sort((a, b) => a.date - b.date);
    return candidates[0];
  }

  /**
   * Führt einen einzelnen Aktivitäts-Lauf aus (analysieren + zeitgestempelt speichern +
   * proaktiv senden). Auch manuell aufrufbar (für Tests/On-Demand).
   */
  async runOnce({ now = new Date() } = {}) {
    const log = this.logger || defaultLogger();
    log.info(`Scheduler-Lauf gestartet`, { tenant: this.tenant, at: now.toISOString() });

    const result = await this.analyzeFn(this.tenant, {
      maxItems: this.maxItems,
      topN: this.topN,
      dir: this.dir
    });

    // Zusätzlich zeitgestempelte Datei schreiben (reports/activity-YYYY-MM-DD-HHmm.md)
    const stamped = writeTimestampedReport(result.report, { now, dir: this.dir });
    log.info(`Zeitgestempelte Berichtsdatei geschrieben`, { filePath: stamped.filePath });

    await sendProactiveReport(result, { sender: this.sender, logger: this.logger });

    return { ...result, stampedFile: stamped.filePath };
  }

  /**
   * Führt einen Lauf aus; fängt Fehler ab, um den Scheduler-Timer am Leben zu halten
   * (kein stilles Scheitern — Fehler werden geloggt, der Job läuft weiter).
   */
  async _safeRun() {
    const log = this.logger || defaultLogger();
    try {
      await this.runOnce();
    } catch (err) {
      log.error(`Scheduler-Lauf fehlgeschlagen (Job bleibt aktiv)`, { error: err.message });
    }
  }

  /**
   * Startet den periodischen Timer. Plant zunächst die nächste Zielzeit und
   * feuert danach zyklisch (Minuten-Takt zum Nachrücken auf die Zielzeit).
   */
  start() {
    if (this._running) return;
    this._running = true;
    const log = this.logger || defaultLogger();
    log.info(`Scheduler gestartet`, { tenant: this.tenant, times: this.times });

    const tick = async () => {
      if (!this._running) return;
      const target = this._nextTarget();
      const now = new Date();
      const key = `${target.date.getFullYear()}-${String(target.date.getMonth() + 1).padStart(2, '0')}-${String(target.date.getDate()).padStart(2, '0')} ${String(target.date.getHours()).padStart(2, '0')}:${String(target.date.getMinutes()).padStart(2, '0')}`;

      // Feuern, wenn die Zielzeit erreicht (jetzt >= Ziel) und noch nicht für diesen Slot gefeuert wurde
      if (now >= target.date && this._lastFired !== key) {
        this._lastFired = key;
        await this._safeRun();
      }

      // Nachrücken im 30s-Takt, damit wir die Zielzeit genau treffen
      this._timer = setTimeout(tick, this.times.length ? 30 * 1000 : 60 * 1000);
    };

    this._timer = setTimeout(tick, 1000);
    return this;
  }

  /**
   * Stoppt den Scheduler sauber.
   */
  stop() {
    this._running = false;
    if (this._timer) {
      clearTimeout(this._timer);
      this._timer = null;
    }
    const log = this.logger || defaultLogger();
    log.info(`Scheduler gestoppt`, { tenant: this.tenant });
    return this;
  }
}

export const activityReportScheduler = new ActivityReportScheduler();

/**
 * Zentrale Steuerungslogik für den Scheduler (status/start/stop/run-once/config).
 *
 * Als pure, exportierte Funktion gebaut, damit sie in Unit-Tests ohne MCP-Kanal
 * prüfbar ist. Der index.js-Handler ruft sie nur noch auf und verpackt das Ergebnis
 * in eine MCP-Text-Antwort.
 *
 * @param {object} opts
 * @param {string} [opts.action] - 'status' | 'start' | 'stop' | 'run-once' | 'config' (Standard 'status')
 * @param {ActivityReportScheduler} [opts.scheduler] - injizierbarer Scheduler (Standard activityReportScheduler)
 * @param {string} [opts.tenant]
 * @param {number} [opts.maxItems]
 * @param {string[]} [opts.times]
 * @param {function} [opts.resolveTenant] - tenant-Resolver (falls nötig)
 * @returns {Promise<object>} serialisierbares Ergebnis
 */
export async function runSchedulerControl({
  action = 'status',
  scheduler = activityReportScheduler,
  tenant,
  maxItems,
  times,
  resolveTenant = null
} = {}) {
  const applyConfig = () => {
    let changed = false;
    const wantTenant = tenant ? (resolveTenant ? resolveTenant(tenant) : tenant) : scheduler.tenant;
    const wantMax = maxItems || scheduler.maxItems;
    if (wantTenant !== scheduler.tenant) { scheduler.tenant = wantTenant; changed = true; }
    if (wantMax !== scheduler.maxItems) { scheduler.maxItems = wantMax; changed = true; }
    if (Array.isArray(times) && times.length) {
      const valid = times.filter((t) => !isNaN(toMinutes(t)));
      if (valid.length) { scheduler.times = valid; changed = true; }
    }
    return changed;
  };

  switch (action) {
    case 'status': {
      return {
        running: scheduler._running,
        tenant: scheduler.tenant,
        times: scheduler.times,
        maxItems: scheduler.maxItems,
        topN: scheduler.topN,
        lastFired: scheduler._lastFired,
        reportsDir: `${scheduler.dir}/reports`
      };
    }

    case 'config': {
      return {
        running: scheduler._running,
        tenant: scheduler.tenant,
        times: scheduler.times,
        maxItems: scheduler.maxItems,
        topN: scheduler.topN,
        dir: scheduler.dir
      };
    }

    case 'start': {
      applyConfig();
      if (scheduler._running) {
        return { started: false, alreadyRunning: true, times: scheduler.times };
      }
      scheduler.start();
      return { started: true, times: scheduler.times, tenant: scheduler.tenant };
    }

    case 'stop': {
      scheduler.stop();
      return { stopped: true };
    }

    case 'run-once': {
      applyConfig();
      const out = await scheduler.runOnce();
      return {
        date: out.date,
        tenant: out.tenant,
        counts: out.counts,
        filePath: out.filePath,
        stampedFile: out.stampedFile,
        report: out.report
      };
    }

    default:
      return { error: `Unbekannte Aktion: ${action}` };
  }
}
