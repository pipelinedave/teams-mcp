#!/usr/bin/env node

import path from 'path';
import fs from 'fs';
import { teamsClient } from '../src/teamsClient.js';
import { browserManager } from '../src/browserManager.js';

function parseArgs(args) {
  const parsed = { command: '', options: {} };
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === '--tenant' || arg === '-t') {
      parsed.options.tenant = args[++i];
    } else if (arg === '--out' || arg === '-o') {
      parsed.options.out = args[++i];
    } else if (arg === '--meeting-url' || arg === '-u') {
      parsed.options.meetingUrl = args[++i];
    } else if (arg === '--speaker' || arg === '-s') {
      parsed.options.speaker = args[++i];
    } else if (arg === '--no-join') {
      parsed.options.noJoin = true;
    } else if (arg === '--help' || arg === '-h') {
      parsed.options.help = true;
    } else if (!arg.startsWith('-') && !parsed.command) {
      parsed.command = arg;
    }
  }
  if (!parsed.command) parsed.command = 'status';
  return parsed;
}

function printUsage() {
  console.log(`
Teams MCP - Active Speaker Tracker CLI

Verwendung:
  node bin/track-speakers.js status [--tenant <name>]
  node bin/track-speakers.js record [--tenant <name>] [--out <pfad/zu/meeting.speakers.json>] [--meeting-url <url>] [--speaker <name>] [--no-join]
  node bin/track-speakers.js start  [--tenant <name>] [--out <pfad/zu/meeting.speakers.json>] [--speaker <name>]
  node bin/track-speakers.js stop   [--tenant <name>] [--out <pfad/zu/meeting.speakers.json>]

Optionen:
  --tenant, -t      Tenant-Name (z.B. adesso, dvelop; Standard: adesso)
  --out, -o         Ausgabepfad für die *.speakers.json (Standard: ./meeting.speakers.json)
  --meeting-url, -u Teams-Meeting-Link direkt angeben (sonst: auto aus Kalender / Chat)
  --speaker, -s     Konkreter Gesprächspartner für 1:1 Calls (z.B. "Theys Schiller")
  --no-join         Nicht automatisch beitreten (nur tracken falls bereits im Meeting)
  --help, -h        Diese Hilfe anzeigen
`);
}

async function main() {
  const { command, options } = parseArgs(process.argv.slice(2));

  if (options.help || command === 'help') {
    printUsage();
    process.exit(0);
  }

  const tenant = options.tenant || 'adesso';

  try {
    switch (command) {
      case 'status': {
        const status = await teamsClient.getMeetingStatus(tenant);
        console.log(JSON.stringify(status, null, 2));
        process.exit(0);
        break;
      }

      case 'record': {
        const outPath = options.out ? path.resolve(options.out) : path.resolve(process.cwd(), `Meeting_${Date.now()}.speakers.json`);
        const pidFile = `${outPath}.pid`;
        const stopFile = `${outPath}.stop`;

        console.log(`[*] Starte Speaker Tracking (Live-Record) für Tenant '${tenant}'...`);
        if (options.speaker) console.log(`[*] 1:1 Gesprächspartner angegeben: ${options.speaker}`);
        if (options.meetingUrl) console.log(`[*] Meeting-URL: ${options.meetingUrl}`);
        else if (!options.noJoin) console.log(`[*] Suche aktives Meeting oder Chat im Teams-Client...`);
        console.log(`[*] Zieldatei: ${outPath}`);

        try {
          fs.writeFileSync(pidFile, String(process.pid), 'utf-8');
        } catch (_) {}

        const trackingOptions = {
          meetingUrl: options.meetingUrl || null,
          noJoin: options.noJoin || false,
          speaker: options.speaker || null,
        };
        const res = await teamsClient.startSpeakerTracking(tenant, outPath, trackingOptions);
        console.log(JSON.stringify(res, null, 2));
        console.log('[*] Tracking aktiv. Beenden via Signaldatei (.stop), Strg+C oder Beenden des Elternprozesses.');

        let isStopping = false;
        const handleShutdown = async () => {
          if (isStopping) return;
          isStopping = true;
          console.log(`\n[*] Beende Speaker Tracking und schreibe Timeline...`);
          try {
            const stopRes = await teamsClient.stopSpeakerTracking(tenant);
            console.log(JSON.stringify(stopRes, null, 2));
          } catch (e) {
            console.error('[!] Fehler beim Stoppen:', e.message);
          } finally {
            try { if (fs.existsSync(pidFile)) fs.unlinkSync(pidFile); } catch (_) {}
            try { if (fs.existsSync(stopFile)) fs.unlinkSync(stopFile); } catch (_) {}
          }
          process.exit(0);
        };

        process.on('SIGINT', handleShutdown);
        process.on('SIGTERM', handleShutdown);

        // Polle auf Stop-Signaldatei alle 300ms (für PowerShell/Batch-Steuerung ohne Signalverlust)
        const checkStopInterval = setInterval(async () => {
          try {
            if (fs.existsSync(stopFile)) {
              clearInterval(checkStopInterval);
              await handleShutdown();
            }
          } catch (_) {}
        }, 300);

        break;
      }

      case 'start': {
        const outPath = options.out ? path.resolve(options.out) : path.resolve(process.cwd(), `Meeting_${Date.now()}.speakers.json`);
        console.log(`[*] Starte Speaker Tracking für Tenant '${tenant}'...`);
        console.log(`[*] Zieldatei: ${outPath}`);
        const trackingOptions = {
          meetingUrl: options.meetingUrl || null,
          noJoin: options.noJoin || false,
          speaker: options.speaker || null,
        };
        const res = await teamsClient.startSpeakerTracking(tenant, outPath, trackingOptions);
        console.log(JSON.stringify(res, null, 2));
        process.exit(0);
        break;
      }

      case 'stop': {
        console.log(`[*] Stoppe Speaker Tracking für Tenant '${tenant}'...`);
        if (options.out) {
          const outPath = path.resolve(options.out);
          const stopFile = `${outPath}.stop`;
          const pidFile = `${outPath}.pid`;

          try {
            fs.writeFileSync(stopFile, 'stop', 'utf-8');
            // Warte bis zu 4s bis Prozess beendet ist
            for (let i = 0; i < 40; i++) {
              await new Promise(r => setTimeout(r, 100));
              if (!fs.existsSync(pidFile)) break;
            }
          } catch (_) {}
        }

        const res = await teamsClient.stopSpeakerTracking(tenant);
        console.log(JSON.stringify(res, null, 2));
        process.exit(0);
        break;
      }

      default:
        console.error(`[!] Unbekannter Befehl: ${command}`);
        printUsage();
        process.exit(1);
    }
  } catch (err) {
    console.error(`[!] Fehler:`, err.message);
    process.exit(1);
  }
}

main();
