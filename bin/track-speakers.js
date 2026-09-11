#!/usr/bin/env node

import path from 'path';
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
  node bin/track-speakers.js record [--tenant <name>] [--out <pfad/zu/meeting.speakers.json>] [--meeting-url <url>] [--no-join]
  node bin/track-speakers.js start  [--tenant <name>] [--out <pfad/zu/meeting.speakers.json>]
  node bin/track-speakers.js stop   [--tenant <name>]

Optionen:
  --tenant, -t   Tenant-Name (z.B. adesso, dvelop; Standard: adesso)
  --out, -o      Ausgabepfad für die *.speakers.json (Standard: ./meeting.speakers.json)
  --meeting-url, -u  Teams-Meeting-Link direkt angeben (sonst: auto aus Kalender)
  --no-join      Nicht automatisch beitreten (nur tracken falls bereits im Meeting)
  --help, -h     Diese Hilfe anzeigen
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
        console.log(`[*] Starte Speaker Tracking (Live-Record) für Tenant '${tenant}'...`);
        if (options.meetingUrl) console.log(`[*] Meeting-URL: ${options.meetingUrl}`);
        else if (!options.noJoin) console.log(`[*] Suche aktuelles Meeting im Teams-Kalender...`);
        console.log(`[*] Zieldatei: ${outPath}`);
        const trackingOptions = {
          meetingUrl: options.meetingUrl || null,
          noJoin: options.noJoin || false,
        };
        const res = await teamsClient.startSpeakerTracking(tenant, outPath, trackingOptions);
        console.log(JSON.stringify(res, null, 2));
        console.log('[*] Tracking aktiv. Beenden mit Strg+C (SIGINT) oder Beenden des Elternprozesses.');

        const handleShutdown = async () => {
          console.log(`\n[*] Beende Speaker Tracking und schreibe Timeline...`);
          try {
            const stopRes = await teamsClient.stopSpeakerTracking(tenant);
            console.log(JSON.stringify(stopRes, null, 2));
          } catch (e) {
            console.error('[!] Fehler beim Stoppen:', e.message);
          }
          process.exit(0);
        };

        process.on('SIGINT', handleShutdown);
        process.on('SIGTERM', handleShutdown);
        // Prozess am Leben halten
        setInterval(() => {}, 10000);
        break;
      }

      case 'start': {
        const outPath = options.out ? path.resolve(options.out) : path.resolve(process.cwd(), `Meeting_${Date.now()}.speakers.json`);
        console.log(`[*] Starte Speaker Tracking für Tenant '${tenant}'...`);
        console.log(`[*] Zieldatei: ${outPath}`);
        const trackingOptions = {
          meetingUrl: options.meetingUrl || null,
          noJoin: options.noJoin || false,
        };
        const res = await teamsClient.startSpeakerTracking(tenant, outPath, trackingOptions);
        console.log(JSON.stringify(res, null, 2));
        process.exit(0);
        break;
      }

      case 'stop': {
        console.log(`[*] Stoppe Speaker Tracking für Tenant '${tenant}'...`);
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
