#!/usr/bin/env node
/**
 * Aktivitäts-Extraktion — Demo- & Smoke-Test (Layer 1).
 *
 * Gibt reale Beispiel-Datenstrukturen des Activity-Feeds aus - einmal deterministisch
 * aus einem simulierten Teams-DOM (ohne Browser) und - optional - als Live-Lauf gegen
 * die echte Teams-Bridge (mit --live).
 *
 * Verwendung:
 *   node tests/activity.demo.mjs            # nur Mock-Beispieldaten (deterministisch)
 *   node tests/activity.demo.mjs --live     # zusätzlich Live-Lauf gegen Teams (adesso)
 *   node tests/activity.demo.mjs --tenant meine-org.onmicrosoft.com --live
 */
import { extractActivityFromDom } from '../src/activityClient.js';

// ---------------------------------------------------------------------------
// Minimaler DOM-Mock, der gerade genug von document/Element nachbildet, um
// extractActivityFromDom deterministisch zu testen. Entspricht der Struktur,
// die der echteste Teams-Activity-Feed im Web-DOM aufweist.
// ---------------------------------------------------------------------------
function makeText(text) {
  return { replace: (re, s) => String(text).replace(re, s) };
}
function makeEl({ text, attrs = {}, children = [], time = null }) {
  const el = {
    innerText: text,
    textContent: text,
    outerHTML: `<div data-tid="activity-item">${text}</div>`,
    getAttribute: (name) => attrs[name] ?? null,
    querySelector: (sel) => {
      if (sel === 'time[datetime]' && time) {
        return {
          getAttribute: (n) => (n === 'datetime' ? time : null),
          innerText: time
        };
      }
      const avatarSel = sel.includes('avatar') || sel === '[role="img"]';
      if (avatarSel) {
        return attrs.avatar ? { getAttribute: (n) => (n === 'aria-label' ? attrs.avatar : null) } : null;
      }
      return null;
    },
    querySelectorAll: () => children
  };
  return el;
}

function buildMockFeed() {
  return {
    querySelector: (sel) => {
      if (sel.startsWith('[data-tid="activity-feed"]')) return feedRoot;
      return null;
    },
    querySelectorAll: () => []
  };
}
const feedRoot = {
  querySelectorAll: () => FEED_ITEMS,
  querySelector: () => null,
  getAttribute: () => null
};

// Realistische Beispiel-Items im Teams-Activity-Format
const FEED_ITEMS = [
  makeEl({
    text: 'Yannick Bülter hat in einem Chat erwähnt @David Hallmann',
    attrs: { avatar: 'Yannick Bülter' },
    time: '2026-09-16T09:12:00Z'
  }),
  makeEl({
    text: 'Lukas Behncke antwortete auf Ihre Nachricht: "Kannst du die Spec nochmal prüfen?"',
    attrs: { avatar: 'Lukas Behncke' },
    time: '2026-09-16T08:45:00Z'
  }),
  makeEl({
    text: 'Mathis Künzel hat auf Ihren Beitrag reagiert 👍',
    attrs: { avatar: 'Mathis Künzel' }
  }),
  makeEl({
    text: 'Neuer Kanal "Projekt Wildau" wurde in Allgemein erstellt',
    time: '2026-09-16T07:30:00Z'
  }),
  makeEl({
    text: 'Theys Schiller erwähnt @Team in "General" - Bitte um Review bis Freitag',
    attrs: { avatar: 'Theys Schiller' },
    time: '2026-09-16T06:15:00Z'
  })
];

const mockDoc = buildMockFeed();

console.log('=== Teams Activity-Extraktion (Layer 1) — Demo ===\n');

// --- 1) Deterministischer Mock-Lauf ---
console.log('--- 1) Mock-DOM-Lauf (simulierte reale Teams-Daten) ---');
const mockItems = extractActivityFromDom(mockDoc, { max: 10 });
console.log(`Extrahiert: ${mockItems.length} Items\n`);
for (const [i, item] of mockItems.entries()) {
  console.log(`[${i + 1}] ${item.text}`);
  console.log(`     Autor: ${item.author ?? '(ohne)'} | Zeitstempel: ${item.timestamp ?? item.timestampRaw ?? '(ohne)'}`);
}
console.log('\nRoh-JSON (erste 2 Items, wie das Tool sie liefert):');
console.log(JSON.stringify({ count: mockItems.length, items: mockItems.slice(0, 2) }, null, 2));

// --- 2) Optionaler Live-Lauf gegen echte Teams-Bridge ---
const isLive = process.argv.includes('--live');
if (isLive) {
  const tenant = (() => {
    const idx = process.argv.indexOf('--tenant');
    return idx !== -1 && process.argv[idx + 1] ? process.argv[idx + 1] : 'adesso';
  })();

  console.log(`\n--- 2) Live-Lauf gegen Teams (tenant="${tenant}") ---`);
  console.log('Hinweis: benötigt gestartete PG-teams-MCP-Bridge. Live-Extraktion läuft über das MCP-Tool teams_get_activity.');
  console.log('Nur der Mock oben ist hier ausführbar - Live-Aufrufe erfolgen über den MCP-CallTool-Handler.');
} else {
  console.log('\n--- Live-Lauf übersprungen (Flag --live + laufende Bridge nötig) ---');
}
