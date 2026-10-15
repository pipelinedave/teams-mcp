import fs from 'fs';
import path from 'path';

/**
 * Normalizes speaker name by removing Teams suffixes like "(Gast)", "(Organizer)", "(External)".
 */
export function cleanSpeakerName(rawName) {
  if (!rawName) return '';
  return rawName
    .replace(/\s*\((Gast|Guest|Organisator|Organizer|Extern|External|Präsentator|Presenter)\)/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Set of known non-human Teams entries that must never be treated as a meeting
 * participant or 1:1 call partner (Teams Copilot/Notebook, transcription bots,
 * busy-status placeholders, UI labels).
 */
const BOT_NAME_PATTERNS = [
  /\bcopilot\b/i,
  /\bnotebook\b/i,
  /\brecording(?: is)? in progress\b/i,
  /\btranscription(?: has)? started\b/i,
  /\blive caption(s)?\b/i,
  /\bmeet(?:ing)? chat\b/i,
  /\bchat\b/i,
  /\baways?\b/i,
  /\bappear(?:ance)?\b/i,
  /\bstaus\b/i,
];

/**
 * True if the name refers to a Team UI pseudo-entry / bot (e.g. "Copilot")
 * rather than a real human participant or 1:1 call partner.
 */
export function isBotOrUiName(name) {
  const n = (name || '').trim();
  if (!n) return true;
  const lower = n.toLowerCase();
  return BOT_NAME_PATTERNS.some((re) => re.test(lower)) || lower.includes('copilot');
}

/**
 * Generische UI-Wörter, die nie ein echter Sprecher-/Teilnehmername sind
 * (Video-Tile-Overlays und aria-labels enthalten oft Status-Wörter wie
 * "Anwesend" oder die eigene Person als "Du"/"You").
 */
const UI_NAME_PATTERNS = [
  /^(du|you|ich|me|sie|er|es)$/i,
  /^(anwesend|present|in the meeting|im meeting|in this meeting|away|abwesend|offline|busy|beschäftigt|do not disturb|nicht stören)$/i,
  /^(guest|gast|organizer|organisator|presenter|präsentator|external|extern)$/i,
  /^(meeting|call|anruf|besprechung|teams)$/i,
  /^\d+$/,
];

/**
 * True if the string looks like a real human speaker/participant name and not
 * a UI label, status word or bot entry.
 */
export function isPlausibleSpeakerName(name) {
  const n = cleanSpeakerName(name || '');
  if (!n || n.length > 60) return false;
  if (n.includes('\\n') || n.includes('\n')) return false;
  if (isBotOrUiName(n)) return false;
  return !UI_NAME_PATTERNS.some((re) => re.test(n.trim()));
}

/**
 * Extrahiert den Sprechernamen aus einer aria-label, die eine
 * Active-Speaker-Anzeige enthaelt (DE + EN):
 *   "Stefan Maciolek spricht"            -> "Stefan Maciolek"
 *   "Stefan Maciolek is speaking"        -> "Stefan Maciolek"
 *   "Stefan Maciolek, speaking"          -> "Stefan Maciolek"
 *   "Stefan Maciolek spricht, Auflegen"  -> "Stefan Maciolek"
 * Liefert '', wenn kein sprechender Name erkannt wird.
 *
 * WICHTIG: Diese Funktion wird als Funktions-Source in page.evaluate injiziert
 * (siehe collectMeetingStateInPage). Regexes hier muessen SINGLE-backslash
 * sein — Playwright serialisiert die Funktion 1:1 in den Browser (doppelte
 * Backslashes wuerden dort literal auf "\s" matchen und nie greifen — genau
 * der Bug, der bis 15.10.2026 die Erkennung lahmlegte).
 */
export function speakerNameFromAriaLabel(aria) {
  if (!aria) return '';
  const a = String(aria).trim();
  if (!a) return '';
  // Name am Anfang, danach optional Separator + "spricht"/"is speaking"/"speaking"
  const m = a.match(/^(.{1,60}?)\s*(?:,\s*|-\s*|:\s*)?(?:spricht|is speaking|speaking)\b(?:[,:].*)?$/i);
  if (m) return m[1].trim();
  // Variante "Spricht: Name" / "Speaking: Name"
  const m2 = a.match(/^(?:spricht|is speaking|speaking)\s*:\s*(.{1,60})$/i);
  if (m2) return m2[1].trim();
  return '';
}

/**
 * Sektions-/Nav-Titel im Teams-Chat-Baum, die nie ein echter 1:1-Chat-Partner
 * sind. Das neue Teams-Web rendert die Chatliste als role="treeitem"-Baum mit
 * Section-Headern ("Copilot", "Favoriten", "Chats", ...) VOR den echten
 * Chat-Eintraegen — blindes chatRows[0] lieferte daher immer "Copilot"
 * (Root Cause der leeren Timelines bis 15.10.2026). Eigene Liste statt Import
 * aus teamsClient.js, um den Import-Zyklus teamsClient <-> speakerTracker
 * zu vermeiden (Bewusstes, kommentiertes Duplikat).
 */
const CHAT_SECTION_TITLES = new Set([
  'copilot', 'quick views', 'schnellansicht', 'drafts', 'entwürfe', 'entwuerfe',
  'favorites', 'favoriten', 'chats', 'teams', 'kanäle', 'kanaele', 'channels',
  'activity', 'aktivität', 'aktivitaet', 'calendar', 'kalender', 'anrufe', 'calls',
  'communities', 'outlook', 'mehr', 'more', 'unread', 'ungelesen', 'mentions',
  'erwähnungen', 'erwaehnungen', 'followed threads', 'gefolgte threads',
  'discover', 'entdecken', 'meine teams', 'your teams', 'hidden teams',
  'ausgeblendete teams', 'versteckte teams', 'manage teams', 'teams verwalten',
  'create a team', 'team erstellen', 'gemeinschaften',
]);

/**
 * Wählt aus DOM-geordneten Chat-Kandidaten ({ title }) den ersten echten
 * (nicht-Bot, keine Sektions-/Nav-Überschrift) Chat-Partner. Liefert den
 * bereinigten Titel oder null, wenn kein echter Partner gefunden wurde.
 *
 * Bewusst PURE gehalten (kein DOM-Zugriff), damit sie ohne Browser testbar ist.
 */
export function pickRealChatPartner(candidates) {
  for (const c of candidates || []) {
    const rawTitle = String((c && (c.title ?? c)) || '');
    const title = cleanSpeakerName(rawTitle.split('\n')[0] || '');
    if (!title) continue;
    if (isBotOrUiName(title)) continue;
    const lower = title.toLowerCase();
    if (CHAT_SECTION_TITLES.has(lower)) continue;
    if (/^(see |alle |mehr anzeigen|show |join |beim |ansehen )/i.test(lower)) continue;
    if (isPlausibleSpeakerName(title) || title.length >= 3) {
      // Gruppenchats/Mehrwort-Titel sind auch gueltige "Partner" — nur einzelne
      // UI-Woerter wurden oben schon gefiltert.
      return title;
    }
  }
  return null;
}

/**
 * Aggregates raw discrete timestamp samples or state transitions into clean, continuous intervals.
 * Merges short speech pauses (< minGapMs) and ignores momentary noise spikes (< minDurationMs).
 */
export function aggregateEvents(samples, options = {}) {
  const { minGapMs = 400, minDurationMs = 250, meetingStartEpochMs = 0 } = options;
  if (!samples || samples.length === 0) return [];

  // Group samples by speaker
  const bySpeaker = new Map();
  for (const sample of samples) {
    const speaker = cleanSpeakerName(sample.speaker);
    if (!speaker) continue;
    if (isBotOrUiName(speaker)) continue; // Bots/UI-Eintraege niemals als Sprecher
    if (!bySpeaker.has(speaker)) bySpeaker.set(speaker, []);
    bySpeaker.get(speaker).push(sample.timestamp);
  }

  const allIntervals = [];

  for (const [speaker, timestamps] of bySpeaker.entries()) {
    // Sort timestamps ascending
    timestamps.sort((a, b) => a - b);

    let intervalStart = timestamps[0];
    let intervalEnd = timestamps[0];

    for (let i = 1; i < timestamps.length; i++) {
      const current = timestamps[i];
      if (current - intervalEnd <= minGapMs) {
        // Continuous speech within gap threshold
        intervalEnd = current;
      } else {
        // Gap detected -> close previous interval if long enough
        const duration = intervalEnd - intervalStart;
        if (duration >= minDurationMs) {
          allIntervals.push({
            speaker,
            start_epoch_ms: intervalStart,
            end_epoch_ms: intervalEnd,
            start_offset_sec: Math.max(0, (intervalStart - meetingStartEpochMs) / 1000),
            end_offset_sec: Math.max(0, (intervalEnd - meetingStartEpochMs) / 1000)
          });
        }
        intervalStart = current;
        intervalEnd = current;
      }
    }

    // Push final interval
    if (intervalEnd - intervalStart >= minDurationMs) {
      allIntervals.push({
        speaker,
        start_epoch_ms: intervalStart,
        end_epoch_ms: intervalEnd,
        start_offset_sec: Math.max(0, (intervalStart - meetingStartEpochMs) / 1000),
        end_offset_sec: Math.max(0, (intervalEnd - meetingStartEpochMs) / 1000)
      });
    }
  }

  // Sort chronologically by start time
  allIntervals.sort((a, b) => a.start_epoch_ms - b.start_epoch_ms);
  return allIntervals;
}

// ---------------------------------------------------------------------------
// IN-PAGE-DOM-LOGIK
// ---------------------------------------------------------------------------
// Die folgende Funktion wird ALS SOURCE in page.evaluate injiziert (gleiche
// Technik wie teamsClient.expandChatList: Funktions-Source + new Function).
// Sie ist bewusst self-contained und bekommt die Node-seitigen Pure-Funktionen
// (speakerNameFromAriaLabel, isPlausibleSpeakerName, isBotOrUiName) als
// helpers-Objekt injiziert — EINE Source of Truth fuer Node-Tests und Browser.
//
// deep=true  -> voller Meeting-State { inMeeting, meetingTitle, participants, activeSpeakers }
// deep=false -> nur { activeSpeakers } (leichtgewichtig, 100ms-Sampler)
// (exportiert fuer Unit-Tests mit Fake-DOM)
export function collectMeetingStateInPage(deep, helpers) {
  const speakerNameFromAriaLabel = helpers.speakerNameFromAriaLabel;
  const isPlausibleSpeakerName = helpers.isPlausibleSpeakerName;
  const isBotOrUiName = helpers.isBotOrUiName;

  // --- Aktive Sprecher erkennen (mehrere Fallback-Layer gegen kosmetische DOM-Aenderungen) ---
  const activeSelector = [
    '[data-tid="active-speaker-border"]',
    '[data-tid="activeSpeaker"]',
    '[data-is-speaking="true"]',
    '[data-speaker-state="speaking"]',
    '[data-tid*="speaking" i]',
    '[class*="activeSpeaker" i]',
    '[class*="active-speaker" i]',
    '[class*="speaking" i]',
    '.speaking',
    '[aria-label*="spricht" i]',
    '[aria-label*="is speaking" i]',
    '[aria-label*=", speaking" i]',
  ].join(', ');

  // Namens-Overlay-Knoten in Tiles (alte + neue UI)
  const NAME_NODE_SEL = [
    '[data-tid="roster-avatar-name"]',
    '[data-tid="video-tile-name"]',
    '[data-tid="participant-name"]',
    '[data-tid*="avatar-name" i]',
    '[data-tid*="tile-name" i]',
    '[class*="participant-name" i]',
    '[class*="displayName" i]',
  ].join(', ');

  // Container, in denen ein aktives Element sitzt (spezifisch -> generisch)
  const TILE_CONTAINER_SEL = [
    '[data-tid*="video-tile" i]',
    '[data-tid*="participant" i]',
    '[data-tid*="roster" i]',
    '[data-tid*="tile" i]',
    '[role="listitem"]',
  ].join(', ');

  function nameFromNode(node) {
    if (!node) return '';
    // 1) Namens-Overlay im Knoten selbst / im Knoten-Baum
    const nameNode = node.matches?.(NAME_NODE_SEL) ? node : node.querySelector?.(NAME_NODE_SEL);
    if (nameNode) {
      const t = (nameNode.innerText || nameNode.textContent || '').split('\n')[0].trim();
      if (t && isPlausibleSpeakerName(t)) return t;
    }
    // 2) aria-label des Knotens (evtl. "X spricht"-Formular)
    const aria = node.getAttribute?.('aria-label') || '';
    if (aria) {
      const fromAria = speakerNameFromAriaLabel(aria);
      if (fromAria && isPlausibleSpeakerName(fromAria)) return fromAria;
      const plain = aria.split('\n')[0].trim();
      if (plain && isPlausibleSpeakerName(plain)) return plain;
    }
    // 3) Textinhalt, erste Zeile
    const text = (node.innerText || node.textContent || '').split('\n')[0].trim();
    if (text && isPlausibleSpeakerName(text)) return text;
    return '';
  }

  function resolveActiveSpeakerName(el) {
    // Layer A: aria-label des aktiven Elements selbst ("X spricht")
    const viaAria = speakerNameFromAriaLabel(el.getAttribute('aria-label') || '');
    if (viaAria && isPlausibleSpeakerName(viaAria)) return viaAria;

    // Layer B: Namens-Overlay im aktiven Element
    const viaNode = nameFromNode(el);
    if (viaNode) return viaNode;

    // Layer C: naechster Tile-/Roster-Container (aufwaerts), dort Overlay oder aria-label
    const container = el.closest?.(TILE_CONTAINER_SEL);
    if (container) {
      const viaContainerNode = nameFromNode(container.querySelector(NAME_NODE_SEL));
      if (viaContainerNode) return viaContainerNode;
      const containerAria = container.getAttribute('aria-label') || '';
      const viaContainerAria = speakerNameFromAriaLabel(containerAria) || containerAria.split('\n')[0].trim();
      if (viaContainerAria && isPlausibleSpeakerName(viaContainerAria)) return viaContainerAria;
    }
    return '';
  }

  const activeSpeakers = [];
  const seen = new Set();
  try {
    const els = document.querySelectorAll(activeSelector);
    for (const el of els) {
      const name = resolveActiveSpeakerName(el);
      if (name && !seen.has(name)) {
        seen.add(name);
        activeSpeakers.push(name);
      }
    }
  } catch (e) {}

  if (!deep) return { activeSpeakers };

  // --- Meeting-/Call-Erkennung (mehrere Indikatoren) ---
  const inMeetingSelector = [
    '[data-tid="call-controls"]',
    '[data-tid="hangup-button"]',
    '[data-tid="hangup-btn"]',
    '[data-tid="calling-screen"]',
    '[data-tid="calling-stage"]',
    '[data-tid="call-status-bar"]',
    '[data-tid*="active-call" i]',
    'div[id*="calling-stage"]',
    'div[id*="calling-container"]',
    'button[aria-label*="Auflegen" i]',
    'button[aria-label*="Hang up" i]',
    'button[aria-label*="Leave" i]',
    'button[aria-label*="Verlassen" i]',
  ].join(', ');

  let inMeeting = false;
  try {
    inMeeting = !!document.querySelector(inMeetingSelector);
  } catch (e) {}

  // --- Meeting-Titel ---
  let meetingTitle = '';
  const titleSel = [
    '[data-tid="calling-header-title"]',
    '[data-tid="call-title"]',
    '[data-tid="header-meeting-title"]',
    '[data-tid="top-bar-title"]',
    '[data-tid*="meeting-title" i]',
    'h1[data-tid*="meeting"]',
  ].join(', ');
  try {
    const titleEl = document.querySelector(titleSel);
    if (titleEl) meetingTitle = (titleEl.innerText || titleEl.textContent || '').split('\n')[0].trim();
  } catch (e) {}
  if (!meetingTitle) {
    // Korrigierte Regex (bis 15.10.2026 stand hier eine doppelte-Backslash-Variante,
    // die nie matchen konnte: "\\s" matcht literalen Backslash, nicht Whitespace).
    const docTitle = document.title || '';
    const match = docTitle.match(/^(.+?)\s*\|\s*Microsoft Teams/i);
    if (match && !match[1].toLowerCase().includes('chat')) {
      meetingTitle = match[1].trim();
    }
  }

  // --- Teilnehmer (sichtbare Tiles/Roster) ---
  const participants = [];
  const pSeen = new Set();
  try {
    const participantEls = document.querySelectorAll(
      '[data-tid="roster-avatar-name"], [data-tid*="video-tile" i] [aria-label], [data-tid*="participant" i] [aria-label], [role="listitem"] [aria-label]'
    );
    for (const el of participantEls) {
      let name = (el.getAttribute('aria-label') || '').split('\n')[0].trim();
      if (!name) name = (el.innerText || '').split('\n')[0].trim();
      const cleaned = speakerNameFromAriaLabel(name) || name;
      if (cleaned && isPlausibleSpeakerName(cleaned) && !pSeen.has(cleaned) && cleaned.length < 50) {
        pSeen.add(cleaned);
        participants.push(cleaned);
      }
      if (participants.length >= 30) break;
    }
  } catch (e) {}

  return {
    inMeeting,
    meetingTitle: meetingTitle || 'Teams Meeting',
    participants,
    activeSpeakers
  };
}

// Helper-Quellen, die bei JEDER page.evaluate-Injektion mitgegeben werden.
// Playwright kann Funktionsargumente NICHT serialisieren (im Gegensatz zu
// Puppeteer) — deshalb werden die Pure-Funktionen als SOURCE uebergeben.
// WICHTIG: new Function materialisiert im GLOBALEN Scope — die internen
// Referenzen (isPlausibleSpeakerName -> cleanSpeakerName/isBotOrUiName/
// UI_NAME_PATTERNS, isBotOrUiName -> BOT_NAME_PATTERNS) wuerden dort
// ReferenceError werfen. Die Factory-Source verdrahtet alle Helfer UND ihre
// Konstanten in EINEM gemeinsamen Scope (EINE Source of Truth fuer
// Node-Tests und Browser).
const IN_PAGE_HELPERS_SRC = '(() => {\n'
  + `const BOT_NAME_PATTERNS = [${BOT_NAME_PATTERNS.map((r) => r.toString()).join(', ')}];\n`
  + `const UI_NAME_PATTERNS = [${UI_NAME_PATTERNS.map((r) => r.toString()).join(', ')}];\n`
  + `const cleanSpeakerName = ${cleanSpeakerName.toString()};\n`
  + `const isBotOrUiName = ${isBotOrUiName.toString()};\n`
  + `const isPlausibleSpeakerName = ${isPlausibleSpeakerName.toString()};\n`
  + `const speakerNameFromAriaLabel = ${speakerNameFromAriaLabel.toString()};\n`
  + 'return { cleanSpeakerName, isBotOrUiName, isPlausibleSpeakerName, speakerNameFromAriaLabel };\n'
  + '})()';

// Fuer Regression-Tests (Factory muss ohne Modul-Scope auswertbar sein)
export const IN_PAGE_HELPERS_SOURCE = IN_PAGE_HELPERS_SRC;

function materializeInPageHelpersSrc() {
  return IN_PAGE_HELPERS_SRC;
}

export class SpeakerTracker {
  constructor() {
    this.activeSessions = new Map(); // tenant -> { startTimeEpochMs, outputPath, timer, samples: [], meetingTitle: '' }
  }

  /**
   * Fuehrt collectMeetingStateInPage im Seitenkontext aus. deep=true liefert
   * den vollen State, deep=false nur die aktiven Sprecher.
   */
  async _collectState(page, deep = true) {
    if (!page || page.isClosed()) {
      return deep
        ? { inMeeting: false, meetingTitle: '', participants: [], activeSpeakers: [] }
        : { activeSpeakers: [] };
    }
    return await page.evaluate(({ src, helperSrc, deep }) => {
      const collect = new Function('return (' + src + ')')();
      const helpers = new Function('return (' + helperSrc + ')')();
      return collect(deep, helpers);
    }, {
      src: collectMeetingStateInPage.toString(),
      helperSrc: materializeInPageHelpersSrc(),
      deep
    }).catch(() => deep
      ? { inMeeting: false, meetingTitle: '', participants: [], activeSpeakers: [] }
      : { activeSpeakers: [] });
  }

  /**
   * Scans current page DOM to detect if Teams is inside an active meeting or call.
   */
  async inspectMeeting(page) {
    return await this._collectState(page, true);
  }

  /**
   * Handles the Teams pre-join screen: turns off mic + camera, clicks "Join now",
   * and handles the "already in meeting from another device" dialog.
   */
  async _handlePreJoinScreen(page) {
    // Wait for pre-join UI to appear
    await page.waitForTimeout(3000);

    // Turn off microphone if it's currently on (aria-pressed="true")
    try {
      const micBtn = await page.$(
        'button[aria-pressed="true"][aria-label*="Mikrofon"], ' +
        'button[aria-pressed="true"][aria-label*="Microphone"], ' +
        'button[aria-pressed="true"][data-tid*="mic"], ' +
        'button[aria-pressed="true"][data-tid*="mute"]'
      );
      if (micBtn) await micBtn.click();
    } catch (_) {}

    // Turn off camera if it's currently on
    try {
      const camBtn = await page.$(
        'button[aria-pressed="true"][aria-label*="Kamera"], ' +
        'button[aria-pressed="true"][aria-label*="Camera"], ' +
        'button[aria-pressed="true"][aria-label*="video"], ' +
        'button[aria-pressed="true"][data-tid*="video"]'
      );
      if (camBtn) await camBtn.click();
    } catch (_) {}

    await page.waitForTimeout(500);

    // Click "Join now" / "Jetzt beitreten"
    const joinNowSelectors = [
      '[data-tid="prejoin-join-button"]',
      'button[data-tid*="prejoin" i][data-tid*="join" i]',
      'button[aria-label="Jetzt beitreten"]',
      'button[aria-label="Join now"]',
      'button:has-text("Jetzt beitreten")',
      'button:has-text("Join now")',
      'button:has-text("Beitreten")',
    ];

    for (const sel of joinNowSelectors) {
      try {
        await page.click(sel, { timeout: 4000 });
        break;
      } catch (_) {}
    }

    // Handle "already in meeting from another device" dialog
    await page.waitForTimeout(2000);
    const joinAnywaySelectors = [
      'button:has-text("Trotzdem beitreten")',
      'button:has-text("Join anyway")',
      'button[aria-label*="Join anyway"]',
      'button[aria-label*="Trotzdem"]',
    ];
    for (const sel of joinAnywaySelectors) {
      try {
        await page.click(sel, { timeout: 3000 });
        break;
      } catch (_) {}
    }

    // Give the meeting UI time to load
    await page.waitForTimeout(4000);
  }

  /**
   * Navigates to a specific Teams meeting URL and joins silently (mic + cam off).
   */
  async joinMeetingByUrl(page, meetingUrl) {
    console.log(`[SpeakerTracker] Joining meeting by URL...`);
    await page.goto(meetingUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await this._handlePreJoinScreen(page);
    console.log(`[SpeakerTracker] Joined meeting.`);
  }

  /**
   * Klickt das erste sichtbare Element, auf das einer der Selektoren passt.
   * Liefert true, wenn ein Klick ausgefuehrt wurde.
   */
  async _clickFirstVisible(page, selectors) {
    for (const sel of selectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible()) {
          await btn.click();
          return true;
        }
      } catch (_) {}
    }
    return false;
  }

  /**
   * Text-basierter Fallback: sucht klickbare Elemente (role=button/a), deren
   * aria-label oder Textinhalt eines der Regex-Muster erfuellt, und klickt das
   * erste sichtbare. Robust gegen geaenderte data-tids (kosmetische DOM-Updates).
   */
  async _clickByText(page, patterns) {
    try {
      const handle = await page.evaluateHandle((pats) => {
        const regexes = pats.map((p) => new RegExp(p, 'i'));
        const clickables = Array.from(document.querySelectorAll('button, a, [role="button"]'));
        for (const el of clickables) {
          const label = (el.getAttribute('aria-label') || '') + ' ' + ((el.innerText || el.textContent || '').split('\n')[0]);
          if (regexes.some((re) => re.test(label))) {
            const rect = el.getBoundingClientRect();
            if (rect.width > 0 && rect.height > 0) return el;
          }
        }
        return null;
      }, patterns);
      const el = handle.asElement();
      if (el) {
        await el.click().catch(async () => { await el.click({ force: true }).catch(() => null); });
        return true;
      }
    } catch (_) {}
    return false;
  }

  /**
   * Nach einem Join-Versuch: verifizieren, dass wir WIRKLICH im Meeting sind.
   */
  async _verifyInMeeting(page) {
    await page.waitForTimeout(2000);
    const info = await this.inspectMeeting(page);
    return info;
  }

  /**
   * Prueft/klickt das "Aktiver Anruf"-Banner ("Zurueck zum Anruf" — erscheint,
   * wenn derselbe Account in einem 1:1-Call auf einem anderen Geraet ist, z.B.
   * Davids Windows-Teams-Client). Liefert Join-Ergebnis oder null.
   */
  async _tryJoinViaCallBanner(page) {
    const bannerSelectors = [
      '[data-tid="call-active-bar"] button',
      '[data-tid="active-call-bar"] button',
      '[data-tid*="active-call" i] button',
      '[data-tid="call-status-bar"] button',
      'button[aria-label*="Zurück zum Anruf" i]',
      'button[aria-label*="Return to call" i]',
      'button[aria-label*="Aktiver Anruf" i]',
      '#calling-header-container button',
    ];
    let clicked = await this._clickFirstVisible(page, bannerSelectors);
    if (!clicked) {
      clicked = await this._clickByText(page, [
        'zurück zum anruf', 'return to call', 'aktiver anruf', 'rejoin'
      ]);
    }
    if (!clicked) return null;

    console.log(`[SpeakerTracker] Active call banner found, clicking to join/return...`);
    await this._handlePreJoinScreen(page);
    const info = await this._verifyInMeeting(page);
    if (info.inMeeting) {
      return { joined: true, meetingTitle: info.meetingTitle || 'Aktiver Teams-Anruf', participants: info.participants };
    }
    return null; // Banner-Klick fuehrte nicht ins Meeting -> weitersuchen
  }

  /**
   * Navigiert in die Kalender-App und versucht, das laufende/naechste Meeting
   * zu betreten. Liefert Join-Ergebnis oder null.
   */
  async _tryJoinViaCalendar(page) {
    console.log(`[SpeakerTracker] Scanning Teams calendar for current/next meeting...`);
    await page.goto('https://teams.microsoft.com', { waitUntil: 'domcontentloaded', timeout: 30000 });
    // App-Shell abwarten (App-Bar rendert erst nach dem SPA-Bootstrap)
    await page.waitForSelector('[data-tid="app-bar-chat"], [role="treeitem"], div[role="main"]', { timeout: 20000 }).catch(() => null);
    await page.waitForTimeout(3000);

    // Nach dem App-Load kann jetzt erst das "Zurueck zum Anruf"-Banner sichtbar sein
    const viaBanner = await this._tryJoinViaCallBanner(page);
    if (viaBanner) return viaBanner;

    // Kalender in der linken App-Bar oeffnen (data-tid-Muster "app-bar-<app>" ist
    // live-verifiziert fuer chat/teams — calendar folgt demselben Schema)
    const calNavSelectors = [
      '[data-tid="app-bar-calendar"]',
      'button[data-tid="app-bar-calendar"]',
      'button[aria-label*="Kalender" i]',
      'button[aria-label*="Calendar" i]',
      'a[aria-label*="Kalender" i]',
      'a[aria-label*="Calendar" i]',
      '[data-app-name="calendar"]',
      'nav [title="Kalender"]',
      'nav [title="Calendar"]',
    ];
    const clickedCal = await this._clickFirstVisible(page, calNavSelectors)
      || await this._clickByText(page, ['^kalender$', '^calendar$']);
    if (clickedCal) {
      await page.waitForTimeout(4000); // Kalender rendert Meeting-Karten langsam
    }

    // "Beitreten"-Button auf Meeting-Karten (jetzt laufende Meetings zeigen einen)
    const joinBtnSelectors = [
      'button[aria-label*="Beitreten" i]',
      'button[aria-label*="Join" i]',
      'button:has-text("Beitreten")',
      'button:has-text("Join")',
      '[data-tid*="join-btn" i]',
      '[data-tid*="joinBtn" i]',
      '[data-tid*="join-button" i]',
    ];
    for (const sel of joinBtnSelectors) {
      try {
        const candidates = await page.$$(sel);
        for (const btn of candidates) {
          if (!(await btn.isVisible().catch(() => false))) continue;
          console.log(`[SpeakerTracker] Found Join button, clicking...`);
          await btn.click();
          await this._handlePreJoinScreen(page);
          const info = await this._verifyInMeeting(page);
          if (info.inMeeting) {
            console.log(`[SpeakerTracker] Successfully joined meeting from calendar.`);
            return { joined: true, meetingTitle: info.meetingTitle, participants: info.participants };
          }
          // Klick war wirkungslos (z.B. falscher Button) -> weiteren Kandidaten probieren
        }
      } catch (_) {}
    }

    // Fallback: Meeting-Karte anklicken -> Detail-Pane enthaelt haeufig den Join-Button
    const cardSelectors = [
      '[data-tid*="calendar-event" i]',
      '[data-tid*="meeting-card" i]',
      '[data-tid*="event-card" i]',
      '[role="gridcell"] [role="button"]',
    ];
    if (await this._clickFirstVisible(page, cardSelectors)) {
      await page.waitForTimeout(2000);
      if (await this._clickFirstVisible(page, joinBtnSelectors)
        || await this._clickByText(page, ['^beitreten$', '^join$', 'jetzt beitreten', 'join now'])) {
        await this._handlePreJoinScreen(page);
        const info = await this._verifyInMeeting(page);
        if (info.inMeeting) {
          return { joined: true, meetingTitle: info.meetingTitle, participants: info.participants };
        }
      }
    }

    // Fallback Kalender-Link: extract a meetup-join link from the DOM
    try {
      const joinUrl = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll(
          'a[href*="meetup-join"], a[href*="/l/meetup"], a[href*="/meeting/"], a[href*="teams.microsoft.com/l/"]'
        ));
        return links.length > 0 ? links[0].href : null;
      });

      if (joinUrl) {
        console.log(`[SpeakerTracker] Found meeting link in DOM, navigating...`);
        await page.goto(joinUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
        await this._handlePreJoinScreen(page);
        const info = await this._verifyInMeeting(page);
        if (info.inMeeting) {
          console.log(`[SpeakerTracker] Successfully joined meeting via link.`);
          return { joined: true, meetingTitle: info.meetingTitle, participants: info.participants };
        }
      }
    } catch (_) {}

    return null;
  }

  /**
   * Fallback fuer spontane 1:1 Anrufe: den obersten ECHTEN Chat aus der
   * Chatliste ermitteln (seit dem neuen Teams-Web sind Chat-Eintraege
   * role="treeitem", aber Section-Header wie "Copilot"/"Favoriten" stehen
   * VOR den echten Chats im Baum — deshalb muss gefiltert werden, siehe
   * pickRealChatPartner). Liefert Join-/Partner-Ergebnis oder null.
   */
  async _detectChatPartner(page) {
    console.log(`[SpeakerTracker] No calendar meeting found. Checking recent chats for 1:1 call partner...`);
    try {
      // Chat-App oeffnen (data-tid app-bar-chat ist live-verifiziert, siehe
      // teamsClient.openChatView). keyboard shortcut Control+Shift+4 ist in der
      // neuen Web-UI unzuverlaessig und wird nur als Fallback behalten.
      const openedChat = await this._clickFirstVisible(page, [
        'button[data-tid="app-bar-chat"]',
        'button[aria-label*="Chat" i]',
        'a[aria-label*="Chat" i]',
      ]);
      if (!openedChat) {
        await page.keyboard.press('Control+Shift+4').catch(() => {});
      }
      await page.waitForTimeout(2000);

      // Alle Chat-Kandidaten in DOM-Reihenfolge sammeln. Bewusst MEHRERE
      // Selektor-Strategien, damit kosmetische DOM-Aenderungen abgefedert sind.
      const candidates = await page.evaluate(() => {
        const sel = [
          '[role="treeitem"]',
          '[data-tid="chat-list-item"]',
          '[data-tid*="chat-list-entry"]',
          '[role="listitem"] [data-tid*="chat-list-entry"]',
        ].join(', ');
        const rows = Array.from(document.querySelectorAll(sel));
        const out = [];
        for (const row of rows) {
          const nameEl = row.querySelector('[data-tid*="title" i], [class*="title" i], span[dir="auto"], h3');
          const title = (nameEl ? (nameEl.innerText || nameEl.textContent || '') : (row.innerText || '')).split('\n')[0].trim();
          if (title) out.push({ title });
          if (out.length >= 40) break;
        }
        return out;
      });

      const partner = pickRealChatPartner(candidates);
      if (partner) {
        console.log(`[SpeakerTracker] Detected active/recent chat partner: '${partner}'`);
        return { joined: false, meetingTitle: `1:1 Call ${partner}`, knownSpeakers: [partner] };
      }

      console.warn(`[SpeakerTracker] Chatliste enthaelt keinen echten Partner (nur Bots/Section-Header).`);
    } catch (e) {
      console.warn(`[SpeakerTracker] Could not inspect chat list: ${e.message}`);
    }
    return null;
  }

  /**
   * Scans Teams for a currently running meeting, spontaneous 1:1 call,
   * or upcoming calendar meeting and joins it silently (mic + cam off).
   */
  async findAndJoinNextMeeting(page) {
    console.log(`[SpeakerTracker] Checking Teams for active call or meeting...`);

    // 1. Bereits in einem Call/Meeting?
    try {
      const currentInfo = await this.inspectMeeting(page);
      if (currentInfo && currentInfo.inMeeting) {
        console.log(`[SpeakerTracker] Already inside active meeting: ${currentInfo.meetingTitle}`);
        return { joined: true, meetingTitle: currentInfo.meetingTitle, participants: currentInfo.participants };
      }
    } catch (_) {}

    // 2. Aktives Call-Banner ("Zurueck zum Anruf" — 1:1-Call laeuft auf anderem Geraet)
    const viaBanner = await this._tryJoinViaCallBanner(page);
    if (viaBanner) return viaBanner;

    // 3. Kalender (geplante Termine) — inkl. Banner-Recheck nach App-Load
    const viaCalendar = await this._tryJoinViaCalendar(page);
    if (viaCalendar) return viaCalendar;

    // 4. Fallback spontane 1:1 Calls: obersten echten Chat ermitteln
    const viaChat = await this._detectChatPartner(page);
    if (viaChat) return viaChat;

    console.warn(`[SpeakerTracker] No joinable meeting or call found. Continuing without auto-join.`);
    return { joined: false, meetingTitle: 'Teams Meeting', knownSpeakers: [] };
  }

  /**
   * Injiziert/aktualisiert den 100ms-Sampler im Seitenkontext. Bewusst
   * idempotent: bestehende Events bleiben erhalten (kein Reset), ein alter
   * Interval wird ersetzt. Nach Navigationen (z.B. Auto-Join-Retry) einfach
   * erneut aufrufen.
   */
  async _injectSampler(page) {
    try {
      await page.evaluate(({ src, helperSrc }) => {
        const collect = new Function('return (' + src + ')')();
        const helpers = new Function('return (' + helperSrc + ')')();
        window.__teamsSampleSpeakers = () => collect(false, helpers).activeSpeakers;
        window.__teamsSpeakerEvents = window.__teamsSpeakerEvents || [];
        if (window.__teamsSpeakerInterval) {
          clearInterval(window.__teamsSpeakerInterval);
        }
        window.__teamsSpeakerInterval = setInterval(() => {
          try {
            const names = window.__teamsSampleSpeakers() || [];
            if (names.length) {
              const now = Date.now();
              for (const name of names) {
                window.__teamsSpeakerEvents.push({ speaker: name, timestamp: now });
              }
            }
          } catch (e) {}
        }, 100);
      }, {
        src: collectMeetingStateInPage.toString(),
        helperSrc: materializeInPageHelpersSrc()
      });
      return true;
    } catch (_) {
      return false;
    }
  }

  /**
   * Starts tracking active speakers in the given tenant's Teams tab.
   */
  async startTracking(page, tenant, outputPath = null, options = {}) {
    if (this.activeSessions.has(tenant)) {
      await this.stopTracking(tenant);
    }

    // Auto-join: navigate to meeting URL, find call/meeting, or detect 1:1 chat partner
    const { meetingUrl, autoJoin = true, noJoin = false, speaker = null } = options;
    let knownSpeakers = speaker ? [cleanSpeakerName(speaker)] : [];
    let detectedMeetingTitle = '';
    let joinAttempts = 0;

    if (!noJoin) {
      if (meetingUrl) {
        await this.joinMeetingByUrl(page, meetingUrl);
        joinAttempts++;
      } else if (autoJoin) {
        const joinRes = await this.findAndJoinNextMeeting(page);
        joinAttempts++;
        if (joinRes) {
          if (joinRes.meetingTitle) detectedMeetingTitle = joinRes.meetingTitle;
          if (joinRes.knownSpeakers && joinRes.knownSpeakers.length > 0) {
            knownSpeakers = Array.from(new Set([...knownSpeakers, ...joinRes.knownSpeakers]));
          }
        }
      }
    }

    const meetingInfo = await this.inspectMeeting(page);
    const startTimeEpochMs = Date.now();
    const samples = [];

    // In-page sampler injizieren (100ms)
    await this._injectSampler(page);

    // Node polling loop to retrieve samples periodically from the browser context
    const pollInterval = setInterval(async () => {
      try {
        if (page.isClosed()) {
          clearInterval(pollInterval);
          return;
        }
        const retrieved = await page.evaluate(() => {
          const evts = window.__teamsSpeakerEvents || [];
          window.__teamsSpeakerEvents = [];
          return evts;
        }).catch(() => []);

        if (retrieved && retrieved.length > 0) {
          samples.push(...retrieved);
        }
      } catch (e) {}
    }, 500);

    // Auto-Join-Retry: Der Tracker startet ZUSAMMEN mit der Aufnahme — der echte
    // Anruf/Meeting-Join von David kann SEKUNDEN bis MINUTEN spaeter beginnen.
    // Ohne Retry bliebe der Tracker fuerever ausserhalb des Meetings (exakt das
    // Ausfallbild der leeren Timelines). Retry nur solange keine Samples vorliegen.
    let lastInMeeting = meetingInfo.inMeeting;
    let lastJoinCheck = startTimeEpochMs;
    const retryJoinInterval = setInterval(async () => {
      try {
        const session = this.activeSessions.get(tenant);
        if (!session || page.isClosed()) {
          clearInterval(retryJoinInterval);
          return;
        }
        if (samples.length > 0 || joinAttempts >= 12) {
          clearInterval(retryJoinInterval);
          return;
        }
        const full = await this.inspectMeeting(page);
        lastInMeeting = full.inMeeting;
        lastJoinCheck = Date.now();
        if (full.inMeeting) return; // drin, aber noch keine aktiven Sprecher -> abwarten
        joinAttempts++;
        console.log(`[SpeakerTracker] Auto-Join-Retry #${joinAttempts} (keine Samples, nicht im Meeting)...`);
        const joinRes = await this.findAndJoinNextMeeting(page);
        if (joinRes) {
          if (joinRes.meetingTitle && (!session.meetingTitle || session.meetingTitle === 'Teams Meeting' || session.meetingTitle.startsWith('1:1 Call (unbekannter'))) {
            session.meetingTitle = joinRes.meetingTitle;
          }
          if (joinRes.knownSpeakers && joinRes.knownSpeakers.length > 0) {
            session.knownSpeakers = Array.from(new Set([...(session.knownSpeakers || []), ...joinRes.knownSpeakers]));
          }
        }
        // Navigation beim Join-Versuch kann den In-Page-Sampler gekillt haben -> neu injizieren
        await this._injectSampler(page);
      } catch (e) {}
    }, 45000);

    const session = {
      tenant,
      page,
      startTimeEpochMs,
      outputPath,
      pollInterval,
      retryJoinInterval,
      samples,
      meetingTitle: detectedMeetingTitle || meetingInfo.meetingTitle || 'Teams Meeting',
      knownSpeakers
    };

    this.activeSessions.set(tenant, session);

    return {
      status: 'tracking_started',
      tenant,
      startTimeEpochMs,
      meetingTitle: session.meetingTitle,
      inMeeting: meetingInfo.inMeeting,
      participants: meetingInfo.participants,
      knownSpeakers
    };
  }

  /**
   * Stops tracking and writes the resulting timeline JSON.
   */
  async stopTracking(tenant) {
    const session = this.activeSessions.get(tenant);
    if (!session) {
      return { status: 'no_active_session', tenant };
    }

    if (session.pollInterval) {
      clearInterval(session.pollInterval);
    }
    if (session.retryJoinInterval) {
      clearInterval(session.retryJoinInterval);
    }

    const { page, startTimeEpochMs, outputPath, samples, meetingTitle } = session;
    this.activeSessions.delete(tenant);

    // Pull remaining events from browser
    try {
      if (page && !page.isClosed()) {
        const remaining = await page.evaluate(() => {
          if (window.__teamsSpeakerInterval) {
            clearInterval(window.__teamsSpeakerInterval);
            delete window.__teamsSpeakerInterval;
          }
          const evts = window.__teamsSpeakerEvents || [];
          window.__teamsSpeakerEvents = [];
          return evts;
        }).catch(() => []);
        if (remaining && remaining.length > 0) {
          samples.push(...remaining);
        }
      }
    } catch (e) {}

    const endTimeEpochMs = Date.now();
    const durationSec = Math.max(0, (endTimeEpochMs - startTimeEpochMs) / 1000);

    // Aggregate samples into clean intervals (aggregateEvents filtert Bots bereits raus)
    const intervals = aggregateEvents(samples, {
      minGapMs: 400,
      minDurationMs: 250,
      meetingStartEpochMs: startTimeEpochMs
    });

    const uniqueSpeakers = Array.from(new Set(intervals.map(i => i.speaker)));
    let finalSpeakers = uniqueSpeakers;
    if (finalSpeakers.length === 0 && session.knownSpeakers && session.knownSpeakers.length > 0) {
      // Nur echte (nicht-Bot) bekannte Partner als Fallback übernehmen, wenn keine
      // DOM-Intervalle erkannt wurden. Bot/UI-Namen (z.B. "Copilot") würden sonst
      // die gesamte Gegenseite fälschlich auf einen Pseudo-Namen mappen.
      finalSpeakers = session.knownSpeakers.filter((s) => !isBotOrUiName(s));
    }

    const result = {
      version: '1.0',
      meeting_title: meetingTitle,
      tenant,
      start_time_epoch_ms: startTimeEpochMs,
      end_time_epoch_ms: endTimeEpochMs,
      duration_sec: durationSec,
      speaker_count: finalSpeakers.length,
      speakers: finalSpeakers,
      intervals
    };

    if (outputPath) {
      try {
        const dir = path.dirname(outputPath);
        if (!fs.existsSync(dir)) {
          fs.mkdirSync(dir, { recursive: true });
        }
        fs.writeFileSync(outputPath, JSON.stringify(result, null, 2), 'utf-8');
      } catch (err) {
        result.write_error = err.message;
      }
    }

    return {
      status: 'tracking_stopped',
      tenant,
      meetingTitle,
      durationSec,
      speakerCount: uniqueSpeakers.length,
      speakers: uniqueSpeakers,
      intervalsCount: intervals.length,
      outputPath: outputPath || null,
      data: result
    };
  }

  isTracking(tenant) {
    return this.activeSessions.has(tenant);
  }
}

export const speakerTracker = new SpeakerTracker();
