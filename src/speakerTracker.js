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

export class SpeakerTracker {
  constructor() {
    this.activeSessions = new Map(); // tenant -> { startTimeEpochMs, outputPath, timer, samples: [], meetingTitle: '' }
  }

  /**
   * Scans current page DOM to detect if Teams is inside an active meeting or call.
   */
  async inspectMeeting(page) {
    if (!page || page.isClosed()) {
      return { inMeeting: false, meetingTitle: '', participants: [], activeSpeakers: [] };
    }

    return await page.evaluate(() => {
      // Common indicators for active call in Teams v2 web client
      const callControls = document.querySelector(
        '[data-tid="call-controls"], [data-tid="hangup-button"], [data-tid="calling-screen"], div[id*="calling-stage"], button[aria-label*="Auflegen"], button[aria-label*="Leave"], button[aria-label*="Verlassen"]'
      );

      const inMeeting = !!callControls;
      let meetingTitle = '';

      const titleEl = document.querySelector(
        '[data-tid="calling-header-title"], [data-tid="call-title"], [data-tid="header-meeting-title"], h1[data-tid*="meeting"]'
      );
      if (titleEl) {
        meetingTitle = (titleEl.innerText || titleEl.textContent || '').trim();
      }
      if (!meetingTitle) {
        const docTitle = document.title || '';
        const match = docTitle.match(/^(.+?)\\s*\\|\\s*Microsoft Teams/i);
        if (match && !match[1].toLowerCase().includes('chat')) {
          meetingTitle = match[1].trim();
        }
      }

      // Collect participants visible in stage/roster
      const participantEls = document.querySelectorAll(
        '[data-tid="roster-avatar-name"], [data-tid="calling-participant-stream"] [aria-label], [data-tid="video-tile"] [aria-label]'
      );
      const participantSet = new Set();
      for (const el of participantEls) {
        const name = (el.innerText || el.getAttribute('aria-label') || '').trim();
        if (name && name.length < 50 && !name.includes('\\n')) {
          participantSet.add(name);
        }
      }

      // Detect active speakers
      const activeSpeakerEls = document.querySelectorAll(
        '[data-tid="active-speaker-border"], [data-is-speaking="true"], [data-speaker-state="speaking"], .speaking, [class*="activeSpeaker"], [class*="active-speaker"], [aria-label*="spricht"], [aria-label*="speaking"]'
      );

      const activeSpeakers = new Set();
      for (const el of activeSpeakerEls) {
        let speakerName = '';
        const nameNode = el.querySelector('[data-tid="roster-avatar-name"], [data-tid="video-tile-name"], [class*="name"]');
        if (nameNode) {
          speakerName = nameNode.innerText || nameNode.textContent || '';
        }
        if (!speakerName) {
          const aria = el.getAttribute('aria-label') || '';
          const m = aria.match(/^(.+?)(?:\\s+(?:spricht|is speaking)|\\s*\\(.*\\))/i);
          if (m) speakerName = m[1];
        }
        if (!speakerName) {
          const parent = el.closest('[data-tid*="participant"], [data-tid*="tile"]');
          if (parent) {
            const pName = parent.querySelector('[data-tid="roster-avatar-name"], [data-tid="video-tile-name"]');
            if (pName) speakerName = pName.innerText || pName.textContent || '';
          }
        }
        if (speakerName) {
          activeSpeakers.add(speakerName.trim());
        }
      }

      return {
        inMeeting,
        meetingTitle: meetingTitle || 'Teams Meeting',
        participants: Array.from(participantSet),
        activeSpeakers: Array.from(activeSpeakers)
      };
    }).catch(() => ({
      inMeeting: false,
      meetingTitle: '',
      participants: [],
      activeSpeakers: []
    }));
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

    // 2. Prüfen auf aktives Call-Banner ("Zurück zum Anruf" / "Return to call" / "Beitreten")
    const activeCallBannerSelectors = [
      '[data-tid="call-active-bar"] button',
      'button[aria-label*="Zurück zum Anruf" i]',
      'button[aria-label*="Return to call" i]',
      'button:has-text("Zurück zum Anruf")',
      'button:has-text("Return to call")',
      'button[aria-label*="Aktiver Anruf" i]',
      '[data-tid*="active-call"] button',
      '#calling-header-container button',
    ];

    for (const sel of activeCallBannerSelectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible()) {
          console.log(`[SpeakerTracker] Found active call banner, clicking to join/return...`);
          await btn.click();
          await this._handlePreJoinScreen(page);
          const info = await this.inspectMeeting(page);
          return { joined: true, meetingTitle: info.meetingTitle || 'Aktiver Teams-Anruf', participants: info.participants };
        }
      } catch (_) {}
    }

    // 3. Kalender prüfen (für geplante Termine)
    console.log(`[SpeakerTracker] Scanning Teams calendar for current/next meeting...`);
    await page.goto('https://teams.microsoft.com', { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2500);

    // Try clicking the Calendar icon in the left nav
    const calNavSelectors = [
      '[data-tid="app-bar-calendar"]',
      'button[aria-label="Kalender"]',
      'button[aria-label="Calendar"]',
      'a[aria-label="Kalender"]',
      'a[aria-label="Calendar"]',
      '[data-app-name="calendar"]',
      'nav [title="Kalender"]',
      'nav [title="Calendar"]',
    ];
    for (const sel of calNavSelectors) {
      try {
        await page.click(sel, { timeout: 3000 });
        await page.waitForTimeout(2000);
        break;
      } catch (_) {}
    }

    // Look for a "Join" button on today's meeting cards (meetings happening now show one)
    const joinBtnSelectors = [
      'button[aria-label*="Beitreten"]',
      'button[aria-label*="Join"]',
      'button:has-text("Beitreten")',
      'button:has-text("Join")',
      '[data-tid*="join-btn"]',
      '[data-tid*="joinBtn"]',
    ];

    for (const sel of joinBtnSelectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible()) {
          console.log(`[SpeakerTracker] Found Join button, clicking...`);
          await btn.click();
          await this._handlePreJoinScreen(page);
          console.log(`[SpeakerTracker] Successfully joined meeting from calendar.`);
          const info = await this.inspectMeeting(page);
          return { joined: true, meetingTitle: info.meetingTitle, participants: info.participants };
        }
      } catch (_) {}
    }

    // Fallback Kalender-Link: extract a meetup-join link from the DOM
    try {
      const joinUrl = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a[href*="meetup-join"], a[href*="teams.microsoft.com/l/"]'));
        return links.length > 0 ? links[0].href : null;
      });

      if (joinUrl) {
        console.log(`[SpeakerTracker] Found meeting link in DOM, navigating...`);
        await page.goto(joinUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
        await this._handlePreJoinScreen(page);
        console.log(`[SpeakerTracker] Successfully joined meeting via link.`);
        const info = await this.inspectMeeting(page);
        return { joined: true, meetingTitle: info.meetingTitle, participants: info.participants };
      }
    } catch (_) {}

    // 4. Fallback für spontane 1:1 Anrufe: Letzten aktiven Chat prüfen
    console.log(`[SpeakerTracker] No calendar meeting found. Checking recent chats for 1:1 call partner...`);
    try {
      await page.keyboard.press('Control+Shift+4'); // Chat-App
      await page.waitForTimeout(2000);

      // Ermittle den obersten Chat in der Liste
      const topChatInfo = await page.evaluate(() => {
        const chatRows = document.querySelectorAll('[role="listitem"] [data-tid*="chat-list-entry"], [role="treeitem"], [data-tid="chat-list-item"]');
        if (chatRows.length > 0) {
          const first = chatRows[0];
          const nameEl = first.querySelector('[data-tid*="title"], [class*="title"], span[dir="auto"], h3');
          const title = (nameEl ? (nameEl.innerText || nameEl.textContent) : first.innerText || '').split('\n')[0].trim();
          return { title };
        }
        return null;
      });

      if (topChatInfo && topChatInfo.title) {
        const cleanName = cleanSpeakerName(topChatInfo.title);
        console.log(`[SpeakerTracker] Detected active/recent chat partner: '${cleanName}'`);
        return { joined: false, meetingTitle: `1:1 Call ${cleanName}`, knownSpeakers: [cleanName] };
      }
    } catch (e) {
      console.warn(`[SpeakerTracker] Could not inspect chat list: ${e.message}`);
    }

    console.warn(`[SpeakerTracker] No joinable meeting or call found. Continuing without auto-join.`);
    return { joined: false, meetingTitle: 'Teams Meeting', knownSpeakers: [] };
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

    if (!noJoin) {
      if (meetingUrl) {
        await this.joinMeetingByUrl(page, meetingUrl);
      } else if (autoJoin) {
        const joinRes = await this.findAndJoinNextMeeting(page);
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

    // Inject in-page sampler
    await page.evaluate(() => {
      window.__teamsSpeakerEvents = [];
      if (window.__teamsSpeakerInterval) {
        clearInterval(window.__teamsSpeakerInterval);
      }

      window.__teamsSpeakerInterval = setInterval(() => {
        try {
          const activeEls = document.querySelectorAll(
            '[data-tid="active-speaker-border"], [data-is-speaking="true"], [data-speaker-state="speaking"], .speaking, [class*="activeSpeaker"], [class*="active-speaker"], [aria-label*="spricht"], [aria-label*="speaking"]'
          );
          const now = Date.now();
          for (const el of activeEls) {
            let name = '';
            const nameNode = el.querySelector('[data-tid="roster-avatar-name"], [data-tid="video-tile-name"], [class*="name"]');
            if (nameNode) name = nameNode.innerText || nameNode.textContent || '';
            if (!name) {
              const aria = el.getAttribute('aria-label') || '';
              const m = aria.match(/^(.+?)(?:\\s+(?:spricht|is speaking)|\\s*\\(.*\\))/i);
              if (m) name = m[1];
            }
            if (!name) {
              const parent = el.closest('[data-tid*="participant"], [data-tid*="tile"]');
              if (parent) {
                const pName = parent.querySelector('[data-tid="roster-avatar-name"], [data-tid="video-tile-name"]');
                if (pName) name = pName.innerText || pName.textContent || '';
              }
            }
            if (name) {
              window.__teamsSpeakerEvents.push({
                speaker: name.trim(),
                timestamp: now
              });
            }
          }
        } catch (e) {}
      }, 100);
    }).catch(() => null);

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

    const session = {
      tenant,
      page,
      startTimeEpochMs,
      outputPath,
      pollInterval,
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

    // Aggregate samples into clean intervals
    const intervals = aggregateEvents(samples, {
      minGapMs: 400,
      minDurationMs: 250,
      meetingStartEpochMs: startTimeEpochMs
    });

    const uniqueSpeakers = Array.from(new Set(intervals.map(i => i.speaker)));
    let finalSpeakers = uniqueSpeakers;
    if (finalSpeakers.length === 0 && session.knownSpeakers && session.knownSpeakers.length > 0) {
      finalSpeakers = session.knownSpeakers;
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
