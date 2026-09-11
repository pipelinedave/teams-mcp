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
   * Starts tracking active speakers in the given tenant's Teams tab.
   */
  async startTracking(page, tenant, outputPath = null) {
    if (this.activeSessions.has(tenant)) {
      await this.stopTracking(tenant);
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
      meetingTitle: meetingInfo.meetingTitle || 'Teams Meeting'
    };

    this.activeSessions.set(tenant, session);

    return {
      status: 'tracking_started',
      tenant,
      startTimeEpochMs,
      meetingTitle: session.meetingTitle,
      inMeeting: meetingInfo.inMeeting,
      participants: meetingInfo.participants
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

    const result = {
      version: '1.0',
      meeting_title: meetingTitle,
      tenant,
      start_time_epoch_ms: startTimeEpochMs,
      end_time_epoch_ms: endTimeEpochMs,
      duration_sec: durationSec,
      speaker_count: uniqueSpeakers.length,
      speakers: uniqueSpeakers,
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
