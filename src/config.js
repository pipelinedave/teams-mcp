import path from 'path';
import os from 'os';
import fs from 'fs';

const HOME = os.homedir();

// Chrome/Edge-Executable automatisch erkennen (WSL/Linux, macOS, Windows).
// Über per-Env TEAMS_MCP_CHROME_PATH überschreibbar.
function detectChromePath() {
  if (process.env.TEAMS_MCP_CHROME_PATH) return process.env.TEAMS_MCP_CHROME_PATH;

  const candidates = [
    // Playwright-Cache (Linux, gängige Versionen)
    ...expandPlaywrightCache(),
    // System-Chrome / Chromium (Linux)
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
    // macOS
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
  ];

  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }
  return null;
}

function expandPlaywrightCache() {
  const base = path.join(HOME, '.cache/ms-playwright');
  const results = [];
  try {
    if (fs.existsSync(base)) {
      const entries = fs.readdirSync(base);
      for (const e of entries) {
        const p = path.join(base, e, 'chrome-linux64', 'chrome');
        if (fs.existsSync(p)) results.push(p);
      }
    }
  } catch (e) {}
  return results;
}

// Basis-Verzeichnis für die Browser-Profile (per Env überschreibbar)
function getProfileBase() {
  return process.env.TEAMS_MCP_PROFILE_BASE || HOME;
}

// Realm-Zuordnung je Tenant. Wenn nicht gesetzt (öffentliches Repo), werden Tenant
// und Realm als identisch behandelt (tenant === realm) – komplett org-neutral.
function getTenantRealms() {
  if (process.env.TEAMS_MCP_TENANT_REALMS) {
    try {
      return JSON.parse(process.env.TEAMS_MCP_TENANT_REALMS);
    } catch (e) {
      throw new Error('TEAMS_MCP_TENANT_REALMS muss ein gültiges JSON-Objekt sein, z.B. {"myorg":"myorg.onmicrosoft.com"}');
    }
  }
  return {};
}

// Optionale Whitelist erlaubter Tenant-Keys (für validated Enum in den Tool-Schemas).
function getTenants() {
  if (process.env.TEAMS_MCP_TENANTS) {
    return process.env.TEAMS_MCP_TENANTS.split(',').map(s => s.trim()).filter(Boolean);
  }
  return Object.keys(getTenantRealms());
}

// Standard-Tenant Fallback (per Env TEAMS_MCP_DEFAULT_TENANT überschreibbar, sonst erster konfigurierter Tenant, sonst "adesso").
function getDefaultTenant() {
  if (process.env.TEAMS_MCP_DEFAULT_TENANT) {
    return process.env.TEAMS_MCP_DEFAULT_TENANT.trim();
  }
  const tenants = getTenants();
  if (tenants.length > 0) {
    return tenants[0];
  }
  return 'adesso';
}

// Eigener Anzeigename für "Ich"-Nachrichten (Fallback im Message-Pane).
function getSelfName() {
  return process.env.TEAMS_MCP_SELF_NAME || 'Ich';
}

// Headless-Default (nur einmalig anwendbar; Login ist immer sichtbar)
function getDefaultHeadless() {
  const v = process.env.TEAMS_MCP_HEADLESS;
  if (v === undefined) return true;
  return String(v).toLowerCase() !== 'false' && v !== '0';
}

// Aktivität-Scheduler: geplanter proaktiver Activity-Bericht (Standard 09:00 + 17:00).
// Per TEAMS_MCP_ACTIVITY_SCHEDULER=0/false deaktivierbar; Times per
// TEAMS_MCP_ACTIVITY_TIMES (Komma-getrennt HH:MM) überschreibbar.
function getActivitySchedulerEnabled() {
  const v = process.env.TEAMS_MCP_ACTIVITY_SCHEDULER;
  if (v === undefined) return true;
  return String(v).toLowerCase() !== 'false' && v !== '0';
}

function getActivityTimes() {
  if (!process.env.TEAMS_MCP_ACTIVITY_TIMES) return null;
  return process.env.TEAMS_MCP_ACTIVITY_TIMES.split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

const detectedChromePath = detectChromePath();
if (!detectedChromePath) {
  console.warn('[teams-mcp] Kein Chrome/Chromium/Edge automatisch gefunden. Setze TEAMS_MCP_CHROME_PATH, falls Playwright keinen eigenen Browser mitbringt.');
}

export const config = {
  get chromePath() {
    return process.env.TEAMS_MCP_CHROME_PATH || detectedChromePath || undefined;
  },
  profileBase: getProfileBase(),
  tenantRealms: getTenantRealms(),
  tenants: getTenants(),
  defaultTenant: getDefaultTenant(),
  selfName: getSelfName(),
  defaultHeadless: getDefaultHeadless(),
  activitySchedulerEnabled: getActivitySchedulerEnabled(),
  activityTimes: getActivityTimes()
};
