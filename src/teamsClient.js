import fs from 'fs';
import path from 'path';
import TurndownService from 'turndown';
import { browserManager } from './browserManager.js';
import { config } from './config.js';
import { speakerTracker } from './speakerTracker.js';
import { activityClient } from './activityClient.js';
import { analyzeItems, buildReportText, writeReport, CATEGORIES } from './activityAnalyzer.js';

const turndown = new TurndownService({
  headingStyle: 'atx',
  hr: '---',
  bulletListMarker: '-'
});
turndown.remove(['script', 'style', 'head']);

export class TeamsClient {
  async getPage(tenant = '', headless = true) {
    const t = browserManager.normalizeTenant(tenant);
    if (!t) {
      throw new Error("Es wurde kein 'tenant' angegeben. Bitte den Tenant-Namen bzw. Realm mitgeben (z.B. teams_list_chats({ tenant: 'meine-org.onmicrosoft.com' })).");
    }
    const { page } = await browserManager.ensureContext(t, headless);
    const url = page.url();

    if (!url.includes('teams.microsoft.com') && !url.includes('login.microsoftonline.com') && !url.includes('teams.cloud.microsoft')) {
      await page.goto(`https://teams.microsoft.com/v2/?realm=${encodeURIComponent(browserManager.realm(t))}`, {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });
      // Warten, bis Teams geladen ist ODER die Session auf die Login-Seite abgelaufen ist
      await Promise.race([
        page.waitForSelector('[data-tid="app-bar-chat"], [role="treeitem"], div[role="main"]', { timeout: 20000 }).catch(() => null),
        page.waitForURL(/login\.microsoftonline\.com/, { timeout: 20000 }).catch(() => null)
      ]).catch(() => null);
    }

    const currentUrl = page.url();
    if (currentUrl.includes('login.microsoftonline.com')) {
      throw new Error(`Nicht in Teams (${t}) angemeldet. Bitte führe teams_login({ tenant: "${t}" }) in einem sichtbaren Fenster aus.`);
    }
    return page;
  }

  async checkStatus(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    try {
      const page = await this.getPage(t, true);
      const url = page.url();

      if (url.includes('login.microsoftonline.com')) {
        return {
          authenticated: false,
          tenant: t,
          currentUrl: url,
          message: `Nicht in Teams (${t}) angemeldet. Bitte führe 'teams_login({ tenant: "${t}" })' aus.`
        };
      }

      await page.waitForSelector('#main, div[role="main"], [role="treeitem"], [data-tid="app-bar-chat"]', { timeout: 12000 }).catch(() => null);

      const title = await page.title();
      const isLoaded = title.toLowerCase().includes('teams') || page.url().includes('teams.microsoft.com') || page.url().includes('teams.cloud.microsoft');

      return {
        authenticated: isLoaded && !url.includes('login.microsoftonline.com'),
        tenant: t,
        currentUrl: page.url(),
        title: title,
        message: isLoaded && !url.includes('login.microsoftonline.com')
          ? `Microsoft Teams (${t}) ist angemeldet und einsatzbereit.`
          : `Seite lädt noch oder erfordert Anmeldung ('teams_login({ tenant: "${t}" })').`
      };
    } catch (err) {
      return {
        authenticated: false,
        tenant: t,
        error: err.message
      };
    }
  }

  async openChatView(page) {
    const chatBtn = await page.$('button[data-tid="app-bar-chat"], button[aria-label*="Chat"], a[aria-label*="Chat"]');
    if (chatBtn) {
      await chatBtn.click().catch(() => null);
      await page.waitForTimeout(2000);
    }
  }

  // Liefert die gefilterte Liste der Chat-Treeitems (gleiche Logik wie listChats),
  // damit chat_index- und chat_name-Lookups konsistent auf denselben Items arbeiten.
  async getFilteredChatRows(page) {
    await this.openChatView(page);
    await page.waitForSelector('[role="treeitem"]', { timeout: 15000 }).catch(() => null);
    const rows = await page.$$('[role="treeitem"]');

    const filtered = [];
    const ignoredTitles = new Set(['copilot', 'quick views', 'drafts', 'favorites', 'chats', 'teams', 'channels', 'unread', 'more options', 'mentions', 'followed threads', 'discover', 'see more', 'see all channels', 'see all your teams', 'communities', 'join communities']);

    for (const row of rows) {
      const text = (await row.innerText())?.trim() || '';
      if (!text) continue;

      const lines = text.split('\n').map(s => s.trim()).filter(Boolean);
      const firstLine = lines[0] || '';
      if (ignoredTitles.has(firstLine.toLowerCase())) continue;

      if (lines.length === 1 && (firstLine.includes('Favorites') || firstLine.includes('Chats') || firstLine.includes('Quick views') || firstLine.includes('d.velop') || firstLine.startsWith('See'))) {
        continue;
      }

      // Kanal-/Gruppen-Sektionen ohne konkreten Chat (z.B. "General", "See all channels")
      if (firstLine.toLowerCase() === 'general' || firstLine.toLowerCase().startsWith('see ')) continue;

      filtered.push({ row, title: firstLine, text });
    }

    return filtered;
  }

  // Wählt einen Chat aus der gefilterten Liste. Bevorzugt exakten Titel-Match,
  // fällt auf Präfix-/Wort-Start-Match zurück (statt blindem Substring 'includes',
  // der z.B. "Yannick" fälschlich auf "Yannick und Kollegen" mappen kann).
  pickChatRow(filtered, chatName, chatIndex) {
    if (typeof chatIndex === 'number' && chatIndex >= 0 && chatIndex < filtered.length) {
      return { row: filtered[chatIndex].row, title: filtered[chatIndex].title };
    }
    if (!chatName) return null;
    const term = chatName.toLowerCase().trim();
    if (!term) return null;

    // 1) exakter Titel-Match
    let exact = filtered.find(c => c.title.toLowerCase() === term);
    if (exact) return { row: exact.row, title: exact.title };

    // 2) Titel beginnt mit dem Suchbegriff (z.B. "Yannick" -> "Yannick Bülter")
    let prefix = filtered.find(c => c.title.toLowerCase().startsWith(term));
    if (prefix) return { row: prefix.row, title: prefix.title };

    // 3) Fallback: Token-Match (alle Wörter des Begriffs kommen im Titel vor)
    const tokens = term.split(/\s+/).filter(Boolean);
    if (tokens.length > 0) {
      const tokenMatch = filtered.find(c => {
        const t = c.title.toLowerCase();
        return tokens.every(tok => t.includes(tok));
      });
      if (tokenMatch) return { row: tokenMatch.row, title: tokenMatch.title };
    }

    return null;
  }

  async listChats(tenant = '', limit = 15) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    const chats = filtered.slice(0, limit).map((chat, idx) => ({
      index: idx,
      title: chat.title,
      preview: chat.text.split('\n').map(s => s.trim()).filter(Boolean).slice(1).join(' | ') || undefined
    }));

    return {
      tenant: t,
      count: chats.length,
      chats: chats
    };
  }

  // Liefert den Titel des aktuell geöffneten Chats. Bevorzugt den Chat-Pane-Header
  // aus dem DOM (stabil, z.B. [data-tid="chat-title"]), fällt auf page.title() zurück.
  async getActiveChatTitle(page) {
    try {
      const domTitle = await page.evaluate(() => {
        // Chat-Header im Message-Pane (Teams/Fluent): [data-tid="chat-title"] ist am stabilsten
        const sels = [
          '[data-tid="chat-title"]',
          '[data-tid="chat-header-title"]',
          '[data-tid="entity-header"] [role="heading"]',
          '[data-tid="thread-pane-header"]',
          'header [aria-level]'
        ];
        for (const sel of sels) {
          const el = document.querySelector(sel);
          if (el && el.innerText && el.innerText.trim() && el.innerText.trim().length < 60) {
            const t = el.innerText.trim();
            if (t && !/^Message List$/i.test(t) && !/^Chat$/i.test(t)) {
              return t;
            }
          }
        }
        return null;
      });
      if (domTitle) return domTitle;
    } catch (e) {}

    // Fallback: page.title() -> "Chat | David Hallmann (You) | Microsoft Teams"
    const raw = (await page.title()).replace(' | Microsoft Teams', '').replace(/^\(\d+\)\s*/, '');
    const parts = raw.split(' | ').map(s => s.trim()).filter(Boolean);
    // Zweites Segment ist meist der Chat-Name (erste = "Chat"), sonst das letzte sichtbare
    if (parts.length >= 2 && /^chat$/i.test(parts[0])) return parts[1];
    return parts[0] || raw;
  }

  async getMessages(tenant = '', { chatIndex, chatName, limit = 20 } = {}) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    if (filtered.length === 0) {
      throw new Error("Keine Chats in Teams gefunden.");
    }

    if (typeof chatIndex === 'number' || chatName) {
      const picked = this.pickChatRow(filtered, chatName, chatIndex);
      if (picked) {
        await picked.row.click();
        await page.waitForTimeout(2500);
      }
    }

    // Aktiven Chat-Titel aus dem Message-Pane verifizieren (stabil)
    const activeChat = await this.getActiveChatTitle(page);
    await page.waitForSelector('[data-tid="chat-pane-message"], .fui-ChatMessage__body, .fui-ChatMyMessage__body', { timeout: 15000 }).catch(() => null);

    const messagesData = await page.evaluate(({ max, selfName, helperStr }) => {
      const toIso = new Function('return (' + helperStr + ')')();
      const msgs = [];
      const msgBodies = Array.from(document.querySelectorAll('[data-tid="chat-pane-message"], .fui-ChatMessage__body, .fui-ChatMyMessage__body'));
      const slice = msgBodies.slice(-max);

      for (const body of slice) {
        let author = '';
        let parent = body.closest('[role="group"], [role="listitem"], .fui-ChatMessage, .fui-ChatMyMessage') || body.parentElement;
        if (parent) {
          const authorEl = parent.querySelector('[data-tid="message-author-name"], span[role="heading"]');
          if (authorEl) author = authorEl.innerText.trim();
        }

        if (!author) {
          let prev = body.previousElementSibling;
          while (prev && !author) {
            const a = prev.querySelector?.('[data-tid="message-author-name"]') || (prev.getAttribute?.('data-tid') === 'message-author-name' ? prev : null);
            if (a) author = a.innerText.trim();
            prev = prev.previousElementSibling;
          }
        }

        const isMe = body.className.includes('ChatMyMessage') || body.className.includes('fui-ChatMyMessage');
        if (!author && isMe) {
          author = selfName;
        }

        const html = body.innerHTML || '';
        const text = body.innerText?.trim() || '';
        if (!text && !html) continue;

        // Timestamp-Extraktion mit mehreren Strategien (robust gegen Teams/UI-Variationen).
        const mid = body.getAttribute('data-mid') || '';
        let timestamp = null;
        let timestampRaw = null;

        // Strategie 1: data-mid ist bei Teams-Web oft ein Unix-Millis-Timestamp.
        if (mid) {
          const isoFromMid = toIso(mid);
          if (isoFromMid) {
            timestamp = isoFromMid;
            timestampRaw = mid;
          }
        }

        // Strategie 2: dediziertes Timestamp-Element (aria-labelledby referenziert "timestamp-{mid}").
        if (!timestamp && mid) {
          const tsEl = document.getElementById(`timestamp-${mid}`);
          if (tsEl) {
            const dT = tsEl.getAttribute('datetime') || tsEl.getAttribute('title') || '';
            if (dT) {
              const iso = toIso(dT);
              if (iso) { timestamp = iso; timestampRaw = dT || (tsEl.innerText || tsEl.textContent || '').trim(); }
            } else {
              timestampRaw = (tsEl.innerText || tsEl.textContent || '').trim();
            }
          }
        }

        // Strategie 3: internes <time datetime> im Body.
        if (!timestamp) {
          const timeEl = body.querySelector('time[datetime]');
          if (timeEl) {
            const iso = toIso(timeEl.getAttribute('datetime'));
            if (iso) { timestamp = iso; timestampRaw = timeEl.getAttribute('datetime'); }
          }
        }

        msgs.push({
          author: author || 'Unbekannt',
          html: html,
          text: text,
          mid: mid || undefined,
          timestamp: timestamp || undefined,
          timestampRaw: timestampRaw || undefined
        });
      }
      return msgs;
    }, { max: limit, selfName: config.selfName, helperStr: messageTimeToIso.toString() });

    return {
      tenant: t,
      chat: activeChat,
      count: messagesData.length,
      messages: messagesData.map(m => ({
        author: m.author,
        content: turndown.turndown(m.html || m.text),
        mid: m.mid,
        timestamp: m.timestamp,
        timestampRaw: m.timestampRaw
      }))
    };
  }

  async search(tenant = '', query, limit = 10) {
    if (!query) throw new Error("Suchbegriff (query) ist erforderlich.");
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    // Suchfeld finden - Teams Web ändert die DOM-Struktur; mehrere Wege probieren,
    // inkl. Öffnen der Suche per Tastaturkürzel (Ctrl+E / Ctrl+K / Alt+Shift+S).
    let searchInput = null;
    const selectors = [
      '#top-search-input',
      'input[aria-label*="Search"]',
      'input[placeholder*="Search"]',
      'input[aria-label*="Suche"]',
      'input[placeholder*="Suche"]',
      '[data-tid="searchbox"] input',
      'input[type="search"]'
    ];

    for (const sel of selectors) {
      searchInput = await page.$(sel);
      if (searchInput) break;
    }

    if (!searchInput) {
      // Suche per Tastatur öffnen und erneut versuchen
      for (const combo of ['Control+E', 'Control+K', 'Alt+Shift+S']) {
        try {
          await page.keyboard.press(combo);
          await page.waitForTimeout(1200);
        } catch (e) {}
        for (const sel of selectors) {
          searchInput = await page.$(sel);
          if (searchInput) break;
        }
        if (searchInput) break;
      }
    }

    if (!searchInput) {
      throw new Error("Suchfeld in Teams Web nicht gefunden (auch nicht nach Tastatur-Öffnung).");
    }

    await searchInput.click();
    await page.waitForTimeout(400);
    await searchInput.fill('');
    // Query per Clipboard-Paste einfügen (robust für Sonderzeichen/Um&laute)
    await page.evaluate((q) => { navigator.clipboard.writeText(q); }, query);
    await page.keyboard.press('ControlOrMeta+V');
    await page.waitForTimeout(400);
    await page.keyboard.press('Enter');

    await page.waitForTimeout(4000);
    await page.waitForSelector('div[data-tid*="search-result"], [role="listitem"], [data-tid*="searchResult"]', { timeout: 12000 }).catch(() => null);

    const results = await page.evaluate((max) => {
      const items = [];
      const rows = document.querySelectorAll('div[data-tid*="search-result"], [data-tid*="searchResult"], [role="listitem"]');
      for (const r of rows) {
        if (items.length >= max) break;
        const text = r.innerText?.trim() || '';
        if (text && text.length > 10) {
          items.push({
            preview: text.split('\n').filter(Boolean).join(' | ')
          });
        }
      }
      return items;
    }, limit);

    return {
      tenant: t,
      query: query,
      count: results.length,
      results: results
    };
  }

  async listTeams(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const teamsButton = await page.$('[data-tid="app-bar-teams"], button[aria-label*="Teams"], a[aria-label*="Teams"]');
    if (teamsButton) {
      await teamsButton.click().catch(() => null);
      await page.waitForTimeout(3000);
    }

    const tree = await page.evaluate(() => {
      const items = [];
      const rows = document.querySelectorAll('div[role="treeitem"], [data-tid="team-channel-item"]');
      for (const r of rows) {
        const text = r.innerText?.trim();
        if (text) {
          items.push({
            title: text.split('\n')[0],
            level: r.getAttribute('aria-level') || '1'
          });
        }
      }
      return items;
    });

    return {
      tenant: t,
      count: tree.length,
      channels: tree
    };
  }

  // Fügt Text per Clipboard-Paste statt keyboard.type ein.
  // Verhindert, dass Zeilenumbrüche (\n) in Teams als Enter/Ab-Senden interpretiert
  // werden und die Nachricht in mehrere Fragmente zersplittert.
  async insertTextViaPaste(page, text) {
    await page.evaluate((t) => {
      navigator.clipboard.writeText(t);
    }, text);
    await page.waitForTimeout(150);
    // Fokus im Compose-Feld + einfügen
    await page.keyboard.press('ControlOrMeta+A');
    await page.keyboard.press('ControlOrMeta+V');
    await page.waitForTimeout(400);
  }

  // Validiert und normalisiert Dateipfade für Anhänge
  validateAttachments(attachments) {
    if (!attachments) return [];
    const list = Array.isArray(attachments) ? attachments : [attachments];
    const resolved = [];
    for (const item of list) {
      if (!item || typeof item !== 'string') continue;
      const trimmed = item.trim();
      if (!trimmed) continue;
      const abs = path.resolve(trimmed);
      if (!fs.existsSync(abs)) {
        throw new Error(`Anhang-Datei nicht gefunden: "${trimmed}" (aufgelöst: "${abs}")`);
      }
      const stat = fs.statSync(abs);
      if (!stat.isFile()) {
        throw new Error(`Anhang ist keine reguläre Datei: "${trimmed}"`);
      }
      resolved.push(abs);
    }
    return resolved;
  }

  // Hängt eine oder mehrere lokale Dateien im Teams Chat-Compose-Bereich an
  async attachFiles(page, filePaths) {
    if (!filePaths || filePaths.length === 0) return;

    // 1. Kaskade: Existiert bereits ein input[type="file"] auf der Seite?
    const fileInputs = await page.$$('input[type="file"]');
    if (fileInputs.length > 0) {
      for (const input of fileInputs) {
        try {
          await input.setInputFiles(filePaths);
          await this.waitForAttachmentUpload(page, filePaths);
          return;
        } catch (e) {
          // Fallback zum nächsten Weg
        }
      }
    }

    // 2. Kaskade: Büroklammer / Attach-Button / Aktionen-Button im Compose-Footer
    const attachButtonSelectors = [
      // Spezifische Teams Test-IDs
      'button[data-tid="newMessageCommands-FilePicker"]',
      'button[data-tid*="FilePicker" i]',
      'button[data-tid*="file-picker" i]',
      'button[data-tid*="file-upload" i]',
      'button[data-tid*="attachment" i]',
      'button[data-tid="compose-attach-button"]',
      'button[data-tid="attach-button"]',
      'button[data-tid*="attach" i]',
      '[data-tid="chat-pane-compose-message-footer"] button[data-tid="attach-button"]',
      '[data-tid="chat-pane-compose-message-footer"] button[data-tid="compose-attach-button"]',
      '[data-tid="chat-pane-compose"] button[data-tid*="attach" i]',

      // Deutsche Labels (Teams Web DE)
      'button[aria-label*="anfügen" i]',
      'button[aria-label*="anhängen" i]',
      'button[aria-label*="anheften" i]',
      'button[aria-label*="Dateien" i]',
      'button[aria-label*="Datei" i]',
      'button[title*="anfügen" i]',
      'button[title*="anhängen" i]',
      'button[title*="anheften" i]',
      'button[title*="Datei" i]',

      // Englische Labels (Teams Web EN)
      'button[aria-label*="Attach" i]',
      'button[aria-label*="File" i]',
      'button[aria-label*="Upload" i]',
      'button[title*="Attach" i]',
      'button[title*="File" i]',

      // Icons (Paperclip / Büroklammer / Attach)
      'button:has(svg[data-icon-name*="Attach" i])',
      'button:has(svg[data-icon-name*="Paperclip" i])',
      'button:has(i[data-icon-name*="Attach" i])',
      'button:has(i[data-icon-name*="Paperclip" i])',

      // Modern Teams v2: "Aktionen und Apps" / "+" Button im Compose-Footer
      'button[aria-label*="Aktionen und Apps" i]',
      'button[aria-label*="Aktionen" i]',
      'button[aria-label*="Actions and apps" i]',
      'button[aria-label*="Add an action" i]',
      'button[aria-label*="Weitere Aktionen" i]',
      'button[data-tid*="action-overflow" i]',
      'button[data-tid*="actions-and-apps" i]',
      'button[data-tid="plus-button"]',
      'button[data-tid="expand-compose-actions-button"]'
    ];

    let attachBtn = null;
    for (const sel of attachButtonSelectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible().catch(() => true)) {
          attachBtn = btn;
          break;
        }
      } catch (e) {}
    }

    // Falls Selektoren keinen Treffer brachten: DOM-Scan über Toolbar-/Footer-Buttons
    if (!attachBtn) {
      const handle = await page.evaluateHandle(() => {
        const containers = Array.from(document.querySelectorAll('[role="toolbar"], [data-tid*="compose"], footer, [data-tid*="chat-pane"]'));
        for (const c of containers) {
          const btns = Array.from(c.querySelectorAll('button'));
          for (const b of btns) {
            const txt = (b.innerText || '').toLowerCase();
            const label = (b.getAttribute('aria-label') || '').toLowerCase();
            const title = (b.getAttribute('title') || '').toLowerCase();
            const tid = (b.getAttribute('data-tid') || '').toLowerCase();
            const props = `${txt} ${label} ${title} ${tid}`;
            if (
              props.includes('anfüg') ||
              props.includes('anhäng') ||
              props.includes('anheft') ||
              props.includes('attach') ||
              props.includes('datei') ||
              props.includes('file') ||
              props.includes('aktion') ||
              props.includes('action') ||
              props.includes('paperclip')
            ) {
              return b;
            }
          }
        }
        return null;
      });
      attachBtn = handle.asElement();
    }

    let fileChooser = null;
    if (attachBtn) {
      // Prüfen, ob Klick direkt den FileChooser öffnet
      const chooserPromise = page.waitForEvent('filechooser', { timeout: 2500 }).catch(() => null);
      await attachBtn.click().catch(() => null);
      fileChooser = await chooserPromise;

      if (!fileChooser) {
        // Möglicherweise hat sich ein Dropdown/Flyout-Menü geöffnet ("Von diesem Gerät hochladen")
        await page.waitForTimeout(600);
        const menuSelectors = [
          '[data-tid="upload-from-computer"]',
          '[data-tid*="upload-from-computer" i]',
          '[data-tid*="attach-upload-from-computer" i]',
          '[data-tid*="upload" i]',
          '[data-tid*="device" i]',
          '[role="menuitem"]:has-text("diesem Gerät")',
          '[role="menuitem"]:has-text("diesem Computer")',
          '[role="menuitem"]:has-text("Computer")',
          '[role="menuitem"]:has-text("this device")',
          '[role="menuitem"]:has-text("computer")',
          '[role="menuitem"]:has-text("Gerät")',
          '[role="menuitem"]:has-text("Upload")',
          '[role="menuitem"]:has-text("Dateien anfügen")',
          '[role="menuitem"]:has-text("Datei anfügen")',
          '[role="menuitem"]:has-text("Dateien hochladen")',
          '[role="menuitem"]:has-text("Datei hochladen")',
          'button:has-text("diesem Gerät")',
          'button:has-text("this device")',
          'button:has-text("Computer")',
          'button:has-text("Upload")'
        ];

        let menuOption = null;
        for (const sel of menuSelectors) {
          try {
            menuOption = await page.$(sel);
            if (menuOption && await menuOption.isVisible().catch(() => true)) break;
          } catch (e) {}
        }

        // Falls menuOption noch nicht gefunden: alle Menüeinträge per Text durchsuchen
        if (!menuOption) {
          const handle = await page.evaluateHandle(() => {
            const items = Array.from(document.querySelectorAll(
              '[role="menu"] [role="menuitem"], [role="menu"] button, [role="listbox"] [role="option"], div[role="menuitem"], div[class*="menu"] [role="menuitem"], div[class*="popover"] button, div[class*="flyout"] button'
            ));
            return items.find(el => {
              const txt = (el.innerText || el.textContent || el.getAttribute('aria-label') || '').toLowerCase();
              return txt.includes('gerät') || txt.includes('device') || txt.includes('computer') || txt.includes('upload') || txt.includes('hochladen') || txt.includes('datei');
            }) || null;
          });
          menuOption = handle.asElement();
        }

        if (menuOption) {
          const [uploadChooser] = await Promise.all([
            page.waitForEvent('filechooser', { timeout: 8000 }).catch(() => null),
            menuOption.click().catch(() => null)
          ]);
          fileChooser = uploadChooser;
        }
      }
    }

    if (fileChooser) {
      await fileChooser.setFiles(filePaths);
      await this.waitForAttachmentUpload(page, filePaths);
      return;
    }

    // 3. Kaskade: Eventuell wurde das input[type="file"] erst durch Interaktion ins DOM gehängt
    const lateInputs = await page.$$('input[type="file"]');
    if (lateInputs.length > 0) {
      for (const lateInput of lateInputs) {
        try {
          await lateInput.setInputFiles(filePaths);
          await this.waitForAttachmentUpload(page, filePaths);
          return;
        } catch (e) {}
      }
    }

    // 4. Kaskade: Drag & Drop Fallback via HTML5 DataTransfer auf das Compose-Feld
    try {
      const filesData = filePaths.map(fp => ({
        name: path.basename(fp),
        type: 'application/octet-stream',
        content: fs.readFileSync(fp).toString('base64')
      }));

      const dropSuccess = await page.evaluate(({ selector, files }) => {
        const target = document.querySelector(selector) || document.querySelector('div[role="textbox"]');
        if (!target) return false;

        const dt = new DataTransfer();
        for (const file of files) {
          const byteCharacters = atob(file.content);
          const byteNumbers = new Array(byteCharacters.length);
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const blob = new Blob([byteArray], { type: file.type });
          const domFile = new File([blob], file.name, { type: file.type });
          dt.items.add(domFile);
        }

        target.dispatchEvent(new DragEvent('dragenter', { dataTransfer: dt, bubbles: true, cancelable: true }));
        target.dispatchEvent(new DragEvent('dragover', { dataTransfer: dt, bubbles: true, cancelable: true }));
        target.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }));
        return true;
      }, { selector: '[data-tid="ckeditor"], div[role="textbox"], [contenteditable]', files: filesData });

      if (dropSuccess) {
        await this.waitForAttachmentUpload(page, filePaths);
        return;
      }
    } catch (dropErr) {
      // Ignorieren und unten verständlichen Fehler werfen
    }

    throw new Error("Dateianhang fehlgeschlagen: Kein Datei-Upload-Button, Datei-Input oder Drop-Ziel im Compose-Bereich ansprechbar.");
  }

  // Wartet auf das Fertigstellen des Uploads in Teams
  async waitForAttachmentUpload(page, filePaths = []) {
    await page.waitForTimeout(1500);

    // Auf Verschwinden von Progressbars warten
    try {
      const progressbars = page.locator('[data-tid*="compose"] [role="progressbar"], div[role="progressbar"], [data-tid*="progress"]');
      const count = await progressbars.count();
      if (count > 0) {
        await progressbars.first().waitFor({ state: 'detached', timeout: 60000 });
      }
    } catch (e) {}

    // Warten, bis Dateikarte(n) oder Anhang-Vorschau im Compose-Bereich sichtbar sind
    try {
      await page.waitForFunction(() => {
        const compose = document.querySelector('[data-tid*="compose"], footer');
        if (!compose) return true;
        const cards = compose.querySelectorAll('[data-tid*="attachment"], [data-tid*="file-card"], [class*="attachment"], [class*="fileCard"]');
        return cards.length > 0;
      }, { timeout: 8000 }).catch(() => null);
    } catch (e) {}

    // Sicherstellen, dass der Senden-Button aktiv / nicht disabled ist
    await page.waitForFunction(() => {
      const btn = document.querySelector('button[data-tid="newMessageCommands-send"], button[data-tid="send-message-button"], button[aria-label*="Send"], button[aria-label*="Senden"]');
      return btn && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true';
    }, { timeout: 15000 }).catch(() => null);

    await page.waitForTimeout(800);
  }

  async inspectCompose(tenant = '') {
    const raw = (!tenant || tenant === 'all') ? config.defaultTenant : tenant;
    const t = browserManager.normalizeTenant(raw);
    const page = await this.getPage(t, true);

    const activeChat = await this.getActiveChatTitle(page);

    const inspectData = await page.evaluate(() => {
      const fileInputs = Array.from(document.querySelectorAll('input[type="file"]')).map(i => ({
        id: i.id || null,
        name: i.name || null,
        dataTid: i.getAttribute('data-tid') || null,
        multiple: i.multiple,
        accept: i.accept || null,
        visible: i.offsetWidth > 0 && i.offsetHeight > 0
      }));

      const composeFooter = document.querySelector('[data-tid="chat-pane-compose-message-footer"], [data-tid="chat-pane-compose"], footer, [role="region"][aria-label*="compose" i]');
      const toolbars = Array.from(document.querySelectorAll('[role="toolbar"], footer, [data-tid*="compose"]'));
      const allButtons = [];
      for (const tb of toolbars) {
        for (const b of tb.querySelectorAll('button')) {
          allButtons.push({
            dataTid: b.getAttribute('data-tid') || null,
            ariaLabel: b.getAttribute('aria-label') || null,
            title: b.getAttribute('title') || null,
            innerText: b.innerText?.trim() || '',
            disabled: b.disabled || b.getAttribute('aria-disabled') === 'true'
          });
        }
      }

      return {
        hasComposeFooter: !!composeFooter,
        fileInputs,
        composeButtons: allButtons.slice(0, 30)
      };
    });

    return {
      tenant: t,
      activeChat,
      ...inspectData
    };
  }

  // Startet einen neuen Chat über das To:-Feld (Adressbuch/GAL), falls kein bestehender Chat vorhanden ist
  async startNewChat(page, personName) {
    // 1. Zuerst sicherstellen, dass wir in der Chat-App sind (Ctrl+Shift+4)
    await page.keyboard.press('Control+Shift+4');
    await page.waitForTimeout(1000);

    // 2. Neuer Chat per Shortcut Alt+Shift+N oder Button
    const newMsgBtn = await page.$('button[aria-label*="New message" i], button[aria-label*="Neuer Chat" i]');
    if (newMsgBtn) {
      await newMsgBtn.click();
    } else {
      await page.keyboard.press('Alt+Shift+N');
    }
    await page.waitForTimeout(1500);

    // 3. To-Input finden
    const toInput = await page.$(
      'input[aria-label*="name" i], input[placeholder*="name" i], input[aria-label*="Namen" i], input[placeholder*="Namen" i], input[data-tid*="people-picker"]'
    );
    if (!toInput) return null;

    await toInput.click();
    await toInput.fill('');
    await toInput.type(personName, { delay: 60 });
    await page.waitForTimeout(2000);

    // 4. Suggestions abwarten
    const optionLocator = page.locator('[role="listbox"] [role="option"], [role="option"], [data-tid*="suggestion"]');
    await optionLocator.first().waitFor({ state: 'visible', timeout: 8000 }).catch(() => null);

    const count = await optionLocator.count();
    if (count === 0) return null;

    // Finde passende Option
    let targetOption = null;
    let pickedTitle = personName;
    const norm = personName.toLowerCase();
    for (let i = 0; i < count; i++) {
      const opt = optionLocator.nth(i);
      const text = (await opt.innerText().catch(() => '')) || '';
      if (text.toLowerCase().includes(norm)) {
        targetOption = opt;
        pickedTitle = text.split('\n')[0].trim();
        break;
      }
    }
    if (!targetOption) {
      targetOption = optionLocator.first();
      const text = (await targetOption.innerText().catch(() => '')) || '';
      if (text) pickedTitle = text.split('\n')[0].trim();
    }

    await targetOption.click();
    await page.waitForTimeout(2500);
    return { title: pickedTitle };
  }

  async sendMessage(tenant = '', { message, chatName, attachments } = {}) {
    if (!message) throw new Error("Nachrichtentext (message) ist erforderlich.");
    if (!chatName) throw new Error("Empfänger (chat_name) ist erforderlich. Es wird bewusst nicht in einen unbestimmten 'aktiven' Chat gesendet.");
    const validAttachments = this.validateAttachments(attachments);
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    let picked = this.pickChatRow(filtered, chatName, undefined);
    if (!picked) {
      // Chat nicht in den geladenen Zeilen -> versuche neuen Chat über das Adressbuch
      const newChatResult = await this.startNewChat(page, chatName);
      if (!newChatResult) {
        throw new Error(`Chat oder Kollege '${chatName}' wurde in Teams (${t}) weder in bestehenden Chats noch im Adressbuch gefunden.`);
      }
      picked = newChatResult;
    } else {
      await picked.row.click();
      await page.waitForTimeout(2500);
    }

    // Compose-Feld GEZIELT im geöffneten Chat-Pane finden. WICHTIG: Darf NICHT das
    // globale div[role="textbox"] verwenden, sonst greift der Klick auf das
    // "New Message"-Compose-Feld und navigiert aus dem Chat heraus.
    const composeCtl = page
      .locator('[data-tid="chat-pane-compose-message-footer"], [data-tid="chat-pane-compose"]')
      .locator('[data-tid="ckeditor"], div[role="textbox"], [contenteditable], textarea, div[id^="new-message-"]')
      .first();

    try {
      await composeCtl.waitFor({ state: 'visible', timeout: 10000 });
    } catch (e) {
      throw new Error("Chat-Komposefeld im geöffneten Chat nicht gefunden. Keine Nachricht gesendet.");
    }
    // ElementHandle holen ist optional (Locator.click funktioniert direkt)
    await composeCtl.click();
    await page.waitForTimeout(500);
    await this.insertTextViaPaste(page, message);
    await page.waitForTimeout(600);

    // Anhänge hinzufügen, falls übergeben
    if (validAttachments.length > 0) {
      await this.attachFiles(page, validAttachments);
    }

    const sendButton = await page.$('button[data-tid="newMessageCommands-send"], button[data-tid="send-message-button"], button[aria-label*="Send"], button[aria-label*="Senden"]');
    if (sendButton) {
      await sendButton.click();
    } else {
      await page.keyboard.press('Enter');
    }
    await page.waitForTimeout(2000);

    const activeChat = await this.getActiveChatTitle(page);

    // Verifizieren, dass der aktive Chat dem Ziel entspricht (Sicherheitsnetz)
    const targetNorm = picked.title.toLowerCase().replace(/ \(you\)$/i, '');
    const activeNorm = activeChat.toLowerCase().replace(/ \(you\)$/i, '');
    const nameParts = targetNorm.split(/\s+/).filter(p => p.length > 2);
    const matchAny = nameParts.some(p => activeNorm.includes(p));
    if (!activeNorm.includes(targetNorm) && !targetNorm.includes(activeNorm) && !matchAny) {
      throw new Error(`Sicherheitsnetz: Der geöffnete Chat-Titel '${activeChat}' stimmt nicht mit Ziel '${picked.title}' überein. Nachricht wurde NICHT gesendet.`);
    }

    return {
      success: true,
      tenant: t,
      recipient: picked.title,
      confirmedActiveChat: activeChat,
      messageSent: message,
      attachments: validAttachments.map(p => path.basename(p)),
      status: validAttachments.length > 0
        ? `Nachricht mit ${validAttachments.length} Anhang/Anhängen erfolgreich gesendet.`
        : "Nachricht erfolgreich gesendet."
    };
  }
  async getMeetingStatus(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);
    const info = await speakerTracker.inspectMeeting(page);
    return {
      tenant: t,
      ...info,
      isTracking: speakerTracker.isTracking(t)
    };
  }

  async startSpeakerTracking(tenant = '', outputPath = null, options = {}) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);
    return await speakerTracker.startTracking(page, t, outputPath, options);
  }

  async stopSpeakerTracking(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    return await speakerTracker.stopTracking(t);
  }

  // Layer 1: Extrahiert den "Aktivität"-Tab (Activity-Feed) aus dem Teams-Web-DOM.
  // Die Roh-Items werden hier nicht klassifiziert (das macht der tim-Agent / die
  // Nachbearbeitung in Layer 2/3) — dieses Tool liefert nur die stabilen Rohdaten.
  async getActivity(tenant = '', options = {}) {
    return await activityClient.getActivity(tenant, options);
  }

  // Layer 2: Extrahiert den Activity-Feed UND klassifiziert die Einträge in
  // Kategorien (Meeting/Task/Entscheidung/Risiko/Sonstiges). Liefert die Gruppen
  // mit Relevanz-Scores, ohne eine Datei zu schreiben (reine Analyse).
  async getAnalyzedActivity(tenant = '', { maxItems = 50 } = {}) {
    const activity = await activityClient.getActivity(tenant, { maxItems });
    const groups = analyzeItems(activity.items);
    const counts = {};
    for (const cat of CATEGORIES) counts[cat] = groups[cat]?.length || 0;
    return {
      tenant: activity.tenant,
      extractedAt: activity.extractedAt,
      total: activity.count,
      counts,
      groups
    };
  }
}

// PURE Hilfsfunktion: konvertiert einen Teams-Zeitwert (data-mid als Unix-Millis/-Sekunden,
// <time datetime> ISO) zu einem ISO-8601-String. Muss reine Funktion bleiben (ohne Zugriff auf
// Modul-Scope), damit sie via .toString() für page.evaluate serialisierbar ist.
export function messageTimeToIso(raw) {
  if (raw == null || raw === '') return null;
  const s = String(raw).trim();
  if (!s) return null;
  // Unix-Epoche: 13-stellig = Millis, 10-stellig = Sekunden
  if (/^\d{10,13}$/.test(s)) {
    const num = Number(s);
    const ms = s.length >= 13 ? num : num * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) {
      return d.toISOString();
    }
  }
  // <time datetime="..."> ISO
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  return null;
}

export const teamsClient = new TeamsClient();
