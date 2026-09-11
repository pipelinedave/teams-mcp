import TurndownService from 'turndown';
import { browserManager } from './browserManager.js';
import { config } from './config.js';
import { speakerTracker } from './speakerTracker.js';

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

    const messagesData = await page.evaluate(({ max, selfName }) => {
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

        if (text) {
          msgs.push({
            author: author || 'Unbekannt',
            html: html,
            text: text
          });
        }
      }
      return msgs;
    }, { max: limit, selfName: config.selfName });

    return {
      tenant: t,
      chat: activeChat,
      count: messagesData.length,
      messages: messagesData.map(m => ({
        author: m.author,
        content: turndown.turndown(m.html || m.text)
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

  async sendMessage(tenant = '', { message, chatName } = {}) {
    if (!message) throw new Error("Nachrichtentext (message) ist erforderlich.");
    if (!chatName) throw new Error("Empfänger (chat_name) ist erforderlich. Es wird bewusst nicht in einen unbestimmten 'aktiven' Chat gesendet.");
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    const picked = this.pickChatRow(filtered, chatName, undefined);
    if (!picked) {
      throw new Error(`Chat '${chatName}' wurde in Teams (${t}) nicht gefunden. Keine Nachricht gesendet (verhindert Fehlversand).`);
    }

    await picked.row.click();
    await page.waitForTimeout(2500);

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

    const sendButton = await page.$('button[data-tid="send-message-button"], button[aria-label*="Send"], button[aria-label*="Senden"]');
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
    if (!activeNorm.includes(targetNorm) && !targetNorm.includes(activeNorm)) {
      throw new Error(`Sicherheitsnetz: Der geöffnete Chat-Titel '${activeChat}' stimmt nicht mit Ziel '${picked.title}' überein. Nachricht wurde NICHT gesendet.`);
    }

    return {
      success: true,
      tenant: t,
      recipient: picked.title,
      confirmedActiveChat: activeChat,
      messageSent: message,
      status: "Nachricht erfolgreich gesendet."
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
}

export const teamsClient = new TeamsClient();
