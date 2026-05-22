import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import { createWriteStream, mkdirSync } from 'fs';
import path from 'path';
import { Ms365Auth, isOAuthAccount } from './ms365-auth.js';
import { CredentialVault } from './vault.js';

// Abort the connection if greeting + login don't complete within this window.
const DEFAULT_LOGIN_TIMEOUT_MS = 20000;

export class ImapManager {
  constructor(accounts, ms365Auth, vault) {
    this.accounts = accounts || [];
    this.ms365 = ms365Auth || new Ms365Auth();
    this.vault = vault || new CredentialVault();
  }

  getAccount(name) {
    const match = this.accounts.find(
      (a) =>
        a.name.toLowerCase() === name.toLowerCase() ||
        a.user.toLowerCase() === name.toLowerCase()
    );
    if (!match) {
      const names = this.accounts.map((a) => `"${a.name}"`).join(', ');
      throw new Error(`Account "${name}" not found. Configured accounts: ${names}`);
    }
    return match;
  }

  // Build and register a new account in-memory. Returns the new entry; the caller
  // is responsible for persisting `this.accounts` back to disk.
  // type: "microsoft365" (OAuth2 SSO) or "imap" (password).
  addAccount({
    type,
    name,
    user,
    host,
    port,
    secure,
    clientId,
    tenantId,
    password,
  }) {
    if (!name) throw new Error('"name" (a friendly label) is required.');
    if (!user) throw new Error('"user" (the email address) is required.');
    if (
      this.accounts.some((a) => a.name.toLowerCase() === name.toLowerCase())
    ) {
      throw new Error(`An account named "${name}" already exists.`);
    }

    const t = (type || '').toLowerCase();
    let entry;

    if (['microsoft365', 'm365', 'exchange', 'oauth2'].includes(t)) {
      if (!clientId) {
        throw new Error(
          'Microsoft 365 accounts need a "clientId" from your Azure AD app registration.'
        );
      }
      entry = {
        _comment: 'Microsoft 365 / Exchange Online — SSO via OAuth2',
        name,
        host: host || 'outlook.office365.com',
        port: port || 993,
        secure: secure !== false,
        user,
        authType: 'oauth2',
        clientId,
        tenantId: tenantId || 'common',
      };
    } else if (t === 'imap') {
      if (!host) throw new Error('IMAP accounts need a "host".');
      entry = {
        _comment: 'IMAP account',
        name,
        host,
        port: port || 993,
        secure: secure !== false,
        user,
      };
      // Store any provided password in the OS keychain rather than plaintext.
      if (password) this.vault.set(name, password);
    } else {
      throw new Error(
        `Unknown account type "${type}". Use "microsoft365" or "imap".`
      );
    }

    this.accounts.push(entry);
    return entry;
  }

  async buildAuth(account) {
    // Microsoft 365 / OAuth2 accounts authenticate with an XOAUTH2 access token
    // obtained via SSO instead of a stored password.
    if (isOAuthAccount(account)) {
      const accessToken = await this.ms365.getAccessToken(account);
      return { user: account.user, accessToken };
    }
    // Password accounts: prefer an explicit plaintext pass (back-compat), else
    // pull the secret from the OS keychain so accounts.json can omit it.
    const pass = account.pass || this.vault.get(account.name);
    if (!pass) {
      throw new Error(
        `No password for "${account.name}". Store one with the "set_credential" tool, ` +
          'or add a "pass" field in accounts.json.'
      );
    }
    return { user: account.user, pass };
  }

  async withClient(account, fn) {
    const auth = await this.buildAuth(account);
    const client = new ImapFlow({
      host: account.host,
      port: account.port || 993,
      secure: account.secure !== false,
      auth,
      logger: false,
      // Increase timeouts for slow servers
      socketTimeout: 30000,
      greetingTimeout: 15000,
    });

    // Time out the connection if login (greeting + authentication) takes too long.
    const loginTimeout = account.loginTimeout ?? DEFAULT_LOGIN_TIMEOUT_MS;
    let timer;
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => {
        client.close(); // force-abort the hung socket
        reject(new Error(`Login to "${account.name}" timed out after ${loginTimeout}ms`));
      }, loginTimeout);
    });
    try {
      await Promise.race([client.connect(), timeout]);
    } finally {
      clearTimeout(timer);
    }

    try {
      return await fn(client);
    } finally {
      try { await client.logout(); } catch (_) { /* ignore logout errors */ }
    }
  }

  async listFolders(accountName) {
    const account = this.getAccount(accountName);
    return this.withClient(account, async (client) => {
      const list = await client.list();
      return list.map((f) => ({
        name: f.name,
        path: f.path,
        flags: [...f.flags],
      }));
    });
  }

  async searchEmails({
    account: accountName,
    query,
    folder = 'INBOX',
    unread_only = false,
    since_days,
    limit = 20,
  }) {
    const targets =
      !accountName || accountName === 'all'
        ? this.accounts
        : [this.getAccount(accountName)];

    const allResults = [];

    for (const account of targets) {
      try {
        const emails = await this.withClient(account, async (client) => {
          await client.mailboxOpen(folder);

          // Build search criteria
          const criteria = [];
          if (unread_only) criteria.push('UNSEEN');
          if (since_days) {
            const since = new Date();
            since.setDate(since.getDate() - since_days);
            criteria.push(['SINCE', since]);
          }
          if (query) {
            // Search subject OR from address
            criteria.push(['OR', ['SUBJECT', query], ['FROM', query]]);
          }

          const searchCriteria = criteria.length ? criteria : ['ALL'];
          const uids = await client.search(searchCriteria, { uid: true });

          // Take most recent N
          const slice = uids.slice(-Math.min(limit, uids.length));
          if (!slice.length) return [];

          const messages = [];
          for await (const msg of client.fetch(
            slice,
            { uid: true, envelope: true, bodyStructure: true, flags: true },
            { uid: true }
          )) {
            const from = msg.envelope.from?.[0];
            const fromStr = from
              ? `${from.name ? from.name + ' ' : ''}<${from.address}>`.trim()
              : 'unknown';

            messages.push({
              uid: String(msg.uid),
              account: account.name,
              email: account.user,
              folder,
              subject: msg.envelope.subject || '(no subject)',
              from: fromStr,
              to: msg.envelope.to?.map((t) => t.address).join(', ') || '',
              date: msg.envelope.date,
              unread: !msg.flags.has('\\Seen'),
              has_attachments:
                msg.bodyStructure?.childNodes?.some(
                  (n) => n.disposition?.toLowerCase() === 'attachment'
                ) || false,
            });
          }
          return messages.reverse(); // newest first
        });
        allResults.push(...emails);
      } catch (err) {
        allResults.push({ account: account.name, error: err.message });
      }
    }

    return allResults;
  }

  async readEmail(accountName, folder, uid) {
    const account = this.getAccount(accountName);
    return this.withClient(account, async (client) => {
      await client.mailboxOpen(folder);
      let parsed = null;
      for await (const msg of client.fetch(uid, { source: true }, { uid: true })) {
        parsed = await simpleParser(msg.source);
      }
      if (!parsed) throw new Error(`Email UID ${uid} not found in ${folder}`);

      return {
        uid,
        account: accountName,
        email: account.user,
        folder,
        subject: parsed.subject || '(no subject)',
        from: parsed.from?.text || 'unknown',
        to: parsed.to?.text || '',
        cc: parsed.cc?.text || '',
        reply_to: parsed.replyTo?.text || '',
        date: parsed.date,
        text: parsed.text || '',
        html: parsed.html || null,
        attachments: parsed.attachments.map((a) => ({
          filename: a.filename || 'unnamed',
          contentType: a.contentType,
          size: a.size,
        })),
      };
    });
  }

  async getAttachments(accountName, folder, uid, downloadDir) {
    const account = this.getAccount(accountName);
    return this.withClient(account, async (client) => {
      await client.mailboxOpen(folder);
      let parsed = null;
      for await (const msg of client.fetch(uid, { source: true }, { uid: true })) {
        parsed = await simpleParser(msg.source);
      }
      if (!parsed) throw new Error(`Email UID ${uid} not found in ${folder}`);

      const meta = parsed.attachments.map((a) => ({
        filename: a.filename || 'unnamed',
        contentType: a.contentType,
        size: a.size,
        saved_to: null,
      }));

      if (downloadDir && parsed.attachments.length > 0) {
        mkdirSync(downloadDir, { recursive: true });
        for (let i = 0; i < parsed.attachments.length; i++) {
          const att = parsed.attachments[i];
          const fname = att.filename || `attachment-${i + 1}`;
          // Sanitize filename
          const safe = fname.replace(/[/\\?%*:|"<>]/g, '-');
          const outPath = path.join(downloadDir, safe);
          await new Promise((resolve, reject) => {
            const ws = createWriteStream(outPath);
            ws.end(att.content);
            ws.on('finish', resolve);
            ws.on('error', reject);
          });
          meta[i].saved_to = outPath;
        }
      }

      return { uid, account: accountName, folder, attachments: meta };
    });
  }

  async getRecentEmails(accountName = 'all', limit = 10, sinceDays = 7) {
    return this.searchEmails({
      account: accountName,
      folder: 'INBOX',
      unread_only: false,
      since_days: sinceDays,
      limit,
    });
  }
}
