import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import MailComposer from 'nodemailer/lib/mail-composer/index.js';
import { createWriteStream, mkdirSync } from 'fs';
import path from 'path';
import { Ms365Auth, isOAuthAccount, SMTP_SCOPES, GRAPH_SCOPES } from './ms365-auth.js';
import { CredentialVault } from './vault.js';

// Abort the connection if greeting + login don't complete within this window.
const DEFAULT_LOGIN_TIMEOUT_MS = 20000;

export class ImapManager {
  // accountsFilePath is optional and used only to make error/help messages
  // point at the right file (e.g. a per-user path set via IMAP_ACCOUNTS_FILE
  // in a multi-user/host-managed deployment) — it's never read here.
  constructor(accounts, ms365Auth, vault, accountsFilePath = null) {
    this.accounts = accounts || [];
    this.ms365 = ms365Auth || new Ms365Auth();
    this.vault = vault || new CredentialVault();
    this.accountsFilePath = accountsFilePath;
  }

  // Shared "nothing configured yet" message for every account-consuming tool,
  // so a fresh/empty accounts file (e.g. a brand-new per-user file that a host
  // app hasn't populated yet) surfaces a helpful instruction instead of a bare
  // "not found" error or a crash.
  noAccountsMessage() {
    const where = this.accountsFilePath
      ? `Add accounts to ${this.accountsFilePath}`
      : 'Add accounts to your accounts file';
    return `No email accounts configured. ${where} (edit it directly, or use the add_account tool), then try again.`;
  }

  getAccount(name) {
    if (this.accounts.length === 0) {
      throw new Error(this.noAccountsMessage());
    }
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

    if (targets.length === 0) {
      throw new Error(this.noAccountsMessage());
    }

    const allResults = [];

    for (const account of targets) {
      try {
        const emails = await this.withClient(account, async (client) => {
          await client.mailboxOpen(folder);

          // Build the search query. IMPORTANT: imapflow uses an OBJECT query
          // ({ seen, since, or, all, ... }), NOT node-imap-style arrays like
          // ['UNSEEN', ['SINCE', date]]. Passing arrays makes client.search()
          // return `false`, which then blows up at `uids.slice(...)` with
          // "uids.slice is not a function". Top-level keys are AND-ed together.
          const searchQuery = {};
          if (unread_only) searchQuery.seen = false;
          if (since_days) {
            const since = new Date();
            since.setDate(since.getDate() - since_days);
            searchQuery.since = since;
          }
          if (query) {
            // Match the keyword in the subject OR the sender.
            searchQuery.or = [{ subject: query }, { from: query }];
          }
          if (Object.keys(searchQuery).length === 0) searchQuery.all = true;

          // imapflow resolves to an array of UIDs, or `false` on a failed/empty
          // search — normalize so .slice is always safe.
          const found = await client.search(searchQuery, { uid: true });
          const uids = Array.isArray(found) ? found : [];

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

  // Mark one or more messages read/unread by setting or clearing the \Seen flag
  // via UID-based IMAP STORE. read=true → +FLAGS (\Seen); read=false → -FLAGS.
  async markRead(accountName, folder, uids, read) {
    if (typeof read !== 'boolean') {
      throw new Error('"read" must be a boolean (true = mark read, false = mark unread).');
    }
    return this._setFlag(accountName, folder, uids, '\\Seen', read, {
      onLabel: 'read',
      offLabel: 'unread',
    });
  }

  // Mark Outlook/Exchange messages read/unread via the Microsoft Graph REST API,
  // by Graph message ID (the `id` returned by outlook_email_search / Graph), not
  // by IMAP UID. This is the path for messages surfaced by the read-only Cowork
  // M365 connector, which exposes Graph IDs rather than IMAP UIDs. Requires a
  // Graph (Mail.ReadWrite) token — granted once via ms365_login_graph; thereafter
  // acquired silently from the shared refresh-token cache.
  async graphMarkRead(accountName, messageIds, read) {
    const account = this.getAccount(accountName);
    if (!isOAuthAccount(account)) {
      throw new Error(
        `mark_read_graph only works with Microsoft 365 / OAuth2 accounts. "${account.name}" is a password/IMAP account — use mark_read with its UID instead.`
      );
    }
    if (!Array.isArray(messageIds) || messageIds.length === 0) {
      throw new Error('"message_ids" must be a non-empty array of Graph message IDs.');
    }
    if (typeof read !== 'boolean') {
      throw new Error('"read" must be a boolean (true = mark read, false = mark unread).');
    }

    let token;
    try {
      token = await this.ms365.getAccessToken(account, GRAPH_SCOPES);
    } catch (e) {
      return {
        ok: false,
        updated: 0,
        error: `Microsoft Graph is not authorized for "${account.name}" (${e.message}). Run ms365_login_graph for this account once to grant Mail.ReadWrite, then retry.`,
      };
    }

    let updated = 0;
    const failures = [];
    for (const id of messageIds) {
      try {
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(id)}`,
          {
            method: 'PATCH',
            headers: {
              Authorization: `Bearer ${token}`,
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ isRead: read }),
          }
        );
        if (res.ok) {
          updated++;
        } else {
          const body = await res.text().catch(() => '');
          failures.push({ id, status: res.status, error: body.slice(0, 300) });
        }
      } catch (err) {
        failures.push({ id, error: err.message });
      }
    }

    if (failures.length === 0) {
      return {
        ok: true,
        updated,
        message: `Marked ${updated} Outlook message(s) as ${read ? 'read' : 'unread'} via Microsoft Graph.`,
      };
    }
    return {
      ok: updated > 0,
      updated,
      failed: failures.length,
      errors: failures,
      message: `Updated ${updated}/${messageIds.length} via Graph; ${failures.length} failed.`,
    };
  }

  // Set or clear the \Flagged flag (the "starred / follow-up" marker) on a batch
  // of UIDs. flagged=true → +FLAGS (\Flagged); flagged=false → -FLAGS.
  async flagEmail(accountName, folder, uids, flagged) {
    if (typeof flagged !== 'boolean') {
      throw new Error('"flagged" must be a boolean (true = add \\Flagged, false = remove it).');
    }
    return this._setFlag(accountName, folder, uids, '\\Flagged', flagged, {
      onLabel: 'flagged',
      offLabel: 'unflagged',
    });
  }

  // Shared UID-based STORE helper for boolean IMAP flags (\Seen, \Flagged, …).
  async _setFlag(accountName, folder, uids, flag, on, { onLabel, offLabel }) {
    const account = this.getAccount(accountName);
    if (!Array.isArray(uids) || uids.length === 0) {
      throw new Error('"uids" must be a non-empty array of email UIDs.');
    }

    return this.withClient(account, async (client) => {
      // STORE requires the folder to be SELECTed read-write. mailboxOpen issues
      // SELECT (read-write) by default — do NOT pass { readOnly: true } (EXAMINE).
      const mbox = await client.mailboxOpen(folder);
      // Comma-separated UID set → single round-trip for the whole batch.
      const range = uids.map((u) => String(u)).join(',');

      try {
        if (mbox && mbox.readOnly) {
          return {
            ok: false,
            updated: 0,
            failed: uids.length,
            message: `Folder "${folder}" opened read-only; cannot change flags.`,
          };
        }
        const applied = on
          ? await client.messageFlagsAdd(range, [flag], { uid: true })
          : await client.messageFlagsRemove(range, [flag], { uid: true });

        if (applied) {
          return {
            ok: true,
            updated: uids.length,
            failed: 0,
            message: `Marked ${uids.length} message(s) as ${on ? onLabel : offLabel} in "${folder}".`,
          };
        }
        return {
          ok: false,
          updated: 0,
          failed: uids.length,
          message: `Server did not apply the ${flag} change in "${folder}" (mailbox may be read-only).`,
        };
      } catch (err) {
        return {
          ok: false,
          updated: 0,
          failed: uids.length,
          message: `Failed to update flags in "${folder}": ${err.message}`,
        };
      }
    });
  }

  // Find a folder by its IMAP special-use attribute (e.g. "\\Sent", "\\Drafts",
  // "\\Trash", "\\Junk"), falling back to common name matches. Returns the path
  // string, or null if nothing matches.
  async _findSpecialFolder(client, specialUse, fallbackNames = []) {
    const list = await client.list();
    const su = specialUse.toLowerCase();
    const bySpecialUse = list.find(
      (f) =>
        (f.specialUse || '').toLowerCase() === su ||
        [...(f.flags || [])].some((fl) => String(fl).toLowerCase() === su)
    );
    if (bySpecialUse) return bySpecialUse.path;
    const wanted = fallbackNames.map((n) => n.toLowerCase());
    const byName = list.find(
      (f) =>
        wanted.includes((f.name || '').toLowerCase()) ||
        wanted.includes((f.path || '').toLowerCase())
    );
    return byName ? byName.path : null;
  }

  // Move messages between folders via UID MOVE (imapflow falls back to
  // COPY + \Deleted + EXPUNGE on servers without MOVE).
  async moveEmail(accountName, sourceFolder, destinationFolder, uids) {
    const account = this.getAccount(accountName);
    if (!sourceFolder) throw new Error('"source_folder" is required.');
    if (!destinationFolder) throw new Error('"destination_folder" is required.');
    if (!Array.isArray(uids) || uids.length === 0) {
      throw new Error('"uids" must be a non-empty array of email UIDs.');
    }

    return this.withClient(account, async (client) => {
      const mbox = await client.mailboxOpen(sourceFolder);
      if (mbox && mbox.readOnly) {
        return {
          ok: false,
          moved: 0,
          failed: uids.length,
          message: `Folder "${sourceFolder}" opened read-only; cannot move messages.`,
        };
      }
      // Don't auto-create the destination — surface a clear error instead.
      const exists = await client.mailboxOpen(destinationFolder).then(
        () => true,
        () => false
      );
      if (!exists) {
        return {
          ok: false,
          moved: 0,
          failed: uids.length,
          message: `Destination folder "${destinationFolder}" does not exist. Create it first (create_folder) or check the name.`,
        };
      }
      // Re-select the source (the existence probe left dest selected).
      await client.mailboxOpen(sourceFolder);
      const range = uids.map((u) => String(u)).join(',');
      try {
        await client.messageMove(range, destinationFolder, { uid: true });
        return {
          ok: true,
          moved: uids.length,
          failed: 0,
          message: `Moved ${uids.length} message(s) from "${sourceFolder}" to "${destinationFolder}".`,
        };
      } catch (err) {
        return {
          ok: false,
          moved: 0,
          failed: uids.length,
          message: `Failed to move messages: ${err.message}`,
        };
      }
    });
  }

  // Delete messages. Default (permanent=false) moves them to the Trash/Deleted
  // Items folder (reversible). permanent=true expunges them irreversibly.
  async deleteEmail(accountName, folder, uids, permanent = false) {
    const account = this.getAccount(accountName);
    if (!folder) throw new Error('"folder" is required.');
    if (!Array.isArray(uids) || uids.length === 0) {
      throw new Error('"uids" must be a non-empty array of email UIDs.');
    }

    return this.withClient(account, async (client) => {
      const mbox = await client.mailboxOpen(folder);
      if (mbox && mbox.readOnly) {
        return {
          ok: false,
          deleted: 0,
          failed: uids.length,
          message: `Folder "${folder}" opened read-only; cannot delete messages.`,
        };
      }
      const range = uids.map((u) => String(u)).join(',');

      if (permanent) {
        try {
          // messageDelete sets \Deleted and EXPUNGEs in one call.
          await client.messageDelete(range, { uid: true });
          return {
            ok: true,
            deleted: uids.length,
            failed: 0,
            permanent: true,
            message: `Permanently deleted ${uids.length} message(s) from "${folder}".`,
          };
        } catch (err) {
          return {
            ok: false,
            deleted: 0,
            failed: uids.length,
            message: `Failed to permanently delete: ${err.message}`,
          };
        }
      }

      // Soft delete = move to Trash. Find the trash folder via special-use.
      const trash = await this._findSpecialFolder(client, '\\Trash', [
        'Trash',
        'Deleted Items',
        'Deleted',
        'Junk',
      ]);
      if (!trash) {
        return {
          ok: false,
          deleted: 0,
          failed: uids.length,
          message:
            'No Trash/Deleted Items folder found. Pass permanent:true to expunge, or move the messages manually.',
        };
      }
      if (trash.toLowerCase() === folder.toLowerCase()) {
        return {
          ok: false,
          deleted: 0,
          failed: uids.length,
          message: `Messages are already in the Trash folder ("${folder}"). Pass permanent:true to expunge them.`,
        };
      }
      // Re-select the source folder (list() doesn't change selection, but be safe).
      await client.mailboxOpen(folder);
      try {
        await client.messageMove(range, trash, { uid: true });
        return {
          ok: true,
          deleted: uids.length,
          failed: 0,
          permanent: false,
          trash_folder: trash,
          message: `Moved ${uids.length} message(s) to "${trash}" (reversible). Pass permanent:true to expunge instead.`,
        };
      } catch (err) {
        return {
          ok: false,
          deleted: 0,
          failed: uids.length,
          message: `Failed to move messages to Trash: ${err.message}`,
        };
      }
    });
  }

  async createFolder(accountName, folderPath) {
    const account = this.getAccount(accountName);
    if (!folderPath) throw new Error('"path" (the folder path to create) is required.');
    return this.withClient(account, async (client) => {
      const res = await client.mailboxCreate(folderPath);
      return {
        ok: true,
        path: res?.path || folderPath,
        message: `Created folder "${res?.path || folderPath}".`,
      };
    });
  }

  async deleteFolder(accountName, folderPath) {
    const account = this.getAccount(accountName);
    if (!folderPath) throw new Error('"path" (the folder path to delete) is required.');
    return this.withClient(account, async (client) => {
      await client.mailboxDelete(folderPath);
      return { ok: true, message: `Deleted folder "${folderPath}".` };
    });
  }

  async renameFolder(accountName, oldPath, newPath) {
    const account = this.getAccount(accountName);
    if (!oldPath) throw new Error('"old_path" is required.');
    if (!newPath) throw new Error('"new_path" is required.');
    return this.withClient(account, async (client) => {
      const res = await client.mailboxRename(oldPath, newPath);
      return {
        ok: true,
        old_path: res?.path || oldPath,
        new_path: res?.newPath || newPath,
        message: `Renamed folder "${res?.path || oldPath}" → "${res?.newPath || newPath}".`,
      };
    });
  }

  // Build the raw RFC 822 bytes of a message from high-level fields, optionally
  // threading a reply (In-Reply-To/References + quoted original) or forwarding an
  // original as a .eml attachment. Returns { raw, envelope, subject }.
  async _composeMessage(account, opts) {
    const {
      to,
      subject,
      body,
      html_body,
      cc,
      bcc,
      reply_to_uid,
      reply_folder = 'INBOX',
      forward_uid,
      forward_folder = 'INBOX',
      attachments = [],
    } = opts;

    if (!Array.isArray(to) || to.length === 0) {
      throw new Error('"to" must be a non-empty array of recipient addresses.');
    }

    const mailOptions = {
      from: account.user,
      to,
      cc: cc && cc.length ? cc : undefined,
      bcc: bcc && bcc.length ? bcc : undefined,
      subject: subject || '',
      text: body || undefined,
      html: html_body || undefined,
      attachments: [],
    };

    // Threading / quoting for replies, and .eml attachment for forwards, both
    // need the original message — fetch it once over IMAP.
    if (reply_to_uid || forward_uid) {
      const fetchFolder = reply_to_uid ? reply_folder : forward_folder;
      const fetchUid = reply_to_uid || forward_uid;
      await this.withClient(account, async (client) => {
        await client.mailboxOpen(fetchFolder);
        const orig = await client.fetchOne(
          String(fetchUid),
          { uid: true, source: true, envelope: true },
          { uid: true }
        );
        if (!orig) {
          throw new Error(
            `Original message UID ${fetchUid} not found in "${fetchFolder}".`
          );
        }

        if (reply_to_uid) {
          const env = orig.envelope || {};
          const origId = env.messageId;
          if (origId) {
            mailOptions.inReplyTo = origId;
            mailOptions.references = origId;
          }
          if (!subject) {
            const s = env.subject || '';
            mailOptions.subject = /^re:/i.test(s) ? s : `Re: ${s}`;
          }
          // Prepend a quoted copy of the original to whatever body was supplied.
          const parsed = await simpleParser(orig.source);
          const when = env.date ? new Date(env.date).toUTCString() : '';
          const who = parsed.from?.text || 'unknown sender';
          const quotedText = (parsed.text || '')
            .split('\n')
            .map((l) => `> ${l}`)
            .join('\n');
          const attribution = `On ${when}, ${who} wrote:`;
          mailOptions.text = `${body || ''}\n\n${attribution}\n${quotedText}`;
          if (html_body || parsed.html) {
            mailOptions.html = `${html_body || ''}<br><br><blockquote>${parsed.html || (parsed.text || '').replace(/\n/g, '<br>')}</blockquote>`;
          }
        }

        if (forward_uid) {
          mailOptions.attachments.push({
            filename: 'forwarded.eml',
            content: orig.source,
            contentType: 'message/rfc822',
          });
          if (!subject) {
            const s = orig.envelope?.subject || '';
            mailOptions.subject = /^fwd?:/i.test(s) ? s : `Fwd: ${s}`;
          }
        }
      });
    }

    // Caller-supplied attachments (base64-encoded).
    for (const a of attachments) {
      if (!a || !a.filename || !a.content_base64) {
        throw new Error(
          'Each attachment needs "filename" and "content_base64".'
        );
      }
      mailOptions.attachments.push({
        filename: a.filename,
        content: Buffer.from(a.content_base64, 'base64'),
        contentType: a.mime_type || undefined,
      });
    }

    const raw = await new Promise((resolve, reject) => {
      new MailComposer(mailOptions).compile().build((err, msg) =>
        err ? reject(err) : resolve(msg)
      );
    });

    const envelope = {
      from: account.user,
      to: [...to, ...(cc || []), ...(bcc || [])],
    };
    return { raw, envelope, subject: mailOptions.subject };
  }

  // Resolve SMTP connection settings + auth for an account. M365/OAuth2 accounts
  // use XOAUTH2 with an SMTP-scoped token; password accounts use SMTP AUTH.
  async _smtpTransport(account) {
    // Bound every phase so a wrong host/port fails fast instead of hanging.
    const timeouts = {
      connectionTimeout: 20000,
      greetingTimeout: 15000,
      socketTimeout: 30000,
    };
    if (isOAuthAccount(account)) {
      const accessToken = await this.ms365.getAccessToken(account, SMTP_SCOPES);
      return nodemailer.createTransport({
        host: account.smtpHost || 'smtp.office365.com',
        port: account.smtpPort || 587,
        secure: account.smtpSecure ?? false, // STARTTLS on 587
        auth: { type: 'OAuth2', user: account.user, accessToken },
        ...timeouts,
      });
    }
    const pass = account.pass || this.vault.get(account.name);
    if (!pass) {
      throw new Error(
        `No password for "${account.name}" to authenticate SMTP. Store one with set_credential.`
      );
    }
    // Default the SMTP host to the IMAP host; most providers share it. Override
    // with smtpHost/smtpPort/smtpSecure in accounts.json if they differ.
    const port = account.smtpPort || 587;
    return nodemailer.createTransport({
      host: account.smtpHost || account.host,
      port,
      secure: account.smtpSecure ?? port === 465,
      auth: { user: account.user, pass },
      ...timeouts,
    });
  }

  // Send a new message / reply / forward via SMTP, optionally saving a copy to
  // the Sent folder over IMAP.
  async sendEmail(accountName, opts) {
    const account = this.getAccount(accountName);
    const { save_to_sent = true } = opts;

    const { raw, envelope, subject } = await this._composeMessage(account, opts);

    const transport = await this._smtpTransport(account);
    let info;
    try {
      info = await transport.sendMail({ envelope, raw });
    } finally {
      transport.close();
    }

    let savedToSent = false;
    if (save_to_sent) {
      try {
        await this.withClient(account, async (client) => {
          const sent = await this._findSpecialFolder(client, '\\Sent', [
            'Sent',
            'Sent Items',
            'Sent Mail',
          ]);
          if (sent) {
            await client.append(sent, raw, ['\\Seen']);
            savedToSent = true;
          }
        });
      } catch (_) {
        // Non-fatal: the message was sent even if the Sent copy failed.
        savedToSent = false;
      }
    }

    return {
      ok: true,
      message_id: info.messageId || null,
      accepted: info.accepted || [],
      rejected: info.rejected || [],
      saved_to_sent: savedToSent,
      subject,
      message: `Sent "${subject || '(no subject)'}" from ${account.user} to ${envelope.to.join(', ')}.`,
    };
  }

  // Save a composed message to the Drafts folder (IMAP APPEND), without sending.
  async createDraft(accountName, opts) {
    const account = this.getAccount(accountName);
    const { raw, subject } = await this._composeMessage(account, opts);

    return this.withClient(account, async (client) => {
      const drafts = await this._findSpecialFolder(client, '\\Drafts', [
        'Drafts',
        'Draft',
      ]);
      if (!drafts) {
        throw new Error(
          'No Drafts folder found for this account. Create one with create_folder first.'
        );
      }
      const res = await client.append(drafts, raw, ['\\Draft', '\\Seen']);
      return {
        ok: true,
        uid: res?.uid ? String(res.uid) : null,
        folder: drafts,
        subject,
        message: `Saved draft "${subject || '(no subject)'}" to "${drafts}".`,
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
