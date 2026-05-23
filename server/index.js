#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { readFileSync, writeFileSync, existsSync, copyFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { ImapManager } from './imap-client.js';
import { Ms365Auth, isOAuthAccount } from './ms365-auth.js';
import { CredentialVault } from './vault.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Load accounts config ──────────────────────────────────────────────────────
// Priority: IMAP_ACCOUNTS_FILE env var → accounts.json next to this server dir.
// The path is plugin-relative (see .mcp.json: ${CLAUDE_PLUGIN_ROOT}/accounts.json),
// so it works on any machine the plugin is cloned to.
const configPath =
  process.env.IMAP_ACCOUNTS_FILE ||
  path.join(__dirname, '..', 'accounts.json');
const examplePath = path.join(__dirname, '..', 'accounts.example.json');

// First-run bootstrap: on a fresh install accounts.json is gitignored and absent,
// so seed it from the template. This gives the user a local (never-committed) file
// to enter their non-SSO IMAP passwords into.
if (!existsSync(configPath) && existsSync(examplePath)) {
  try {
    copyFileSync(examplePath, configPath);
    process.stderr.write(
      `[claude-email] Created ${configPath} from the template. ` +
        'Edit it to add your accounts, then either fill in "pass" or use the ' +
        'set_credential tool (OS keychain) for non-SSO accounts. Microsoft 365 uses SSO.\n'
    );
  } catch (e) {
    process.stderr.write(`[claude-email] Could not create ${configPath}: ${e.message}\n`);
  }
}

let accounts = [];
if (existsSync(configPath)) {
  try {
    accounts = JSON.parse(readFileSync(configPath, 'utf8'));
    if (!Array.isArray(accounts)) throw new Error('accounts.json must be an array');
  } catch (e) {
    process.stderr.write(`[claude-email] Failed to load ${configPath}: ${e.message}\n`);
  }
} else {
  process.stderr.write(
    `[claude-email] No accounts.json found at ${configPath} and no template to seed it. ` +
      'Create accounts.json (an array of account objects) to get started.\n'
  );
}

const ms365 = new Ms365Auth();
const vault = new CredentialVault();
const manager = new ImapManager(accounts, ms365, vault);

// ── MCP Server ────────────────────────────────────────────────────────────────
const server = new Server(
  { name: 'claude-email', version: '0.1.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    {
      name: 'list_accounts',
      description:
        'List all configured IMAP email accounts (name, address, host). Use this to discover which mailboxes are available before calling other tools.',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'list_folders',
      description: 'List all folders/mailboxes for a specific account.',
      inputSchema: {
        type: 'object',
        required: ['account'],
        properties: {
          account: {
            type: 'string',
            description: 'Account name as configured in accounts.json',
          },
        },
      },
    },
    {
      name: 'search_emails',
      description:
        'Search emails in a folder. Supports filtering by keyword, unread status, and date range. Returns metadata (subject, from, date, uid) — use read_email for the full body.',
      inputSchema: {
        type: 'object',
        required: ['account'],
        properties: {
          account: {
            type: 'string',
            description: 'Account name or "all" to search across every account',
          },
          query: {
            type: 'string',
            description: 'Keyword to match against subject or sender',
          },
          folder: {
            type: 'string',
            description: 'Folder to search (default: INBOX)',
          },
          unread_only: {
            type: 'boolean',
            description: 'Return only unread emails',
          },
          since_days: {
            type: 'number',
            description: 'Only emails from the last N days',
          },
          limit: {
            type: 'number',
            description: 'Max results per account (default: 20)',
          },
        },
      },
    },
    {
      name: 'read_email',
      description:
        'Fetch the full content (headers, plain text body, HTML body, attachment list) of a specific email by its UID.',
      inputSchema: {
        type: 'object',
        required: ['account', 'folder', 'uid'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          folder: {
            type: 'string',
            description: 'Folder the email is in (e.g. INBOX)',
          },
          uid: {
            type: 'string',
            description: 'Email UID returned by search_emails',
          },
        },
      },
    },
    {
      name: 'get_recent_emails',
      description:
        'Get recent emails from INBOX across all or a specific account. Shortcut for the most common "what\'s new" check.',
      inputSchema: {
        type: 'object',
        properties: {
          account: {
            type: 'string',
            description: 'Account name or "all" (default)',
          },
          limit: {
            type: 'number',
            description: 'Max emails per account (default: 10)',
          },
          since_days: {
            type: 'number',
            description: 'Days to look back (default: 7)',
          },
        },
      },
    },
    {
      name: 'get_attachments',
      description:
        'List attachments on an email, or download them to a local directory.',
      inputSchema: {
        type: 'object',
        required: ['account', 'folder', 'uid'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          folder: { type: 'string', description: 'Folder the email is in' },
          uid: { type: 'string', description: 'Email UID' },
          download_dir: {
            type: 'string',
            description:
              'Absolute path to a directory to save attachments. If omitted, only metadata is returned.',
          },
        },
      },
    },
    {
      name: 'ms365_login',
      description:
        'Begin Microsoft 365 single sign-on (SSO) for an OAuth2 account. Returns a short code and a URL — show them to the user so they can sign in and approve in a browser. After approval, the token is cached and IMAP tools work without a password.',
      inputSchema: {
        type: 'object',
        required: ['account'],
        properties: {
          account: {
            type: 'string',
            description: 'Name of the OAuth2/Microsoft 365 account in accounts.json',
          },
        },
      },
    },
    {
      name: 'login_accounts',
      description:
        'Sign in to every Microsoft 365 (SSO/OAuth2) account that is not already authenticated. Returns a device-code + URL for each account that needs login — show them to the user to approve in a browser. Accounts already signed in, and non-SSO (password/IMAP) accounts, are skipped. Optionally pass a single account name to log in just that one.',
      inputSchema: {
        type: 'object',
        properties: {
          account: {
            type: 'string',
            description:
              'Optional: only log in this one account. Omit to process all SSO accounts that need login.',
          },
        },
      },
    },
    {
      name: 'add_account',
      description:
        'Add a new email account to accounts.json. Supports type "microsoft365" (Exchange Online via SSO/OAuth2 — needs clientId, optional tenantId) or "imap" (generic IMAP — needs host; a provided password is stored in the OS keychain, never plaintext). After adding a microsoft365 account, run login_accounts (or ms365_login) to sign in.',
      inputSchema: {
        type: 'object',
        required: ['type', 'name', 'user'],
        properties: {
          type: {
            type: 'string',
            enum: ['microsoft365', 'imap'],
            description: '"microsoft365" for Exchange Online SSO, or "imap" for a password account',
          },
          name: { type: 'string', description: 'A friendly label, e.g. "Support Inbox"' },
          user: { type: 'string', description: 'The email address / login' },
          host: {
            type: 'string',
            description:
              'IMAP host. Required for imap; defaults to outlook.office365.com for microsoft365.',
          },
          port: { type: 'number', description: 'IMAP port (default 993)' },
          secure: { type: 'boolean', description: 'Use implicit TLS (default true)' },
          clientId: {
            type: 'string',
            description: 'Azure AD app client ID (required for microsoft365)',
          },
          tenantId: {
            type: 'string',
            description: 'Azure AD tenant ID for microsoft365 (default "common")',
          },
          password: {
            type: 'string',
            description:
              'For imap accounts only: stored securely in the OS keychain, not written to accounts.json.',
          },
        },
      },
    },
    {
      name: 'set_credential',
      description:
        'Securely store an IMAP password in the OS keychain (Windows Credential Manager) for an account, so it no longer needs to live in plaintext in accounts.json. Overwrites any existing stored password for that account.',
      inputSchema: {
        type: 'object',
        required: ['account', 'password'],
        properties: {
          account: {
            type: 'string',
            description: 'Account name as configured in accounts.json',
          },
          password: {
            type: 'string',
            description: 'The password (or app password) to store in the OS keychain',
          },
        },
      },
    },
    {
      name: 'list_credentials',
      description:
        'List configured accounts and how each one authenticates (OAuth2 SSO, keychain-stored password, plaintext pass in accounts.json, or none). Never reveals any password.',
      inputSchema: { type: 'object', properties: {} },
    },
    {
      name: 'test_credential',
      description:
        'Verify that an account can actually log in to its IMAP server using its currently resolved credential (keychain, plaintext, or OAuth token). Returns success or the server error.',
      inputSchema: {
        type: 'object',
        required: ['account'],
        properties: {
          account: { type: 'string', description: 'Account name to test' },
        },
      },
    },
    {
      name: 'delete_credential',
      description:
        'Remove an account\'s password from the OS keychain. Does not touch accounts.json.',
      inputSchema: {
        type: 'object',
        required: ['account'],
        properties: {
          account: { type: 'string', description: 'Account name whose stored password to delete' },
        },
      },
    },
    {
      name: 'mark_read',
      description:
        'Mark one or more emails as read or unread by setting or clearing the IMAP \\Seen flag via UID-based STORE. Accepts a batch of UIDs in one call. Reversible and non-destructive. Works for all IMAP/Exchange/M365 accounts (uses the existing OAuth2 token; no extra scope needed).',
      inputSchema: {
        type: 'object',
        required: ['account', 'folder', 'uids', 'read'],
        properties: {
          account: { type: 'string', description: 'Account name as configured in accounts.json' },
          folder: { type: 'string', description: 'Folder the emails are in (e.g. INBOX)' },
          uids: {
            type: 'array',
            items: { type: 'string' },
            description: 'One or more email UIDs to update (as returned by search_emails)',
          },
          read: {
            type: 'boolean',
            description: 'true = mark as read (+FLAGS \\Seen); false = mark as unread (-FLAGS \\Seen)',
          },
        },
      },
    },
    {
      name: 'move_email',
      description:
        'Move one or more emails from one folder to another via IMAP UID MOVE (falls back to COPY+delete on servers without MOVE). The destination folder must already exist — it is never auto-created.',
      inputSchema: {
        type: 'object',
        required: ['account', 'source_folder', 'destination_folder', 'uids'],
        properties: {
          account: { type: 'string', description: 'Account name as configured in accounts.json' },
          source_folder: { type: 'string', description: 'Current folder (e.g. INBOX)' },
          destination_folder: {
            type: 'string',
            description: 'Target folder (e.g. "2 CT", "Junk Email"). Must already exist.',
          },
          uids: {
            type: 'array',
            items: { type: 'string' },
            description: 'One or more UIDs to move (as returned by search_emails)',
          },
        },
      },
    },
    {
      name: 'send_email',
      description:
        'Send a new email, reply, or forward via SMTP. Microsoft 365 accounts use the cached OAuth2 token (SMTP.Send scope — re-run login_accounts once after updating to grant it); password accounts use SMTP AUTH (set smtpHost/smtpPort in accounts.json if they differ from the IMAP host). Set reply_to_uid to thread a reply (adds In-Reply-To/References and quotes the original), or forward_uid to forward the original as a .eml attachment. A copy is saved to the Sent folder by default.',
      inputSchema: {
        type: 'object',
        required: ['account', 'to'],
        properties: {
          account: { type: 'string', description: 'Account name (must support SMTP — M365/OAuth2 or an IMAP account with SMTP)' },
          to: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses' },
          subject: { type: 'string', description: 'Subject line (auto "Re:"/"Fwd:" if omitted on a reply/forward)' },
          body: { type: 'string', description: 'Plain-text body' },
          html_body: { type: 'string', description: 'HTML body (rich-text alternative)' },
          cc: { type: 'array', items: { type: 'string' }, description: 'CC recipients' },
          bcc: { type: 'array', items: { type: 'string' }, description: 'BCC recipients' },
          reply_to_uid: {
            type: 'string',
            description: 'UID of the message being replied to (sets threading headers + quotes the original)',
          },
          reply_folder: {
            type: 'string',
            description: 'Folder containing reply_to_uid (default INBOX)',
          },
          forward_uid: {
            type: 'string',
            description: 'UID of the message to forward (attached as a .eml / message/rfc822)',
          },
          forward_folder: {
            type: 'string',
            description: 'Folder containing forward_uid (default INBOX)',
          },
          attachments: {
            type: 'array',
            description: 'Additional file attachments',
            items: {
              type: 'object',
              required: ['filename', 'content_base64'],
              properties: {
                filename: { type: 'string' },
                content_base64: { type: 'string', description: 'Base64-encoded file contents' },
                mime_type: { type: 'string', description: 'Optional MIME type' },
              },
            },
          },
          save_to_sent: {
            type: 'boolean',
            description: 'Append a copy to the Sent folder after sending (default true)',
          },
        },
      },
    },
    {
      name: 'flag_email',
      description:
        'Set or clear the IMAP \\Flagged flag on one or more emails (the "starred / follow-up" marker) via UID-based STORE. Reversible and non-destructive.',
      inputSchema: {
        type: 'object',
        required: ['account', 'folder', 'uids', 'flagged'],
        properties: {
          account: { type: 'string', description: 'Account name as configured in accounts.json' },
          folder: { type: 'string', description: 'Folder the emails are in (e.g. INBOX)' },
          uids: {
            type: 'array',
            items: { type: 'string' },
            description: 'One or more UIDs to update',
          },
          flagged: {
            type: 'boolean',
            description: 'true = add \\Flagged; false = remove it',
          },
        },
      },
    },
    {
      name: 'delete_email',
      description:
        'Delete one or more emails. By default (permanent:false) they are moved to the Trash/Deleted Items folder, which is reversible. Pass permanent:true to expunge them irreversibly. The safe move-to-trash default is used unless permanent is explicitly true.',
      inputSchema: {
        type: 'object',
        required: ['account', 'folder', 'uids'],
        properties: {
          account: { type: 'string', description: 'Account name as configured in accounts.json' },
          folder: { type: 'string', description: 'Current folder the emails are in' },
          uids: {
            type: 'array',
            items: { type: 'string' },
            description: 'One or more UIDs to delete',
          },
          permanent: {
            type: 'boolean',
            description: 'false (default) = move to Trash/Deleted Items; true = expunge immediately (irreversible)',
          },
        },
      },
    },
    {
      name: 'create_folder',
      description: 'Create a new IMAP folder/mailbox. Nested paths use the server hierarchy separator (e.g. "Archive/2026").',
      inputSchema: {
        type: 'object',
        required: ['account', 'path'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          path: { type: 'string', description: 'Full folder path to create (e.g. "2 CT", "INBOX/Receipts")' },
        },
      },
    },
    {
      name: 'delete_folder',
      description: 'Delete an IMAP folder/mailbox. Errors (surfaced as-is) if the folder has children or is a system folder.',
      inputSchema: {
        type: 'object',
        required: ['account', 'path'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          path: { type: 'string', description: 'Folder path to delete' },
        },
      },
    },
    {
      name: 'rename_folder',
      description: 'Rename or move an IMAP folder/mailbox to a new path.',
      inputSchema: {
        type: 'object',
        required: ['account', 'old_path', 'new_path'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          old_path: { type: 'string', description: 'Current folder path' },
          new_path: { type: 'string', description: 'New folder path' },
        },
      },
    },
    {
      name: 'create_draft',
      description:
        'Compose a message and save it to the Drafts folder without sending (IMAP APPEND with the \\Draft flag). Accepts the same fields as send_email (to/subject/body/html_body/cc/bcc/reply_to_uid/attachments).',
      inputSchema: {
        type: 'object',
        required: ['account', 'to'],
        properties: {
          account: { type: 'string', description: 'Account name' },
          to: { type: 'array', items: { type: 'string' }, description: 'Recipient email addresses' },
          subject: { type: 'string', description: 'Subject line' },
          body: { type: 'string', description: 'Plain-text body' },
          html_body: { type: 'string', description: 'HTML body' },
          cc: { type: 'array', items: { type: 'string' }, description: 'CC recipients' },
          bcc: { type: 'array', items: { type: 'string' }, description: 'BCC recipients' },
          reply_to_uid: { type: 'string', description: 'UID of a message this draft replies to (sets threading headers + quotes original)' },
          reply_folder: { type: 'string', description: 'Folder containing reply_to_uid (default INBOX)' },
          attachments: {
            type: 'array',
            description: 'File attachments',
            items: {
              type: 'object',
              required: ['filename', 'content_base64'],
              properties: {
                filename: { type: 'string' },
                content_base64: { type: 'string', description: 'Base64-encoded file contents' },
                mime_type: { type: 'string', description: 'Optional MIME type' },
              },
            },
          },
        },
      },
    },
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  try {
    let result;

    switch (name) {
      case 'list_accounts':
        result = accounts.map((a) => ({
          name: a.name,
          email: a.user,
          host: a.host,
          port: a.port || 993,
          secure: a.secure !== false,
          auth: isOAuthAccount(a) ? 'oauth2 (Microsoft 365 SSO)' : 'password',
        }));
        break;

      case 'list_folders':
        result = await manager.listFolders(args.account);
        break;

      case 'search_emails':
        result = await manager.searchEmails(args);
        break;

      case 'read_email':
        result = await manager.readEmail(args.account, args.folder, args.uid);
        break;

      case 'get_recent_emails':
        result = await manager.getRecentEmails(
          args.account,
          args.limit,
          args.since_days
        );
        break;

      case 'get_attachments':
        result = await manager.getAttachments(
          args.account,
          args.folder,
          args.uid,
          args.download_dir
        );
        break;

      case 'ms365_login': {
        const account = manager.getAccount(args.account);
        const info = await ms365.beginDeviceLogin(account);
        result = {
          status: 'awaiting_user',
          instructions: `Open ${info.verificationUri} and enter code ${info.userCode} to sign in as ${info.user}. Approval is required within ${Math.round(info.expiresInSeconds / 60)} minutes; the token is cached automatically once you approve.`,
          ...info,
        };
        break;
      }

      case 'login_accounts': {
        // Process one named account, or every OAuth2 account that needs login.
        const targets = args.account
          ? [manager.getAccount(args.account)].filter(isOAuthAccount)
          : accounts.filter(isOAuthAccount);

        const logins = [];
        for (const a of targets) {
          // Already signed in? acquireTokenSilent succeeds → skip.
          let signedIn = false;
          try {
            await ms365.getAccessToken(a);
            signedIn = true;
          } catch (_) {
            signedIn = false;
          }
          if (signedIn) {
            logins.push({ account: a.name, user: a.user, status: 'already_signed_in' });
            continue;
          }
          try {
            const info = await ms365.beginDeviceLogin(a);
            logins.push({
              account: a.name,
              user: a.user,
              status: 'awaiting_user',
              userCode: info.userCode,
              verificationUri: info.verificationUri,
              expiresInSeconds: info.expiresInSeconds,
            });
          } catch (e) {
            logins.push({ account: a.name, user: a.user, status: 'error', error: e.message });
          }
        }

        const pending = logins.filter((l) => l.status === 'awaiting_user');
        result = {
          logins,
          instructions: pending.length
            ? `For each account below: open its verificationUri and enter its userCode, signing in with that account's address and approving. ${pending.length} account(s) need sign-in. Then run test_credential (or login_accounts again) to confirm.`
            : 'No accounts need sign-in. (Non-SSO/password accounts are not handled here.)',
        };
        break;
      }

      case 'add_account': {
        const entry = manager.addAccount(args);
        // Persist the updated in-memory array back to accounts.json.
        writeFileSync(configPath, JSON.stringify(accounts, null, 2) + '\n');
        const { pass, clientId, ...rest } = entry;
        result = {
          status: 'added',
          account: { ...rest, ...(clientId ? { clientId } : {}) },
          password_in_keychain: args.type === 'imap' && !!args.password,
          file: configPath,
          next:
            entry.authType === 'oauth2'
              ? `Run login_accounts (or ms365_login for "${entry.name}") to complete SSO.`
              : `Run test_credential for "${entry.name}" to verify the login.`,
        };
        break;
      }

      case 'set_credential': {
        const account = manager.getAccount(args.account);
        vault.set(account.name, args.password);
        result = {
          status: 'stored',
          account: account.name,
          message: `Password for "${account.name}" stored in the OS keychain. You can now remove the "pass" field from accounts.json. Use test_credential to verify it works.`,
        };
        break;
      }

      case 'list_credentials':
        result = accounts.map((a) => {
          let source;
          if (isOAuthAccount(a)) source = 'oauth2 (Microsoft 365 SSO)';
          else if (a.pass) source = 'plaintext (accounts.json)';
          else if (vault.has(a.name)) source = 'keychain';
          else source = 'none (not set)';
          return { name: a.name, email: a.user, auth: source };
        });
        break;

      case 'test_credential': {
        const account = manager.getAccount(args.account);
        try {
          const folders = await manager.listFolders(account.name);
          result = {
            account: account.name,
            ok: true,
            message: `Login succeeded; ${folders.length} folders visible.`,
          };
        } catch (e) {
          result = { account: account.name, ok: false, error: e.message };
        }
        break;
      }

      case 'mark_read':
        result = await manager.markRead(
          args.account,
          args.folder,
          args.uids,
          args.read
        );
        break;

      case 'move_email':
        result = await manager.moveEmail(
          args.account,
          args.source_folder,
          args.destination_folder,
          args.uids
        );
        break;

      case 'send_email':
        result = await manager.sendEmail(args.account, args);
        break;

      case 'flag_email':
        result = await manager.flagEmail(
          args.account,
          args.folder,
          args.uids,
          args.flagged
        );
        break;

      case 'delete_email':
        result = await manager.deleteEmail(
          args.account,
          args.folder,
          args.uids,
          args.permanent === true
        );
        break;

      case 'create_folder':
        result = await manager.createFolder(args.account, args.path);
        break;

      case 'delete_folder':
        result = await manager.deleteFolder(args.account, args.path);
        break;

      case 'rename_folder':
        result = await manager.renameFolder(
          args.account,
          args.old_path,
          args.new_path
        );
        break;

      case 'create_draft':
        result = await manager.createDraft(args.account, args);
        break;

      case 'delete_credential': {
        const account = manager.getAccount(args.account);
        const removed = vault.delete(account.name);
        result = {
          account: account.name,
          removed,
          message: removed
            ? `Removed the stored password for "${account.name}" from the OS keychain.`
            : `No keychain password was stored for "${account.name}".`,
        };
        break;
      }

      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    return {
      content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
    };
  } catch (err) {
    return {
      content: [{ type: 'text', text: `Error: ${err.message}` }],
      isError: true,
    };
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────
const transport = new StdioServerTransport();
await server.connect(transport);
