# imap-mail plugin

Connect Claude to any IMAP mailbox — search, read, generate reports, get reply
suggestions, and surface follow-up reminders across all your email accounts.

Supports Microsoft Exchange, Yahoo Mail, Apple iCloud, cPanel hosting, and any
standard IMAP server.

---

## Setup (one-time)

### 1. Install Node.js dependencies

Open a terminal in the `server/` directory and run:

```bash
cd server
npm install
```

This installs `imapflow`, `mailparser`, `@modelcontextprotocol/sdk`,
`@azure/msal-node` (Microsoft 365 SSO), and `@napi-rs/keyring` (secure
credential storage).

### 2. Create your accounts config

**You don't have to create this file by hand.** The first time the MCP server
starts, if `accounts.json` is missing it is automatically created next to the
plugin (from `accounts.example.json`) — this is your local, git-ignored
credentials file where you enter non-SSO IMAP passwords. Just edit it.

To set it up manually instead:

```bash
cp accounts.example.json accounts.json
```

Edit `accounts.json` — add one entry per mailbox. Each entry needs:

| Field    | Description                                       |
|----------|---------------------------------------------------|
| `name`   | A friendly label (e.g. "Work Exchange")           |
| `host`   | IMAP hostname (see provider notes below)          |
| `port`   | Usually `993` (SSL/TLS)                           |
| `secure` | `true` for SSL/TLS (recommended)                 |
| `user`   | Your email address (or `DOMAIN\username` for on-prem Exchange) |
| `pass`   | *(optional)* password or app password — **better to store it in the keychain instead, see below** |

You can have up to 8+ accounts — just add more entries to the array.

> **Tip:** Leave `pass` out of `accounts.json` entirely. Add the account with
> just its `name`/`host`/`user`, then ask Claude to **"store the password for
> &lt;account&gt;"** (the `set_credential` tool). The secret goes into the OS
> keychain and never touches disk in plaintext.

### 3. Install the plugin

In Claude Cowork, go to **Settings → Plugins** and install this `.plugin` file.
Claude will automatically start the MCP server and load the skills.

---

## Provider-specific notes

### Microsoft 365 / Exchange Online — SSO (OAuth2) ✅ recommended
Microsoft has **disabled basic-auth (password) IMAP** for Exchange Online, so a
password will not work. Use single sign-on instead:

- **Host:** `outlook.office365.com`, port `993`, secure `true`
- In `accounts.json`, set `"authType": "oauth2"` and supply `clientId` /
  `tenantId` instead of `pass`. See the "Microsoft 365" entry in
  `accounts.example.json`.

**One-time Azure setup:**
1. Go to the [Azure / Entra portal](https://entra.microsoft.com) → **App registrations** → **New registration**. Give it a name; leave redirect URI blank.
2. Copy the **Application (client) ID** → that's your `clientId`. Copy the **Directory (tenant) ID** → that's your `tenantId` (or use `"common"`).
3. Under **API permissions** → **Add a permission** → **APIs my organization uses** → search **Office 365 Exchange Online** → **Delegated** → add **IMAP.AccessAsUser.All**. Also add **offline_access** (under Microsoft Graph delegated).
4. Under **Authentication** → **Advanced settings** → set **Allow public client flows** to **Yes** (required for the device-code flow).
5. Put `clientId` and `tenantId` in `accounts.json`, then ask Claude to **"sign in to my Microsoft 365 account"** (the `ms365_login` tool). It returns a code + URL — open it, enter the code, and approve. The refresh token is cached, so you only do this once.

The OAuth token cache is stored at `~/.imap-mail/ms365-token-cache.json` (file
mode `600`), **outside** this plugin folder — override with the
`IMAP_TOKEN_CACHE_FILE` env var.

### Microsoft Exchange (legacy / on-premises)
- **Host:** `mail.yourcompany.com` (on-premises Exchange — ask your IT team)
- Where basic auth is still permitted, use your password or an app password
  with the `pass` field.

### Yahoo Mail
- **Host:** `imap.mail.yahoo.com`
- **Required:** You must use an **App Password**, not your Yahoo login password.
  Generate one at: https://myaccount.yahoo.com/security → App passwords.

### Apple iCloud Mail
- **Host:** `imap.mail.me.com`
- **Required:** You must use an **App-Specific Password**, not your Apple ID password.
  Generate one at: https://appleid.apple.com → Sign-In and Security → App-Specific Passwords.
- Your username is your iCloud email (e.g. `you@icloud.com` or `you@me.com`).

---

## Skills

| Skill | Trigger phrases |
|-------|----------------|
| **email-report** | "summarize my inbox", "what emails do I have today", "inbox digest" |
| **reply-suggestion** | "draft a reply to X", "help me respond to [sender]" |
| **email-reminders** | "what needs follow-up", "any pending replies", "action items in my email" |

---

## Available tools (for Claude)

| Tool | What it does |
|------|-------------|
| `list_accounts` | Show all configured mailboxes |
| `list_folders` | List folders for an account |
| `search_emails` | Search by keyword, sender, date, unread status |
| `read_email` | Fetch full email body by UID |
| `get_recent_emails` | Get latest emails across all accounts |
| `get_attachments` | List or download attachments to a local folder |
| `ms365_login` | Start Microsoft 365 SSO (device-code flow) for an OAuth2 account |
| `set_credential` | Store an account's password securely in the OS keychain |
| `list_credentials` | Show how each account authenticates (never reveals passwords) |
| `test_credential` | Verify an account can actually log in with its stored credential |
| `delete_credential` | Remove an account's password from the OS keychain |

---

## Password manager (secure credential storage)

Passwords don't have to live in `accounts.json`. The plugin can store them in
your **OS keychain** — Windows Credential Manager, macOS Keychain, or Linux
Secret Service — via `@napi-rs/keyring`. Secrets are encrypted by the OS and
never written to disk in plaintext.

**Credential resolution order** for a password account, at connect time:
1. A plaintext `pass` in `accounts.json` (if present — kept for back-compat).
2. Otherwise, the password stored in the OS keychain under service `imap-mail`.
3. If neither exists, the tool returns a clear error telling you to run `set_credential`.

**Typical flow:**
```
"store the password for Tu Dang FF"   → set_credential   (saves to keychain)
"test the Tu Dang FF login"           → test_credential  (verifies it works)
"how are my accounts authenticating?" → list_credentials (oauth2 / keychain / plaintext / none)
"forget the stored password for X"    → delete_credential
```

After storing a password in the keychain, **delete the `pass` field from
`accounts.json`** so no plaintext remains. (Microsoft 365 accounts use SSO and
need no password at all.)

The keyring service name defaults to `imap-mail`; override with the
`IMAP_VAULT_SERVICE` env var.

---

## Security

- Prefer the keychain (`set_credential`) or Microsoft 365 SSO over plaintext —
  then `accounts.json` holds **no secrets at all**.
- Credentials are stored **only** in `accounts.json` on your local machine.
- The MCP server runs as a local process — nothing is sent to the cloud.
- Keep `accounts.json` private. It is listed in `.gitignore` if you version
  this plugin directory.
- Use app-specific passwords wherever possible (Yahoo, iCloud) to limit exposure.
- **Prefer SSO (OAuth2) for Microsoft 365** — no password is ever stored; only a
  refresh token in `~/.imap-mail/ms365-token-cache.json` (mode `600`), which you
  can revoke any time from your Microsoft account's app permissions.
- The login step aborts automatically if authentication hangs (default 20s;
  override per account with `"loginTimeout": <ms>`).
