# claude-email plugin

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

This plugin is a directory (identified by `.claude-plugin/plugin.json`), not a
packaged file. Install it in Claude Code by adding this folder as a marketplace,
then installing from it:

```
/plugin marketplace add "$env:OneDrive\\GitHub\\claude-email-plugin"
/plugin install claude-email@local-marketplace
```

Or, to load it for a single session without installing:

```bash
claude --plugin-dir "$env:OneDrive\\GitHub\\claude-email-plugin"
```

Once enabled, Claude automatically starts the MCP server and loads the skills.
Verify with `/mcp` — the **claude-email** server should appear with its tools.

---

## Running modes: stdio vs HTTP (plugin vs connector vs gallery)

The server speaks MCP over two transports. **stdio is the default**; HTTP is opt-in.

| You want it to appear as… | Use | How it runs |
|---|---|---|
| A **plugin** (tools **+ skills**, under `/plugin` and `/mcp`) | Install via marketplace (above) | stdio, launched per-session by Claude Code |
| A **connector** (tools only, under `/mcp`) | `claude mcp add -s user claude-email -- node "<path>/server/index.js"` | stdio, launched per-session |
| A **custom connector in the remote Connectors gallery** | Run in HTTP mode (below) + add its URL | HTTP, persistent local service |

> The plugin and the stdio connector are the same server; don't enable both at
> once or every tool appears twice. A **stdio** server can never show up in the
> remote Connectors *gallery* (the page listing Notion/Gmail/etc.) — that page
> only lists **URL-based** connectors, which is what HTTP mode is for.

### HTTP mode

Set `MCP_HTTP_PORT` and the server runs as a persistent service exposing the
Streamable HTTP transport at `http://127.0.0.1:<port>/mcp` instead of stdio:

```bash
MCP_HTTP_PORT=3939 node server/index.js
```

| Env var | Default | Purpose |
|---|---|---|
| `MCP_HTTP_PORT` | *(unset → stdio)* | Port to listen on. Setting it enables HTTP mode. |
| `MCP_HTTP_HOST` | `127.0.0.1` | Bind address. Keep it loopback. |
| `MCP_HTTP_PATH` | `/mcp` | URL path for the MCP endpoint. |
| `MCP_HTTP_TOKEN` | *(none)* | If set, every request must send `Authorization: Bearer <token>`. |
| `MCP_HTTP_ALLOW_ORIGIN` | *(none → CORS off)* | Set to a specific origin to enable CORS for it. |

On Windows, copy [`start-http-server.vbs.example`](start-http-server.vbs.example)
to `start-http-server.vbs`, fill in the path/port/token, and double-click it to
run hidden (add a shortcut in `shell:startup` to auto-start at login). The real
`start-http-server.vbs` is git-ignored because the token is a secret.

Then register the URL (`http://127.0.0.1:3939/mcp`) via the Connectors page's
"Add custom connector". Supply the `Authorization: Bearer` header if the form
allows it; if it only supports "no auth", omit `MCP_HTTP_TOKEN`.

> **Security (HTTP mode):** these tools can read/send email and manage stored
> credentials. The server binds to loopback only and disables CORS by default,
> but any local process can still reach an un-tokened port — prefer setting
> `MCP_HTTP_TOKEN`, and stop the server when not in use.

---

## Multi-user / host-managed deployments

Beyond the single-user "one plugin folder on your machine" setup described
above, this server also works as one **stdio process per user**, spawned by a
host application that gives each instance its own environment. Two env vars
make that isolation possible:

| Env var | Purpose |
|---|---|
| `IMAP_ACCOUNTS_FILE` | Absolute path to that user's accounts file. Overrides the plugin-relative default (`accounts.json` next to this repo). |
| `MS365_TOKEN_CACHE_FILE` | Absolute path to that user's Microsoft 365 OAuth token cache. Overrides the default (`~/.imap-mail/ms365-token-cache.json`). The older name `IMAP_TOKEN_CACHE_FILE` still works as a back-compat alias. |

A host app should point both at a per-user location (e.g. under that user's
own data directory) before spawning the server, so each user's accounts and
OAuth tokens stay isolated from every other user's.

**Empty-file behavior.** When `IMAP_ACCOUNTS_FILE` is set explicitly and that
file doesn't exist yet, the server does **not** seed it from
`accounts.example.json` and does **not** exit — it starts with zero accounts
configured, and every account-consuming tool (`list_accounts`, `list_folders`,
`search_emails`, `read_email`, `send_email`, etc.) returns a clear message —
*"No email accounts configured. Add accounts to \<path\> …"* — instead of
erroring out or crashing. (The auto-seed-from-template convenience stays
reserved for the legacy plugin-relative default path, aimed at a single
interactive user setting the plugin up by hand.) The host app can create the
file itself — as a JSON array, `[]` to start or pre-populated — at any point,
including after the server has already started.

**Live reload, no restart.** The server re-reads `IMAP_ACCOUNTS_FILE` whenever
its mtime changes, checked with one cheap `stat` on every tool call (no
polling loop, no file watcher). This lets an external account manager add,
edit, or remove accounts on disk while the server keeps running — it picks up
the change on the very next tool call.

**Unknown fields tolerated.** An account entry only needs `name` and `user` to
be accepted; anything else a host app wants to stash on the entry (e.g. a
`_meta` block for its own bookkeeping) passes through untouched and is simply
ignored by the server. Entries missing `name`/`user` are skipped individually
(logged to stderr) rather than failing the whole file.

**Token cache permissions.** The OAuth token cache directory is created mode
`0700` and the cache file `0600` the first time a token is written, so only
the process owner can read cached refresh tokens.

### Per-provider setup, in short (for a host app's own onboarding UI)

| Provider | Account shape | What the user needs |
|---|---|---|
| **Microsoft 365 / Exchange Online** | `"authType": "oauth2"` + `clientId` (+ optional `tenantId`) | An Azure AD app registration — see "Microsoft 365 / Exchange Online — SSO" below. Then call the `ms365_login` tool (device-code SSO); no password is ever stored. |
| **Gmail** | `"imap"`, `host: "imap.gmail.com"` | **Not** wired up as an OAuth2 SSO provider here (there's no Google equivalent of the Microsoft 365 device-code flow in this server) — add it as a plain IMAP account with a Google **App Password**, which requires 2-Step Verification to be enabled: https://myaccount.google.com/apppasswords |
| **Yahoo Mail** | `"imap"`, `host: "imap.mail.yahoo.com"` | An **App Password** (not the Yahoo account password): https://myaccount.yahoo.com/security |
| **Apple iCloud Mail** | `"imap"`, `host: "imap.mail.me.com"` | An **App-Specific Password** (not the Apple ID password): https://appleid.apple.com → Sign-In and Security |
| **Generic cPanel / web hosting** | `"imap"`, `host: "mail.<domain>"` | The mailbox password |

For every non-SSO provider above: add the account (via `add_account`, or by
writing the entry into the accounts file directly) with no `pass` field, then
store its password with the `set_credential` tool — it goes to the OS
keychain, never into the accounts file in plaintext.

---

## Provider-specific notes

> ### ⚠️ Prefer plain IMAP over Exchange/Microsoft 365 where you can
> If a mailbox can be reached over **standard IMAP with an app password**, that
> path is **recommended over the Microsoft 365 / Exchange (OAuth2) path**. IMAP is
> simpler and more robust here:
> - **No Azure setup** — no app registration, client/tenant IDs, admin consent, or
>   "allow public client flows" toggles.
> - **No interactive sign-in or token expiry** — you store one app password in the
>   OS keychain once; there's no device-code dance and no refresh-token cache to go
>   stale or get revoked.
> - **Fewer moving parts to break** — the Exchange path depends on Microsoft tenant
>   policy that can change underneath you.
>
> Use the Microsoft 365 / Exchange (OAuth2) path **only when you have to** — i.e.
> for Exchange Online mailboxes, where Microsoft has disabled basic-auth IMAP and
> OAuth2 is the *only* option. For Gmail, Yahoo, iCloud, cPanel/web hosting, or any
> mailbox that still allows app passwords, add it as a plain `imap` account.

### Microsoft 365 / Exchange Online — SSO (OAuth2)
Microsoft has **disabled basic-auth (password) IMAP** for Exchange Online, so a
password will not work and OAuth2 SSO is the only option for these mailboxes. If
your mailbox is *not* on Exchange Online, prefer a plain IMAP account (see the
warning above). To set up Exchange Online SSO:

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

The OAuth token cache is stored at `~/.imap-mail/ms365-token-cache.json`
(directory mode `700`, file mode `600`), **outside** this plugin folder —
override with the `MS365_TOKEN_CACHE_FILE` env var (`IMAP_TOKEN_CACHE_FILE`
still works as an older alias). See "Multi-user / host-managed deployments"
above for per-user cache paths.

### Sending email (`send_email` / `create_draft`)
Sending uses SMTP, which is separate from the IMAP read path:

- **Microsoft 365 (OAuth2):** sending needs the `SMTP.Send` scope in addition to
  `IMAP.AccessAsUser.All`. Both live on the same `outlook.office365.com` resource,
  so a single sign-in consents to both. **Existing users must re-run
  `login_accounts` once** after this update so the cached token gains `SMTP.Send`
  — until then `send_email` returns a "not signed in" error. (Also add the
  delegated **SMTP.Send** permission to your Azure app registration alongside
  IMAP.AccessAsUser.All.) SMTP host defaults to `smtp.office365.com:587` (STARTTLS).
- **Password / plain IMAP accounts:** SMTP host defaults to the account's IMAP
  `host`. If your provider's SMTP host or port differ, add `smtpHost`, `smtpPort`
  (default 587), and `smtpSecure` (default `true` only on port 465) to the
  account in `accounts.json`.

A copy of every sent message is appended to the **Sent** folder by default
(`save_to_sent: false` to skip). Replies (`reply_to_uid`) add threading headers
and quote the original; forwards (`forward_uid`) attach the original as a `.eml`.

### Marking Outlook read/unread by Graph ID (`mark_read_graph`)
Messages surfaced by Microsoft's **Graph API** (e.g. via a read-only connector)
are addressed by a **Graph message ID** (`AAMkAG…`), not an IMAP UID, so the
UID-based `mark_read` can't act on them. `mark_read_graph` calls
`PATCH /v1.0/me/messages/{id}` with `{ isRead }` instead.

Graph is a **different Azure resource** (`graph.microsoft.com`) from IMAP/SMTP
(`outlook.office365.com`), and a single device-code sign-in cannot consent to
scopes across two resources. So Graph needs its **own one-time consent**: run
`ms365_login_graph` for the account once (grants `Mail.ReadWrite`), after which
`mark_read_graph` acquires Graph tokens silently from the shared refresh-token
cache. Add the delegated **Mail.ReadWrite** permission to your Azure app
registration alongside IMAP.AccessAsUser.All. M365/OAuth2 accounts only — IMAP
accounts already toggle read state via `mark_read`.

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
| `mark_read` | Mark one or more emails read/unread (sets/clears the IMAP `\Seen` flag) |
| `mark_read_graph` | Mark Outlook/Exchange messages read/unread by **Graph message ID** (for M365 messages from the Cowork connector); needs `ms365_login_graph` once |
| `ms365_login_graph` | One-time Microsoft Graph (Mail.ReadWrite) consent for an M365 account, enabling `mark_read_graph` |
| `move_email` | Move emails between folders (IMAP UID MOVE; destination must already exist) |
| `send_email` | Send a new email, reply, or forward via SMTP (M365 OAuth2 or password auth) |
| `flag_email` | Set/clear the `\Flagged` (starred / follow-up) flag on emails |
| `delete_email` | Delete emails — move to Trash by default, or expunge with `permanent:true` |
| `create_folder` | Create a new IMAP folder/mailbox |
| `delete_folder` | Delete an IMAP folder/mailbox |
| `rename_folder` | Rename or move an IMAP folder/mailbox |
| `create_draft` | Save a composed message to the Drafts folder without sending |
| `add_account` | Add a new Microsoft 365 (SSO) or IMAP account to `accounts.json` |
| `login_accounts` | Sign in every Microsoft 365 account that isn't authenticated yet |
| `ms365_login` | Start Microsoft 365 SSO (device-code flow) for one OAuth2 account |
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
