// Microsoft 365 SSO for IMAP via OAuth2 (XOAUTH2).
//
// Exchange Online no longer accepts basic-auth (password) IMAP, so M365 accounts
// authenticate with an OAuth2 access token obtained through Azure AD. This module
// uses MSAL's device-code flow: the user opens a URL, enters a short code, and
// approves once. The refresh token is cached to disk, so subsequent IMAP calls
// acquire access tokens silently without re-prompting.
import { PublicClientApplication, LogLevel } from '@azure/msal-node';
import { readFileSync, writeFileSync, existsSync, mkdirSync, chmodSync } from 'fs';
import os from 'os';
import path from 'path';

// IMAP delegated scope for Outlook/Exchange Online. offline_access yields a
// refresh token so we don't re-prompt every hour.
const IMAP_SCOPES = [
  'https://outlook.office365.com/IMAP.AccessAsUser.All',
  'offline_access',
];

// SMTP send scope (used by send_email). It lives on the SAME resource as the
// IMAP scope (outlook.office365.com), so a single device-code login can consent
// to both at once and one cached refresh token covers IMAP reads and SMTP sends.
export const SMTP_SCOPES = [
  'https://outlook.office365.com/SMTP.Send',
  'offline_access',
];

// Scopes requested at sign-in time. Bundling IMAP + SMTP here means existing
// users who re-run login_accounts after this update grant both in one approval.
// NOTE: only outlook.office365.com scopes belong here — see GRAPH_SCOPES below
// for why Microsoft Graph can't be added to this list.
const LOGIN_SCOPES = [
  'https://outlook.office365.com/IMAP.AccessAsUser.All',
  'https://outlook.office365.com/SMTP.Send',
  'offline_access',
];

// Microsoft Graph scopes for toggling Outlook/Exchange read state via the Graph
// REST API (mark_read_graph). Graph is a DIFFERENT resource (graph.microsoft.com)
// than the IMAP/SMTP scopes above (outlook.office365.com). Azure AD will not issue
// one access token spanning two resources, and MSAL's device-code flow rejects a
// request that mixes resources — so Graph CANNOT be merged into LOGIN_SCOPES
// without breaking the existing IMAP/SMTP sign-in. Instead Graph gets its own
// one-time device-code consent (the ms365_login_graph tool); afterwards the shared
// cached refresh token lets us acquire Graph tokens silently via
// getAccessToken(account, GRAPH_SCOPES).
export const GRAPH_SCOPES = [
  'https://graph.microsoft.com/Mail.ReadWrite',
  'offline_access',
];

// Scopes presented at the dedicated Graph sign-in (adds Mail.Read for parity).
export const GRAPH_LOGIN_SCOPES = [
  'https://graph.microsoft.com/Mail.ReadWrite',
  'https://graph.microsoft.com/Mail.Read',
  'offline_access',
];

// Token cache lives outside the (OneDrive-synced) plugin folder by default.
// MS365_TOKEN_CACHE_FILE is the primary override (documented, for multi-user /
// host-managed deployments — a host app sets one per-user path per spawned
// server instance). IMAP_TOKEN_CACHE_FILE is kept as a back-compat alias for
// anyone already relying on the older name.
const DEFAULT_CACHE_FILE =
  process.env.MS365_TOKEN_CACHE_FILE ||
  process.env.IMAP_TOKEN_CACHE_FILE ||
  path.join(os.homedir(), '.imap-mail', 'ms365-token-cache.json');

export function isOAuthAccount(account) {
  const t = (account.authType || account.auth || '').toString().toLowerCase();
  return t === 'oauth2' || t === 'xoauth2' || !!account.oauth2;
}

function oauthConfig(account) {
  const o = account.oauth2 || {};
  const clientId = account.clientId || o.clientId || process.env.MS365_CLIENT_ID;
  const tenantId =
    account.tenantId || o.tenantId || process.env.MS365_TENANT_ID || 'common';
  if (!clientId) {
    throw new Error(
      `Account "${account.name}" is configured for OAuth2 but has no clientId. ` +
        'Add "clientId" (and optionally "tenantId") from your Azure AD app registration to accounts.json.'
    );
  }
  return { clientId, tenantId };
}

export class Ms365AuthRequiredError extends Error {
  constructor(accountName) {
    super(
      `Microsoft 365 account "${accountName}" is not signed in. ` +
        `Run the "ms365_login" tool for this account to complete SSO.`
    );
    this.name = 'Ms365AuthRequiredError';
    this.accountName = accountName;
  }
}

export class Ms365Auth {
  constructor(cacheFile = DEFAULT_CACHE_FILE) {
    this.cacheFile = cacheFile;
    this.apps = new Map(); // key: `${tenantId}:${clientId}` → PublicClientApplication
    this.pendingLogins = new Map(); // account.name → device-code info
  }

  _cachePlugin() {
    const file = this.cacheFile;
    return {
      beforeCacheAccess: async (ctx) => {
        try {
          if (existsSync(file)) ctx.tokenCache.deserialize(readFileSync(file, 'utf8'));
        } catch (_) { /* start with an empty cache on read failure */ }
      },
      afterCacheAccess: async (ctx) => {
        if (ctx.cacheHasChanged) {
          try {
            const dir = path.dirname(file);
            mkdirSync(dir, { recursive: true });
            // Restrict the cache directory to the owner (0700), same intent as
            // the file's 0600 below — this file holds a refresh token.
            try { chmodSync(dir, 0o700); } catch (_) { /* e.g. unsupported on this fs */ }
            writeFileSync(file, ctx.tokenCache.serialize(), { mode: 0o600 });
            try { chmodSync(file, 0o600); } catch (_) { /* mode above already applied on create */ }
          } catch (_) { /* non-fatal: tokens just won't persist */ }
        }
      },
    };
  }

  _getApp(account) {
    const { clientId, tenantId } = oauthConfig(account);
    const key = `${tenantId}:${clientId}`;
    let app = this.apps.get(key);
    if (!app) {
      app = new PublicClientApplication({
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${tenantId}`,
        },
        cache: { cachePlugin: this._cachePlugin() },
        system: {
          loggerOptions: { logLevel: LogLevel.Error, piiLoggingEnabled: false },
        },
      });
      this.apps.set(key, app);
    }
    return app;
  }

  async _findCachedAccount(app, username) {
    const accounts = await app.getTokenCache().getAllAccounts();
    return accounts.find(
      (a) => a.username?.toLowerCase() === username.toLowerCase()
    );
  }

  // Returns a valid access token, refreshing silently from cache. Throws
  // Ms365AuthRequiredError if the user hasn't completed the device-code login yet.
  async getAccessToken(account, scopes = IMAP_SCOPES) {
    const app = this._getApp(account);
    const cached = await this._findCachedAccount(app, account.user);
    if (!cached) throw new Ms365AuthRequiredError(account.name);
    try {
      const result = await app.acquireTokenSilent({
        account: cached,
        scopes,
      });
      return result.accessToken;
    } catch (_) {
      // Refresh token expired/revoked — user must sign in again.
      throw new Ms365AuthRequiredError(account.name);
    }
  }

  // Starts the device-code flow and resolves immediately with the code + URL to
  // show the user. Token acquisition continues in the background and is cached on
  // completion, so later IMAP calls succeed silently.
  async beginDeviceLogin(account, scopes = LOGIN_SCOPES) {
    if (!isOAuthAccount(account)) {
      throw new Error(`Account "${account.name}" is not configured for OAuth2/SSO.`);
    }
    if (this.pendingLogins.has(account.name)) {
      return this.pendingLogins.get(account.name);
    }

    const app = this._getApp(account);

    let resolveCode, rejectCode;
    const codeReady = new Promise((res, rej) => {
      resolveCode = res;
      rejectCode = rej;
    });

    const loginPromise = app
      .acquireTokenByDeviceCode({
        scopes,
        deviceCodeCallback: (resp) => {
          resolveCode({
            account: account.name,
            user: account.user,
            userCode: resp.userCode,
            verificationUri: resp.verificationUri,
            expiresInSeconds: resp.expiresIn,
            message: resp.message,
          });
        },
      })
      .then(() => true)
      .catch((err) => {
        rejectCode(err); // surfaces if the flow fails before a code is issued
        throw err;
      })
      .finally(() => {
        this.pendingLogins.delete(account.name);
      });

    // Don't leave unhandled rejections if no one awaits the background promise.
    loginPromise.catch(() => {});

    const codeInfo = await codeReady;
    this.pendingLogins.set(account.name, codeInfo);
    return codeInfo;
  }
}
