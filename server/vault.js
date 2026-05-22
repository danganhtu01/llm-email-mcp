// Secure credential store backed by the OS keychain (Windows Credential Manager,
// macOS Keychain, Linux Secret Service) via @napi-rs/keyring. Lets accounts.json
// hold no plaintext password — the secret lives encrypted in the OS vault and is
// fetched at connect time.
import { Entry } from '@napi-rs/keyring';

// Keyring "service" namespace; the account name is the per-entry key.
const SERVICE = process.env.IMAP_VAULT_SERVICE || 'imap-mail';

function entryFor(accountName) {
  return new Entry(SERVICE, accountName);
}

export class CredentialVault {
  // Store (or overwrite) a password for an account.
  set(accountName, password) {
    if (!password) throw new Error('Password must be a non-empty string.');
    entryFor(accountName).setPassword(password);
  }

  // Return the stored password, or null if none is set.
  get(accountName) {
    try {
      return entryFor(accountName).getPassword();
    } catch (_) {
      // keyring throws when no entry exists for this service/account.
      return null;
    }
  }

  has(accountName) {
    return this.get(accountName) !== null;
  }

  // Remove a stored password. Returns true if one was deleted.
  delete(accountName) {
    try {
      return entryFor(accountName).deletePassword();
    } catch (_) {
      return false;
    }
  }
}
