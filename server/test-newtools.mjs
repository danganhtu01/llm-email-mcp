// Tests login_accounts + add_account against a TEMP accounts file (never the real one).
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { writeFileSync, readFileSync, rmSync } from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const serverPath = path.join(__dirname, 'index.js');
const tmp = path.join(os.tmpdir(), `newtools-${Date.now()}.json`);
writeFileSync(tmp, JSON.stringify([
  { name: 'TmpImap', host: 'imap.example.com', port: 993, secure: true, user: 'a@example.com' },
]));

async function call(c, name, args = {}) {
  const r = await c.callTool({ name, arguments: args });
  let p; try { p = JSON.parse(r.content?.[0]?.text ?? ''); } catch { p = r.content?.[0]?.text; }
  return { ok: !r.isError, p };
}

const t = new StdioClientTransport({ command: 'node', args: [serverPath], env: { ...process.env, IMAP_ACCOUNTS_FILE: tmp } });
const c = new Client({ name: 'tn', version: '1' }, { capabilities: {} });
await c.connect(t);

const tools = (await c.listTools()).tools.map((x) => x.name);
console.log('Tool count:', tools.length);
console.log('Has login_accounts:', tools.includes('login_accounts'));
console.log('Has add_account:', tools.includes('add_account'));
console.log('─'.repeat(60));

// login_accounts with no OAuth accounts → nothing to do
let r = await call(c, 'login_accounts');
console.log('login_accounts (no oauth):', r.p.instructions);

// add_account imap (password → keychain, not file)
r = await call(c, 'add_account', { type: 'imap', name: 'TmpImap2', user: 'b@example.com', host: 'mail.example.com', password: 'pw-secret-xyz' });
console.log('add imap:', r.p.status, '| pw in keychain:', r.p.password_in_keychain);

// add_account microsoft365
r = await call(c, 'add_account', { type: 'microsoft365', name: 'TmpM365', user: 'c@example.com', clientId: '00000000-test', tenantId: 'common' });
console.log('add m365:', r.p.status, '| next:', r.p.next);

// Verify the file now has 3 accounts and the m365 one is oauth2; imap password NOT in file
const written = JSON.parse(readFileSync(tmp, 'utf8'));
console.log('accounts in file:', written.map((a) => a.name).join(', '));
const m365 = written.find((a) => a.name === 'TmpM365');
const imap2 = written.find((a) => a.name === 'TmpImap2');
console.log('TmpM365 authType:', m365?.authType, '| clientId present:', !!m365?.clientId);
console.log('TmpImap2 has plaintext pass in file:', Object.prototype.hasOwnProperty.call(imap2 || {}, 'pass'), '(should be false)');

// list_credentials should show TmpImap2 as keychain
r = await call(c, 'list_credentials');
const cred = r.p.find((x) => x.name === 'TmpImap2');
console.log('TmpImap2 auth source:', cred?.auth, '(should be keychain)');

// Duplicate name should error
r = await call(c, 'add_account', { type: 'imap', name: 'TmpImap2', user: 'd@example.com', host: 'h' });
console.log('duplicate add rejected:', !r.ok, '→', r.p);

// Cleanup keychain entry + temp file
await call(c, 'delete_credential', { account: 'TmpImap2' });
await c.close();
rmSync(tmp, { force: true });
console.log('─'.repeat(60));
console.log('cleanup done');
process.exit(0);
