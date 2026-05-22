// Drives the credential-manager tools through the live MCP server.
// Uses a throwaway test account so it never touches real secrets.
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { writeFileSync, rmSync } from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const serverPath = path.join(__dirname, 'index.js');

// Temp accounts file with a fake account so we don't depend on real config.
const tmpAccounts = path.join(os.tmpdir(), `imap-vault-test-${Date.now()}.json`);
writeFileSync(tmpAccounts, JSON.stringify([
  { name: 'VaultTest', host: 'imap.invalid.example', port: 993, secure: true, user: 'probe@example.com' },
]));

async function call(client, name, args = {}) {
  const res = await client.callTool({ name, arguments: args });
  let parsed;
  try { parsed = JSON.parse(res.content?.[0]?.text ?? ''); } catch { parsed = res.content?.[0]?.text; }
  return { ok: !res.isError, parsed };
}

const transport = new StdioClientTransport({
  command: 'node', args: [serverPath],
  env: { ...process.env, IMAP_ACCOUNTS_FILE: tmpAccounts },
});
const client = new Client({ name: 'test-vault', version: '1.0.0' }, { capabilities: {} });
await client.connect(transport);

const tools = (await client.listTools()).tools.map((t) => t.name);
console.log('Tools:', tools.join(', '));
console.log('Count:', tools.length);
console.log('─'.repeat(60));

// 1. list before storing → should be "none (not set)"
let r = await call(client, 'list_credentials');
console.log('list (before):', JSON.stringify(r.parsed));

// 2. store a password
r = await call(client, 'set_credential', { account: 'VaultTest', password: 's3cr3t-test-pw' });
console.log('set:', r.parsed.status, '-', r.ok ? 'PASS' : 'FAIL');

// 3. list after storing → should be "keychain"
r = await call(client, 'list_credentials');
const src = r.parsed[0]?.auth;
console.log('list (after):', src, src === 'keychain' ? 'PASS' : 'FAIL');

// 4. test_credential → should fail to connect (bad host) but resolve the keychain pw
r = await call(client, 'test_credential', { account: 'VaultTest' });
console.log('test:', r.parsed.ok === false && /invalid|ENOTFOUND|getaddrinfo|timed out/i.test(r.parsed.error || '')
  ? 'PASS (resolved pw, expected connect failure)' : 'CHECK', '→', r.parsed.error);

// 5. delete
r = await call(client, 'delete_credential', { account: 'VaultTest' });
console.log('delete:', r.parsed.removed === true ? 'PASS' : 'FAIL', '-', r.parsed.message);

// 6. list after delete → "none (not set)"
r = await call(client, 'list_credentials');
console.log('list (final):', r.parsed[0]?.auth, r.parsed[0]?.auth === 'none (not set)' ? 'PASS' : 'FAIL');

await client.close();
rmSync(tmpAccounts, { force: true });
process.exit(0);
