// Live end-to-end test: spawns the actual MCP server over stdio and exercises
// all 6 tools against the configured accounts. Run: node test-live.mjs
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const serverPath = path.join(__dirname, 'index.js');
const accountsFile = path.join(__dirname, '..', 'accounts.json');

function hr() { console.log('─'.repeat(70)); }
async function call(client, name, args = {}) {
  const t0 = Date.now();
  try {
    const res = await client.callTool({ name, arguments: args });
    const ms = Date.now() - t0;
    const text = res.content?.[0]?.text ?? '';
    let parsed;
    try { parsed = JSON.parse(text); } catch { parsed = text; }
    return { ok: !res.isError, ms, parsed, raw: text };
  } catch (e) {
    return { ok: false, ms: Date.now() - t0, parsed: null, raw: e.message };
  }
}

const transport = new StdioClientTransport({
  command: 'node',
  args: [serverPath],
  env: { ...process.env, IMAP_ACCOUNTS_FILE: accountsFile },
});
const client = new Client({ name: 'test-live', version: '1.0.0' }, { capabilities: {} });
await client.connect(transport);

const summary = [];
function record(label, r, extra = '') {
  const tag = r.ok ? 'PASS' : 'FAIL';
  summary.push(`${tag}  ${label}  (${r.ms}ms) ${extra}`);
  console.log(`[${tag}] ${label} (${r.ms}ms) ${extra}`);
}

hr();
const tools = await client.listTools();
console.log('Tools advertised:', tools.tools.map((t) => t.name).join(', '));
console.log('Tool count:', tools.tools.length);
hr();

// 1. list_accounts
const accts = await call(client, 'list_accounts');
record('list_accounts', accts, `→ ${Array.isArray(accts.parsed) ? accts.parsed.map((a) => a.name).join(' | ') : ''}`);
const accountNames = Array.isArray(accts.parsed) ? accts.parsed.map((a) => a.name) : [];

// Per-account deep test for the IMAP-backed tools
for (const acct of accountNames) {
  hr();
  console.log(`### Account: ${acct}`);

  // 2. list_folders
  const folders = await call(client, 'list_folders', { account: acct });
  const folderList = Array.isArray(folders.parsed) ? folders.parsed.map((f) => f.path) : [];
  record(`list_folders [${acct}]`, folders, `→ ${folderList.length} folders`);

  // 3. get_recent_emails
  const recent = await call(client, 'get_recent_emails', { account: acct, limit: 3, since_days: 30 });
  const recentArr = Array.isArray(recent.parsed) ? recent.parsed.filter((m) => !m.error) : [];
  record(`get_recent_emails [${acct}]`, recent, `→ ${recentArr.length} msgs`);

  // 4. search_emails
  const search = await call(client, 'search_emails', { account: acct, folder: 'INBOX', limit: 3, since_days: 60 });
  const searchArr = Array.isArray(search.parsed) ? search.parsed.filter((m) => !m.error) : [];
  record(`search_emails [${acct}]`, search, `→ ${searchArr.length} hits`);

  // Pick a UID for read/attachment tests
  const sample = searchArr[0] || recentArr[0];
  if (sample && sample.uid) {
    // 5. read_email
    const read = await call(client, 'read_email', { account: acct, folder: sample.folder || 'INBOX', uid: sample.uid });
    record(`read_email [${acct}] uid=${sample.uid}`, read,
      `→ subject="${(read.parsed?.subject || '').slice(0, 40)}"`);

    // 6. get_attachments (metadata only)
    const att = await call(client, 'get_attachments', { account: acct, folder: sample.folder || 'INBOX', uid: sample.uid });
    record(`get_attachments [${acct}] uid=${sample.uid}`, att,
      `→ ${att.parsed?.attachments?.length ?? 0} attachments`);
  } else {
    console.log(`  (no message found for ${acct} — skipping read_email/get_attachments)`);
  }
}

hr();
console.log('SUMMARY');
summary.forEach((s) => console.log('  ' + s));
hr();

await client.close();
process.exit(0);
