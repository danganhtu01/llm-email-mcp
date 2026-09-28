// Tests the write tools (move/send/flag/delete/folders/draft) WITHOUT touching a
// real mailbox: it checks tool registration + input validation against a TEMP
// accounts file, and unit-tests the offline MIME composer directly.
// Run: node test-writetools.mjs
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { writeFileSync, rmSync } from 'fs';
import path from 'path';
import os from 'os';
import { fileURLToPath } from 'url';
import { ImapManager } from './imap-client.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const serverPath = path.join(__dirname, 'index.js');
const tmp = path.join(os.tmpdir(), `writetools-${Date.now()}.json`);
writeFileSync(
  tmp,
  JSON.stringify([
    { name: 'TmpImap', host: 'imap.example.com', port: 993, secure: true, user: 'a@example.com' },
  ])
);

let failures = 0;
function check(label, cond, extra = '') {
  const tag = cond ? 'PASS' : 'FAIL';
  if (!cond) failures++;
  console.log(`[${tag}] ${label} ${extra}`);
}

async function call(c, name, args = {}) {
  const r = await c.callTool({ name, arguments: args });
  let p;
  try { p = JSON.parse(r.content?.[0]?.text ?? ''); } catch { p = r.content?.[0]?.text; }
  return { isError: !!r.isError, p };
}

// ── Part 1: tool registration + validation over the MCP server ────────────────
const t = new StdioClientTransport({
  command: 'node',
  args: [serverPath],
  env: { ...process.env, IMAP_ACCOUNTS_FILE: tmp },
});
const c = new Client({ name: 'tw', version: '1' }, { capabilities: {} });
await c.connect(t);

const tools = (await c.listTools()).tools.map((x) => x.name);
console.log('Tool count:', tools.length);
check('tool count is 24', tools.length === 24, `(got ${tools.length})`);
for (const name of [
  'move_email', 'send_email', 'flag_email', 'delete_email',
  'create_folder', 'delete_folder', 'rename_folder', 'create_draft',
  'mark_read_graph', 'ms365_login_graph',
]) {
  check(`tool registered: ${name}`, tools.includes(name));
}
console.log('─'.repeat(60));

// Validation errors should be raised BEFORE any network call.
let r;
r = await call(c, 'move_email', { account: 'TmpImap', source_folder: 'INBOX', destination_folder: 'X', uids: [] });
check('move_email empty uids rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'send_email', { account: 'TmpImap', to: [] });
check('send_email empty to rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'flag_email', { account: 'TmpImap', folder: 'INBOX', uids: ['1'], flagged: 'yes' });
check('flag_email non-boolean rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'delete_email', { account: 'TmpImap', folder: 'INBOX', uids: [] });
check('delete_email empty uids rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'create_folder', { account: 'TmpImap' });
check('create_folder missing path rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'rename_folder', { account: 'TmpImap', old_path: 'A' });
check('rename_folder missing new_path rejected', r.isError, `→ ${r.p}`);

r = await call(c, 'move_email', { account: 'Nope', source_folder: 'INBOX', destination_folder: 'X', uids: ['1'] });
check('unknown account rejected', r.isError, `→ ${r.p}`);

// mark_read_graph: rejects non-OAuth (password/IMAP) accounts before any network call.
r = await call(c, 'mark_read_graph', { account: 'TmpImap', message_ids: ['AAMkAG'], read: true });
check('mark_read_graph rejects non-M365 account', r.isError, `→ ${r.p}`);

r = await call(c, 'mark_read_graph', { account: 'TmpImap', message_ids: [], read: true });
check('mark_read_graph empty message_ids rejected', r.isError, `→ ${r.p}`);

await c.close();
console.log('─'.repeat(60));

// ── Part 2: offline MIME composer unit test (no IMAP/SMTP needed) ──────────────
const mgr = new ImapManager(
  [{ name: 'TmpImap', host: 'imap.example.com', user: 'a@example.com' }],
  null,
  null
);
const account = { name: 'TmpImap', user: 'a@example.com' };

const composed = await mgr._composeMessage(account, {
  to: ['bob@example.com'],
  cc: ['carol@example.com'],
  bcc: ['dave@example.com'],
  subject: 'Hello there',
  body: 'This is the plain body.',
  attachments: [
    { filename: 'note.txt', content_base64: Buffer.from('hi file').toString('base64'), mime_type: 'text/plain' },
  ],
});
const raw = composed.raw.toString('utf8');
check('compose: From header', /^From: .*a@example\.com/m.test(raw));
check('compose: To header', /^To: .*bob@example\.com/m.test(raw));
check('compose: Cc header', /^Cc: .*carol@example\.com/m.test(raw));
check('compose: Subject header', /^Subject: Hello there/m.test(raw));
check('compose: body present', raw.includes('This is the plain body.'));
check('compose: attachment filename', raw.includes('note.txt'));
// envelope.to should include to + cc + bcc (so the MTA actually delivers bcc).
check('compose: envelope has all 3 recipients', composed.envelope.to.length === 3,
  `→ ${composed.envelope.to.join(', ')}`);

// Empty "to" must throw.
let threw = false;
try { await mgr._composeMessage(account, { to: [], subject: 'x' }); } catch { threw = true; }
check('compose: empty to throws', threw);

// A forward attaches the original as message/rfc822, which RFC 2046 § 5.2.1 allows only
// in 7bit, 8bit or binary. Encoded as base64 (nodemailer's default), Microsoft 365 read the
// base64 text as the message itself and delivered a 25-byte scrap (2026-09-28).
const original = 'From: shop@example.com\r\nSubject: Invoice 42\r\nMIME-Version: 1.0\r\n\r\nTotal: 50.00\r\n';
mgr.withClient = async (_acct, fn) => fn({
  mailboxOpen: async () => {},
  fetchOne: async () => ({ source: Buffer.from(original), envelope: { subject: 'Invoice 42' } }),
});
const fwd = (await mgr._composeMessage(account, { to: ['bob@example.com'], body: 'See attached.', forward_uid: '7' })).raw.toString('utf8');
const part = fwd.slice(fwd.indexOf('Content-Type: message/rfc822'));
check('forward: subject gets Fwd:', /^Subject: Fwd: Invoice 42/m.test(fwd));
check('forward: rfc822 part is 8bit, not base64', /^Content-Transfer-Encoding: 8bit/m.test(part.split('\r\n\r\n')[0]));
check('forward: original carried verbatim', part.includes(original));

rmSync(tmp, { force: true });
console.log('─'.repeat(60));
console.log(failures === 0 ? 'ALL PASS' : `${failures} FAILURE(S)`);
process.exit(failures === 0 ? 0 : 1);
