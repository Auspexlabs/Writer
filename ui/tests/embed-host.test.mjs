// node --test ui/tests/ — the embed page's side of the protocol (ui/embed/host.js): hello, the site's documents going in, ready
// once the editor shows one, calls answered only for the site that sent init, change and open events, the assistant's model
// calls sent to the site. The engine is a stand-in (as in embed-server.test.mjs); the shell is played by the test.
import { test } from 'node:test';
import assert from 'node:assert/strict';

// the embed page in an iframe: its window, the site around it (parent), the options embed.dc.html reads from the URL
const bus = new EventTarget(), posted = [];
globalThis.window = globalThis;
globalThis.addEventListener = bus.addEventListener.bind(bus);
globalThis.removeEventListener = bus.removeEventListener.bind(bus);
globalThis.dispatchEvent = bus.dispatchEvent.bind(bus);
const parent = { postMessage(msg, origin, transfer) { posted.push({ msg: structuredClone(msg), origin, transfer }); } };
globalThis.parent = parent;
globalThis.location = new URL('https://cdn.example/writer/embed.dc.html?id=w1&mode=docx&ai=1&model=m');
globalThis.fetch = async u => new Response('static ' + u);
globalThis.EventSource = class { };
URL.createObjectURL = () => 'blob:x';
URL.revokeObjectURL = () => { };
window.__WRITER_EMBED__ = { id: 'w1', mode: 'docx', ai: true, aiModel: 'm', blankName: 'Draft' };

function fakeEngine() {
  const files = new Map(), changed = [];
  const bump = p => { if (!changed.includes(p)) changed.push(p); };
  return {
    files, calls: [],
    Run(body) { const { argv, command } = JSON.parse(body); const a = argv || command.split(' '); this.calls.push(a); if (a[0] === 'create') { files.set(a[1], new Uint8Array([1])); bump(a[1]); } if (a[0] === 'add') bump(a[1]); return JSON.stringify({ code: 0, output: 'ran ' + a.join(' ') }); },
    Files() { return JSON.stringify({ files: [...files.keys()].map(p => ({ path: p })) }); },
    Stat(p) { return JSON.stringify({ path: p, mtime: 1, size: files.get(p).length }); },
    Read(p) { if (!files.has(p)) throw new Error(JSON.stringify({ status: 404, error: { code: 'FILE_NOT_FOUND', message: p + ' not found' } })); return files.get(p); },
    Write(p, bytes) { files.set(p, bytes); bump(p); return JSON.stringify({ path: p, size: bytes.length }); },
    Json(p) { return JSON.stringify({ type: 'document', path: '/', file: p }); },
    Drain() { const out = JSON.stringify(changed); changed.length = 0; return out; },
    ChatSystem(p) { return 'SYSTEM for ' + p; },
    ChatTools() { return JSON.stringify([{ name: 'writer', description: 'd', input_schema: {} }]); },
    ChatTool(name, input) { this.calls.push([name, JSON.parse(input)]); return JSON.stringify({ display: 'writer ' + JSON.parse(input).command, code: 0, output: 'done', wrote: true }); }
  };
}
const S = await import('../embed/server.js');
const E = S.useEngine(fakeEngine());
await import('../embed/host.js');

const SITE = 'https://site.example';
/** A message to the page: from the site by default. */
const say = (data, from = {}) => dispatchEvent(Object.assign(new Event('message'), { data: Object.assign({ writer: 1, id: 'w1' }, data), origin: from.origin || SITE, source: 'source' in from ? from.source : parent }));
const until = async (what, ms = 3000) => { for (const end = Date.now() + ms; Date.now() < end; await new Promise(r => setTimeout(r, 5))) { const x = what(); if (x) return x; } throw new Error('timed out'); };
const find = (type, more = () => true) => posted.find(p => p.msg.type === type && more(p.msg));
/** The shell (index.dc.html) as the page reports it: open a document, and it becomes the one shown. */
const embed = window.__WRITER_EMBED__, shown = [];
const api = { async openPath(name) { shown.push(name); setTimeout(() => embed.shell({ path: name, isDoc: true }, api), 5); } };
let seq = 0;
const call = async (method, ...args) => { const n = ++seq; say({ type: 'call', seq: n, method, args }); const r = await until(() => find('result', m => m.seq === n)); return r; };

test('hello, then the site\'s documents go in quietly and ready comes once the editor shows the one to open', async () => {
  assert.deepEqual(posted[0], { msg: { writer: 1, id: 'w1', type: 'hello' }, origin: '*', transfer: [] }, 'hello before the page knows who embedded it');
  say({ type: 'init', files: [{ name: 'a.docx', data: new Uint8Array([80, 75]).buffer }, { name: 'b.md', data: new TextEncoder().encode('# B').buffer }], open: 'b.md' }, { origin: 'https://evil.example', source: {} });
  assert.equal(E.files.size, 0, 'an init from another window is not the site');
  say({ type: 'init', files: [{ name: 'a.docx', data: new Uint8Array([80, 75]).buffer }, { name: 'b.md', data: new TextEncoder().encode('# B').buffer }], open: 'b.md' });
  await until(() => E.files.size === 2);
  embed.shell({ path: '', isDoc: false }, api);
  const ready = await until(() => find('ready'));
  assert.deepEqual([ready.msg.file, ready.origin, shown], ['b.md', SITE, ['b.md']], 'answers go to the site\'s origin only');
  assert.deepEqual(posted.filter(p => p.msg.type === 'event' && p.msg.name === 'change'), [], 'the documents the site gave are no change to tell it about');
  assert.deepEqual(find('event', m => m.name === 'open').msg.data, { file: 'b.md' });
  say({ type: 'init', files: [{ name: 'c.docx', data: new Uint8Array([1]).buffer }] });
  await new Promise(r => setTimeout(r, 20));
  assert.equal(E.files.has('c.docx'), false, 'init once');
});

test('calls: run, put, get, list, tree, current, open, theme; an unknown method fails; others cannot call', async () => {
  const run = await call('run', ['add', 'a.docx', '/body']);
  assert.deepEqual(run.msg, { writer: 1, id: 'w1', type: 'result', seq: run.msg.seq, ok: true, value: { code: 0, output: 'ran add a.docx /body' } });
  const change = await until(() => find('event', m => m.name === 'change'));
  assert.deepEqual(change.msg.data, { files: ['a.docx'], from: 'api' }, 'the site hears its own change, and that it made it');
  await call('put', 'n.md', 'text');
  assert.equal(new TextDecoder().decode(E.files.get('n.md')), 'text');
  const got = await call('get', 'a.docx');
  assert.deepEqual([[...new Uint8Array(got.msg.value)], got.transfer.length], [[80, 75], 1], 'bytes moved to the site');
  assert.deepEqual((await call('list')).msg.value.map(f => f.path), ['a.docx', 'b.md', 'n.md']);
  assert.deepEqual((await call('tree', 'a.docx')).msg.value, { type: 'document', path: '/', file: 'a.docx' });
  assert.equal((await call('current')).msg.value, 'b.md');
  assert.equal((await call('open', 'a.docx')).msg.value, true);
  assert.equal((await call('current')).msg.value, 'a.docx');
  let settings = 0; addEventListener('writer-settings', () => settings++);
  assert.equal((await call('theme', 'dark')).msg.value, true);
  assert.deepEqual([window.__WRITER_PREFS.theme, settings], ['dark', 1]);
  const bad = await call('rm', 'a.docx');
  assert.deepEqual([bad.msg.ok, bad.msg.error, bad.msg.code], [false, 'Unknown method rm', '']);
  const missing = await call('get', 'nope.docx');
  assert.deepEqual([missing.msg.ok, missing.msg.error, missing.msg.code], [false, 'nope.docx not found', 'FILE_NOT_FOUND'], 'the engine\'s error, readable');
  const before = posted.length;
  say({ type: 'call', seq: 99, method: 'list', args: [] }, { origin: 'https://evil.example' });
  await new Promise(r => setTimeout(r, 20));
  assert.equal(posted.length, before, 'a call from another origin gets nothing');
});

test('the assistant: /chat in the editor runs the loop here, each model call goes to the site and its answer comes back', async () => {
  const files = await (await fetch('https://cdn.example/files')).json();
  assert.deepEqual([files.chat, files.model], [true, 'm'], 'the editor sees an assistant, with the model the site named');
  const res = fetch('https://cdn.example/chat', { method: 'POST', body: JSON.stringify({ file: 'a.docx', messages: [{ role: 'user', content: '加一段' }] }) });
  const first = await until(() => find('chat', m => m.seq === 1));
  assert.equal(first.origin, SITE);
  assert.deepEqual([first.msg.request.system, first.msg.request.tools.map(t => t.name), first.msg.request.messages], ['SYSTEM for a.docx', ['writer'], [{ role: 'user', content: '加一段' }]]);
  say({ type: 'chat-event', seq: 1, name: 'delta', data: '好' });
  say({ type: 'chat-event', seq: 1, name: 'done', data: { text: '好', toolCalls: [{ id: 't', name: 'writer', input: { command: 'add a.docx /body' } }] } });
  const second = await until(() => find('chat', m => m.seq === 2));
  assert.deepEqual(second.msg.request.messages.at(-1), { role: 'tool', toolCallId: 't', name: 'writer', content: 'done', isError: false });
  say({ type: 'chat-event', seq: 2, name: 'done', data: { text: '好了', toolCalls: [] } });
  const text = await (await res).text();
  assert.deepEqual(text.split('\n\n').filter(Boolean).map(b => b.split('\n')[0]), ['event: delta', 'event: text', 'event: tool', 'event: text', 'event: done']);
  assert.deepEqual(E.calls.at(-1), ['writer', { command: 'add a.docx /body' }]);
  const ac = new AbortController();
  const stopped = fetch('https://cdn.example/chat', { method: 'POST', body: JSON.stringify({ file: 'a.docx', messages: [{ role: 'user', content: '再来' }] }), signal: ac.signal });
  await until(() => find('chat', m => m.seq === 3));
  ac.abort();
  await until(() => find('chat-abort', m => m.seq === 3));
  assert.equal(await (await stopped).text(), '', 'stopped: nothing more, no error');
});
