// node --test ui/tests/ — the embed page's side of the protocol (ui/embed/host.js): hello, the site's documents going in, ready
// once the editor shows one, calls answered only for the site that sent init, change and open events, the assistant's model
// calls sent to the site. The page is embed-page.mjs's: an engine stand-in, the shell played by the test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { page, SITE } from './embed-page.mjs';

const { E, posted, say, until, find, embed, api, shown, call } = await page({ id: 'w1', mode: 'docx', ai: true, aiModel: 'm', blankName: 'Draft' });

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
