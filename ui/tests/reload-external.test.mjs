// node --test ui/tests/   — reloading a document changed outside Writer, and keeping local edits when that races a save
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const EN = await import('../engine.js');

/** fetch stub: answers /files and /files?drafts=1, records every request. Same shape as storage.test.mjs's. */
function stub({ workspace, drafts, files = [], draftFiles = [] }) {
  const calls = [];
  globalThis.fetch = async (url, opts = {}) => {
    calls.push({ url, method: opts.method || 'GET', body: opts.body });
    const json = x => new Response(JSON.stringify(x), { status: 200, headers: { 'Content-Type': 'application/json' } });
    if (url === '/files') return json({ workspace, drafts, files: files.map(path => ({ path, format: path.split('.').pop() })) });
    if (url === '/files?drafts=1') return json({ workspace, drafts, files: draftFiles.map(path => ({ path, format: path.split('.').pop() })) });
    if (url === '/run') return json({ code: 0, output: '' });
    return json({ path: 'x', size: 0 });
  };
  return calls;
}

/** The shell's logic in a vm over the real bridge; only EN is faked (per-test). */
function shell(docs, cur) {
  const code = readFileSync(new URL('../index.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, document: { querySelector: () => null, visibilityState: 'visible' },
    $t: (s, v) => v ? String(s).replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : s, // i18n.js's $t, always in Chinese here (no dictionary loaded)
    React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); } } };
  vm.runInNewContext(code + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell();
  Object.assign(c, { props: {}, EN: Object.assign({}, EN, { save: async () => 0, watch: () => () => { } }) });
  Object.assign(c.state, { docs, cur });
  return c;
}

test('reloadDecision: reload only when nothing local is pending, conflict when it is, none when unset or unchanged', () => {
  assert.equal(EN.reloadDecision(null, 100, false), 'none', 'no tracked mtime yet: nothing to compare');
  assert.equal(EN.reloadDecision(100, 100, false), 'none', 'unchanged');
  assert.equal(EN.reloadDecision(100, 100, true), 'none', 'unchanged even with something pending');
  assert.equal(EN.reloadDecision(100, 200, false), 'reload', 'changed, nothing local pending: safe to load in place');
  assert.equal(EN.reloadDecision(100, 200, true), 'conflict', 'changed while local edits are pending: keep both');
});

test('saveConflictCopy: names the copy beside the file, falling back to the drafts folder when that is refused', async () => {
  let calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' });
  await EN.files();
  const path1 = await EN.saveConflictCopy({ path: 'a.md', type: 'md', text: 'local edits' });
  assert.equal(path1, 'a（冲突副本）.md');
  const put1 = calls.find(c => c.method === 'PUT');
  assert.equal(put1.url, '/file?file=' + encodeURIComponent('a（冲突副本）.md'));
  assert.equal(put1.body, 'local edits', 'the copy holds the live (unsaved) text, not what was last read from disk');

  calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' });
  await EN.files();
  let puts = 0;
  const realFetch = globalThis.fetch;
  globalThis.fetch = async (url, opts = {}) => {
    if ((opts.method || 'GET') === 'PUT' && ++puts === 1) return new Response('{"error":{"code":"VALIDATION"}}', { status: 400 });
    return realFetch(url, opts);
  };
  const path2 = await EN.saveConflictCopy({ path: '/U/Docs/b.md', type: 'md', text: 'local edits' });
  assert.equal(path2, '/U/app/Drafts/b（冲突副本）.md', 'the sibling location was refused, so the copy went to drafts instead');
});

test('the shell: a save does not make its own poll reload the document (mtime matches what the save returned)', async () => {
  stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' });
  await EN.files();
  const doc = { id: 'd1', title: 'a', type: 'md', path: 'a.md', loaded: true, text: 'v2', _orig: 'v1', _mtime: 100 };
  const c = shell([doc], 'd1');
  let statCalls = 0;
  c.EN = Object.assign({}, EN, {
    save: async d => { d._mtime = 200; d._orig = d.text; return 1; }, // what saveNow really does: writes, then remembers its own new mtime
    stat: async () => { statCalls++; return { mtime: 200 }; },
    watch: () => () => { },
  });
  await c.flushSave('d1');
  assert.equal(c.state.docs[0]._mtime, 200, 'adopt carried the new mtime onto the live doc');

  await c.checkExternal('d1');
  assert.equal(statCalls, 1);
  assert.equal(c.state.toast, null, 'no reload toast: the disk mtime is exactly the one this save just produced');
  assert.equal(c.state.docs[0].text, 'v2', 'the document in the editor was left alone');
});
