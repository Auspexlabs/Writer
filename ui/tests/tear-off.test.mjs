// node --test ui/tests/   — a tab dragged off the title strip opens its document in a window of its own: where it has to be
// let go (tabs.js), and the shell saving it before its tab closes (index.dc.html detach).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { tornOff } from '../tabs.js';

test('a tab tears off when let go outside the window or well below the strip, and a single tab never does', () => {
  const at = (x, y, tabs = 2) => tornOff({ x, y }, { bottom: 33 }, { w: 1280, h: 782 }, tabs);
  assert.equal(at(600, 20), false, 'on the strip: a click, or a drag along it');
  assert.equal(at(600, 70), false, 'just below the strip');
  assert.equal(at(600, 90), true, 'over the document');
  assert.equal(at(-4, 20), true, 'past the left edge');
  assert.equal(at(1280, 20), true, 'past the right edge');
  assert.equal(at(600, -2), true, 'above the window');
  assert.equal(at(600, 900), true, 'below the window');
  assert.equal(at(600, 900, 1), false, 'one tab: nothing to separate');
});

/** The shell's logic in a vm, as storage.test.mjs runs it; the engine bridge is faked. */
function shell(docs, save) {
  const code = readFileSync(new URL('../index.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, document: { querySelector: () => null }, $t: s => s,
    React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); } } };
  vm.runInNewContext(code + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell();
  Object.assign(c, { props: { onSaveDialog: async () => null }, EN: { isDraft: p => p.startsWith('/D/'), save, watch: () => () => { } } });
  Object.assign(c.state, { docs, cur: docs[0].id });
  return c;
}

test('detach saves the document first, closes a draft without the 存储 question once it is open elsewhere, keeps it when the save fails', async () => {
  const docs = () => [{ id: 'd', title: '未命名', type: 'md', path: '/D/未命名.md', loaded: true, text: '改过' }, { id: 'a', title: 'a', type: 'md', path: 'a.md', loaded: true, text: '' }];
  const log = [], c = shell(docs(), async d => { log.push('save ' + d.path); return 0; });
  const done = await Promise.race([c.detach('d', async path => { log.push('open ' + path); return true; }), new Promise(r => setTimeout(r, 500, 'asked'))]);
  assert.equal(done, true);
  assert.deepEqual(log, ['save /D/未命名.md', 'open /D/未命名.md']);
  assert.deepEqual(c.state.docs.map(d => d.id), ['a']);
  const f = shell(docs(), async () => { throw new Error('磁盘已满'); });
  assert.equal(await f.detach('d', async () => assert.fail('nothing opens after a failed save')), false);
  assert.deepEqual(f.state.docs.map(d => d.id), ['d', 'a']);
});
