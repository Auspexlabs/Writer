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

test('dragging a tab selects no text: the press is cancelled and selectstart / dragstart are held off until the tab is let go', async () => {
  const { dragTab } = await import('../tabs.js');
  const L = () => { const m = new Map(); return { m, addEventListener: (t, f) => m.set(t, f), removeEventListener: (t, f) => { if (m.get(t) === f) m.delete(t); } }; };
  const win = L(), doc = L(); Object.assign(globalThis, { window: win, document: doc, innerWidth: 1280, innerHeight: 800 });
  let prevented = false;
  const el = { parentElement: { getBoundingClientRect: () => ({ bottom: 40 }) }, setPointerCapture() { }, addEventListener() { }, removeEventListener() { } };
  dragTab({ button: 0, target: { closest: () => null }, currentTarget: el, clientX: 10, clientY: 10, pointerId: 1, preventDefault() { prevented = true; } }, 2, () => { }, () => { });
  assert.equal(prevented, true, 'the press starts no selection');
  const sel = { prevented: false, preventDefault() { this.prevented = true; } };
  doc.m.get('selectstart')(sel); assert.equal(sel.prevented, true, 'no selection while it moves');
  assert.ok(doc.m.has('dragstart'));
  win.m.get('pointerup')({ type: 'pointerup', clientX: 12, clientY: 12 });
  assert.deepEqual([doc.m.has('selectstart'), doc.m.has('dragstart'), win.m.has('pointermove')], [false, false, false], 'all taken off when it is let go');
  for (const page of ['mac.dc.html', 'win.dc.html']) { // the tab icon is a background: nothing to drag, and the page's raw template asks for no image at "{{ t.icon }}"
    const src = readFileSync(new URL('../' + page, import.meta.url), 'utf8');
    assert.match(src, /<span aria-hidden="true" style="[^"]*background:url\('\{\{ t\.icon \}\}'\)/, page);
    assert.doesNotMatch(src.slice(src.indexOf('<x-dc>'), src.indexOf('</x-dc>')), /<img[^>]*src="\{\{/, page + ': no <img src="{{ … }}"> in the page itself');
  }
});

test('with the desktop bridge: a window\'s only tab carries its window and moves into the window whose tab strip it is let go on; elsewhere a tab tears off as before', async () => {
  const { dragTab } = await import('../tabs.js');
  const L = () => { const m = new Map(); return { m, addEventListener: (t, f) => m.set(t, f), removeEventListener: (t, f) => { if (m.get(t) === f) m.delete(t); } }; };
  const win = L(), doc = L(); Object.assign(globalThis, { window: win, document: doc, innerWidth: 1280, innerHeight: 800, requestAnimationFrame: f => { f(); return 1; } });
  const el = { parentElement: { getBoundingClientRect: () => ({ bottom: 40 }) }, setPointerCapture() { }, addEventListener() { }, removeEventListener() { } };
  const press = () => ({ button: 0, target: { closest: () => null }, currentTarget: el, clientX: 10, clientY: 10, pointerId: 1, preventDefault() { } });
  const calls = [], dock = target => ({ start: carry => calls.push('start ' + carry), move: () => calls.push('move'), end: async () => { calls.push('end'); return target; }, to: t => calls.push('to ' + t) });
  const moves = [];
  dragTab(press(), 1, (dx, dy) => moves.push([dx, dy]), () => calls.push('tear'), dock('main-2'));
  win.m.get('pointermove')({ clientX: 60, clientY: 40 });
  await win.m.get('pointerup')({ type: 'pointerup', clientX: 60, clientY: 40 });
  assert.deepEqual(calls, ['start true', 'move', 'end', 'to main-2'], 'the window went along and its tab went into main-2');
  assert.deepEqual(moves, [[0, 0]], 'a carried window moves, not its tab');
  calls.length = 0; dragTab(press(), 2, () => { }, () => calls.push('tear'), dock(null));
  await win.m.get('pointerup')({ type: 'pointerup', clientX: 600, clientY: 900 });
  assert.deepEqual(calls, ['start false', 'end', 'tear'], 'no strip under the pointer: a window of its own');
  calls.length = 0; dragTab(press(), 2, () => { }, () => calls.push('tear'), dock('main-3'));
  await win.m.get('pointerup')({ type: 'pointerup', clientX: 600, clientY: 900 });
  assert.deepEqual(calls, ['start false', 'end', 'to main-3'], 'over another window\'s strip: into it, not a new window');
  for (const page of ['mac.dc.html', 'win.dc.html']) {
    const src = readFileSync(new URL('../' + page, import.meta.url), 'utf8');
    assert.match(src, /this\.tearOff\(x\.id\), this\.dockBridge\(x\.id\)\)/, page + ': the tabs use the bridge');
    assert.match(src, /window\.__writerDockHint = on =>/, page + ': the strip lights up for a tab from another window');
  }
});
