// node --test ui/tests/ — 纯净模式: the shell keeps its own panel flags (格式 / 缩略图 / AI 助手 open over the page and start closed
// each time), Esc closes 格式 before it leaves, and the Word editor shows its toolbar over the page only while the shell asks.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const script = f => readFileSync(new URL('../' + f, import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const base = () => ({ location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, getComputedStyle: () => ({}), window: { innerWidth: 1200 },
  document: { querySelector: () => null, querySelectorAll: () => [], documentElement: { dataset: {} }, activeElement: null }, localStorage: { getItem: () => null, setItem() { } },
  $t: (s, v) => { s = String(s).replace(/@@.*$/, ''); return v ? s.replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : s; },
  React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } });

function shell() {
  const ctx = base(); vm.runInNewContext(script('index.dc.html') + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell(); c.props = { mac: true }; c.prefs = {};
  Object.assign(c.state, { ready: true, view: 'doc', cur: 'a', docs: [{ id: 'a', type: 'docx', title: '报告', path: 'sub/报告.docx', loaded: true, html: '' }] });
  return c;
}

test('the shell: pure mode starts with every panel closed; 格式 / 缩略图 / AI 助手 open over the page without moving it', () => {
  const c = shell(); c.setState({ showThumbs: true, showAI: true });
  let v = c.vals(); assert.equal(v.pureAttr, '0'); assert.equal(v.showThumbs, true); assert.equal(v.showPanel, true); assert.notEqual(v.edRight, '0px');
  c.setPure(true); v = c.vals();
  assert.deepEqual([v.pureAttr, v.pure, v.pureName, v.showThumbs, v.showPanel, v.edRight, v.pTools], ['1', true, '报告.docx', false, false, '0px', false], 'the sidebar and panel of the normal view stay out of it');
  v.toggleThumbs(); v.toggleAI(); v.toggleTools(); v = c.vals();
  assert.deepEqual([v.showThumbs, v.showPanel, v.edRight, v.pTools, v.aiOn], [true, true, '0px', true, true], 'the panels overlay: the editor keeps the whole width');
  c.setPure(false); v = c.vals();
  assert.deepEqual([v.showThumbs, v.showPanel, c.state.pTools, c.state.pThumbs, c.state.pAI], [true, true, false, false, false], 'leaving restores the normal view and forgets the overlays');
  c.setPure(true); assert.equal(c.vals().showPanel, false, 'the next time starts closed again');
});

test('the shell: Esc closes 格式 first, then leaves; ⌘. toggles; 改写 opens the assistant and sends the selection', () => {
  const c = shell(); c.setPure(true); c.setState({ pTools: true });
  const key = (k, o = {}) => { const e = Object.assign({ key: k, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, preventDefault() { this.prevented = true; } }, o); c.onKey(e); return !!e.prevented; };
  assert.equal(key('Escape'), true); assert.deepEqual([c.state.pure, c.state.pTools], [true, false]);
  assert.equal(key('Escape'), true); assert.equal(c.state.pure, false);
  assert.equal(key('.', { metaKey: true }), true); assert.equal(c.state.pure, true);
  const sent = []; c.send = t => sent.push(t); c.vals().onRewrite('  这一段  '); c.vals().onRewrite('');
  assert.equal(c.state.pAI, true); assert.deepEqual(sent, ['改写这段文字，保持原意，让它更通顺简洁：\n这一段']);
});

test('the Word editor: in pure mode the toolbar shows only for the shell\'s 格式, thumbnails overlay without a column, the bar ends in ✦ 改写', () => {
  const ctx = base(); ctx.getComputedStyle = () => ({ marginBottom: '28px' });
  vm.runInNewContext(script('WordEditor.dc.html') + '\nglobalThis.WordEditor = Component;', ctx);
  const ed = props => { const c = new ctx.WordEditor(); c.props = Object.assign({ doc: { id: 'd', html: '' }, onChange() { } }, props); c.pgCss = { textContent: '' }; c.edRef.current = { innerText: '', children: [], querySelectorAll: () => [] }; c.refreshInfo(); return c.renderVals(); };
  let v = ed({ pure: true, showThumbs: true, pureTools: false });
  assert.deepEqual([v.showBar, v.notPure, v.showNav, v.bodyCols.startsWith('0px'), v.rootBg], [false, false, true, true, 'transparent']);
  v = ed({ pure: true, showThumbs: false, pureTools: true, onRewrite() { } });
  assert.deepEqual([v.showBar, v.showNav], [true, false]);
  const labels = btns => Array.from(btns, b => b.sep ? '|' : b.label); // arrays from the editor's realm, compared by value
  assert.deepEqual(labels(v.bubbleBtns), ['B', 'I', 'U', '|', '标题', '|', '链接', '|', '改写']);
  assert.equal(v.bubbleBtns[v.bubbleBtns.length - 1].star, true);
  assert.deepEqual(labels(ed({ pure: true }).bubbleBtns), ['B', 'I', 'U', '|', '标题', '|', '链接'], 'no 改写 without an assistant to send to');
  v = ed({ pure: false, showThumbs: true });
  assert.deepEqual([v.showBar, v.showNav, v.bodyCols.startsWith('224px')], [true, true, true], 'the normal view is as before');
});
