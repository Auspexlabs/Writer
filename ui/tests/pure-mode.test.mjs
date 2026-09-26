// node --test ui/tests/ — 纯净模式: the shell keeps its own panel flags (格式 / 缩略图 / AI 助手 open over the page and start closed
// each time), Esc closes 格式 before it leaves, and Aa controls the format panel for each editor.
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

test('沉浸书写: going in leaves the document alone; Aa and AI open the same panels as outside it; coming out restores them', () => {
  const c = shell(); c.setState({ showThumbs: true, showAI: true, showFormat: false });
  let v = c.vals(); assert.deepEqual([v.pureAttr, v.aiOn, v.panelOpen, v.showThumbs], ['0', true, true, true]);
  c.setPure(true); v = c.vals();
  assert.deepEqual([v.pureAttr, v.pure, v.showThumbs, v.aiOn, v.formatOpen, v.panelOpen], ['1', true, false, false, false, false], 'no panel, no thumbnails');
  c.toggleFormat(); v = c.vals();
  assert.deepEqual([v.formatOpen, v.aiOn, v.panelOpen, v.panelEnter], [true, false, true, 'open'], 'Aa: the format panel slides in');
  c.toggleAIPanel(); v = c.vals();
  assert.deepEqual([v.formatOpen, v.aiOn, v.panelOpen, v.panelEnter], [false, true, true, 'swap'], 'AI takes its place: same card, new content');
  c.setPure(false); v = c.vals();
  assert.deepEqual([v.pure, v.showThumbs, v.aiOn, v.formatOpen], [false, true, true, false], 'coming out: the panels and thumbnails from before');
  c.setPure(true); assert.equal(c.vals().panelOpen, false, 'the next time starts clear again');
});

test('the shell: Esc closes 格式 first, then leaves 沉浸书写; ⌘. toggles; 改写 opens the assistant and sends the selection', () => {
  const c = shell(); c.setPure(true); c.toggleFormat();
  const key = (k, o = {}) => { const e = Object.assign({ key: k, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, preventDefault() { this.prevented = true; } }, o); c.onKey(e); return !!e.prevented; };
  assert.equal(key('Escape'), true); assert.deepEqual([c.state.pure, c.state.showFormat], [true, false]);
  assert.equal(key('Escape'), true); assert.equal(c.state.pure, false);
  assert.equal(key('.', { metaKey: true }), true); assert.equal(c.state.pure, true);
  const sent = []; c.send = t => sent.push(t); c.vals().onRewrite('  这一段  '); c.vals().onRewrite('');
  assert.deepEqual([c.state.showAI, c.state.showFormat], [true, false]); assert.deepEqual(sent, ['改写这段文字，保持原意，让它更通顺简洁：\n这一段']);
});

test('the side panels: one at a time in one place; opening slides in, taking the other\'s place swaps, and the first Aa after a launch with AI open works', () => {
  const c = shell(); c.setState({ showFormat: false, showAI: false });
  let v = c.vals(); assert.deepEqual([v.panelOpen, v.panelW], [false, '304px']);
  c.toggleAIPanel(); v = c.vals(); assert.deepEqual([v.aiOn, v.panelOpen, v.panelEnter], [true, true, 'open']);
  c.toggleFormat(); v = c.vals(); assert.deepEqual([v.formatOpen, v.aiOn, v.panelEnter], [true, false, 'swap']);
  v = c.vals(); assert.equal(v.panelEnter, 'swap', 'sticky while the same panel stays open (a new value would restart its animation)');
  c.toggleFormat(); v = c.vals(); assert.deepEqual([v.panelOpen, v.formatOpen], [false, false]);
  c.toggleFormat(); assert.equal(c.vals().panelEnter, 'open');
  // a launch that restores the assistant: Aa, pressed once, shows the format panel
  const d = shell(); d.setState({ showAI: true, showFormat: true }); d.toggleFormat();
  assert.deepEqual([d.state.showFormat, d.state.showAI], [true, false]);
});

test('the Word editor: the selection bar comes with a right-click, at the pointer, in 沉浸书写 as outside it; a press elsewhere or typing takes it away', () => {
  const ctx = base(); ctx.getComputedStyle = () => ({ marginBottom: '28px' }); ctx.window.getSelection = () => ({ rangeCount: 0 });
  vm.runInNewContext(script('WordEditor.dc.html') + '\nglobalThis.WordEditor = Component;', ctx);
  const make = props => { const c = new ctx.WordEditor(); c.props = Object.assign({ doc: { id: 'd', html: '' }, onChange() { } }, props); c.pgCss = { textContent: '' }; c.edRef.current = { innerText: '', children: [], querySelectorAll: () => [], contains: () => true, closest: () => null }; c.refreshInfo(); return c; };
  const right = c => { const e = { target: {}, clientX: 400, clientY: 300, preventDefault() { this.prevented = true; } }; c.onCtx(e); return e.prevented; };
  for (const pure of [false, true]) {
    const c = make({ pure, onRewrite() { } });
    assert.equal(c.renderVals().hasWordSelection, false, 'no bar for a selection by itself');
    assert.equal(right(c), true, 'the system menu gives way');
    let v = c.renderVals(); assert.deepEqual([v.hasWordSelection, v.bx, v.by, v.selBelow], [true, '400px', '288px', '0'], 'the bar above the pointer');
    c.ctxAway({ button: 0, target: { closest: () => null } }); assert.equal(c.renderVals().hasWordSelection, false, 'a press elsewhere');
    right(c); c.ctxAway({ button: 0, target: { closest: s => /data-word-selection/.test(s) ? {} : null } }); assert.equal(c.renderVals().hasWordSelection, true, 'a press on the bar keeps it');
    c.ctxKey({ key: 'b', metaKey: true }); assert.equal(c.renderVals().hasWordSelection, true, '⌘B keeps it');
    c.ctxKey({ key: 'a' }); assert.equal(c.renderVals().hasWordSelection, false, 'typing');
  }
  const v = make({ pure: false, showThumbs: true, formatOpen: true }).renderVals();
  assert.deepEqual([v.showNav, v.formatOpen, v.bodyCols], [true, true, '240px minmax(0,1fr) 312px'], 'the sidebar and the panel take their columns (format-panel.test.mjs has the panel)');
});

test('Word format and AI share one right panel', () => {
  const c = shell();
  assert.equal(c.state.showFormat, true);
  c.toggleAIPanel(); assert.deepEqual([c.state.showAI, c.state.showFormat], [true, false]);
  c.toggleFormat(); assert.deepEqual([c.state.showAI, c.state.showFormat], [false, true]);
  c.toggleFormat(); assert.equal(c.state.showFormat, false);
});

test('Aa controls the shared format panel for every document type', () => {
  const c = shell();
  for (const type of ['docx', 'xlsx', 'pptx', 'md', 'mm', 'pdf']) {
    Object.assign(c.state.docs[0], { type, sheets: [], slides: [], pages: [], map: { text: 'Root', children: [] } });
    c.setState({ showFormat: true, showAI: false });
    assert.equal(c.vals().formatOpen, true, `${type}: panel opens`);
    c.toggleFormat();
    assert.equal(c.vals().formatOpen, false, `${type}: Aa closes it`);
    c.toggleFormat();
    c.toggleAIPanel();
    assert.deepEqual([c.vals().formatOpen, c.vals().aiOn], [false, true], `${type}: AI replaces it`);
    c.toggleAIPanel();
  }
});

test('the title bar\'s tabs follow the order the documents were opened in', async () => {
  const c = shell(); let info; c.props.onShell = i => { info = i; };
  c.state.docs = [{ id: 'x', type: 'xlsx', title: '预算', path: '预算.xlsx' }, { id: 'a', type: 'docx', title: '方案', path: '方案.docx' }, { id: 'p', type: 'pptx', title: '路演', path: '路演.pptx' }];
  c.EN = { open: async d => Object.assign({}, d, { loaded: true }), isDraft: () => false, watch: () => () => { } }; c.toastMsg = () => { }; c.checkExternal = () => { };
  for (const id of ['a', 'p', 'x']) await c.open(id);
  c.emitShell();
  assert.deepEqual(info.docs.filter(d => d.open).sort((a, b) => a.seq - b.seq).map(d => d.id), ['a', 'p', 'x'], 'not the folder\'s order');
});

test('every open document keeps its editor: one shown, the others hidden with the props they last had, six at most, in the documents\' order', () => {
  const c = shell();
  c.state.docs = 'abcdefgh'.split('').map((id, i) => ({ id, type: ['docx', 'xlsx', 'pptx', 'md', 'mm', 'pdf', 'docx', 'md'][i], title: id, path: id, loaded: true, html: '', sheets: [], slides: [], pages: [], map: { text: 'Root', children: [] } }));
  const show = id => { c.setState({ cur: id, view: 'doc' }); return c.renderVals(); };
  let v; for (const id of 'abcdefgh') v = show(id);
  assert.deepEqual(Array.from(v.live, L => L.id), ['c', 'd', 'e', 'f', 'g', 'h'], 'the six used last, in the documents\' order (a and b let go)');
  assert.deepEqual(Array.from(v.live.filter(L => L.onA === '1'), L => L.id), ['h'], 'one shown');
  assert.ok(v.live.filter(L => L.onA !== '1').every(L => /content-visibility:hidden|visibility:hidden/.test(L.vis)), 'the others hidden');
  assert.deepEqual([v.live.find(L => L.id === 'g').isDocx, v.live.find(L => L.id === 'h').isMd], [true, true]);
  // a panel opened over h: g, hidden, keeps what it had; shown again, it takes the current values
  c.setState({ showFormat: false, showAI: false }); v = c.renderVals();
  assert.equal(v.live.find(L => L.id === 'g').formatOpen, true, 'g still has the panel it was last shown with');
  assert.equal(v.live.find(L => L.id === 'h').formatOpen, false);
  v = show('g'); assert.equal(v.live.find(L => L.id === 'g').formatOpen, false, 'shown again: the current state');
  c.setState({ docs: c.state.docs.map(d => d.id === 'e' ? Object.assign({}, d, { loaded: false }) : d) }); v = c.renderVals();
  assert.ok(!Array.from(v.live, L => L.id).includes('e'), 'a closed document lets its editor go');
});
