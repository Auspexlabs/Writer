// node --test ui/tests/ — the mind map editor's own logic without a browser: keys add and remove topics, drops reorder,
// reparent, float and copy, pasted outlines become branches, markers toggle, and the topic popup writes labels and links.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import * as M from '../mindmap.js';

const html = readFileSync(new URL('../MindMapEditor.dc.html', import.meta.url), 'utf8');
const at = html.search(/<script type="text\/x-dc" data-dc-script/), start = html.indexOf("'>", at) + 2; // data-props holds '>' characters
const script = html.slice(start, html.indexOf('</script>', start));
const opened = [];
const ctx = { React: { createRef: () => ({ current: null }) }, navigator: { platform: 'MacIntel' }, window: { open: u => opened.push(u) }, performance, setTimeout, clearTimeout, structuredClone,
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), // i18n.js stand-in: Chinese passthrough, placeholders filled
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(script + '\nglobalThis.MindMapEditor = Component;', ctx);

const T = (id, text, children = [], extra = {}) => Object.assign({ id, text, children }, extra);
const map = () => T('R', 'Root', [T('A', 'Alpha', [T('A1', 'a one'), T('A2', 'a two')]), T('B', 'Beta'), T('C', 'Gamma', [T('C1', 'c one')])]);
/** An editor over `map` the way the shell mounts it: every change comes back as the new doc. */
function editor(m = map()) {
  const ed = new ctx.MindMapEditor(); ed.M = M; ed.measure = (t, s) => M.estimate(t) * s.fs; ed.toasts = [];
  ed.props = { doc: { id: 'd', type: 'mm', title: 'm', path: 'm.mm', loaded: true, map: m }, onChange: d => { ed.props.doc = d; }, toast: t => ed.toasts.push(t) };
  ed.state.box = { w: 1200, h: 800 }; ed.canvasRef.current = { getBoundingClientRect: () => ({ left: 0, top: 0 }) };
  return ed;
}
const key = (k, o = {}) => Object.assign({ key: k, keyCode: 0, nativeEvent: {}, preventDefault() {}, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false }, o);
const texts = n => n.children.map(c => c.text);
/** Drops the topics `ids` where the pointer (map coordinates) is, the way mouseup does after a drag. */
function drop(ed, ids, p, alt) {
  const t = M.dropAt(ed.L(), ed.map, p, ids, { dx: 0, dy: 0 });
  ed.op = { kind: 'node', id: ids[0], ids, moved: true }; ed.state.drag = { ids, target: t }; ed.onWU({ altKey: !!alt });
  return t.kind;
}
const box = (ed, id) => ed.L().nodes.find(n => n.id === id);

test('Tab adds a child in edit mode, Enter commits then adds a sibling, Backspace removes it', () => {
  const ed = editor(); ed.select('A');
  ed.onKey(key('Tab'));
  const a = ed.map.children[0]; assert.equal(a.children.length, 3); assert.equal(ed.state.editing, a.children[2].id);
  ed.state.editText = 'hello'; ed.onKey(key('Enter'));
  assert.equal(ed.map.children[0].children[2].text, 'hello'); assert.equal(ed.state.editing, null);
  ed.onKey(key('Enter')); assert.equal(texts(ed.map.children[0]).length, 4); assert.equal(ed.map.children[0].children[3].text, '新主题');
  ed.onKey(key('Escape')); ed.onKey(key('Backspace'));
  assert.deepEqual(texts(ed.map.children[0]), ['a one', 'a two', 'hello']); assert.equal(ed.selId, 'hello' && ed.map.children[0].children[2].id);
  ed.select('R'); ed.onKey(key('Backspace')); assert.equal(ed.toasts.pop(), '中心主题不能删除');
  ed.onKey(key('Enter', { metaKey: true })); assert.equal(ed.toasts.pop(), '中心主题不能插入父主题');
  ed.select('B'); ed.onKey(key('Enter', { metaKey: true })); assert.equal(ed.map.children[1].children[0].id, 'B');
});

test('drops: between siblings reorders, onto a topic reparents, on empty canvas floats, ⌥ leaves a copy', () => {
  const ed = editor(); let N = Object.fromEntries(ed.L().nodes.map(n => [n.id, n]));
  assert.equal(drop(ed, ['B'], { x: N.A1.x + 5, y: (N.A1.y + N.A1.h + N.A2.y) / 2 }), 'between');
  assert.deepEqual(texts(ed.map.children[0]), ['a one', 'Beta', 'a two']); assert.equal(ed.selIds.join(), 'B');
  N = Object.fromEntries(ed.L().nodes.map(n => [n.id, n]));
  assert.equal(drop(ed, ['A1'], { x: N.C.x + 3, y: N.C.y + 3 }), 'into');
  assert.deepEqual(texts(ed.map.children[1]), ['c one', 'a one']); assert.deepEqual(texts(ed.map.children[0]), ['Beta', 'a two']);
  assert.equal(drop(ed, ['A2'], { x: 700, y: 400 }), 'free');
  const a2 = ed.map.children.find(c => c.id === 'A2'); assert.equal(a2.free, '700,400'); assert.equal(box(ed, 'A2').x, 700);
  assert.equal(drop(ed, ['A2'], { x: 500, y: 100 }), 'free'); assert.equal(ed.map.children.find(c => c.id === 'A2').free, '500,100'); assert.equal(ed.map.children.length, 3);
  N = Object.fromEntries(ed.L().nodes.map(n => [n.id, n]));
  assert.equal(drop(ed, ['A2'], { x: N.C1.x + 2, y: N.C1.y + 2 }), 'into');
  assert.equal(ed.map.children.length, 2); const c1 = M.find(ed.map, 'C1'); assert.equal(c1.children[0].id, 'A2'); assert.equal(c1.children[0].free, undefined);
  N = Object.fromEntries(ed.L().nodes.map(n => [n.id, n]));
  assert.equal(drop(ed, ['C1'], { x: N.A.x + 2, y: N.A.y + 2 }, true), 'into');
  assert.ok(M.find(ed.map, 'C1') && M.find(ed.map, 'C').children.length === 2, 'the original stays');
  const copy = M.find(ed.map, 'A').children.at(-1); assert.equal(copy.text, 'c one'); assert.notEqual(copy.id, 'C1'); assert.equal(copy.children[0].text, 'a two');
  const k = box(ed, copy.id); assert.equal(drop(ed, ['A'], { x: k.x + 2, y: k.y + 2 }), 'free', 'nothing lands inside its own branch');
  assert.equal(drop(ed, ['R'], { x: 0, y: 900 }), 'free'); assert.equal(ed.map.id, 'R'); assert.equal(ed.map.free, undefined, 'the centre never moves');
});

test('a pasted outline becomes branches under the selection; copies travel whole', () => {
  const ed = editor(); ed.select('B');
  ed.paste('- x\n  - y\n  - z\n- w');
  const b = M.find(ed.map, 'B'); assert.deepEqual(texts(b), ['x', 'w']); assert.deepEqual(texts(b.children[0]), ['y', 'z']); assert.equal(ed.selIds.length, 2);
  ed.select('A'); const clip = ed.copySel(); assert.equal(clip.text, 'Alpha\n\ta one\n\ta two');
  ed.select('C'); ed.paste(clip.text);
  const c = M.find(ed.map, 'C'); assert.equal(c.children[1].text, 'Alpha'); assert.notEqual(c.children[1].id, 'A'); assert.equal(c.children[1].children.length, 2);
});

test('markers toggle per topic: one priority and one progress at a time, flags and icons stack', () => {
  const ed = editor(); ed.select('A');
  ed.toggleIcon('full-2'); ed.toggleIcon('full-5'); assert.equal(M.find(ed.map, 'A').icon, 'full-5');
  ed.toggleIcon('flag'); ed.toggleIcon('50%'); ed.toggleIcon('idea'); assert.equal(M.find(ed.map, 'A').icon, 'full-5,flag,50%,idea');
  ed.toggleIcon('100%'); ed.toggleIcon('flag'); assert.equal(M.find(ed.map, 'A').icon, 'full-5,idea,100%');
  ed.select('B'); ed.select('A', 'add'); ed.toggleIcon('star'); assert.equal(M.find(ed.map, 'B').icon, 'star'); assert.equal(M.find(ed.map, 'A').icon, 'full-5,idea,100%,star');
  assert.ok(box(ed, 'A').iw > 0 && box(ed, 'A').icons.length === 4);
  const svg = ed.renderVals().svgMap.__html; assert.ok(svg.includes('data-nid="A"') && svg.includes('>5</text>') && svg.includes('💡'));
});

test('relationships, boundaries and summaries: relate two topics or click the second, style the line, wrap a branch, sum up siblings', () => {
  const ed = editor(), J = v => JSON.stringify(v); ed.select('A'); ed.select('C1', 'add'); ed.relate(); // objects made inside the editor's realm compare as JSON
  assert.equal(J(M.find(ed.map, 'A').rels), J([{ to: 'C1', label: '', color: '', arrows: 'end' }])); assert.equal(J(ed.state.selRel), J({ from: 'A', i: 0 })); assert.equal(ed.selId, null);
  ed.openRelPop(ed.state.selRel); ed.onPopText({ target: { value: 'because' } }); ed.onPopKey(key('Enter'));
  ed.patchRel({ from: 'A', i: 0 }, { arrows: 'both', color: 'B5563A' });
  assert.equal(J(M.find(ed.map, 'A').rels[0]), J({ to: 'C1', label: 'because', color: 'B5563A', arrows: 'both' }));
  const svg = ed.renderVals().svgMap.__html; assert.ok(svg.includes('data-rel="A|0"') && svg.includes('>because</text>') && svg.includes('#B5563A'));
  ed.select('B'); ed.relate(); assert.equal(ed.state.relFrom, 'B'); assert.equal(ed.toasts.pop(), '点击另一个主题，画出关联线');
  ed.onCanvasDown({ button: 0, target: { closest: s => s.includes('data-nid') ? { getAttribute: () => 'A2' } : null }, preventDefault() {}, metaKey: false, ctrlKey: false, shiftKey: false, clientX: 0, clientY: 0 });
  assert.equal(M.find(ed.map, 'B').rels[0].to, 'A2'); assert.equal(ed.state.relFrom, null);
  ed.onKey(key('Backspace')); assert.equal(M.find(ed.map, 'B').rels, undefined); assert.equal(ed.state.selRel, null);
  ed.select('C1'); ed.onKey(key('Backspace')); assert.equal(M.find(ed.map, 'A').rels, undefined, 'a line to a deleted topic goes with it');
  ed.select('A'); ed.cloudToggle(); assert.equal(M.find(ed.map, 'A').cloud, 'CFE2F3'); assert.equal(ed.L().clouds.length, 1);
  ed.cloudToggle('FFE8A3'); assert.equal(M.find(ed.map, 'A').cloud, 'FFE8A3'); ed.cloudToggle(); assert.equal(M.find(ed.map, 'A').cloud, undefined);
  ed.select('R'); ed.cloudToggle(); assert.equal(ed.toasts.pop(), '先选中一个主题'); assert.equal(ed.map.cloud, undefined);
  ed.select('A1'); ed.select('A2', 'add'); ed.summarize();
  const a = M.find(ed.map, 'A'), s = a.children[2]; assert.equal(s.summary, 'A1:A2'); assert.equal(s.text, '概要'); assert.equal(ed.state.editing, s.id);
  assert.equal(ed.L().braces[0].id, s.id); assert.ok(box(ed, s.id).x > box(ed, 'A2').x + box(ed, 'A2').w);
  ed.commitEdit(); ed.select('A1'); ed.select('C', 'add'); ed.summarize(); assert.equal(ed.toasts.pop(), '概要只能加在同一层的相邻主题上');
});

test('search walks the hits and unfolds them; the outliner rebuilds the map from edited text; the view tab restructures and themes', () => {
  const ed = editor(); M.find(ed.map, 'A').collapsed = true;
  ed.onKey(key('f', { metaKey: true })); assert.ok(ed.state.search); ed.onSearch({ target: { value: 'ONE ' } });
  assert.deepEqual([...ed.matches()], ['A1', 'C1']); assert.equal(ed.renderVals().searchStat, '2');
  ed.onSearchKey(key('Enter')); assert.equal(ed.selId, 'A1'); assert.equal(M.find(ed.map, 'A').collapsed, undefined, 'the folded branch opens to show the hit');
  ed.onSearchKey(key('Enter')); assert.equal(ed.selId, 'C1'); assert.equal(ed.renderVals().searchStat, '2/2'); ed.onSearchKey(key('Enter', { shiftKey: true })); assert.equal(ed.selId, 'A1');
  assert.ok(ed.renderVals().svgOver.__html.includes('#C28A1A')); ed.onSearchKey(key('Escape')); assert.equal(ed.state.search, null);
  ed.toggleOutline(); assert.equal(ed.state.view, 'outline'); assert.equal(ed.state.outline.split('\n')[1], '\tAlpha');
  ed.onOutlineText({ target: { value: 'Root\n\tBeta\n\tAlpha\n\t\ta two\n\t\tnew one\n\tGamma\n\t\tc one' } }); ed.onOutlineKey(key('Escape'));
  assert.equal(ed.state.view, 'map'); assert.deepEqual(texts(ed.map), ['Beta', 'Alpha', 'Gamma']); assert.equal(ed.map.children[1].id, 'A'); assert.deepEqual(texts(ed.map.children[1]), ['a two', 'new one']); assert.equal(ed.map.children[1].children[0].id, 'A2');
  ed.select('R'); const tools = () => ed.renderVals().tools;
  ed.state.tab = 'view'; tools().find(t => t.label === '组织结构图').onClick(); assert.equal(ed.map.structure, 'org'); assert.equal(box(ed, 'R').mode, 'org');
  tools().find(t => t.title === 'ocean').onClick(); assert.equal(ed.map.theme, 'ocean'); tools().find(t => t.label === '彩虹分支').onClick(); assert.equal(ed.map.mono, true);
  tools().find(t => t.label === '折线').onClick(); assert.equal(ed.map.lines, 'elbow'); tools().find(t => t.label === '思维导图').onClick(); assert.equal(ed.map.structure, undefined);
  ed.select('A'); tools().find(t => t.label === '树状图').onClick(); assert.equal(M.find(ed.map, 'A').structure, 'tree'); assert.equal(box(ed, 'A').mode, 'tree'); assert.ok(tools().find(t => t.label === '时间轴').color.includes('C7C7CC'), 'timeline is for the whole map');
  const v = ed.renderVals(); assert.ok(v.svgMap.__html.includes(M.THEMES.ocean.root));
});

test('exports: a standalone SVG with resolved colours and a Markdown outline with links, labels and notes', () => {
  const ed = editor(); ed.patch('A', { labels: 'urgent', link: 'https://x.y', note: 'first\nsecond' }); ed.patch('B', { cloud: 'CFE2F3' });
  const svg = ed.svgDoc();
  assert.ok(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"') && svg.includes('viewBox="') && svg.includes('>Alpha</tspan>') && svg.includes('#CFE2F3'));
  assert.ok(!svg.includes('var(--'), 'CSS variables are resolved for other programs'); assert.ok(svg.includes('fill:#1D1D1F'), 'the centre takes the light colours');
  assert.equal(ed.mdOutline(), '# Root\n\n- Alpha (https://x.y) `urgent`\n  > first\n  > second\n  - a one\n  - a two\n- Beta\n- Gamma\n  - c one\n');
});

test('the topic popup writes labels and links, the note saves as typed, and the link glyph opens it', () => {
  const ed = editor(); ed.openPop('labels', 'A'); assert.equal(ed.state.pop.kind, 'labels'); assert.equal(ed.selId, 'A');
  ed.onPopText({ target: { value: ' urgent， Q4 ' } }); ed.onPopKey(key('Enter'));
  assert.equal(M.find(ed.map, 'A').labels, 'urgent,Q4'); assert.equal(ed.state.pop, null);
  const plain = editor(); assert.ok(box(ed, 'A').h > box(plain, 'A').h && box(ed, 'A').labels.length === 2);
  ed.onKey(key('k', { metaKey: true })); assert.equal(ed.state.pop.kind, 'link');
  ed.onPopText({ target: { value: 'example.com' } }); ed.onPopBlur(); assert.equal(M.find(ed.map, 'A').link, 'example.com');
  ed.doAct('link', 'A'); assert.deepEqual(opened, ['https://example.com']);
  ed.doAct('note', 'A'); ed.onPopText({ target: { value: 'remember' } }); assert.equal(M.find(ed.map, 'A').note, 'remember'); assert.equal(ed.state.pop.val, 'remember');
  ed.onPopKey(key('Escape')); assert.equal(ed.state.pop, null);
  const v = ed.renderVals(); assert.ok(v.svgMap.__html.includes('data-act="note"') && v.svgMap.__html.includes('data-act="link"'));
  ed.patch('A', { image: 'data:image/png;base64,AA', imageSize: '400,300' });
  const b = box(ed, 'A'); assert.deepEqual([b.img.w, b.img.h], [240, 180]); assert.ok(ed.renderVals().svgMap.__html.includes('<image href="data:image/png;base64,AA"'));
});
