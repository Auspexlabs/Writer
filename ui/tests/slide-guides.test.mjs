// node --test ui/tests/ — smart guides: a dragged object snaps to the slide's centre and edges and to other objects' edges and
// centres, red lines show where, gaps to the neighbours are labelled; ⇧ keeps a move on one axis and ⌥ drags a copy.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const K = await import('../office-io.js');

test('snapMove: the nearest edge or centre within the threshold wins, per axis, and says where the guide goes', () => {
  const others = [{ x: 100, y: 100, w: 200, h: 100 }];
  const cands = K.snapCands(others, 1600, 900);
  assert.deepEqual(K.snapMove({ x: 306, y: 380, w: 100, h: 50 }, cands, 8), { x: 300, y: 380, guides: [{ axis: 'v', at: 300 }] }, 'left edge to the other\'s right edge');
  assert.deepEqual(K.snapMove({ x: 745, y: 120, w: 100, h: 60 }, cands, 8), { x: 750, y: 120, guides: [{ axis: 'v', at: 800 }, { axis: 'h', at: 150 }] }, 'centre to the slide\'s centre, middle to the other\'s middle');
  assert.deepEqual(K.snapMove({ x: 500, y: 500, w: 100, h: 50 }, cands, 8).guides, [], 'nothing near: no guide');
  assert.deepEqual(K.snapResize({ x: 100, y: 300, w: 195, h: 100 }, 'e', cands, 8), { x: 100, y: 300, w: 200, h: 100, guides: [{ axis: 'v', at: 300 }] }, 'a dragged right edge snaps, the left stays');
  assert.deepEqual(K.snapResize({ x: 96, y: 300, w: 204, h: 100 }, 'nw', cands, 8), { x: 100, y: 300, w: 200, h: 100, guides: [{ axis: 'v', at: 100 }] }, 'the left edge snaps and the width follows; the top finds nothing');
});

test('gapHints: the gap to the nearest neighbour on each side the box overlaps with', () => {
  const others = [{ x: 100, y: 100, w: 200, h: 100 }, { x: 800, y: 120, w: 100, h: 100 }, { x: 0, y: 0, w: 50, h: 50 }];
  const hints = K.gapHints({ x: 400, y: 110, w: 100, h: 80 }, others);
  assert.deepEqual(hints.map(h => [h.axis, h.from, h.to, h.text]), [['h', 300, 400, '100'], ['h', 500, 800, '300']], 'left and right neighbours; the corner square does not overlap');
  assert.equal(hints[0].at, 150, 'the line runs through the middle of the shared band');
});

test('dragging: guides and hints show while moving, ⇧ locks the axis, ⌥ drags a copy, ⌘ drags free, a resize snaps its edge', () => {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const a = K.shape({ id: 'a', x: 100, y: 100, w: 200, h: 100 }), b = K.shape({ id: 'b', x: 500, y: 380, w: 100, h: 50 });
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [{ id: 's1', layout: 'blank', decor: [], objs: [a, b], notes: '', trans: 'none', hidden: false, bg: null }] };
  const c = new ctx.SlideEditor(); c.props = { get doc() { return doc; }, onChange: d => { doc = d; }, toast() { } }; c.K = K;
  c.setState({ box: { w: 1664, h: 948 } }); // k = 1: slide units are pixels
  const md = (id, mods = {}) => c.renderVals().objs.find(o => o.id === id).onMD(Object.assign({ clientX: 0, clientY: 0, stopPropagation() { }, preventDefault() { }, currentTarget: { closest: () => null } }, mods));
  const mv = (x, y, mods = {}) => c.onWM(Object.assign({ clientX: x, clientY: y }, mods));
  md('b'); mv(-195, 0); // b's left edge lands 5 from a's right edge (300): it snaps there
  assert.equal(c.state.lives.b.x, 300);
  let v = c.renderVals();
  assert.deepEqual(v.guides.map(g => [g.left, g.top]), [['300px', 0]], 'a vertical guide at the snapped edge');
  assert.equal(v.hints.length, 0, 'b sits below a: no neighbour on its rows');
  mv(-190, -275); // up beside a: left edge 310 (no snap), top 105 snaps to a's top 100
  assert.deepEqual([c.state.lives.b.x, c.state.lives.b.y], [310, 100]);
  v = c.renderVals();
  assert.deepEqual(v.guides.map(g => [g.left, g.top]), [[0, '100px']]);
  assert.deepEqual(v.hints.map(h => [h.text, h.left, h.width, h.top]), [['10', '300px', '10px', '125px']], 'the gap to a, drawn through the shared band');
  c.onWU();
  assert.deepEqual([doc.slides[0].objs[1].x, doc.slides[0].objs[1].y, c.state.lives], [310, 100, null]);
  md('b'); mv(50, 3, { shiftKey: true }); // ⇧: the small vertical drift is dropped
  assert.deepEqual([c.state.lives.b.x, c.state.lives.b.y], [360, 100]); c.onWU();
  md('b'); mv(-55, 0, { metaKey: true }); // ⌘: 305 stays 305, no guides
  assert.equal(c.state.lives.b.x, 305); assert.deepEqual(JSON.parse(JSON.stringify(c.state.guides)), { lines: [], hints: [] }); c.onWU();
  md('b', { altKey: true }); // ⌥: a copy is dragged, b stays
  assert.equal(doc.slides[0].objs.length, 3);
  const copy = doc.slides[0].objs[2];
  assert.equal(c.state.sel, copy.id);
  mv(100, 200); c.onWU();
  assert.deepEqual([doc.slides[0].objs[1].x, doc.slides[0].objs[2].x, doc.slides[0].objs[2].y], [305, 405, 300]);
  // a resize snaps the dragged edge: the copy's left edge from 405 to 297 lands on a's right edge, 300
  c.setState({ sel: copy.id, sels: [copy.id] });
  const west = c.renderVals().handles.find(h => h.cursor === 'ew-resize' && h.left === '0%');
  west.onMD({ clientX: 0, clientY: 0, stopPropagation() { }, preventDefault() { } });
  mv(-108, 0);
  assert.deepEqual([c.state.lives[copy.id].x, c.state.lives[copy.id].w], [300, 205]);
  assert.deepEqual(c.renderVals().guides.map(g => g.left), ['300px']);
  c.onWU();
});
