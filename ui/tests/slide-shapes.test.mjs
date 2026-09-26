// node --test ui/tests/ — the slide kit's shapes: the gallery's polygons drawn as SVG with dashed outlines and gradients, lines
// with arrowheads that stick to shapes, and groups whose members move with them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const K = await import('../office-io.js');
const th = K.THEMES.paper;

test('polygon shapes draw as an SVG polygon with the outline\'s width and dash, a gradient as its own linearGradient', () => {
  const star = K.shape({ id: 's1', shape: 'star5', x: 10, y: 20, w: 300, h: 300, fill: '#E3B25A', stroke: '#1D1D1F', sw: 4, dash: 'dash', shadow: true });
  const v = K.objView(star, th);
  assert.match(v.svg.__html, /<polygon points="([\d.]+,[\d.]+ ?){10}"/, 'ten points');
  assert.match(v.svg.__html, /fill="#E3B25A" stroke="#1D1D1F" stroke-width="4" stroke-dasharray="16 12"/);
  assert.equal(v.bg, 'transparent'); assert.equal(v.border, 'none');
  assert.equal(v.flt, 'drop-shadow(0 3px 6px rgba(0,0,0,0.35))');
  const grad = K.objView(K.shape({ id: 'g/1', shape: 'hexagon', fill: 'grad:#FF0000,#0000FF,45' }), th);
  assert.match(grad.svg.__html, /<linearGradient id="gg_1" gradientTransform="rotate\(45 .5 .5\)">.*fill="url\(#gg_1\)"/);
  const rect = K.objView(K.shape({ shape: 'round', w: 200, h: 100, fill: 'grad:#FF0000,#0000FF,90', stroke: '#333333', sw: 2, dash: 'dot' }), th);
  assert.equal(rect.bg, 'linear-gradient(180deg,#FF0000,#0000FF)', 'CSS shapes take a CSS gradient (the file\'s 90° is top to bottom)');
  assert.equal(rect.border, '2px dotted #333333');
  assert.equal(rect.svg.__html, '');
  assert.equal(Object.keys(K.POLY).length, 21);
  for (const [, items] of K.SHAPE_GALLERY.slice(0, 3)) for (const [key] of items) assert.ok(K.POLY[key] || ['rect', 'round', 'ellipse', 'pill'].includes(key), key + ' is drawn');
});

test('lines: the box runs start to end with flips, arrowheads point along the line, ends stuck to shapes follow them', () => {
  const l = K.line({ id: 'l', x: 100, y: 100, w: 200, h: 0, sw: 3, tail: 'triangle', stroke: '#C00000' });
  const v = K.objView(l, th);
  assert.ok(v.isLine);
  assert.match(v.svg.__html, /<path d="M0,0 L200,0" fill="none" stroke="#C00000" stroke-width="3"/);
  assert.match(v.svg.__html, /<polygon points="200.0,0.0 184.0,7.3 184.0,-7.3" fill="#C00000"\/>/, 'the tail arrowhead sits at the end and points right');
  assert.deepEqual(K.lineBox([300, 300], [100, 100]), { x: 100, y: 100, w: 200, h: 200, flipH: true, flipV: true });
  assert.deepEqual(K.lineEnds({ w: 200, h: 200, flipH: true, flipV: true }), [[200, 200], [0, 0]]);
  const box = K.shape({ id: 'b', x: 400, y: 100, w: 200, h: 100 }), stuck = K.line({ id: 's', x: 0, y: 0, w: 10, h: 10, start: { id: 'b', idx: 3 } });
  const s = { objs: [box, stuck, K.line({ id: 'free', x: 5, y: 5, w: 50, h: 0 })] };
  K.relinkLines(s);
  assert.deepEqual([stuck.x, stuck.y, stuck.w, stuck.h, stuck.flipH, stuck.flipV], [10, 10, 590, 140, true, true], 'the start moved to the box\'s right side (600,150); the free end stayed at (10,10)');
  box.x = 800; K.relinkLines(s);
  assert.deepEqual([stuck.x + (stuck.flipH ? stuck.w : 0), stuck.y + (stuck.flipV ? stuck.h : 0)], [1000, 150]);
  assert.deepEqual([s.objs[2].x, s.objs[2].w], [5, 50], 'an unconnected line is left alone');
});

test('groups: the box around the members, kids drawn at their offsets; moving or resizing the group carries them', () => {
  const a = K.shape({ id: 'a', x: 100, y: 100, w: 200, h: 100 }), b = K.shape({ id: 'b', x: 400, y: 300, w: 100, h: 100, shape: 'ellipse' });
  const g = K.group([a, b]);
  assert.deepEqual([g.t, g.x, g.y, g.w, g.h], ['group', 100, 100, 400, 300]);
  const v = K.objView(g, th);
  assert.ok(v.isGroup);
  assert.deepEqual(v.kids.map(k => [k.left, k.top, k.width]), [['0px', '0px', '200px'], ['300px', '200px', '100px']]);
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [{ id: 's1', layout: 'blank', decor: [], objs: [a, b, K.line({ id: 'l', x: 300, y: 150, w: 100, h: 0, end: { id: 'b', idx: 1 } })], notes: '', trans: 'none', hidden: false, bg: null }] };
  const c = new ctx.SlideEditor(); c.props = { get doc() { return doc; }, onChange: d => { doc = d; }, toast() { } }; c.K = K;
  c.setState({ sel: 'b', sels: ['a', 'b'] });
  c.groupSels();
  let s = doc.slides[0];
  assert.deepEqual(s.objs.map(o => o.t), ['group', 'line'], 'the group stands where the topmost member was, under the line that was above it');
  const gid = s.objs[0].id;
  assert.deepEqual([c.state.sel, Array.from(c.state.sels)], [gid, [gid]]);
  c.patchSel(c.moveKids(s.objs[0], { x: 200, w: 800 })); // moved right by 100 and doubled in width: kids follow
  s = doc.slides[0];
  assert.deepEqual(s.objs[0].kids.map(k => [k.x, k.y, k.w, k.h]), [[200, 100, 400, 100], [800, 300, 200, 100]]);
  assert.deepEqual([s.objs[1].x + s.objs[1].w, s.objs[1].y + s.objs[1].h], [800, 350], 'the line stuck to b follows it into the group (its end is b\'s left side)');
  c.ungroup();
  s = doc.slides[0];
  assert.deepEqual(s.objs.map(o => o.t), ['shape', 'shape', 'line']);
  assert.deepEqual(Array.from(c.state.sels), ['a', 'b']);
  const v2 = c.renderVals();
  assert.equal(v2.frames.length, 1, 'a second frame for the other selected object');
  const fmt = () => { c.setState({ tab: 'format', bubble: true }); return c.renderVals().ribbon; };
  assert.ok(fmt().some(i => i.label === '组合'), '形状格式 offers 组合 for a multiple selection');
  c.setState({ sel: 'l', sels: ['l'] });
  const lineTools = fmt().filter(i => i.isMenu || i.isColor || i.isNum).map(i => i.label);
  assert.deepEqual(Array.from(lineTools), ['线条', '粗细', '虚线', '起点箭头', '终点箭头', '线型', '排列', '对齐', 'X', 'Y', '宽', '高']); // the editor's arrays come from its vm
  c.setState({ sel: 'a', sels: ['a'] });
  const numX = fmt().find(i => i.isNum && i.label === 'X');
  assert.equal(numX.value, 4.23, '200 slide units is 4.23 cm on a 16:9 deck');
  numX.onChange({ target: { value: '2.117' } });
  assert.equal(doc.slides[0].objs.find(o => o.id === 'a').x, 100, 'the exact field moves the shape');
  c.insertShape('doubleArrowLine');
  const ln = doc.slides[0].objs[doc.slides[0].objs.length - 1];
  assert.deepEqual([ln.t, ln.head, ln.tail, ln.h], ['line', 'triangle', 'triangle', 0]);
  c.insertShape('wedgeRectCallout');
  assert.equal(doc.slides[0].objs[doc.slides[0].objs.length - 1].shape, 'wedgeRectCallout');
});

test('right-click on the slide: the selected objects\' menu (clipboard, order, grouping), or the slide\'s when nothing is selected', () => {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const a = K.shape({ id: 'a', shape: 'rect', x: 100, y: 100, w: 200, h: 100 }), b = K.shape({ id: 'b', shape: 'ellipse', x: 400, y: 100, w: 200, h: 100 });
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [{ id: 's1', layout: 'blank', decor: [], objs: [a, b], notes: '', trans: 'none', hidden: false, bg: null }] };
  const c = new ctx.SlideEditor(); c.props = { get doc() { return doc; }, onChange: d => { doc = d; }, toast() { } }; c.K = K;
  const right = () => { const e = { clientX: 300, clientY: 200, preventDefault() { this.prevented = true; } }; c.onStageCtx(e); return e.prevented; };
  const items = () => Array.from(c.renderVals().popItems).filter(i => i.isItem !== false && i.label && !i.isHead && !i.head).map(i => i.label);
  c.setState({ sel: 'a', sels: ['a'] });
  assert.equal(right(), true, 'no system menu'); assert.equal(c.state.pop.id, 'stagectx');
  let l = items(); for (const k of ['剪切', '复制', '复制一份', '删除', '置于顶层', '置于底层']) assert.ok(l.includes(k), k);
  assert.ok(!l.includes('粘贴') && !l.includes('组合'), 'nothing to paste yet, one object');
  Array.from(c.renderVals().popItems).find(i => i.label === '复制').onClick();
  right(); assert.ok(items().includes('粘贴'), 'after 复制');
  c.setState({ sel: 'b', sels: ['a', 'b'] }); right(); assert.ok(items().includes('组合'), 'two objects group');
  c.setState({ sel: null, sels: [] }); right(); l = items();
  assert.ok(l.includes('从此页放映') && l.includes('粘贴') && !l.includes('删除'), 'the slide\'s menu: ' + l.join(' '));
});
