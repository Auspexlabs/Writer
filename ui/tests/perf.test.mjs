// node --test ui/tests/   (Node 20) — the pure parts of the frame-time diagnostic and of the render memoisation.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { frameStats, windowStats } from '../perf.js';
import '../memo.js'; // a classic script: it installs globalThis.dcMemo

const { propsEqual, stable, stableAll } = globalThis.dcMemo;
// support.js is an IIFE that needs a DOM and React; its expression/encoding sections are plain functions, cut out of
// the repository's own file and evaluated here (no outside input reaches the Function body).
const rt = readFileSync(new URL('../support.js', import.meta.url), 'utf8');
const { resolve, compileExpr, compileAttr, cssToObj } = new Function(rt.slice(rt.indexOf('// src/expr.ts'), rt.indexOf('// src/compile.ts')) + '\nreturn { resolve, compileExpr, compileAttr, cssToObj };')();

test('frameStats: fps from the mean delta, worst frame, hitches over 50 ms, vsyncs missed against the median', () => {
  assert.deepEqual(frameStats([]), { frames: 0, fps: 0, avg: 0, worst: 0, long: 0, dropped: 0 });
  assert.deepEqual(frameStats([16.7, 16.7, 16.6]), { frames: 3, fps: 60, avg: 16.7, worst: 17, long: 0, dropped: 0 });
  assert.deepEqual(frameStats([16, 84, 16, 16]), { frames: 4, fps: 30, avg: 33, worst: 84, long: 1, dropped: 4 });
  assert.deepEqual(frameStats([8.3, 8.3, 16.7, 8.4, 8.3]), { frames: 5, fps: 100, avg: 10, worst: 17, long: 0, dropped: 1 });
  assert.deepEqual(frameStats([38, 38, 38], 16.7), { frames: 3, fps: 26, avg: 38, worst: 38, long: 0, dropped: 3 }); // every frame slow
});

test('windowStats only counts frames inside the window', () => {
  const times = [0, 16, 32, 132, 148, 164, 180];
  assert.deepEqual(windowStats(times, 20, 150), { frames: 3, fps: 23, avg: 44, worst: 100, long: 1, dropped: 5, ms: 130 });
  assert.deepEqual(windowStats(times, 150, 200), { frames: 2, fps: 63, avg: 16, worst: 16, long: 0, dropped: 0, ms: 50 });
});

test('propsEqual: identical values are equal; the host-style object compares by value', () => {
  const f = () => 1, doc = { id: 'a' };
  assert.equal(propsEqual({ doc, onChange: f, n: 1 }, { doc, onChange: f, n: 1 }), true);
  assert.equal(propsEqual({ doc, onChange: f }, { doc, onChange: () => 1 }), false); // a fresh closure is a change
  assert.equal(propsEqual({ doc }, { doc: { id: 'a' } }), false);
  assert.equal(propsEqual({ n: 1 }, { n: 1, m: undefined }), false);
  assert.equal(propsEqual({ __hostStyle: { position: 'absolute', inset: '0' } }, { __hostStyle: { position: 'absolute', inset: '0' } }), true);
  assert.equal(propsEqual({ __hostStyle: { inset: '0' } }, { __hostStyle: { inset: '1px' } }), false);
  assert.equal(propsEqual({ n: NaN }, { n: NaN }), true);
});

test('stable: one identity per key and owner, always calling the latest function', () => {
  const owner = {}, calls = [];
  const a = stable(owner, 'onChange', x => calls.push(['first', x]));
  const b = stable(owner, 'onChange', x => calls.push(['second', x]));
  assert.equal(a, b);
  a(1);
  assert.deepEqual(calls, [['second', 1]]);
  assert.notEqual(stable(owner, 'onUndo', () => 0), a);
  assert.notEqual(stable({}, 'onChange', () => 0), a);
});

test('sc-for memo="1": an item that is the same object as at the last render keeps its elements; another component keeps its own', () => {
  const src = rt.slice(rt.indexOf('  function walkFor(el, host) {'), rt.indexOf('  function walkIf(el, host) {')), drawn = [];
  const walkFor = new Function('compileAttr', 'walkChildren', 'h', 'getReact', 'warnUnresolved', src + '\nreturn walkFor;')(compileAttr,
    () => [sub => { drawn.push(sub.c.v); return { v: sub.c.v }; }], (type, props, kids) => ({ type, key: props.key, kids }), () => ({ Fragment: 'F' }), () => { });
  const el = attrs => ({ getAttribute: k => (k in attrs ? attrs[k] : null), hasAttribute: k => k in attrs });
  const list = walkFor(el({ list: '{{ cells }}', as: 'c', memo: '1' }), {}), a = { v: 'a' }, b = { v: 'b' }, owner = {};
  const r1 = list({ cells: [a, b] }, owner, 'k'), r2 = list({ cells: [a, { v: 'b2' }] }, owner, 'k');
  assert.equal(r2.kids[0], r1.kids[0], 'the same item, the same element: React skips it');
  assert.notEqual(r2.kids[1], r1.kids[1]);
  assert.deepEqual(drawn, ['a', 'b', 'b2'], 'only the changed item is drawn again');
  list({ cells: [a] }, {}, 'k'); assert.deepEqual(drawn.slice(3), ['a'], 'another instance of the component draws its own');
  const plainList = walkFor(el({ list: '{{ cells }}', as: 'c' }), {}); drawn.length = 0;
  plainList({ cells: [a] }, owner, 'k'); plainList({ cells: [a] }, owner, 'k');
  assert.deepEqual(drawn, ['a', 'a'], 'without memo every render draws every item');
});

test('compiled template expressions give exactly what resolve() gives', () => {
  const vals = { a: { b: { c: 3 }, b1: 'x', 0: 'zero', 1: 'one' }, n: 0, s: '', t: true, arr: ['p', 'q'], i: 1, obj: { k: 'v' }, nul: null, $index: 4, 'true': 'shadow' };
  const exprs = ['a', ' a ', 'a.b', 'a.b.c', 'a.0', 'a.b1', 'a.1b', 'missing', 'missing.x', 'nul.x', 'n', 's', '$index', 'true', 'false', 'null',
    'undefined', '150', '-1', '"str"', "'s'", '!t', '!!n', 'n === 0', 'a.b.c == 3', 'i !== 1', '(a.b.c)', 'arr[0]', 'arr[i]', 'obj["k"]', 'a..b', 'a.'];
  for (const e of exprs) assert.deepEqual(compileExpr(e)(vals), resolve(vals, e), e);
  const item = Object.create(vals); item.c = { d: 5 }; // what sc-for hands each list item
  for (const e of ['c.d', 'a.b.c', '$index', 'n === 0']) assert.deepEqual(compileExpr(e)(item), resolve(item, e), e);
});

test('compileAttr: a whole binding keeps the value, mixed text stringifies with empty holes', () => {
  const doc = { id: 'd' };
  assert.equal(compileAttr('{{ doc }}')({ doc }), doc);
  assert.equal(compileAttr('left:{{ x }};top:{{ y }}px;{{ z }}')({ x: 1.5, y: null }), 'left:1.5;top:px;');
  assert.equal(compileAttr('plain')({}), 'plain');
});

test('cssToObj parses once and hands back the same object', () => {
  const a = cssToObj('position:absolute;inset:0;--k:1;background-color:red'), b = cssToObj('position:absolute;inset:0;--k:1;background-color:red');
  assert.equal(a, b);
  assert.deepEqual(a, { position: 'absolute', inset: '0', '--k': '1', backgroundColor: 'red' });
});

test('stableAll swaps top-level functions only, so two renders give equal props', () => {
  const owner = {}, row = { onClick: () => 0 };
  const r1 = stableAll(owner, { toggle: () => 'one', label: 'AI', row });
  const r2 = stableAll(owner, { toggle: () => 'two', label: 'AI', row });
  assert.equal(r1.toggle, r2.toggle);
  assert.equal(r2.toggle(), 'two');
  assert.equal(r2.row.onClick, row.onClick);
  assert.equal(propsEqual(r1, r2), true);
});
