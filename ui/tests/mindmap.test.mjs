// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layout, wrap, flatten, plan, markAi, count, estimate } from '../mindmap.js';

const T = (id, text, children = [], extra = {}) => Object.assign({ id, text, children }, extra);
const map = () => T('R', 'Root', [T('A', 'Alpha', [T('A1', 'a one'), T('A2', 'a two')]), T('B', 'Beta'), T('C', 'Gamma', [T('C1', 'c one')])]);
/** Runs a plan against a fake engine that hands out ids ID_1, ID_2… and returns the argv list it saw. */
function exec(orig, model, file = 'm.mm') {
  const { steps, ids } = plan(orig, model, file); const seen = []; let n = 0;
  for (const s of steps) { const argv = s.argv; seen.push(argv); if (argv[0] === 'add') s.onResult && s.onResult({ kind: 'topic', path: '/x', props: { id: 'ID_' + (++n) } }); }
  return { seen, ids };
}
const byId = L => Object.fromEntries(L.nodes.map(n => [n.id, n]));

test('wrap breaks CJK per character and Latin per word', () => {
  const m = t => estimate(t) * 10;
  assert.deepEqual(wrap('hello world foo', 60, m), ['hello', 'world foo']);
  assert.deepEqual(wrap('一二三四五六七八', 40, m), ['一二三四', '五六七八']);
  assert.deepEqual(wrap('a\nb', 100, m), ['a', 'b']);
  assert.deepEqual(wrap('', 100, m), ['']);
});

test('layout: root centred, first level alternates right/left, side honoured, no sibling overlap', () => {
  const L = layout(map()), N = byId(L);
  assert.equal(N.R.x, -N.R.w / 2); assert.equal(N.R.y, -N.R.h / 2);
  assert.ok(N.A.x > N.R.w / 2, 'A goes right'); assert.ok(N.B.x + N.B.w < -N.R.w / 2, 'B goes left'); assert.ok(N.C.x > 0, 'C goes right');
  assert.ok(N.A1.x > N.A.x + N.A.w && N.A2.x > N.A.x + N.A.w, 'children beyond the parent');
  assert.ok(N.A1.y + N.A1.h <= N.A2.y, 'siblings stacked without overlap');
  assert.ok(Math.abs((N.A1.y + N.A2.y + N.A2.h) / 2 - (N.A.y + N.A.h / 2)) < 1, 'children centred on the parent');
  const m = map(); m.children[0].side = 'left';
  const N2 = byId(layout(m)); assert.ok(N2.A.x < 0, 'explicit side=left wins'); assert.ok(N2.A1.x < N2.A.x, 'left branch children grow leftwards');
  assert.equal(L.edges.length, 6);
  assert.ok(L.bounds.minX < 0 && L.bounds.maxX > 0);
});

test('layout: collapsed branch hides children and carries a descendant count badge', () => {
  const m = map(); m.children[0].collapsed = true;
  const L = layout(m), N = byId(L);
  assert.equal(N.A1, undefined); assert.equal(N.A.badge, 2); assert.equal(N.B.badge, 0);
  assert.equal(count(m), 7);
});

test('flatten records parent, index and normalized props', () => {
  const f = flatten(map());
  assert.deepEqual(f.A2, { text: 'a two', note: '', collapsed: false, side: '', link: '', color: '', fill: '', icon: '', parentId: 'A', index: 1, depth: 2 });
  assert.equal(f.R.parentId, null);
});

test('plan: unchanged model → no steps; prop changes → one set with only the changed props', () => {
  const orig = flatten(map());
  assert.deepEqual(exec(orig, map()).seen, []);
  const m = map(); m.children[1].text = 'Beta 2'; m.children[1].note = 'hi'; m.children[1].fill = 'FFE8A3'; m.text = 'Root!';
  assert.deepEqual(exec(orig, m).seen, [
    ['set', 'm.mm', '//topic[@id=R]', '--prop', 'text=Root!'],
    ['set', 'm.mm', '//topic[@id=B]', '--prop', 'text=Beta 2', '--prop', 'note=hi', '--prop', 'fill=FFE8A3']
  ]);
  const m2 = map(); m2.children[0].collapsed = true; m2.children[0].children[0].color = '';
  const o2 = flatten(map()); o2.A1.color = 'FF0000'; o2.A1.note = 'old';
  assert.deepEqual(exec(o2, m2).seen, [
    ['set', 'm.mm', '//topic[@id=A]', '--prop', 'collapsed=true'],
    ['set', 'm.mm', '//topic[@id=A1]', '--prop', 'note=', '--prop', 'color=none']
  ]);
});

test('plan: side only for first-level nodes', () => {
  const m = map(); m.children[2].side = 'left'; m.children[0].children[0].side = 'left';
  assert.deepEqual(exec(flatten(map()), m).seen, [['set', 'm.mm', '//topic[@id=C]', '--prop', 'side=left']]);
});

test('plan: new nodes are added with a 1-based index, get their engine id, and children use it', () => {
  const m = map(); const nn = T('tmp_1', 'New', [T('tmp_2', 'Grandchild', [], { icon: 'idea' })]);
  m.children[0].children.splice(1, 0, nn);
  const { seen, ids } = exec(flatten(map()), m);
  assert.deepEqual(seen, [
    ['add', 'm.mm', '//topic[@id=A]', '--type', 'topic', '--prop', 'text=New', '--index', '2'],
    ['add', 'm.mm', '//topic[@id=ID_1]', '--type', 'topic', '--prop', 'text=Grandchild', '--index', '1'],
    ['set', 'm.mm', '//topic[@id=ID_2]', '--prop', 'icon=idea']
  ]);
  assert.deepEqual(ids, { tmp_1: 'ID_1', tmp_2: 'ID_2' });
  assert.equal(nn.id, 'ID_1', 'the model node carries the real id afterwards');
  assert.equal(nn.children[0].id, 'ID_2');
});

test('plan: removing a subtree is one remove for the topmost node', () => {
  const m = map(); m.children.splice(0, 1);
  assert.deepEqual(exec(flatten(map()), m).seen, [['remove', 'm.mm', '//topic[@id=A]']]);
});

test('plan: reparenting and reordering become moves with the target index', () => {
  const m = map(); const [a1] = m.children[0].children.splice(0, 1); m.children[1].children.push(a1); // A1 → under B
  assert.deepEqual(exec(flatten(map()), m).seen, [['move', 'm.mm', '//topic[@id=A1]', '--to', '//topic[@id=B]', '--index', '1']]);
  const m2 = map(); const [b] = m2.children.splice(1, 1); m2.children.splice(0, 0, b); // B before A
  assert.deepEqual(exec(flatten(map()), m2).seen, [['move', 'm.mm', '//topic[@id=B]', '--to', '//topic[@id=R]', '--index', '1']]);
  const m3 = map(); m3.children[0].children.reverse(); // swap A1/A2 → one move
  assert.deepEqual(exec(flatten(map()), m3).seen, [['move', 'm.mm', '//topic[@id=A2]', '--to', '//topic[@id=A]', '--index', '1']]);
});

test('plan: a survivor whose parent is removed is moved out before the removal', () => {
  const m = map(); const [a] = m.children.splice(0, 1); m.children[1].children.push(a.children[1]); // delete A, keep A2 under C
  assert.deepEqual(exec(flatten(map()), m).seen, [
    ['move', 'm.mm', '//topic[@id=A2]', '--to', '//topic[@id=C]'],
    ['remove', 'm.mm', '//topic[@id=A]']
  ]);
});

test('plan: indexes count siblings that are still in the file but about to leave', () => {
  const m = map(); const [a1] = m.children[0].children.splice(0, 1); m.children[1].children.push(a1); m.children[0].children.push(T('tmp_x', 'X')); // A: [A2, X], A1 → B
  assert.deepEqual(exec(flatten(map()), m).seen, [
    ['add', 'm.mm', '//topic[@id=A]', '--type', 'topic', '--prop', 'text=X', '--index', '3'],
    ['move', 'm.mm', '//topic[@id=A1]', '--to', '//topic[@id=B]', '--index', '1']
  ]);
});

test('plan: a node moved under a brand-new parent waits for that parent id', () => {
  const m = map(); const [c1] = m.children[2].children.splice(0, 1); m.children.push(T('tmp_9', 'Fresh', [c1]));
  assert.deepEqual(exec(flatten(map()), m).seen, [
    ['add', 'm.mm', '//topic[@id=R]', '--type', 'topic', '--prop', 'text=Fresh', '--index', '4'],
    ['move', 'm.mm', '//topic[@id=C1]', '--to', '//topic[@id=ID_1]', '--index', '1']
  ]);
});

test('markAi flags new, changed, moved nodes and counts topmost removals', () => {
  const after = map(); after.children[0].text = 'Alpha!'; after.children.push(T('N', 'New')); const [c1] = after.children[2].children.splice(0, 1); after.children[1].children.push(c1);
  const before = map(); before.children.push(T('X', 'x', [T('X1', 'x1')]));
  const items = markAi(before, after);
  assert.deepEqual(items, [['修改', '主题 · Alpha!'], ['移动', '主题 · c one'], ['新增', '主题 · New'], ['删除', '1 个主题']]);
  assert.equal(after.children[0].ai, true); assert.equal(after.children[1].ai, undefined);
});
