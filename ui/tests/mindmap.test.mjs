// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { layout, wrap, flatten, plan, markAi, count, estimate, topicStyle, cssFont, toOutline, parseOutline, cloneBranch, within, dropAt, iconList, marker, markerSvg, bracePath, retarget, THEMES, BRANCH, fromOutline } from '../mindmap.js';

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
  assert.deepEqual(f.A2, { text: 'a two', note: '', collapsed: false, side: '', link: '', color: '', fill: '', icon: '', bold: false, italic: false, strike: false, size: '', font: '', free: '', labels: '', image: '', imageSize: '', cloud: '', summary: '', rels: '', structure: '', theme: '', lines: '', mono: false, parentId: 'A', index: 1, depth: 2 });
  assert.equal(f.R.parentId, null);
});

test('layout: a floating topic sits at its free position with no line from the root; the sides ignore it', () => {
  const m = map(); m.children[1].free = '300,-200';
  const L = layout(m), N = byId(L);
  assert.equal([N.B.x, N.B.y, N.B.free].join(), '300,-200,true');
  assert.ok(N.A.x > 0 && N.C.x < 0, 'A and C take the two sides');
  assert.equal(L.edges.length, 5); assert.ok(!L.edges.some(e => e.to === 'B'));
  assert.ok(N.C1.x < N.C.x, 'C is a left branch now, so its children grow left');
  assert.deepEqual(exec(flatten(map()), m).seen, [['set', 'm.mm', '//topic[@id=B]', '--prop', 'free=300,-200']]);
  const back = map(); assert.deepEqual(exec(flatten(m), back).seen, [['set', 'm.mm', '//topic[@id=B]', '--prop', 'free=']]);
});

test('markers draw by name, and a topic with markers, labels and a picture grows to fit them', () => {
  assert.deepEqual(iconList({ icon: ' full-3, flag ,star' }), ['full-3', 'flag', 'star']);
  assert.equal(marker('full-3').text, '3'); assert.equal(marker('50%').shapes.length, 2); assert.equal(marker('0%').shapes.length, 1); assert.ok(marker('flag-blue') && marker('star-red')); assert.equal(marker('idea'), null);
  assert.ok(/<text[^>]*>3<\/text>/.test(markerSvg('full-3', 0, 0, 16)) && markerSvg('idea', 0, 0, 16).includes('💡') && markerSvg('nothing', 0, 0, 16).includes('•'));
  const m = map(); Object.assign(m.children[1], { labels: 'urgent,Q4', image: 'data:image/png;base64,AA', imageSize: '400,300', icon: 'full-1,flag' });
  const N = byId(layout(m)), P = byId(layout(map()));
  assert.ok(N.B.w > P.B.w && N.B.h > P.B.h + 180, 'the picture and the labels make the box bigger');
  assert.deepEqual([N.B.img.w, N.B.img.h, N.B.labels.length, N.B.icons.length], [240, 180, 2, 2]);
  assert.ok(N.B.iw > 30 && N.B.ty > 180 && N.B.labels[0].w > 20);
  assert.deepEqual(exec(flatten(map()), m).seen, [['set', 'm.mm', '//topic[@id=B]', '--prop', 'icon=full-1,flag', '--prop', 'labels=urgent,Q4', '--prop', 'image=data:image/png;base64,AA', '--prop', 'imageSize=400,300']]);
});

test('boundaries pad the branch, summaries sit beside a brace over their range, relationships curve between boxes', () => {
  const m = map(); m.children[0].cloud = 'CFE2F3'; m.children[2].children.push(T('S', 'sum', [], { summary: 'C1:C1' })); m.children[1].rels = [{ to: 'C1', label: 'why', arrows: 'both' }];
  const L = layout(m), N = byId(L), P = byId(layout(map()));
  assert.equal(L.clouds.length, 1); const c = L.clouds[0];
  assert.ok(c.x < N.A.x && c.y < N.A1.y && c.x + c.w > N.A2.x + N.A2.w && c.y + c.h > N.A2.y + N.A2.h, 'the cloud wraps A and its children');
  assert.ok(N.C.y > P.C.y + 10 || N.B.y !== P.B.y, 'neighbours make room for the padding');
  assert.equal(L.braces.length, 1); const br = L.braces[0];
  assert.equal(br.id, 'S'); assert.ok(br.x > N.C1.x + N.C1.w && N.S.x > br.x, 'brace beyond the summed topic, the summary beyond the brace');
  assert.ok(Math.abs((br.y0 + br.y1) / 2 - (N.S.y + N.S.h / 2)) < 1, 'the summary is centred on the brace'); assert.ok(bracePath(br).startsWith('M'));
  assert.ok(!L.edges.some(e => e.to === 'S'), 'no branch line to a summary'); assert.ok(N.S.isSummary);
  assert.equal(L.rels.length, 1); const r = L.rels[0];
  assert.deepEqual([r.from, r.to, r.label, r.arrows, r.color], ['B', 'C1', 'why', 'both', '']); assert.ok(/^M[-\d.]+ [-\d.]+C/.test(r.d) && r.heads.end.endsWith('z') && isFinite(r.mid.x));
  assert.deepEqual(exec(flatten(map()), m).seen.map(a => a.slice(2)), [
    ['//topic[@id=A]', '--prop', 'cloud=CFE2F3'], ['//topic[@id=C]', '--type', 'topic', '--prop', 'text=sum', '--index', '2'],
    ['//topic[@id=B]', '--prop', 'rels=[{"to":"C1","label":"why","color":"","arrows":"both"}]'], ['//topic[@id=ID_1]', '--prop', 'summary=C1:C1']]);
  const o = flatten(m); m.children[1].rels = [];
  assert.deepEqual(exec(o, m).seen.map(a => a.slice(2)), [['//topic[@id=B]', '--prop', 'rels=[]']]);
});

test('plan: rels and summaries that point at topics added in the same save get the engine ids', () => {
  const m = map(); m.children.push(T('tmp_n', 'New')); m.children[1].rels = [{ to: 'tmp_n' }]; m.children[0].children.push(T('tmp_s', 'S', [], { summary: 'A1:tmp_n' }));
  const { seen } = exec(flatten(map()), m);
  assert.deepEqual(seen.find(a => a[2] === '//topic[@id=B]').slice(3), ['--prop', 'rels=[{"to":"ID_2","label":"","color":"","arrows":"end"}]']);
  assert.deepEqual(seen.find(a => a[2] === '//topic[@id=ID_1]').slice(3), ['--prop', 'summary=A1:ID_2']);
  const n = { rels: [{ to: 'tmp_n' }], summary: 'tmp_s:x' }; retarget(n, { tmp_n: 'ID_9', tmp_s: 'ID_8' }); assert.deepEqual([n.rels[0].to, n.summary], ['ID_9', 'ID_8:x']);
});

test('structures: logic keeps one side, org rows children below, tree indents them, timeline alternates along an axis; themes and line styles colour and route the lines', () => {
  const at = (extra, m = map()) => { Object.assign(m, extra); const L = layout(m); return { L, N: byId(L) }; };
  const lg = at({ structure: 'logic' }); assert.ok(lg.N.A.x > 0 && lg.N.B.x > 0 && lg.N.C.x > 0 && lg.N.A.y < lg.N.B.y && lg.N.B.y < lg.N.C.y);
  const ll = at({ structure: 'logic-left' }); assert.ok(ll.N.A.x + ll.N.A.w < 0 && ll.N.A1.x < ll.N.A.x, 'left logic grows leftwards');
  const og = at({ structure: 'org' });
  assert.ok(og.N.A.y > og.N.R.y + og.N.R.h && og.N.A.x < og.N.B.x && og.N.B.x < og.N.C.x, 'first level in a row below the centre');
  assert.ok(og.N.A1.y > og.N.A.y + og.N.A.h && og.N.A1.x < og.N.A2.x && og.N.A1.x + og.N.A1.w <= og.N.A2.x, 'grandchildren row below their parent');
  assert.ok(og.L.edges.find(e => e.to === 'A').d.includes('V') && !og.L.edges.find(e => e.to === 'A').d.includes('C'), 'org lines are elbows');
  assert.equal(og.N.R.mode, 'org');
  const tr = at({ structure: 'tree' });
  assert.ok(tr.N.A.x === tr.N.R.x + 28 && tr.N.A.y > tr.N.R.y + tr.N.R.h && tr.N.A1.x === tr.N.A.x + 28 && tr.N.A2.y > tr.N.A1.y + tr.N.A1.h && tr.N.B.y > tr.N.A2.y + tr.N.A2.h, 'an indented list');
  const tl = at({ structure: 'timeline' });
  assert.ok(tl.N.A.x > tl.N.R.x + tl.N.R.w && tl.N.B.x > tl.N.A.x && tl.N.C.x > tl.N.B.x, 'milestones march right');
  const ay = tl.N.R.y + tl.N.R.h / 2; assert.ok(tl.N.A.y > ay && tl.N.B.y + tl.N.B.h < ay && tl.N.C.y > ay, 'below, above, below the axis');
  assert.ok(tl.N.A1.y > tl.N.A.y + tl.N.A.h, 'a milestone below the axis grows down'); assert.ok(tl.L.edges[0].d.startsWith('M') && tl.L.edges[0].d.includes('H'), 'the axis is the first line');
  const m2 = map(); m2.children[0].structure = 'org'; const br = byId(layout(m2)); assert.ok(br.A1.y > br.A.y + br.A.h && br.A2.x > br.A1.x && br.B.x < 0, 'one branch can have its own structure');
  const th = at({ theme: 'ocean', mono: true }); assert.equal(th.N.A.color, THEMES.ocean.branches[0]); assert.equal(th.N.C.color, THEMES.ocean.branches[0]); assert.equal(th.N.R.color, THEMES.ocean.root);
  assert.equal(at({}).N.C.color, BRANCH[2]);
  assert.ok(at({ lines: 'straight' }).L.edges[0].d.includes('L') && at({ lines: 'elbow' }).L.edges[0].d.includes('H') && at({}).L.edges[0].d.includes('C'));
  const m3 = map(); m3.structure = 'org'; const L3 = layout(m3), N3 = byId(L3);
  const gap = dropAt(L3, m3, { x: (N3.A.x + N3.A.w + N3.B.x) / 2, y: N3.A.y + 4 }, ['C1']);
  assert.deepEqual([gap.kind, gap.horiz, gap.parent, gap.after, gap.before], ['between', true, 'R', 'A', 'B']);
});

test('fromOutline: an edited outline becomes the map again, topics keeping their ids and looks by text', () => {
  const m = map(); m.children[0].icon = 'flag'; m.children[0].children[0].note = 'n'; m.children[1].text = 'two\nlines';
  const out = toOutline([m]); assert.equal(out.split('\n')[0], 'Root'); assert.equal(out.split('\n')[4], '\ttwo lines');
  const edited = 'Root!\n\tBeta\n\tAlpha\n\t\ta one\n\t\tfresh\n\tGamma\nExtra top';
  const r = fromOutline(m, edited.replace('Beta', 'two lines'));
  assert.equal(r.id, 'R'); assert.equal(r.text, 'Root!');
  assert.deepEqual(r.children.map(c => [c.id, c.text]), [['B', 'two\nlines'], ['A', 'Alpha'], ['C', 'Gamma'], [r.children[3].id, 'Extra top']]);
  assert.equal(r.children[1].icon, 'flag'); assert.equal(r.children[1].children[0].note, 'n'); assert.ok(r.children[1].children[1].id.startsWith('tmp_')); assert.equal(r.children[1].children.length, 2);
  assert.equal(fromOutline(m, '  \n'), null);
});

test('dropAt: onto a topic reparents, the gap between siblings orders, empty canvas floats; never inside its own branch', () => {
  const m = map(), L = layout(m), N = byId(L), mid = b => ({ x: b.x + b.w / 2, y: b.y + b.h / 2 });
  assert.deepEqual(dropAt(L, m, mid(N.C), ['B']), { kind: 'into', id: 'C', side: null });
  assert.equal(dropAt(L, m, mid(N.A1), ['A']).kind, 'free');
  const gap = dropAt(L, m, { x: N.A1.x + 5, y: (N.A1.y + N.A1.h + N.A2.y) / 2 }, ['B']);
  assert.deepEqual([gap.kind, gap.parent, gap.after, gap.before], ['between', 'A', 'A1', 'A2']);
  const top = dropAt(L, m, { x: N.A1.x + 5, y: N.A1.y - 10 }, ['B']);
  assert.deepEqual([top.parent, top.after, top.before, top.side], ['A', null, 'A1', null]);
  const side = dropAt(L, m, { x: N.A.x + 5, y: N.A.y - 12 }, ['C']);
  assert.deepEqual([side.kind, side.parent, side.before, side.side], ['between', 'R', 'A', 'right']);
  assert.deepEqual(dropAt(L, m, { x: -5, y: 0 }, ['B']), { kind: 'into', id: 'R', side: 'left' });
  assert.deepEqual(dropAt(L, m, { x: 900, y: 900 }, ['B'], { dx: 10, dy: 5 }), { kind: 'free', x: 890, y: 895 });
});

test('plan: bold, italic, strike, size and font travel as props; clearing them writes false / 0 / empty', () => {
  const m = map(); Object.assign(m.children[0], { bold: true, italic: true, size: 20, font: 'Georgia' });
  assert.deepEqual(exec(flatten(map()), m).seen, [['set', 'm.mm', '//topic[@id=A]', '--prop', 'bold=true', '--prop', 'italic=true', '--prop', 'size=20', '--prop', 'font=Georgia']]);
  const o = flatten(m); m.children[0] = T('A', 'Alpha', m.children[0].children, { strike: true });
  assert.deepEqual(exec(o, m).seen, [['set', 'm.mm', '//topic[@id=A]', '--prop', 'bold=false', '--prop', 'italic=false', '--prop', 'strike=true', '--prop', 'size=0', '--prop', 'font=']]);
});

test('topicStyle: a bold, resized topic keeps its level metrics but changes the font', () => {
  const s = topicStyle({ text: 'x', bold: true, size: 20, italic: true, font: 'SansSerif' }, 2);
  assert.equal(s.fw, 700); assert.equal(s.fs, 20); assert.equal(s.fi, true); assert.equal(s.family, '', 'Java logical names mean the default');
  assert.equal(cssFont(s), "italic 700 20px 'IBM Plex Sans','Noto Sans SC',sans-serif");
  assert.equal(topicStyle({ text: 'x', font: 'Georgia' }, 0).family, 'Georgia');
  const L = layout(map(), null, { edit: { id: 'B', text: 'a much longer text while typing' } }), N = byId(L);
  assert.ok(N.B.w > byId(layout(map())).B.w, 'the topic being edited is laid out with the typed text');
});

test('clipboard: branches go out as an indented outline and any outline comes back as topics', () => {
  const m = map();
  assert.equal(toOutline([m.children[0], m.children[1]]), 'Alpha\n\ta one\n\ta two\nBeta');
  assert.deepEqual(parseOutline('- Alpha\n  - a one\n    1. deep\n  - [ ] a two\n# Beta\n\n'), [
    { text: 'Alpha', children: [{ text: 'a one', children: [{ text: 'deep', children: [] }] }, { text: 'a two', children: [] }] }, { text: 'Beta', children: [] }]);
  const c = cloneBranch(Object.assign(m.children[0], { side: 'left', ai: true }));
  assert.notEqual(c.id, 'A'); assert.notEqual(c.children[0].id, 'A1'); assert.equal(c.side, undefined); assert.equal(c.ai, undefined); assert.equal(c.children[1].text, 'a two');
  assert.equal(within(m, 'A', 'A2'), true); assert.equal(within(m, 'B', 'A2'), false); assert.equal(within(m, 'A', 'A'), true);
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
