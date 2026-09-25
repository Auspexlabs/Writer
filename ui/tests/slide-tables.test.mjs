// node --test ui/tests/ — tables: PowerPoint's styles drawn cell by cell, merges, rows and columns added and removed around
// them, columns resized by dragging, and the bridge writing widths, style flags, merges and cell styles to the file.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const K = await import('../office-io.js'), EN = await import('../engine.js');
const th = K.THEMES.paper;
const plain = x => JSON.parse(JSON.stringify(x)); // the editor's objects come from its vm
const grid = (r, c) => Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => `${i}${j}`));

test('tableCells: the style decides fills and lines, merges span and hide cells, a cell\'s own fill wins', () => {
  const t = K.txt({ id: 't', t: 'table', w: 900, h: 300, rows: grid(3, 3), merges: [{ r: 1, c: 0, rs: 2, cs: 2 }], cells: { '0:2': { fill: '#FF0000', align: 'center' } } });
  let cells = K.tableCells(t, th);
  assert.equal(cells.length, 9 - 3, 'the merge hides three cells');
  const at = (r, c) => cells.find(x => x.r === r && x.c === c);
  assert.deepEqual([at(0, 0).bg, at(0, 0).color, at(0, 0).fw, at(1, 2).bg, at(2, 2).bg], [th.acc, '#FFFFFF', 700, th.card, 'transparent'], 'Medium Style 2: accent header, banded rows');
  assert.deepEqual([at(1, 0).gc, at(1, 0).gr, at(0, 2).bg, at(0, 2).jc], ['1 / span 2', '2 / span 2', '#FF0000', 'center']);
  assert.equal(at(1, 1), undefined);
  cells = K.tableCells(Object.assign({}, t, { tstyle: 'DarkStyle1', header: false, banded: false }), th);
  assert.deepEqual([cells[0].bg, cells[0].color, cells[0].fw], [th.fg, th.bg, 400]);
  cells = K.tableCells(Object.assign({}, t, { tstyle: 'TableGrid', firstCol: true }), th);
  assert.deepEqual([cells[0].bg, cells[0].border, cells[0].fw, cells[1].fw], ['transparent', '1px solid rgba(128,128,128,0.35)', 700, 700]);
  assert.deepEqual(K.colWidths(t), [300, 300, 300]);
  assert.equal(K.objView(t, th).gtc, '300px 300px 300px');
});

test('row and column edits keep merges and cell styles in place', () => {
  const t = K.txt({ id: 't', t: 'table', w: 900, h: 300, rows: grid(3, 3), merges: [{ r: 1, c: 1, rs: 2, cs: 2 }], cells: { '2:0': { fill: '#00FF00' } } });
  K.tableInsertRow(t, 0);
  assert.deepEqual([t.rows.length, t.rows[0], t.merges, t.cells, t.h], [4, ['', '', ''], [{ r: 2, c: 1, rs: 2, cs: 2 }], { '3:0': { fill: '#00FF00' } }, 400]);
  K.tableInsertRow(t, 3); // inside the merge: it grows
  assert.deepEqual([t.rows.length, t.merges[0].rs], [5, 3]);
  K.tableDeleteRow(t, 0);
  assert.deepEqual([t.rows.length, t.merges, Object.keys(t.cells)], [4, [{ r: 1, c: 1, rs: 3, cs: 2 }], ['3:0']]);
  K.tableInsertCol(t, 0);
  assert.deepEqual([t.rows[0].length, t.colW, t.w, t.merges[0].c, Object.keys(t.cells)], [4, [300, 300, 300, 300], 1200, 2, ['3:1']]);
  K.tableDeleteCol(t, 3); // the merge's last column: it narrows
  assert.deepEqual([t.rows[0].length, t.w, t.merges[0].cs], [3, 900, 1]);
  K.tableDeleteCol(t, 2); // the merge's only column: the merge goes… (rs 3 stays: still a merge)
  assert.deepEqual(t.merges, [{ r: 1, c: 1, rs: 3, cs: 1 }]);
  K.tableMerge(t, 0, 0, 1, 1);
  assert.deepEqual([t.merges, t.rows[0][0], t.rows[1][1]], [[{ r: 0, c: 0, rs: 2, cs: 2 }], '00 10', ''], 'the texts join the anchor (the empty new column adds nothing), the merge it touched went');
  K.tableSplit(t, 0, 0);
  assert.deepEqual(t.merges, []);
});

test('the editor: 表格工具 works at the current cell, ⇧-click ranges merge, a column boundary drags', () => {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [{ id: 's1', layout: 'blank', decor: [], objs: [], notes: '', trans: 'none', hidden: false, bg: null }] };
  const c = new ctx.SlideEditor(); c.props = { get doc() { return doc; }, onChange: d => { doc = d; }, toast() { } }; c.K = K;
  c.setState({ box: { w: 1664, h: 948 } });
  c.insertTable(2, 3);
  const t = () => doc.slides[0].objs[0];
  assert.deepEqual([t().rows.length, t().rows[0][0], t().tstyle, t().header], [2, '标题 1', 'MediumStyle2Accent1', true]);
  const md = (r, cc, mods = {}) => c.renderVals().objs.find(o => o.id === t().id).tcells.find(x => x.r === r && x.c === cc).onMD(Object.assign({ stopPropagation() { }, preventDefault() { } }, mods));
  md(1, 1); // a click in the selected table puts the caret in the cell
  assert.deepEqual(plain([c.state.editing, c.state.cell, c.focusEdit]), [t().id, { r: 1, c: 1 }, t().id + ':1:1']);
  md(1, 2, { shiftKey: true });
  assert.deepEqual(plain(c.state.cellSel), { r: 1, c: 2 });
  c.setState({ tab: 'table', bubble: true, editing: null });
  const rib = () => c.renderVals().ribbon;
  rib().find(i => i.label === '合并单元格').onClick();
  assert.deepEqual(plain([t().merges, c.state.cellSel, c.state.cell]), [[{ r: 1, c: 1, rs: 1, cs: 2 }], null, { r: 1, c: 1 }]);
  rib().find(i => i.label === '下方插入行').onClick();
  assert.deepEqual([t().rows.length, t().h], [3, 240]);
  rib().find(i => i.label === '左侧插入列').onClick();
  assert.deepEqual([t().rows[0].length, t().w, t().merges[0].c], [4, 4 * Math.round(1280 / 3), 2]);
  const fill = rib().find(i => i.isColor && i.label === '填充'); fill.onChange({ target: { value: '#123456' } });
  assert.deepEqual(plain(t().cells), { '1:1': { fill: '#123456' } });
  rib().find(i => i.label === '表格样式'); c.menus.tstyle[3].onClick();
  rib().find(i => i.label === '镶边行').onClick();
  assert.deepEqual([t().tstyle, t().banded], ['DarkStyle1', false]);
  // dragging the first column boundary 100 to the right widens that column and the table
  c.setState({ tab: 'home', bubble: false, editing: null, sel: t().id, sels: [t().id] });
  const before = K.colWidths(t()), v = c.renderVals();
  assert.equal(v.colHandles.length, 3);
  v.colHandles[0].onMD({ stopPropagation() { }, preventDefault() { }, clientX: 0, clientY: 0 });
  c.onWM({ clientX: 100, clientY: 0 }); c.onWU();
  assert.deepEqual([t().colW[0], t().colW[1], t().w], [before[0] + 100, before[1], before.reduce((a, b) => a + b, 0) + 100]);
});

test('the bridge: widths, style flags, merges and cell styles are read from the file and the changes written back', async () => {
  const cell = (path, text, extra) => ({ kind: 'cell', path, props: Object.assign({ text }, extra) });
  const tree = { kind: 'document', props: { width: '33.867cm', height: '19.05cm' }, children: [{ kind: 'slide', path: '/slide[@id=256]', props: { id: '256', layout: 'Blank' }, children: [
    { kind: 'table', path: '/slide[@id=256]/table[@id=5]', props: { id: '5', rows: '2', cols: '2', x: '2cm', y: '2cm', w: '10cm', h: '3cm', widths: '["4cm","6cm"]', style: 'LightStyle1', header: 'true' }, children: [
      { kind: 'row', path: '/slide[@id=256]/table[@id=5]/row[1]', children: [cell('/slide[@id=256]/table[@id=5]/row[1]/cell[1]', 'a', { colspan: '2', fill: 'D9E2F3' }), cell('/slide[@id=256]/table[@id=5]/row[1]/cell[2]', '', { covered: 'true' })] },
      { kind: 'row', path: '/slide[@id=256]/table[@id=5]/row[2]', children: [cell('/slide[@id=256]/table[@id=5]/row[2]/cell[1]', 'c', {}), cell('/slide[@id=256]/table[@id=5]/row[2]/cell[2]', 'd', { align: 'right' })] }] }] }] };
  const prev = globalThis.fetch, commands = [];
  globalThis.fetch = async (url, opts) => { if (String(url).startsWith('/json?')) return { ok: true, json: async () => tree }; const { argv } = JSON.parse(opts.body); commands.push(argv); return { ok: true, json: async () => ({ code: 0, output: '{}' }) }; };
  try {
    const doc = await EN.open({ id: 'p1', path: 'deck.pptx', type: 'pptx' }), t = doc.slides[0].objs[0];
    assert.deepEqual([t.t, t.colW, t.tstyle, t.header, t.banded, t.merges, t.cells], ['table', [189, 283], 'LightStyle1', true, false, [{ r: 0, c: 0, rs: 1, cs: 2 }], { '0:0': { fill: '#D9E2F3' }, '1:1': { align: 'right' } }]);
    assert.equal(await EN.save(doc), 0);
    K.tableSplit(t, 0, 0); K.tableMerge(t, 0, 1, 1, 1); t.cells['0:0'] = { fill: '#FF0000', line: '#000000' }; delete t.cells['1:1']; t.colW = [236, 236]; t.w = 472; t.tstyle = 'TableGrid'; t.banded = true;
    await EN.save(doc);
    const sets = commands.filter(c => c[0] === 'set').map(c => [c[2], ...c.slice(3).filter(x => x !== '--prop')]);
    assert.deepEqual(sets[0].slice(0, 1).concat(sets[0].slice(1).sort()), ['/slide[@id=256]/table[@id=5]', 'banded=true', 'data=[["a","d"],["c",""]]', 'style=TableGrid', 'widths=["4.995cm","4.995cm"]'], 'the merged text moved into its anchor; the width is the same, the columns are not');
    assert.deepEqual(sets.slice(1), [
      ['/slide[@id=256]/table[@id=5]/row[1]/cell[1]', 'colspan=1', 'rowspan=1', 'fill=FF0000', 'line=000000'],
      ['/slide[@id=256]/table[@id=5]/row[1]/cell[2]', 'colspan=1', 'rowspan=2'],
      ['/slide[@id=256]/table[@id=5]/row[2]/cell[2]', 'align=left']], 'the old anchor lets go before the new merge takes the cell; a cleared style goes back to its default');
  } finally { globalThis.fetch = prev; }
});
