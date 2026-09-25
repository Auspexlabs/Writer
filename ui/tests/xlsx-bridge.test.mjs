// node --test ui/tests — the xlsx bridge: engine tree → sheet model → mutated model → planned commands.
import { test } from 'node:test';
import assert from 'node:assert/strict';
globalThis.window = globalThis; // pdf-kit touches window at import time
const { sheetModel, cellModel, cellProps, fmtOf, codeOf, planXlsx, sheetProps, chartProps } = await import('../engine.js');

const cell = (ref, props) => ({ kind: 'cell', path: `/sheet[1]/row[${ref.replace(/\D/g, '')}]/cell[${ref}]`, props });
const row = (n, cells) => ({ kind: 'row', path: `/sheet[1]/row[${n}]`, props: {}, children: cells });
const TREE_SHEET = {
  kind: 'sheet', path: '/sheet[1]',
  props: { name: 'Sheet1', range: 'A1:C3', merges: '["A1:C1"]', widths: '{"A":12.5,"B":20}', heights: '{"2":30}', freeze: 'B2', filter: 'A2:C3' },
  children: [
    row(1, [cell('A1', { value: 'Title', type: 'string', bold: 'true', italic: 'true', underline: 'true', strike: 'true', wrap: 'true', color: 'FFFFFF', fill: '1F4E79', size: '14', font: 'Georgia', align: 'center', valign: 'top', indent: '2', border: 'thin', borderColor: 'FF0000', link: 'https://example.com', note: 'hello', format: '0.00%' })]),
    row(2, [cell('A2', { value: '12.5', type: 'number', format: '"¥"#,##0.00' }), cell('B2', { formula: 'A2*2', value: '25', borders: '{"bottom":"double"}' })]),
    row(3, [cell('A3', { value: '2024-01-01', type: 'date', format: 'yyyy-mm-dd' })]),
    { kind: 'chart', path: '/sheet[1]/chart[1]', props: { id: '7', type: 'column', title: 'Sales', categories: 'A2:A3', series: '[{"name":"B1","values":"B2:B3"}]', legend: 'bottom', stacked: 'true', x: '2.54cm', y: '5.08cm', w: '12.7cm', h: '7.62cm' } }
  ]
};
const clone = x => JSON.parse(JSON.stringify(x));
const stub = () => { const cmds = []; let id = 100; const exec = async argv => { cmds.push(argv); if (argv[0] === 'add') { id++; return argv[4] === 'sheet' ? { path: '/sheet[9]', props: { id: String(id) } } : { path: argv[2] + '/chart[1]', props: { id: String(id) } }; } return {}; }; return { cmds, exec }; };
const propsOf = argv => { const p = {}; for (let i = 3; i < argv.length; i += 2) { const [k, ...v] = argv[i + 1].split('='); p[k] = v.join('='); } return p; };
const plan = async (orig, cur) => { const s = stub(); const n = await planXlsx('f.xlsx', orig, cur, s.exec); assert.equal(n, s.cmds.length); return s.cmds; };

test('tree → model maps every cell prop', () => {
  const m = sheetModel(TREE_SHEET);
  assert.deepEqual(m.cells.A1, { v: 'Title', s: { b: true, i: true, u: true, st: true, wrap: true, color: '#FFFFFF', fill: '#1F4E79', fs: 14, font: 'Georgia', align: 'center', va: 'top', indent: 2, bd: 'thin', bdc: '#FF0000', link: 'https://example.com', note: 'hello', fmt: 'pct', dec: 2, code: '0.00%' } });
  assert.deepEqual(m.cells.A2, { v: '12.5', s: { fmt: 'money', dec: 2, code: '"¥"#,##0.00' } });
  assert.deepEqual(m.cells.B2, { v: '=A2*2', s: { bd: { bottom: 'double' } } });
  assert.deepEqual(m.cells.A3.s, { fmt: 'date', code: 'yyyy-mm-dd' });
});

test('tree → model maps sheet props and charts', () => {
  const m = sheetModel(TREE_SHEET);
  assert.equal(m.name, 'Sheet1'); assert.equal(m.path, '/sheet[1]');
  assert.deepEqual(m.merges, [{ r: 0, c: 0, rs: 1, cs: 3 }]);
  assert.deepEqual(m.colW, { A: 93, B: 145 });
  assert.deepEqual(m.rowH, { 2: 40 });
  assert.equal(m.frR, 1); assert.equal(m.frC, 1);
  assert.equal(m.filter, 'A2:C3');
  assert.deepEqual(m.charts, [{ id: '/sheet[1]/chart[@id=7]', eid: '7', path: '/sheet[1]/chart[1]', type: 'column', title: 'Sales', cat: 'A2:A3', ser: [{ name: 'B1', values: 'B2:B3' }], legend: 'bottom', stacked: true, x: 96, y: 192, w: 480, h: 288 }]);
});

test('freeze none / no sheet props → defaults', () => {
  const m = sheetModel({ kind: 'sheet', path: '/sheet[2]', props: { name: 'S2', freeze: 'none', filter: 'none' }, children: [] });
  assert.deepEqual(m, { name: 'S2', path: '/sheet[2]', cells: {}, colW: {}, rowH: {}, merges: [], frR: 0, frC: 0, filter: null, filters: {}, frows: [], cf: [], dv: [], hiddenRows: [], hiddenCols: [], color: null, charts: [], images: [] });
});

test('sheet rules: cf, validations, filter criteria, hidden lines and the tab colour map both ways; filter-hidden rows are the filter\'s, not the user\'s', async () => {
  const props = { name: 'R', id: '3', filter: 'A1:C9', filters: { A: { values: ['East', ''] }, C: { operator: 'greaterThan', value: '5' } }, hidden: { rows: [2, 4, 12], cols: ['B'] }, color: 'C0392B',
    cf: [{ range: 'A2:A9', type: 'cellIs', operator: 'between', value: '1', value2: '9', fill: 'FFEB9C', color: '9C5700' }, { range: 'B2:B9', type: 'colorScale', colors: ['F8696B', '63BE7B'] }],
    validations: [{ range: 'A2:A9', type: 'list', values: ['East', 'West'], error: 'East or West' }] };
  const m = sheetModel({ kind: 'sheet', path: '/sheet[3]', props, children: [row(1, [cell('A1', { value: '1', rotate: '45' })])] });
  assert.deepEqual(m.cf, [{ range: 'A2:A9', type: 'cellIs', operator: 'between', value: '1', value2: '9', fill: '#FFEB9C', color: '#9C5700' }, { range: 'B2:B9', type: 'colorScale', colors: ['#F8696B', '#63BE7B'] }]);
  assert.deepEqual(m.dv, props.validations); assert.deepEqual(m.filters, props.filters); assert.equal(m.color, '#C0392B');
  assert.deepEqual([m.frows, m.hiddenRows, m.hiddenCols], [[1, 3], [11], [1]], 'rows 2 and 4 sit in the filtered range: the filter hid them; row 12 was hidden by hand');
  assert.deepEqual(m.cells.A1, { v: '1', s: { rotate: 45 } });
  const orig = [m], cur = clone(orig), s = cur[0];
  assert.deepEqual(await plan(orig, clone(orig)), [], 'nothing edited → nothing written');
  s.cf[0].fill = '#C6EFCE'; s.dv = []; s.filters = { A: { values: ['West'] } }; s.frows = [2]; s.hiddenRows = []; s.hiddenCols = [1, 3]; s.color = null; s.cells.A1.s.rotate = -30;
  const cmds = await plan(orig, cur);
  assert.deepEqual(cmds, [
    ['set', 'f.xlsx', '/sheet[@id=3]/cell[A1]', '--prop', 'rotate=-30'],
    ['set', 'f.xlsx', '/sheet[@id=3]', '--prop', 'filters={"A":{"values":["West"]}}', '--prop', 'cf=' + JSON.stringify([{ range: 'A2:A9', type: 'cellIs', operator: 'between', value: '1', value2: '9', fill: 'C6EFCE', color: '9C5700' }, { range: 'B2:B9', type: 'colorScale', colors: ['F8696B', '63BE7B'] }]),
      '--prop', 'validations=[]', '--prop', 'hidden={"rows":[3],"cols":["B","D"]}', '--prop', 'color=none']
  ]);
  assert.deepEqual(cellProps({ v: '1' }, { v: '1', s: { fmt: 'eur', dec: 0 } }), { format: '"€"#,##0' });
  assert.deepEqual(fmtOf('"€"#,##0.00'), { fmt: 'eur', dec: 2, code: '"€"#,##0.00' });
});

test('format codes: kind detection and round trip', () => {
  assert.deepEqual(fmtOf('General'), {});
  assert.deepEqual(fmtOf('0.0%'), { fmt: 'pct', dec: 1, code: '0.0%' });
  assert.deepEqual(fmtOf('"$"#,##0.00'), { fmt: 'usd', dec: 2, code: '"$"#,##0.00' });
  assert.equal(fmtOf('_("¥"* #,##0.00_);_("¥"* (#,##0.00);_("¥"* "-"??_);_(@_)').fmt, 'acct');
  assert.equal(fmtOf('hh:mm:ss').fmt, 'time');
  assert.equal(fmtOf('yyyy-mm-dd hh:mm').fmt, 'date');
  assert.equal(fmtOf('@').fmt, 'text');
  assert.deepEqual(fmtOf('[Red]General'), { fmt: 'custom', code: '[Red]General' });
  assert.equal(codeOf(fmtOf('#,##0.00_);[Red](#,##0.00)')), '#,##0.00_);[Red](#,##0.00)');
  assert.equal(codeOf({ fmt: 'pct', dec: 0 }), '0%'); assert.equal(codeOf({ fmt: 'number', dec: 0 }), '#,##0'); assert.equal(codeOf({ fmt: 'usd' }), '"$"#,##0.00'); assert.equal(codeOf({ fmt: 'time' }), 'h:mm:ss'); assert.equal(codeOf({ fmt: 'custom', code: '0000' }), '0000');
  assert.deepEqual(cellModel({ value: 'x' }), { v: 'x' });
});

test('format round trip: an untouched code is never rewritten, a chosen preset writes its exact code', () => {
  const o = { v: '1', s: { fmt: 'number', dec: 2, code: '#,##0.00_);[Red](#,##0.00)' } };
  assert.deepEqual(fmtOf(o.s.code), o.s);
  const bolded = clone(o); bolded.s.b = true;
  assert.deepEqual(cellProps(o, bolded), { bold: 'true' }, 'only bold changed → no format=');
  assert.deepEqual(cellProps(o, { v: '2', s: clone(o.s) }), { value: '2' }, 'value edit keeps the code');
  const picked = { v: '1', s: { fmt: 'plain', dec: 2 } };
  assert.deepEqual(cellProps(o, picked), { format: '0.00' }, '数值 → 0.00');
  assert.deepEqual(fmtOf('0.00'), { fmt: 'plain', dec: 2, code: '0.00' });
  const exact = { plain: '0.00', number: '#,##0.00', money: '"¥"#,##0.00', usd: '"$"#,##0.00', acct: '_("¥"* #,##0.00_)', pct: '0.00%', date: 'yyyy/m/d', time: 'h:mm:ss', text: '@' };
  for (const [fmt, code] of Object.entries(exact)) { assert.equal(codeOf({ fmt, dec: 2 }), code, fmt); assert.equal(fmtOf(code).fmt, fmt, 'and back: ' + code); }
  assert.equal(codeOf({ fmt: 'number', dec: 0 }), '#,##0', '千分位'); assert.equal(codeOf({ fmt: 'pct', dec: 0 }), '0%'); assert.equal(codeOf({ fmt: 'plain', dec: 0 }), '0');
  assert.equal(codeOf({ fmt: 'custom', code: '[Red]0.0' }), '[Red]0.0', '自定义 = typed code');
  const custom = sheetModel({ kind: 'sheet', path: '/sheet[1]', props: { name: 'S' }, children: [row(1, [cell('A1', { value: '5', format: '[Blue]General' })])] });
  assert.deepEqual(custom.cells.A1.s, { fmt: 'custom', code: '[Blue]General' });
  assert.deepEqual(cellProps(custom.cells.A1, { v: '5', s: { fmt: 'custom', code: '[Blue]General', i: true } }), { italic: 'true' });
});

test('cell diff emits only the changed props, empties remove note/link, border none', () => {
  const o = sheetModel(TREE_SHEET).cells.A1;
  const n = clone(o); Object.assign(n.s, { b: false, i: false, note: '', link: '', bd: null, indent: 0, align: null, va: null, fs: 12, font: '', fmt: 'number', dec: 0, code: null });
  assert.deepEqual(cellProps(o, n), { bold: 'false', italic: 'false', size: '12', font: 'Calibri', align: 'general', valign: 'bottom', indent: '0', border: 'none', link: '', note: '', format: '#,##0' });
  assert.deepEqual(cellProps(o, clone(o)), {});
  assert.deepEqual(cellProps({ v: '' }, { v: 'x', s: { fmt: 'text' } }), { value: 'x', type: 'string', format: '@' });
  assert.deepEqual(cellProps({ v: '1' }, { v: '=A1+1', s: { bd: { top: 'thin', bottom: 'thick' }, bdc: '#00FF00' } }), { formula: 'A1+1', borders: '{"top":"thin","bottom":"thick"}', borderColor: '00FF00' });
});

test('plan: cell sets, removals and a sheet prop set', async () => {
  const orig = [sheetModel(TREE_SHEET)], cur = clone(orig);
  const s = cur[0];
  s.cells.A2.v = '13'; s.cells.A1.s.note = 'changed'; delete s.cells.B2; s.cells.A3 = { v: '' };
  s.cells.C3 = { v: 'new', s: { b: true } };
  s.merges = []; s.colW.A = 100; s.rowH = { 2: 40, 5: 48 }; s.frR = 0; s.frC = 0; s.filter = null;
  const cmds = await plan(orig, cur);
  assert.deepEqual(cmds, [
    ['remove', 'f.xlsx', '/sheet[1]/cell[B2]'],
    ['remove', 'f.xlsx', '/sheet[1]/cell[A3]'],
    ['set', 'f.xlsx', '/sheet[1]/cell[A1]', '--prop', 'note=changed'],
    ['set', 'f.xlsx', '/sheet[1]/cell[A2]', '--prop', 'value=13'],
    ['set', 'f.xlsx', '/sheet[1]/cell[C3]', '--prop', 'value=new', '--prop', 'bold=true'],
    ['set', 'f.xlsx', '/sheet[1]', '--prop', 'merges=[]', '--prop', 'widths={"A":13.57}', '--prop', 'heights={"5":36}', '--prop', 'freeze=none', '--prop', 'filter=none']
  ]);
  assert.deepEqual(await plan(orig, clone(orig)), []);
});

test('plan: ≥ 8 identical style changes over a rectangle become one range set', async () => {
  const orig = [{ path: '/sheet[1]', name: 'S', cells: {}, charts: [] }], cur = clone(orig);
  for (let r = 1; r <= 4; r++) for (const c of ['A', 'B']) cur[0].cells[c + r] = { v: c + r, s: { b: true, fill: '#EEEEEE' } };
  cur[0].cells.D9 = { v: 'x', s: { b: true, fill: '#EEEEEE' } }; // same style but outside the block: on its own
  const cmds = await plan(orig, cur);
  assert.equal(cmds.filter(c => c[2].includes('range')).length, 0, 'the D9 straggler makes the group ragged → no range');
  delete cur[0].cells.D9;
  const cmds2 = await plan(orig, cur);
  assert.deepEqual(cmds2[0], ['set', 'f.xlsx', '/sheet[1]/range[A1:B4]', '--prop', 'bold=true', '--prop', 'fill=EEEEEE']);
  assert.equal(cmds2.length, 9);
  assert.deepEqual(cmds2[1], ['set', 'f.xlsx', '/sheet[1]/cell[A1]', '--prop', 'value=A1']);
  assert.ok(cmds2.slice(1).every(c => c.length === 5 && c[4].startsWith('value=')));
});

test('plan: charts add / set / remove, positions as cm strings, ids written back', async () => {
  const orig = [sheetModel(TREE_SHEET)], cur = clone(orig);
  const s = cur[0];
  s.charts[0].title = 'Revenue'; s.charts[0].x = 192; s.charts[0].legend = 'none';
  s.charts.push({ id: 'chabc', type: 'pie', title: 'Share', cat: 'A2:A3', ser: [{ name: 'B1', values: 'B2:B3' }], legend: 'right', stacked: false, x: 0, y: 96, w: 460, h: 300 });
  const cmds = await plan(orig, cur);
  assert.deepEqual(cmds, [
    ['set', 'f.xlsx', '/sheet[1]/chart[@id=7]', '--prop', 'title=Revenue', '--prop', 'legend=none', '--prop', 'x=5.08cm'],
    ['add', 'f.xlsx', '/sheet[1]', '--type', 'chart', '--prop', 'type=pie', '--prop', 'title=Share', '--prop', 'categories=A2:A3', '--prop', 'series=[{"name":"B1","values":"B2:B3"}]', '--prop', 'legend=right', '--prop', 'stacked=false', '--prop', 'x=0cm', '--prop', 'y=2.54cm', '--prop', 'w=12.171cm', '--prop', 'h=7.938cm']
  ]);
  assert.equal(s.charts[1].eid, '101'); assert.equal(s.charts[1].path, '/sheet[1]/chart[1]'); assert.equal(s.charts[1].id, 'chabc', 'the editor id stays');
  assert.deepEqual(chartProps(s.charts[1]).x, '0cm');
  const cur2 = clone(cur); cur2[0].charts = [];
  const orig2 = clone(cur);
  assert.deepEqual(await plan(orig2, cur2), [['remove', 'f.xlsx', '/sheet[1]/chart[@id=101]'], ['remove', 'f.xlsx', '/sheet[1]/chart[@id=7]']]);
});

test('plan: sheets added, renamed and removed', async () => {
  const orig = [sheetModel(TREE_SHEET), { path: '/sheet[2]', name: 'Old', cells: {}, charts: [] }], cur = clone(orig);
  cur[0].name = 'Renamed'; cur.pop(); cur.push({ name: 'Fresh', cells: { A1: { v: '1' } } });
  const cmds = await plan(orig, cur);
  assert.deepEqual(cmds, [
    ['set', 'f.xlsx', '/sheet[1]', '--prop', 'name=Renamed'],
    ['add', 'f.xlsx', '/', '--type', 'sheet', '--prop', 'name=Fresh', '--after', '/sheet[1]'],
    ['set', 'f.xlsx', '/sheet[@id=101]/cell[A1]', '--prop', 'value=1'],
    ['remove', 'f.xlsx', '/sheet[2]']
  ]);
  assert.equal(cur[1].path, '/sheet[@id=101]', 'a new sheet is addressed by its id, which moves leave alone');
  assert.deepEqual(sheetProps({ merges: [], colW: {}, rowH: {} }, { merges: [{ r: 0, c: 0, rs: 2, cs: 2 }], frR: 2, frC: 0, filter: 'A1:B9' }), { merges: '["A1:B2"]', freeze: 'A3', filter: 'A1:B9' });
  assert.deepEqual(propsOf(cmds[0]), { name: 'Renamed' });
});

// The live /json tree (what openXlsx really receives): json props arrive parsed, font sizes carry "pt", dates come as ISO text.
const LIVE_SHEET = {
  kind: 'sheet', path: '/sheet[2]',
  props: { name: 'Live', range: 'A1:C3', id: '2', merges: ['A1:C1'], widths: { A: 22, B: 32 }, heights: { 2: 30 }, freeze: 'A2', filter: 'A2:C3' },
  children: [
    row(1, [cell('A1', { text: 'T', value: 'T', type: 'string', bold: 'true', size: '14pt', border: 'thin', borders: { top: 'thin', bottom: 'double' }, borderColor: 'FF0000' })]),
    row(2, [cell('A2', { value: '2024-03-05', type: 'date', format: 'yyyy/m/d' }), cell('B2', { value: '1899-12-30 12:30:00', type: 'date', format: 'h:mm:ss' }), cell('C2', { value: '007', type: 'string' })]),
    { kind: 'chart', path: '/sheet[2]/chart[1]', props: { id: '2', type: 'column', title: 'S', categories: 'A2:A3', series: [{ name: 'B1', values: 'B2:B3' }], legend: 'bottom', stacked: true, x: '0cm', y: '0cm', w: '12.7cm', h: '7.62cm' } }
  ]
};
test('live tree → model: parsed json props, "14pt" sizes and chart series survive (they were dropped before)', async () => {
  const m = sheetModel(LIVE_SHEET);
  assert.deepEqual(m.merges, [{ r: 0, c: 0, rs: 1, cs: 3 }]);
  assert.deepEqual(m.colW, { A: 159, B: 229 }); assert.deepEqual(m.rowH, { 2: 40 });
  assert.deepEqual(m.cells.A1.s, { b: true, fs: 14, bd: { top: 'thin', bottom: 'double' }, bdc: '#FF0000' });
  assert.deepEqual(m.cells.C2, { v: '007', s: { fmt: 'text' } }, 'a numeric-looking string stays text');
  assert.deepEqual(m.charts[0].ser, [{ name: 'B1', values: 'B2:B3' }]); assert.equal(m.charts[0].stacked, true);
  const orig = [m];
  assert.deepEqual(await plan(orig, clone(orig)), [], 'nothing edited → nothing written, so nothing read can be lost on save');
});

test('borders: sides taken away are written as none, whole-cell styles as border', () => {
  assert.deepEqual(cellProps({ v: '', s: { bd: { top: 'thin', bottom: 'thick' } } }, { v: '', s: { bd: { top: 'thin' } } }), { borders: '{"bottom":"none"}' });
  assert.deepEqual(cellProps({ v: '', s: { bd: 'thin' } }, { v: '', s: { bd: { top: 'thin' } } }), { borders: '{"right":"none","bottom":"none","left":"none"}' });
  assert.deepEqual(cellProps({ v: '', s: { bd: { top: 'thin' } } }, { v: '', s: { bd: 'medium' } }), { border: 'medium' });
  assert.deepEqual(cellProps({ v: '', s: { bd: 'thin', bdc: '#FF0000' } }, { v: '' }), { border: 'none', borderColor: 'none' });
});

test('widths and heights: only changed keys, null resets a removed one', () => {
  assert.deepEqual(sheetProps({ colW: { A: 100, B: 145 }, rowH: { 3: 40 } }, { colW: { B: 145 }, rowH: { 3: 40, 4: 20 } }), { widths: '{"A":null}', heights: '{"4":15}' });
  assert.deepEqual(sheetProps({ colW: { A: 100 } }, { colW: { A: 100 } }), {});
});

test('charts on two sheets with the same engine id are addressed under their own sheet', async () => {
  const a = sheetModel(TREE_SHEET), b = sheetModel(LIVE_SHEET); b.charts[0].eid = b.charts[0].id = '7';
  const orig = [a, b], cur = clone(orig);
  cur[1].charts[0].title = 'Only the second';
  cur[0].charts = [];
  assert.deepEqual(await plan(orig, cur), [
    ['remove', 'f.xlsx', '/sheet[1]/chart[@id=7]'],
    ['set', 'f.xlsx', '/sheet[@id=2]/chart[@id=7]', '--prop', 'title=Only the second']
  ]);
});

test('the file\'s look: grid lines off both ways, date codes behind a locale tag, the engine\'s date type wins over the code\'s shape', () => {
  assert.equal(fmtOf('[$-409]mmmm d, yyyy;@').fmt, 'date');
  assert.equal(fmtOf('[$-F400]h:mm:ss AM/PM').fmt, 'time');
  assert.equal(fmtOf('[$$-409]#,##0.00').fmt, 'usd');
  const d = cellModel({ value: '1899-12-30 00:12:34', type: 'date', format: 'mm:ss.0' });
  assert.equal(d.s.fmt, 'date'); assert.equal(d.s.code, 'mm:ss.0');
  const m = sheetModel({ kind: 'sheet', path: '/sheet[1]', props: { name: 'S', gridlines: 'false' }, children: [] });
  assert.equal(m.noGrid, true);
  const on = clone(m); on.noGrid = false;
  assert.deepEqual(sheetProps(m, on), { gridlines: 'true' });
  assert.deepEqual(sheetProps(on, m), { gridlines: 'false' });
  assert.deepEqual(sheetProps(m, clone(m)), {});
});
