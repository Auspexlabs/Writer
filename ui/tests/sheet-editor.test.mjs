import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import * as E from '../sheet-engine.js';
globalThis.window = globalThis; // pdf-kit is imported by the bridge and expects a browser global.
const { planXlsx } = await import('../engine.js');

const html = readFileSync(new URL('../SheetEditor.dc.html', import.meta.url), 'utf8');
const script = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)?.[1];
assert.ok(script, 'sheet editor script exists');
const ctx = { document: { documentElement: { dataset: {} } }, // renderVals reads the theme (light here)
  React: { createRef: () => ({ current: null }), createElement: (...a) => a }, DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} }, structuredClone,
  $t: (s, v) => v ? String(s).replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : s, $lang: () => 'zh' }; // i18n globals: a no-op stub, tests assert against the Chinese source text
vm.runInNewContext(script + '\nglobalThis.SheetEditor = Component;', ctx);

const plain = x => JSON.parse(JSON.stringify(x)); // the editor runs in another vm realm: compare values, not prototypes
function editor(sheets) {
  const component = new ctx.SheetEditor();
  let doc = { active: 0, sheets }, changed = 0;
  const messages = [];
  component.E = E;
  component.props = { get doc() { return doc; }, onChange(next) { doc = next; changed++; }, toast(msg) { messages.push(msg); } };
  return { component, get doc() { return doc; }, get changed() { return changed; }, messages };
}

test('inserting a row preserves cells shifted past the 80-row viewport and adjusts references', () => {
  const x = editor([
    { name: 'Data', cells: { A1: { v: '=A80' }, A80: { v: 'last visible' }, A81: { v: 'already offscreen' } } },
    { name: 'Calc', cells: { A1: { v: '=Data!A80' } } }
  ]);
  x.component.insDel('r', 79, 1);
  assert.equal(x.changed, 1);
  assert.equal(x.doc.sheets[0].cells.A81.v, 'last visible');
  assert.equal(x.doc.sheets[0].cells.A82.v, 'already offscreen');
  assert.equal(x.doc.sheets[0].cells.A1.v, '=A81');
  assert.equal(x.doc.sheets[1].cells.A1.v, '=Data!A81');
});

test('inserting a column preserves cells shifted past the 20-column viewport', () => {
  const x = editor([{ name: 'Data', cells: { T1: { v: 'last visible' }, U1: { v: 'already offscreen' } } }]);
  x.component.insDel('c', 19, 1);
  assert.equal(x.changed, 1);
  assert.equal(x.doc.sheets[0].cells.U1.v, 'last visible');
  assert.equal(x.doc.sheets[0].cells.V1.v, 'already offscreen');
});

test('the XLSX save plan writes both cells moved beyond the viewport', async () => {
  const x = editor([{ name: 'Data', path: '/sheet[1]', cells: { A80: { v: 'visible' }, A81: { v: 'offscreen' } } }]);
  const before = structuredClone(x.doc.sheets);
  x.component.insDel('r', 79, 1);
  const commands = [];
  await planXlsx('book.xlsx', before, x.doc.sheets, async argv => { commands.push(argv); return {}; });
  const setCells = commands.filter(cmd => cmd[0] === 'set' && cmd[2].includes('/cell['));
  assert.ok(setCells.some(cmd => cmd[2].endsWith('/cell[A81]') && cmd.includes('value=visible')));
  assert.ok(setCells.some(cmd => cmd[2].endsWith('/cell[A82]') && cmd.includes('value=offscreen')));
});

test('an insertion that exceeds Excel limits is rejected without changing the document', () => {
  const x = editor([{ name: 'Data', cells: { A1048576: { v: 'edge' } } }]);
  x.component.insDel('r', 0, 1);
  assert.equal(x.changed, 0);
  assert.equal(x.doc.sheets[0].cells.A1048576.v, 'edge');
  assert.match(x.messages[0], /上限/);
});

test('an unparseable source address is rejected without changing the document', () => {
  const x = editor([{ name: 'Data', cells: { XFD1: { v: 'edge' } } }]);
  x.component.insDel('r', 0, 1);
  assert.equal(x.changed, 0);
  assert.equal(x.doc.sheets[0].cells.XFD1.v, 'edge');
  assert.match(x.messages[0], /无法移动/);
});

test('formula reference coloring skips quoted text and reuses a color for repeated references', () => {
  const x = editor([{ name: 'Data', cells: {} }]);
  const refs = x.component.refsOf('=SUM(A1:B2,A1:B2)+"C3"+D4');
  assert.deepEqual(Array.from(refs, r => r.sh), [null, null, null]);
  assert.equal(refs[0].col, refs[1].col);
  assert.notEqual(refs[0].col, refs[2].col);
});

test('reference colouring covers whole columns/rows and quoted sheet names, not function names', () => {
  const x = editor([{ name: 'Data', cells: {} }]);
  const refs = Array.from(x.component.refsOf("=SUM(A:A)+COUNT(3:4)+'My Sheet'!B2+LOG10(C1)"), r => [r.sh, r.r1, r.c1, r.r2, r.c2]);
  assert.deepEqual(refs, [[null, 0, 0, 79, 0], [null, 2, 0, 3, 19], ['My Sheet', 1, 1, 1, 1], [null, 0, 2, 0, 2]]);
});

test('typed input gets the value and format Excel would store', () => {
  const x = editor([{ name: 'Data', cells: {} }]), p = (v, s) => JSON.parse(JSON.stringify(x.component.parseInput(v, s)));
  assert.deepEqual(p('2024/3/5'), { v: '2024-03-05', patch: { fmt: 'date' } }, 'dates go to the engine as ISO text');
  assert.deepEqual(p('2024年3月5日'), { v: '2024-03-05', patch: { fmt: 'date' } });
  assert.deepEqual(p('9:30'), { v: '0.395833333333', patch: { fmt: 'time' } });
  assert.deepEqual(p('2024-03-05 12:00'), { v: '2024-03-05 12:00:00', patch: { fmt: 'date', code: 'yyyy/m/d h:mm' } }, 'date-times too: the engine writes the workbook\'s own serial (1900 or 1904)');
  assert.deepEqual(p('12.5%', { fmt: 'pct', dec: 2, code: '0.00%' }), { v: '0.125', patch: null }, 'a percent cell keeps its code, as in Excel');
  assert.deepEqual(p('$7', { fmt: 'usd', dec: 2, code: '"$"#,##0.00_);("$"#,##0.00)' }), { v: '7', patch: null });
  assert.deepEqual(p('12.55%'), { v: '0.1255', patch: { fmt: 'pct', dec: 2 } });
  assert.deepEqual(p('$1,200.50'), { v: '1200.50', patch: { fmt: 'usd', dec: 2 } });
  assert.deepEqual(p('¥1,200'), { v: '1200', patch: { fmt: 'money', dec: 0 } });
  assert.deepEqual(p('2024/3/5', { fmt: 'text' }), { v: '2024/3/5', patch: null }, 'a text cell keeps what was typed');
  assert.deepEqual(p('2024/3/5', { fmt: 'date', code: 'yyyy-mm-dd' }), { v: '2024-03-05', patch: null }, 'an existing date code stays');
  assert.deepEqual(p('hello-world'), { v: 'hello-world', patch: null });
  x.component.setCellRaw(x.doc.sheets[0], 0, 0, '45356.5', { fmt: 'date', code: 'yyyy/m/d h:mm' });
  assert.deepEqual(plain(x.doc.sheets[0].cells.A1.s), { fmt: 'date', code: 'yyyy/m/d h:mm' }, 'a code carried by the patch is kept');
});

test('dark mode draws the file\'s colours as they are; only automatic colours follow the theme', async () => {
  const x = editor([{ name: 'Data', path: '/sheet[1]', cells: { A1: { v: 'red', s: { color: '#FF0000' } }, A2: { v: 'dark red', s: { color: '#C00000' } },
    A3: { v: 'red on yellow', s: { color: '#FF0000', fill: '#FFFF00' } }, A4: { v: 'on yellow', s: { fill: '#FFFF00' } }, A5: { v: 'plain' },
    A6: { v: 'blue border', s: { bd: 'thin', bdc: '#0000FF' } }, A7: { v: 'auto border', s: { bd: 'thin' } } } }]);
  const look = theme => { ctx.document.documentElement.dataset.theme = theme; return Object.fromEntries(x.component.renderVals().cells.filter(c => c.text).map(c => [c.text, [c.color, c.bg, c.bt]])); };
  const dark = look('dark'), light = look('light');
  assert.deepEqual(dark, light, 'the theme reaches the cells only through the CSS variables of their automatic colours');
  assert.deepEqual(dark['red'], ['#FF0000', 'var(--k0, #FFFFFF)', 'none']);
  assert.equal(dark['dark red'][0], '#C00000', 'dark red used to turn white');
  assert.deepEqual(dark['red on yellow'], ['#FF0000', '#FFFF00', 'none'], 'the fill used to be dimmed and the text turned white');
  assert.deepEqual(dark['on yellow'], ['#1D1D1F', '#FFFF00', 'none'], 'automatic text on a colour from the file keeps its light value');
  assert.deepEqual(dark['plain'], ['var(--k7, #1D1D1F)', 'var(--k0, #FFFFFF)', 'none']);
  assert.equal(dark['blue border'][2], '1px solid #0000FF');
  assert.equal(dark['auto border'][2], '1px solid var(--k13, #3A3A3C)', 'an automatic border stays visible on the dark sheet');
  // the border colour menu writes a colour the file can hold (it wrote the theme variable var(--k59, …), which the engine rejects)
  const before = structuredClone(x.doc.sheets), commands = [];
  x.component.state.sel = x.component.state.anc = { r: 6, c: 0 };
  x.component.renderVals(); x.component.menus.bd.find(i => i.label === '绿色').onClick();
  assert.equal(x.doc.sheets[0].cells.A7.s.bdc, '#3F7D5C');
  await planXlsx('book.xlsx', before, x.doc.sheets, async argv => { commands.push(argv); return {}; });
  assert.ok(commands.some(cmd => cmd[2] === '/sheet[1]/cell[A7]' && cmd.includes('borderColor=3F7D5C')));
});

test('every template binding resolves in every ribbon tab, and every ribbon menu has items', () => {
  const names = [...new Set([...html.split('<script type="text/x-dc"')[0].matchAll(/\{\{\s*([A-Za-z_$][\w$]*)\s*\}\}/g)].map(m => m[1]))].filter(n => n !== 'true' && n !== 'false');
  const x = editor([{ name: 'Data', cells: { A1: { v: '1' }, B1: { v: '=A1*2', s: { note: 'n', link: 'https://e.com' } } }, charts: [{ id: 'c1', type: 'column', title: 'T', cat: 'A1:A1', ser: [{ values: 'B1:B1' }], x: 0, y: 0, w: 460, h: 300 }] }]);
  for (const tab of ['home', 'insert', 'formula', 'data', 'view']) {
    x.component.state.tab = tab; x.component.state.pop = null;
    const v = x.component.renderVals();
    assert.deepEqual(names.filter(n => !(n in v)), [], 'unresolved in ' + tab);
    for (const id of Object.keys(x.component.menus)) {
      x.component.state.pop = { id, x: 0, y: 0 };
      assert.ok(x.component.renderVals().popItems.length > 0, `menu ${id} in ${tab} renders items`);
    }
  }
});

test('the active cell is where the selection started, a merge\'s top-left — not the hidden corner of the merge', () => {
  const x = editor([{ name: 'Data', cells: { C8: { v: 'm', s: { b: true } }, D8: { v: '' } }, merges: [{ r: 7, c: 2, rs: 1, cs: 2 }] }]);
  Object.assign(x.component.state, { anc: { r: 7, c: 2 }, sel: { r: 7, c: 3 } }); // what clicking the merged C8:D8 selects
  assert.deepEqual(plain(x.component.act()), { r: 7, c: 2 });
  assert.deepEqual(plain(x.component.actStyle()), { b: true });
  Object.assign(x.component.state, { anc: { r: 1, c: 1 }, sel: { r: 4, c: 3 } }); // a drag from B2 to D5 types into B2, as in Excel
  assert.deepEqual(plain(x.component.act()), { r: 1, c: 1 });
  assert.equal(x.component.renderVals().nameBox, 'B2:D5');
});

test('a new formula takes the date/currency/percent format of the first cell it refers to, like Excel', () => {
  const x = editor([{ name: 'Data', cells: { C10: { v: '2024-03-05', s: { fmt: 'date', code: 'yyyy/m/d' } }, D1: { v: '5', s: { fmt: 'money', dec: 2 } }, E1: { v: '7' } } }, { name: 'Other', cells: { A1: { v: '0.5', s: { fmt: 'pct', dec: 0 } } } }]);
  const p = (v, s) => plain(x.component.parseInput(v, s));
  assert.deepEqual(p('=C10+30'), { v: '=C10+30', patch: { fmt: 'date', code: 'yyyy/m/d' } });
  assert.deepEqual(p('=sum(d1,e1)'), { v: '=SUM(D1,E1)', patch: { fmt: 'money', dec: 2 } });
  assert.deepEqual(p('=Other!A1*2'), { v: '=Other!A1*2', patch: { fmt: 'pct', dec: 0 } });
  assert.deepEqual(p('=E1*2'), { v: '=E1*2', patch: null });
  assert.deepEqual(p('=C10+30', { fmt: 'number', dec: 0 }), { v: '=C10+30', patch: null }, 'a cell that has a format keeps it');
});

test('Enter commits and moves at once, so fast typing fills consecutive cells (a timer used to lose them)', () => {
  const x = editor([{ name: 'Data', cells: {} }]), c = x.component;
  for (const v of ['East', 'West', 'North']) { const a = c.act(); c.state.edit = { r: a.r, c: a.c, val: v, mode: 'type' }; c.commitEdit(1, 0); }
  assert.deepEqual(['A1', 'A2', 'A3'].map(k => x.doc.sheets[0].cells[k] && x.doc.sheets[0].cells[k].v), ['East', 'West', 'North']);
  assert.deepEqual(plain(c.act()), { r: 3, c: 0 });
});

test('chart anchors map between Excel\'s grid (64 × 20 px cells) and the editor\'s roomier grid, both ways', () => {
  const x = editor([{ name: 'Data', cells: {}, colW: { A: 100 } }]), g = (v, a, back) => x.component.gridMap(v, a, back);
  assert.equal(g(100 + 64 * 7, 'x'), 100 + 96 * 7, 'column H in Excel is column H in the editor');
  assert.equal(g(100 + 96 * 7, 'x', true), 100 + 64 * 7);
  assert.equal(g(100 + 64 * 2 + 32, 'x'), 100 + 96 * 2 + 48, 'halfway into a column stays halfway');
  assert.equal(g(40, 'y'), 52); assert.equal(g(52, 'y', true), 40);
});

test('row/column insert and delete keep merges and the autofilter in step (they were dropped or left stale)', () => {
  const x = editor([{ name: 'Data', cells: { A2: { v: 'h' } }, merges: [{ r: 0, c: 0, rs: 5, cs: 1 }, { r: 9, c: 0, rs: 1, cs: 2 }], filter: 'A2:C13' }]);
  x.component.insDel('r', 3, 1); // a row inside the merge and the filter
  assert.deepEqual(plain(x.doc.sheets[0].merges), [{ r: 0, c: 0, rs: 6, cs: 1 }, { r: 10, c: 0, rs: 1, cs: 2 }]);
  assert.equal(x.doc.sheets[0].filter, 'A2:C14');
  x.component.insDel('r', 2, -2); // two rows out of both
  assert.deepEqual(plain(x.doc.sheets[0].merges), [{ r: 0, c: 0, rs: 4, cs: 1 }, { r: 8, c: 0, rs: 1, cs: 2 }]);
  assert.equal(x.doc.sheets[0].filter, 'A2:C12');
  x.component.insDel('c', 1, -1); // column B: the 1×2 merge collapses to one cell and goes
  assert.deepEqual(plain(x.doc.sheets[0].merges), [{ r: 0, c: 0, rs: 4, cs: 1 }]);
  assert.equal(x.doc.sheets[0].filter, 'A2:B12');
  x.component.insDel('r', 0, -20);
  assert.equal(x.doc.sheets[0].filter, null, 'deleting every filtered row removes the filter');
});

test('chart labels reach the SVG as plain text: SVG draws no HTML <span>, so axis labels and pie percentages never showed', () => {
  // walkText is the runtime's {{ … }} text binding; support.js needs a DOM, so the function is cut out of the repository's
  // own file (no outside input reaches the Function body) and given a stub h
  const rt = readFileSync(new URL('../support.js', import.meta.url), 'utf8');
  const walkText = new Function('h', 'getReact', 'compileExpr', rt.slice(rt.indexOf('function walkText('), rt.indexOf('function walkFor(')) + '\nreturn walkText;')(
    (type, props, ...kids) => ({ type, kids }), () => ({ Fragment: 'Fragment', isValidElement: () => false }), e => vals => vals[e.trim()]);
  const label = walkText({ nodeValue: '{{ t }}', parentNode: { namespaceURI: 'http://www.w3.org/2000/svg' } }); // <text>{{ t.text }}</text>
  assert.deepEqual(label({ t: '51%' }).kids, ['', '51%', '']);
  const cell = walkText({ nodeValue: '{{ t }}', parentNode: { namespaceURI: 'http://www.w3.org/1999/xhtml' } });
  assert.equal(cell({ t: 'x' }).kids[1].type, 'span', 'HTML text keeps its sc-interp span');
});

test('a doughnut\'s hole and the gaps between slices take the chart card\'s background, so dark mode shows no white disc', () => {
  const x = editor([{ name: 'Data', cells: { A1: { v: 'a' }, A2: { v: 'b' }, B1: { v: '1' }, B2: { v: '3' } }, charts: [{ id: 'c1', type: 'doughnut', title: 'T', cat: 'A1:A2', ser: [{ values: 'B1:B2' }], x: 0, y: 0, w: 460, h: 300 }] }]);
  const card = html.match(/"\{\{ ch\.onSel \}\}" style="[^"]*?background:([^;]+);/)[1], gap = html.match(/"fill:\{\{ s\.fill \}\};stroke:([^;]+);/)[1];
  assert.equal(card, 'var(--k0, #FFFFFF)', 'the card follows the theme');
  assert.equal(x.component.renderVals().charts[0].slices.at(-1).fill, card, 'the hole');
  assert.equal(gap, card, 'the gaps between slices');
});

test('the grid\'s layers and panes: what scrolls passes under the frozen rows and columns, what is anchored in them sticks with them', () => {
  // rows 1–2 and column A frozen: a chart and a picture in the scrolling pane, a chart in the frozen rows, a picture in the frozen
  // column, one in the frozen corner (A1:A2)
  const chart = (id, x, y) => ({ id, type: 'column', title: 'T', cat: 'A1:A1', ser: [{ values: 'B1:B1' }], x, y, w: 200, h: 150 });
  const pic = (id, x, y) => ({ id, path: '/sheet[1]/image[@id=' + id + ']', src: '/binary?' + id, x, y, w: 40, h: 40, look: {} });
  const x = editor([{ name: 'Data', frR: 2, frC: 1, cells: { A1: { v: '1' }, D8: { v: '2' } }, charts: [chart('body', 300, 200), chart('top', 300, 10)], images: [pic('2', 300, 400), pic('3', 10, 400), pic('4', 10, 10)] }]), c = x.component;
  // the template tag that carries `marker`, or the wrapper around it when it has no z-index of its own
  const tagOf = marker => { const i = html.indexOf(marker), at = html.lastIndexOf('<', i), own = html.slice(at, html.indexOf('>', i)); return /z-index:/.test(own) ? own : html.slice(html.lastIndexOf('<', at - 1), at); };
  // its z-index, a {{ binding }} read from `vals` (the render values, plus a list item)
  const z = (marker, vals) => { const m = /z-index:(?:\{\{ ([\w.]+) \}\}|(\d+))/.exec(tagOf(marker)); return m[2] ? +m[2] : m[1].split('.').reduce((o, k) => o[k], vals); };
  const chartZ = (v, ch) => z('onMouseDown="{{ ch.onSel }}"', { ...v, ch }), picZ = (v, pc) => z('data-pic="{{ pc.id }}"', { ...v, pc });
  // every kind of frame at a cell: the selection and its fill handle, a copy marquee, a formula's reference box, a fill preview
  const frames = (r, col) => {
    c.clip = { si: 0, r1: r, c1: col, r2: r, c2: col };
    Object.assign(c.state, { anc: { r, c: col }, sel: { r, c: col }, edit: { r, c: col, val: '=' + E.A(r, col), mode: 'type' }, fillTo: { r: r + 1, c: col } });
    const v = c.renderVals();
    assert.ok(v.hasClip && v.refBoxes.length === 1 && v.hasFillPrev, 'every frame is drawn');
    return ['left:{{ selL }}', 'onMouseDown="{{ onFillDown }}"', 'left:{{ clipL }}', 'left:{{ fpL }}'].map(m => z(m, v)).concat(z('left:{{ rb.L }}', { ...v, rb: v.refBoxes[0] }));
  };
  const inScrolling = frames(5, 3), inFrozen = frames(0, 1), inCorner = frames(0, 0), v = c.renderVals();
  const pane = k => (k.top !== 'auto' ? 1 : 0) + (k.left !== 'auto' ? 1 : 0), cellsIn = n => v.cells.filter(k => pane(k) === n).map(k => k.z);
  const tiers = [
    ['cells of the scrolling pane', cellsIn(0)],
    ['its frames, picture and chart', [...inScrolling, chartZ(v, v.charts[0]), picZ(v, v.pics[0])]],
    ['cells of the frozen rows and column', cellsIn(1)],
    ['their frames, chart and picture', [...inFrozen, chartZ(v, v.charts[1]), picZ(v, v.pics[1])]],
    ['cells of the frozen corner', cellsIn(2)],
    ['its frames and picture', [...inCorner, picZ(v, v.pics[2])]],
    ['row and column headers', [...v.colHeads, ...v.rowHeads].map(h => h.z)],
    ['corner', [z('onMouseDown="{{ selAll }}"', v)]],
    ['cell editor, fill hint, note card, find bar, ribbon, menus', ['ref="{{ inMirRef }}"', 'ref="{{ inRef }}"', 'ref="{{ tipRef }}"', 'data-note="1"', 'top:104px;right:24px', 'data-bubble="1"', 'data-menu="1"'].map(m => z(m, v))]];
  const zs = t => [...new Set(t[1])].join();
  for (let i = 1; i < tiers.length; i++) assert.ok(Math.max(...tiers[i - 1][1]) < Math.min(...tiers[i][1]), `${tiers[i - 1][0]} (${zs(tiers[i - 1])}) go under ${tiers[i][0]} (${zs(tiers[i])})`);
  // sticking: every frame and object sits in a zero-size wrapper, sticky along the frozen axes it lies in (objects by their top-left corner)
  for (const m of ['left:{{ selL }}', 'onMouseDown="{{ onFillDown }}"', 'left:{{ clipL }}', 'left:{{ fpL }}', 'left:{{ rb.L }}', 'data-pic="{{ pc.id }}"', 'onMouseDown="{{ ch.onSel }}"'])
    assert.match(tagOf(m), /^<div style="position:sticky;top:\{\{ [\w.]+ \}\};left:\{\{ [\w.]+ \}\};width:0;height:0;z-index:/, m);
  const stick = o => [String(o.st), String(o.sl)];
  assert.deepEqual([v.charts[0], v.charts[1], v.pics[1], v.pics[2]].map(stick), [['auto', 'auto'], ['0', 'auto'], ['auto', '0'], ['0', '0']], 'scrolling, frozen rows, frozen column, corner');
  const sel = (r1, c1, r2, c2) => { Object.assign(c.state, { anc: { r: r1, c: c1 }, sel: { r: r2, c: c2 }, edit: null, fillTo: null }); c.clip = null; const w = c.renderVals(); return [w.selST, w.selSL, w.selZ, w.selL, w.selT, w.selW, w.selH].map(String); };
  assert.deepEqual(sel(5, 3, 5, 3).slice(0, 2), ['auto', 'auto']);
  assert.deepEqual(sel(0, 1, 0, 3).slice(0, 3), ['0', 'auto', String(inFrozen[0])], 'B1:D1 sticks with the frozen rows');
  assert.deepEqual(sel(0, 1, 9, 1).slice(0, 3), ['auto', 'auto', String(inFrozen[0])], 'B1:B10 reaches into the scrolling rows: it scrolls, over the frozen cells');
  // at rest nothing reaches under what covers it: a frame straddles the gridlines, but where the headers or a higher pane's cells
  // cover its top or left edge (the headers are 44 × 26 px) it starts inside the cell; its far edges stay put
  assert.deepEqual(sel(0, 0, 0, 0).slice(3), ['44px', '26px', '97px', '27px'], 'A1, at the headers');
  assert.deepEqual(sel(0, 1, 0, 1).slice(3), ['140px', '26px', '97px', '27px'], 'B1, beside the frozen corner');
  assert.deepEqual(sel(2, 1, 2, 1).slice(3), ['140px', '78px', '97px', '27px'], 'B3, the first cell of the scrolling pane');
});

test('chart grid lines take the chart card\'s own line colour, which follows the theme (#EFEFF4 glared on the dark sheet)', () => {
  const x = editor([{ name: 'Data', cells: { A1: { v: 'a' }, B1: { v: '5' } }, charts: [{ id: 'c1', type: 'column', title: 'T', cat: 'A1:A1', ser: [{ values: 'B1:B1' }], x: 0, y: 0, w: 460, h: 300 }] }]);
  const ch = x.component.renderVals().charts[0], strokes = [...new Set(ch.grid.map(g => g.stroke))];
  assert.deepEqual(strokes, ['var(--k47, #C7C7CC)', ch.border], 'the zero line, then the grid in the card border\'s colour');
  assert.equal(ch.border, 'var(--k6, #E5E5EA)');
});

test('the selected cell\'s row and column headers stay opaque, so cells scrolled under them (zooming scrolls) never show through', () => {
  const x = editor([{ name: 'Data', cells: { F2: { v: '357' } } }]), dark = readFileSync(new URL('../theme-dark.css', import.meta.url), 'utf8');
  Object.assign(x.component.state, { anc: { r: 10, c: 5 }, sel: { r: 10, c: 5 } });
  const v = x.component.renderVals();
  for (const h of [v.colHeads[5], v.rowHeads[10], v.colHeads[0]]) {
    const base = /var\((--k\d+), #[0-9A-F]{6}\)$/i.exec(h.bg); // the bottom layer: a solid colour in light mode…
    assert.ok(base, h.bg);
    assert.match(dark.match(new RegExp(base[1] + ':([^;]+);'))[1], /^#[0-9A-F]{6}$/i, '…and in dark mode');
  }
});

test('conditional formats: Excel-shaped rules colour cells, the first rule that holds wins, presets come from the menu dialog, rules move with inserted rows and clear by selection', () => {
  const x = editor([{ name: 'Data', cells: { A1: { v: '5' }, A2: { v: '50' }, A3: { v: '50' }, B1: { v: 'an error here' }, C1: { v: '1' }, C2: { v: '3' } },
    cf: [{ range: 'A1:A3', type: 'cellIs', operator: 'greaterThan', value: '10', fill: '#C6EFCE', color: '#006100' }, { range: 'A1:A3', type: 'duplicateValues' }, { range: 'B1', type: 'containsText', text: 'ERROR', bold: true }, { range: 'C1:C2', type: 'colorScale', colors: ['#FFFFFF', '#63BE7B'] }] }]), c = x.component;
  const look = () => Object.fromEntries(c.renderVals().cells.filter(k => k.text).map(k => [k.text, [k.bg, k.color, k.fw]]));
  const L = look();
  assert.deepEqual(L['50'], ['#C6EFCE', '#006100', 400], 'greater-than (green, first) wins over duplicate values (the default red look)');
  assert.deepEqual(L['5'], ['var(--k0, #FFFFFF)', 'var(--k7, #1D1D1F)', 400]);
  assert.deepEqual(L['an error here'], ['var(--k0, #FFFFFF)', 'var(--k7, #1D1D1F)', 700], 'text contains is case-insensitive; a look of only bold changes only the weight');
  assert.deepEqual([L['1'][0], L['3'][0]], ['#ffffff', '#63be7b'], 'a two-colour scale runs from the smallest to the largest');
  Object.assign(c.state, { anc: { r: 0, c: 2 }, sel: { r: 1, c: 2 } }); c.renderVals(); c.menus.cf.find(i => i.label === '大于…').onClick();
  c.state.dlg.fields[0].value = '2'; c.state.dlg.fields[1].value = 'yellow'; c.dlgOk();
  assert.deepEqual(plain(x.doc.sheets[0].cf.at(-1)), { range: 'C1:C2', type: 'cellIs', operator: 'greaterThan', value: '2', fill: '#FFEB9C', color: '#9C5700' });
  assert.equal(look()['3'][0], '#63be7b', 'the earlier colour scale still wins');
  c.cfManage(); c.state.dlg.fields[0].value = '4'; c.state.dlg.fields[1].value = 'up'; c.dlgOk();
  assert.equal(look()['3'][0], '#FFEB9C', 'moved up, the new rule wins');
  c.insDel('r', 0, 1);
  assert.deepEqual(plain(x.doc.sheets[0].cf.map(r => r.range)), ['A2:A4', 'A2:A4', 'B2', 'C2:C3', 'C2:C3']);
  Object.assign(c.state, { anc: { r: 1, c: 1 }, sel: { r: 1, c: 1 } }); c.renderVals(); c.menus.cf.find(i => i.label === '清除所选单元格的规则').onClick();
  assert.deepEqual(plain(x.doc.sheets[0].cf.map(r => r.type)), ['cellIs', 'duplicateValues', 'cellIs', 'colorScale']);
});

test('data validation: rules by range as the file keeps them, the list arrow and picker on the active cell, Excel\'s dialog, error texts', () => {
  const x = editor([{ name: 'Data', cells: { H1: { v: 'East' }, H2: { v: 'West' } }, dv: [{ range: 'A1:A5', type: 'list', values: ['Yes', 'No'], error: 'Yes or No' }, { range: 'B1:B5', type: 'whole', operator: 'between', value: '1', value2: '10' }, { range: 'C1:C5', type: 'list', source: 'H1:H2' }, { range: 'D1', type: 'date', operator: 'greaterThan', value: '45292' }] }]), c = x.component;
  assert.equal(c.dvBad(0, 0, 'Maybe'), 'Yes or No'); assert.equal(c.dvBad(0, 0, 'Yes'), null); assert.equal(c.dvBad(0, 0, ''), null, 'blank passes');
  assert.match(c.dvBad(0, 1, '11'), /1 到 10/); assert.match(c.dvBad(0, 1, '2.5'), /整数/); assert.equal(c.dvBad(0, 1, '7'), null);
  assert.equal(c.dvBad(0, 2, 'West'), null); assert.match(c.dvBad(0, 2, 'North'), /East、West/);
  assert.equal(c.dvBad(0, 3, '2024-03-05'), null); assert.match(c.dvBad(0, 3, '2023-12-31'), /大于 2024\/1\/1/);
  Object.assign(c.state, { anc: { r: 0, c: 2 }, sel: { r: 0, c: 2 } });
  let v = c.renderVals();
  assert.ok(v.cells.find(k => k.gr === '2' && k.gc === '4').hasF, 'the arrow on the active list cell'); assert.ok(!v.cells.find(k => k.gr === '3' && k.gc === '4').hasF, 'not on the others');
  c.state.pop = { id: 'dv', x: 0, y: 0 }; v = c.renderVals(); assert.deepEqual(plain(v.popItems.map(i => i.label)), ['East', 'West']);
  v.popItems[1].onClick(); assert.equal(x.doc.sheets[0].cells.C1.v, 'West');
  Object.assign(c.state, { anc: { r: 0, c: 1 }, sel: { r: 4, c: 1 }, pop: null }); c.dvSet();
  const f = k => c.state.dlg.fields.find(y => y.key === k); assert.deepEqual([f('t').value, f('o').value, f('v').value, f('v2').value], ['whole', 'between', '1', '10']);
  f('o').value = 'greaterThan'; f('v').value = '0'; f('e').value = 'Positive only'; c.dlgOk();
  assert.deepEqual(plain(x.doc.sheets[0].dv.filter(r => r.range === 'B1:B5')), [{ range: 'B1:B5', type: 'whole', operator: 'greaterThan', value: '0', error: 'Positive only' }], 'the rule the selection covers is replaced');
  assert.equal(c.dvBad(0, 1, '-3'), 'Positive only');
});

test('filter criteria hide rows until the filter is applied again, the dropdown offers sort, text/number criteria and values, the sort dialog takes three keys', () => {
  const x = editor([{ name: 'Data', filter: 'A1:C5', cells: { A1: { v: 'Region' }, B1: { v: 'Item' }, C1: { v: 'Sales' }, A2: { v: 'East' }, B2: { v: 'Pen' }, C2: { v: '120' }, A3: { v: 'West' }, B3: { v: 'Pencil' }, C3: { v: '80' }, A4: { v: 'East' }, B4: { v: 'Ink' }, C4: { v: '200' }, A5: { v: 'West' }, B5: { v: 'Pad' }, C5: { v: '20' } } }]), c = x.component;
  c.setFilters({ A: { values: ['East'] } }); assert.deepEqual(plain(x.doc.sheets[0].frows), [2, 4]);
  c.setFilters({ A: { values: ['East'] }, C: { operator: 'greaterThan', value: '150' } }); assert.deepEqual(plain(x.doc.sheets[0].frows), [1, 2, 4]);
  c.setFilters({ B: { operator: 'contains', value: 'pen' } }); assert.deepEqual(plain(x.doc.sheets[0].frows), [3, 4], 'text criteria ignore case');
  assert.ok(!c.renderVals().rowHeads.some(h => h.n === 4 || h.n === 5), 'hidden rows are not drawn');
  c.state.pop = { id: 'flt0', x: 0, y: 0 }; let items = c.renderVals().popItems;
  assert.deepEqual(plain(items.filter(i => i.isItem && i.check === '✓').map(i => i.label)), ['（全选）', 'East', 'West'], 'no criterion on A: every value shows');
  items.find(i => i.label === 'West').onClick(); assert.deepEqual(plain(x.doc.sheets[0].filters), { B: { operator: 'contains', value: 'pen' }, A: { values: ['East'] } });
  c.state.pop = { id: 'flt2', x: 0, y: 0 }; items = c.renderVals().popItems; items.find(i => i.label === '介于…').onClick();
  c.state.dlg.fields[0].value = '50'; c.state.dlg.fields[1].value = '150'; c.dlgOk();
  assert.deepEqual(plain(x.doc.sheets[0].filters.C), { operator: 'between', value: '50', value2: '150' }); assert.deepEqual(plain(x.doc.sheets[0].frows), [2, 3, 4]);
  c.commit(sh => { sh.cells.A5.v = 'East'; sh.cells.B5.v = 'Pen'; sh.cells.C5.v = '100'; });
  assert.deepEqual(plain(x.doc.sheets[0].frows), [2, 3, 4], 'an edit moves no row until 重新应用');
  c.setFilters(x.doc.sheets[0].filters); assert.deepEqual(plain(x.doc.sheets[0].frows), [2, 3]);
  c.toggleFilter(); assert.deepEqual([x.doc.sheets[0].filter, plain(x.doc.sheets[0].filters), plain(x.doc.sheets[0].frows)], [null, {}, []], 'turning the filter off clears its criteria and rows');
  Object.assign(c.state, { anc: { r: 0, c: 0 }, sel: { r: 4, c: 2 } }); c.sortCustom();
  assert.deepEqual(plain(c.state.dlg.fields[0].options.map(o => o[1])), ['Region', 'Item', 'Sales'], 'keys are named by the header row');
  c.state.dlg.fields[0].value = 'A'; c.state.dlg.fields[2].value = 'C'; c.state.dlg.fields[3].value = 'desc'; c.dlgOk();
  assert.deepEqual([2, 3, 4, 5].map(r => x.doc.sheets[0].cells['A' + r].v + ' ' + x.doc.sheets[0].cells['C' + r].v), ['East 200', 'East 120', 'East 100', 'West 80']);
});

test('freeze panes: 首行 / 首列 / 至当前单元格 / 取消 from the view tab', () => {
  const x = editor([{ name: 'Data', cells: {} }]), c = x.component; c.state.tab = 'view';
  Object.assign(c.state, { anc: { r: 2, c: 1 }, sel: { r: 2, c: 1 } }); c.renderVals();
  const item = l => c.menus.frz.find(i => i.label === l), panes = () => [x.doc.sheets[0].frR, x.doc.sheets[0].frC];
  item('冻结至当前单元格').onClick(); assert.deepEqual(panes(), [2, 1]);
  c.renderVals(); item('冻结首行').onClick(); assert.deepEqual(panes(), [1, 0]);
  c.renderVals(); item('冻结首列').onClick(); assert.deepEqual(panes(), [0, 1]);
  c.renderVals(); assert.equal(item('冻结首列').check, '✓'); item('取消冻结').onClick(); assert.deepEqual(panes(), [0, 0]);
});

test('number formats: 货币 €, the date and time codes offered by example, all written as codes', async () => {
  const before = [{ name: 'Data', path: '/sheet[1]', cells: { A1: { v: '1234.5' }, A2: { v: '45292.75' } } }];
  const x = editor(structuredClone(before)), c = x.component;
  c.renderVals(); const item = l => c.menus.fmt.find(i => i.label === l), text = a => c.renderVals().cells.find(k => k.gr === String(E.parseA(a).r + 2) && k.gc === String(E.parseA(a).c + 2)).text;
  item('货币 €').onClick(); assert.deepEqual(plain(x.doc.sheets[0].cells.A1.s), { fmt: 'eur' }); assert.equal(text('A1'), '€1,234.50');
  Object.assign(c.state, { anc: { r: 1, c: 0 }, sel: { r: 1, c: 0 } }); c.renderVals();
  item('2024年1月1日').onClick(); assert.deepEqual(plain(x.doc.sheets[0].cells.A2.s), { fmt: 'date', code: 'yyyy年m月d日' }); assert.equal(text('A2'), '2024年1月1日');
  c.renderVals(); item('12:00 PM').onClick(); assert.deepEqual(plain(x.doc.sheets[0].cells.A2.s), { fmt: 'time', code: 'h:mm AM/PM' }); assert.equal(text('A2'), '6:00 PM');
  const commands = []; await planXlsx('book.xlsx', before, x.doc.sheets, async argv => { commands.push(argv); return {}; });
  assert.ok(commands.some(cmd => cmd[2] === '/sheet[1]/cell[A1]' && cmd.includes('format="€"#,##0.00')));
  assert.ok(commands.some(cmd => cmd[2] === '/sheet[1]/cell[A2]' && cmd.includes('format=h:mm AM/PM')));
});

test('cell styling: inner borders, text rotation and the fill colour picker, all written to the file', async () => {
  const before = [{ name: 'Data', path: '/sheet[1]', cells: { A1: { v: 'a' } } }], x = editor(structuredClone(before)), c = x.component;
  Object.assign(c.state, { anc: { r: 0, c: 0 }, sel: { r: 1, c: 1 } }); c.renderVals(); c.menus.bd.find(i => i.label === '内部框线').onClick();
  const bd = a => plain(x.doc.sheets[0].cells[a].s.bd);
  assert.deepEqual([bd('A1'), bd('B1'), bd('A2'), bd('B2')], [{ right: 'thin', bottom: 'thin' }, { left: 'thin', bottom: 'thin' }, { top: 'thin', right: 'thin' }, { top: 'thin', left: 'thin' }]);
  c.renderVals(); c.menus.rot.find(i => i.label === '逆时针 45°').onClick(); assert.equal(x.doc.sheets[0].cells.A1.s.rotate, 45);
  assert.equal(c.renderVals().cells.find(k => k.gr === '2' && k.gc === '2').tr, 'rotate(-45deg)');
  c.renderVals().ribbon.find(i => i.isColor && i.title === '填充颜色').onChange({ target: { value: '#ffff00' } }); assert.equal(x.doc.sheets[0].cells.A1.s.fill, '#ffff00');
  const commands = []; await planXlsx('book.xlsx', before, x.doc.sheets, async argv => { commands.push(argv); return {}; });
  const a1 = commands.find(cmd => cmd[2] === '/sheet[1]/cell[A1]');
  assert.ok(a1.includes('rotate=45') && a1.includes('fill=FFFF00') && a1.includes('borders={"right":"thin","bottom":"thin"}'), a1.join(' '));
});

test('rows and columns: the header menu hides, shows, sizes and fits lines; hidden lines skip drawing and arrow moves, shift with insertions and reach the file', async () => {
  const x = editor([{ name: 'Data', path: '/sheet[1]', cells: { A1: { v: 'x' }, B1: { v: 'hidden col' }, C1: { v: 'z' } } }]), c = x.component;
  Object.assign(c.state, { anc: { r: 0, c: 1 }, sel: { r: 79, c: 1 }, pop: { id: 'colctx', x: 0, y: 0 } });
  let v = c.renderVals(); assert.deepEqual(plain(v.popItems.filter(i => i.isItem).map(i => i.label)), ['插入列', '删除列', '清除内容', '列宽…', '自动调整列宽', '隐藏', '取消隐藏']);
  v.popItems.find(i => i.label === '隐藏').onClick(); assert.deepEqual(plain(x.doc.sheets[0].hiddenCols), [1]);
  v = c.renderVals(); assert.ok(!v.colHeads.some(h => h.l === 'B') && !v.cells.some(k => k.text === 'hidden col'), 'a hidden column is not drawn'); assert.equal(v.gtc.split(' ')[2], '0px');
  Object.assign(c.state, { anc: { r: 0, c: 0 }, sel: { r: 0, c: 0 } }); c.move(0, 1); assert.deepEqual(plain(c.state.sel), { r: 0, c: 2 }, 'the arrow skips it');
  c.insDel('c', 0, 1); assert.deepEqual(plain(x.doc.sheets[0].hiddenCols), [2], 'an inserted column shifts it');
  Object.assign(c.state, { anc: { r: 0, c: 1 }, sel: { r: 0, c: 3 } }); c.hideLines('c', false); assert.deepEqual(plain(x.doc.sheets[0].hiddenCols), []);
  Object.assign(c.state, { anc: { r: 2, c: 0 }, sel: { r: 3, c: 0 } }); c.hideLines('r', true); assert.deepEqual(plain(x.doc.sheets[0].hiddenRows), [2, 3]);
  c.sizeDialog('r'); c.state.dlg.fields[0].value = '40'; c.dlgOk(); assert.deepEqual(plain(x.doc.sheets[0].rowH), { 3: 40, 4: 40 });
  Object.assign(c.state, { anc: { r: 0, c: 2 }, sel: { r: 0, c: 2 } }); c.fitCols(); assert.equal(x.doc.sheets[0].colW.C, 'hidden col'.length * 8 + 18);
  const commands = []; await planXlsx('book.xlsx', [{ name: 'Data', path: '/sheet[1]', cells: {} }], x.doc.sheets, async argv => { commands.push(argv); return {}; });
  const sheet = commands.find(cmd => cmd[2] === '/sheet[1]'); assert.ok(sheet.includes('hidden={"rows":[3,4],"cols":[]}'), sheet.join(' '));
});

test('find & replace: match case, whole cell and workbook scope', () => {
  const x = editor([{ name: 'One', cells: { A1: { v: 'Apple pie' }, A2: { v: 'apple' } } }, { name: 'Two', cells: { B2: { v: 'apple' } } }]), c = x.component;
  c.state.fq = 'apple'; c.findNext(); assert.deepEqual([c.state.fcount, plain(c.state.sel)], [2, { r: 1, c: 0 }], 'the next hit after the active cell');
  c.state.fopt = { mc: true }; c.findNext(); assert.deepEqual([c.state.fcount, plain(c.state.sel)], [1, { r: 1, c: 0 }]);
  c.state.fopt = { whole: true }; Object.assign(c.state, { anc: { r: 0, c: 0 }, sel: { r: 0, c: 0 } }); c.findNext(); assert.deepEqual([c.state.fcount, plain(c.state.sel)], [1, { r: 1, c: 0 }]);
  c.state.fopt = { book: true }; c.findNext(); assert.deepEqual([c.state.fcount, x.doc.active, plain(c.state.sel)], [3, 1, { r: 1, c: 1 }], 'from A2 the next hit is on the second sheet');
  c.state.fr = 'pear'; c.replace(true); assert.deepEqual([x.doc.sheets[0].cells.A1.v, x.doc.sheets[0].cells.A2.v, x.doc.sheets[1].cells.B2.v], ['pear pie', 'pear', 'pear']);
});

test('a chart resizes from its corner handle: the drag shows live, the drop writes the size', () => {
  const x = editor([{ name: 'Data', cells: { A1: { v: 'a' }, B1: { v: '5' } }, charts: [{ id: 'c1', type: 'column', title: 'T', cat: 'A1:A1', ser: [{ values: 'B1:B1' }], x: 0, y: 0, w: 460, h: 300 }] }]), c = x.component;
  c.renderVals().charts[0].onResize({ preventDefault() { }, stopPropagation() { }, clientX: 100, clientY: 100 });
  c.onWM({ clientX: 160, clientY: 140 }); assert.deepEqual(plain(c.state.live), { kind: 'chartsz', id: 'c1', w: 520, h: 340 });
  assert.equal(c.renderVals().charts[0].w, '520px', 'drawn at the live size');
  c.onWU({}); assert.deepEqual([x.doc.sheets[0].charts[0].w, x.doc.sheets[0].charts[0].h, c.state.live], [520, 340, null]);
});

test('sheet tabs: a colour stripe, drag reorder, delete asks first', () => {
  const x = editor([{ name: 'A', cells: {} }, { name: 'B', cells: {} }, { name: 'C', cells: {} }]), c = x.component;
  c.state.pop = { id: 'tabctx', i: 1, x: 0, y: 0 }; let items = c.renderVals().popItems;
  items.find(i => i.label === '绿色').onClick(); assert.equal(x.doc.sheets[1].color, '#27AE60');
  assert.match(c.renderVals().sheetTabs[1].line, /inset 0 -3px 0 #27AE60/);
  const tabs = c.renderVals().sheetTabs; tabs[0].onMD({ button: 0 }); tabs[2].onME(); assert.deepEqual([plain(x.doc.sheets.map(s => s.name)), x.doc.active], [['B', 'C', 'A'], 2]); c.op = null;
  c.state.pop = { id: 'tabctx', i: 0, x: 0, y: 0 }; items = c.renderVals().popItems; items.find(i => i.label === '删除工作表').onClick();
  assert.equal(x.doc.sheets.length, 3, 'nothing goes before the dialog is confirmed'); assert.match(c.state.dlg.title, /删除工作表「B」/);
  c.dlgOk(); assert.deepEqual(plain(x.doc.sheets.map(s => s.name)), ['C', 'A']);
});

test('Excel\'s keys: Enter after a run of Tabs returns to the column the run began in; Ctrl+Space and Shift+Space select columns and rows; Ctrl+Shift+L toggles the filter', () => {
  const x = editor([{ name: 'Data', cells: { A1: { v: '1' }, B1: { v: '2' } } }]), c = x.component, key = (k, o = {}) => c.renderVals().onInKey(Object.assign({ key: k, preventDefault() { }, nativeEvent: {} }, o));
  Object.assign(c.state, { anc: { r: 0, c: 1 }, sel: { r: 0, c: 1 } });
  key('Tab'); key('Tab'); key('Enter'); assert.deepEqual(plain(c.state.sel), { r: 1, c: 1 });
  key('ArrowRight'); key('Enter'); assert.deepEqual(plain(c.state.sel), { r: 2, c: 2 }, 'an arrow ends the run');
  key(' ', { ctrlKey: true }); assert.deepEqual([plain(c.state.anc), plain(c.state.sel)], [{ r: 0, c: 2 }, { r: 79, c: 2 }]);
  Object.assign(c.state, { anc: { r: 3, c: 1 }, sel: { r: 4, c: 1 } }); key(' ', { shiftKey: true }); assert.deepEqual([plain(c.state.anc), plain(c.state.sel)], [{ r: 3, c: 0 }, { r: 4, c: 19 }]);
  Object.assign(c.state, { anc: { r: 0, c: 0 }, sel: { r: 0, c: 0 } }); key('l', { ctrlKey: true, shiftKey: true }); assert.equal(x.doc.sheets[0].filter, 'A1:B1');
});

test('the fill handle runs dates on as dates: ISO text in and out, across a month end and by whole months', () => {
  const d = v => ({ v, s: { fmt: 'date', code: 'yyyy-mm-dd' } });
  const x = editor([{ name: 'Data', cells: { A1: d('2024-01-30'), A2: d('2024-01-31'), B1: d('2024-01-15'), B2: d('2024-02-15'), C1: d('2024-02-28') } }]);
  const fill = (c, r2, to) => plain(x.component.previewFill(x.doc.sheets[0].cells, { r1: 0, c1: c, r2, c2: c }, { r: to, c }, false).items.map(i => i.cell.v));
  assert.deepEqual(fill(0, 1, 3), ['2024-02-01', '2024-02-02']);
  assert.deepEqual(fill(1, 1, 3), ['2024-03-15', '2024-04-15']);
  assert.deepEqual(fill(2, 0, 2), ['2024-02-29', '2024-03-01']);
});
