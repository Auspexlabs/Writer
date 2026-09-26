// node --test ui/tests/ — the right-hand format panel of the 定稿 design: the Word editor's tabs and the groups each shows, the
// controls running the editor's own commands, the page and spacing helpers behind the fields, and FormatPanel.dc.html turning
// the plain items into its controls.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const FP = await import('../panel.js');
const K = await import('../office-io.js');
const script = f => readFileSync(new URL('../' + f, import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const base = () => ({ location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, getComputedStyle: () => ({ marginBottom: '28px' }), window: { innerWidth: 1200 },
  document: { querySelector: () => null, querySelectorAll: () => [], documentElement: { dataset: {} }, activeElement: null }, localStorage: { getItem: () => null, setItem() { } },
  $t: (s, v) => { s = String(s).replace(/@@.*$/, ''); return v ? s.replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : s; }, $lang: () => 'zh',
  React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } });

const ctx = base();
vm.runInNewContext(script('WordEditor.dc.html') + '\nglobalThis.WordEditor = Component; globalThis.__T = { marginsCm, marginText, marginPresets, marginPreset, spaceText, spaceStep, spaceParse, rgbHex, Z100, colName };', ctx);
const H = ctx.__T;
/** A Word editor with the panel open (no page on screen): what it calls is recorded in `calls`. */
function word(doc = {}, props = {}) {
  const c = new ctx.WordEditor(); c.props = Object.assign({ doc: Object.assign({ id: 'd', html: '' }, doc), onChange: d => c.calls.push(['change', d]), formatOpen: true, showThumbs: false }, props);
  c.FP = FP; c.calls = []; c.pgCss = { textContent: '' };
  c.edRef.current = { innerText: '', children: [], querySelectorAll: () => [], querySelector: () => null, contains: () => false };
  for (const name of ['exec', 'para', 'tbl', 'setNoteFormat', 'setTocStyle', 'insertToc', 'setZoom']) c[name] = (...a) => c.calls.push([name, ...a]);
  c.refreshInfo();
  return c;
}
const plain = x => JSON.parse(JSON.stringify(x)); // the editor runs in another realm: compare values
const titles = v => Array.from(v.panelGroups, g => g.title);
const group = (v, title) => v.panelGroups.find(g => g.title === title);
const items = g => g.rows.flatMap(r => Array.from(r.items));
const find = (v, title, pick) => items(group(v, title)).find(pick);
const at = (c, tab, fmt) => { c.state.tab = tab; if (fmt) c.state.fmt = Object.assign({}, c.state.fmt, fmt); return c.renderVals(); };

test('the Word panel has the design\'s six tabs and, on each, its groups in order; 表格 joins in a table as a context tab', () => {
  const c = word();
  const design = { home: ['样式', '字体', '颜色', '对齐与缩进', '间距', '列表', '边框与底纹', '换行和分页', '工具'], insert: ['常用', '页面', '符号与批注'], layout: ['纸张', '页边距', '分栏与分隔', '页面'],
    refs: ['目录', '脚注和尾注', '题注', '交叉引用', '书签'], review: ['字数统计', '批注', '修订', '查找与替换'], view: ['视图', '缩放', '显示', '翻页'] };
  for (const [tab, want] of Object.entries(design)) {
    const v = at(c, tab);
    assert.deepEqual(Array.from(v.panelTabs, t => t.label), ['开始', '插入', '布局', '引用', '审阅', '视图']);
    assert.equal(v.panelTabs.find(t => t.on).key, tab);
    assert.deepEqual(titles(v), want, tab);
  }
  const v = at(c, 'table', { inTable: true, tbl: { style: 'GridTable4Accent1', header: 'true' }, row: {}, cell: {} });
  assert.deepEqual(plain(v.panelTabs.slice(-1).map(t => [t.label, t.on, t.ctx])), [['表格', true, true]]);
  assert.deepEqual(titles(v), ['表格样式', '行和列', '单元格', '更多']);
  assert.deepEqual(plain(items(group(v, '表格样式')).filter(x => x.t === 'tsty').map(x => [x.title, x.on])), [['网格表 4（彩色标题行）', true], ['简明表格（隔行底纹）', false], ['网格表 4（橙色标题行）', false]], 'the design\'s three');
  assert.deepEqual(plain(c.menus['cbd-p'].slice(-4).map(m => m.label)), ['整个表格', '网格', '三线表', '无框线'], 'the other styles are in 边框');
  assert.deepEqual(plain(c.menus.tprops.map(m => m.label)), ['表格属性…', '删除表格']);
  assert.deepEqual(plain(items(group(v, '更多')).map(x => x.label)), ['排序', '转文本', '属性'], 'no extra link under the group');
  assert.deepEqual(plain(items(group(at(c, 'review'), '查找与替换')).filter(x => x.t === 'btn').map(x => x.label)), ['下一个', '全部替换'], 'Return in 替换为 replaces one');
  assert.equal(find(at(c, 'view'), '缩放', x => x.t === 'size').value, '100%', '100% is the design\'s page, A4 680px wide');
  assert.equal(c.state.zoom, H.Z100);
  const closed = word({}, { formatOpen: false }).renderVals();
  assert.deepEqual([closed.formatOpen, closed.panelTabs.length, closed.panelGroups.length], [false, 0, 0], 'a closed panel builds nothing');
});

test('the panel\'s controls run the editor\'s commands', () => {
  const c = word({ noteFormat: 'lowerRoman' });
  let v = at(c, 'home', { b: true, pat: {} });
  const bius = find(v, '字体', x => x.t === 'seg' && x.opts.some(o => o.title === '加粗'));
  assert.deepEqual(plain(bius.opts.map(o => [o.title, !!o.on])), [['加粗', true], ['斜体', false], ['下划线', false], ['删除线', false]]);
  bius.opts[1].onClick(); assert.deepEqual(c.calls.pop(), ['exec', 'italic']);
  // 孤行控制: the paragraph's own setting, else its style's; a click writes only what differs from the style
  let widow = find(v, '换行和分页', x => x.label === '孤行控制');
  assert.equal(widow.on, true); widow.onChange(false); assert.deepEqual(c.calls.pop(), ['para', 'widowControl', 'false']);
  v = at(c, 'home', { widowOff: true, pat: {} }); widow = find(v, '换行和分页', x => x.label === '孤行控制');
  assert.equal(widow.on, false); widow.onChange(true); assert.deepEqual(c.calls.pop(), ['para', 'widowControl', 'true']);
  v = at(c, 'home', { widowOff: true, pat: { widowcontrol: 'true' } }); widow = find(v, '换行和分页', x => x.label === '孤行控制');
  assert.equal(widow.on, true); widow.onChange(false); assert.deepEqual(c.calls.pop(), ['para', 'widowControl', ''], 'back to what the style says');
  // 段前 in lines: the stepper moves half a line, typing takes 行 or 磅
  const before = find(at(c, 'home', { pat: { spacebefore: '0.5lines' }, widowOff: false }), '间距', x => x.t === 'num' && x.title === '段前');
  assert.equal(before.value, '0.5 行');
  before.onStep(1); assert.deepEqual(c.calls.pop(), ['para', 'spaceBefore', '1lines']);
  before.onSet('12 磅'); assert.deepEqual(c.calls.pop(), ['para', 'spaceBefore', '12pt']);
  // 页边距: the four sides, stepping 0.1 cm; a preset's values come back as its name
  v = at(c, 'layout');
  const top = find(v, '页边距', x => x.t === 'num' && x.lab === '上');
  assert.equal(top.value, '2.54 厘米');
  top.onStep(1); assert.deepEqual(plain(c.calls.pop()[1].page.margin), '2.64cm 2.54cm 2.54cm 2.54cm');
  find(v, '页边距', x => x.t === 'num' && x.lab === '左').onSet('1.91'); assert.equal(c.calls.pop()[1].page.margin, '2.54cm 2.54cm 2.54cm 1.91cm');
  // 脚注和尾注 › 编号: the document's format, a menu of the formats Word has
  v = at(c, 'refs');
  const nfmt = find(v, '脚注和尾注', x => x.t === 'sel' && x.title === '脚注和尾注的编号样式');
  assert.equal(nfmt.label, 'i, ii, iii');
  const menu = c.menus.nfmt; assert.deepEqual(plain(menu.map(m => m.label)), ['1, 2, 3', 'i, ii, iii', 'I, II, III', 'a, b, c', 'A, B, C', '①, ②, ③', '一, 二, 三']);
  menu[5].onClick(); assert.deepEqual(c.calls.pop(), ['setNoteFormat', 'decimalEnclosedCircleChinese']);
  c.state.pop = { id: 'nfmt', x: 0, y: 0 };
  assert.deepEqual(plain(c.renderVals().popItems.map(m => m.label)).slice(0, 2), ['1, 2, 3', 'i, ii, iii'], 'a menu only the panel has opens with its items');
  c.state.pop = null;
  c.menus.tocsty[2].onClick(); assert.deepEqual(c.calls.pop(), ['setTocStyle', 'plain']);
  // 表格 › 标题行 is the table's header look
  v = at(c, 'table', { inTable: true, tbl: { header: 'false' }, row: {}, cell: {} });
  const head = find(v, '表格样式', x => x.label === '标题行');
  assert.equal(head.on, false); head.onChange(true); assert.deepEqual(c.calls.pop(), ['tbl', 'look', 'true']);
});

test('the page and spacing helpers: margins as presets or four lengths, spacing in lines or points', () => {
  assert.deepEqual(plain(H.marginsCm('moderate')), { top: 2.54, right: 1.905, bottom: 2.54, left: 1.905 });
  assert.deepEqual(plain(H.marginsCm('2.54cm 3.18cm 2cm 3.18cm')), { top: 2.54, right: 3.18, bottom: 2, left: 3.18 }, 'the engine\'s four lengths');
  assert.deepEqual(plain(H.marginsCm('25.4mm 1in')), { top: 2.54, right: 2.54, bottom: 2.54, left: 2.54 });
  assert.deepEqual(plain(H.marginsCm('sideways')), plain(H.marginsCm('normal')), 'what cannot be read is Word\'s normal');
  assert.equal(H.marginText({ top: 2.54, right: 5.08, bottom: 2.54, left: 5.08 }), 'wide');
  // 普通 is the language's Word's: Chinese 上下 2.54 · 左右 3.18 (the design's), English 2.54 all round; both read as 普通
  assert.deepEqual(plain(H.marginPresets().map(p => [p[0], p[1]])), [['2.54cm 3.175cm 2.54cm 3.175cm', '普通'], ['narrow', '窄'], ['moderate', '适中'], ['wide', '宽']]);
  assert.equal(H.marginPreset(H.marginsCm('2.54cm 3.18cm 2.54cm 3.18cm'))[1], '普通');
  assert.equal(H.marginPreset(H.marginsCm('normal'))[1], '普通');
  assert.equal(H.marginPreset(H.marginsCm('2cm 2cm 2cm 2cm')), undefined);
  assert.equal(H.marginText(H.marginsCm('2.54cm 3.175cm 2.54cm 3.175cm')), '2.54cm 3.18cm 2.54cm 3.18cm', 'a side typed next to them keeps the others as Word shows them');
  assert.deepEqual([H.colName(0), H.colName(4), H.colName(25), H.colName(26), H.colName(27)], ['A', 'E', 'Z', 'AA', 'AB']);
  assert.equal(H.marginText({ top: 2.5, right: 3.176, bottom: 2.5, left: 3.176 }), '2.5cm 3.18cm 2.5cm 3.18cm');
  assert.deepEqual([H.spaceText('0.5lines'), H.spaceText('6pt'), H.spaceText('')], ['0.5 行', '6 磅', '0 行']);
  assert.deepEqual([H.spaceStep('0.5lines', 1), H.spaceStep('0.5lines', -1), H.spaceStep('0lines', -1), H.spaceStep('6pt', 1), H.spaceStep('', 1)], ['1lines', '0lines', '0lines', '12pt', '0.5lines']);
  assert.deepEqual([H.spaceParse('1.5'), H.spaceParse('1.5 行'), H.spaceParse('12磅'), H.spaceParse('6pt'), H.spaceParse('abc')], ['1.5lines', '1.5lines', '12pt', '6pt', null]);
  assert.deepEqual([H.rgbHex('rgb(29, 29, 31)'), H.rgbHex('#007aff'), H.rgbHex('red')], ['#1D1D1F', '#007AFF', '']);
  // what the fields take, and print and export use the same margins
  assert.deepEqual([FP.parseCm('2.5'), FP.parseCm('25mm'), FP.parseCm('1in'), FP.parseCm('2.5 厘米'), FP.parseCm('x')].map(x => Number.isNaN(x) ? 'NaN' : +x.toFixed(3)), [2.5, 2.5, 2.54, 2.5, 'NaN']);
  assert.equal(FP.cmLabel(2.539), '2.54 厘米');
  assert.equal(FP.snap(2.6400000001, 0.01), 2.64);
  assert.deepEqual(K.marginsCm('wide'), [2.54, 5.08, 2.54, 5.08]);
  assert.deepEqual(K.marginsCm('2.54cm 3.18cm 2cm 3.18cm'), [2.54, 3.18, 2, 3.18]);
});

test('FormatPanel.dc.html turns each item into its control', () => {
  const pctx = base();
  vm.runInNewContext(script('FormatPanel.dc.html') + '\nglobalThis.FormatPanel = Component;', pctx);
  const k = FP.kit({ state: {}, menus: {}, openPop() { } }), got = [];
  const p = new pctx.FormatPanel();
  p.props = { tab: 'home', tabs: FP.tabs([['home', '开始'], ['table', '表格', true]], 'table', key => got.push(key)),
    groups: [k.G('字体', k.R(k.sel('font', '宋体', []), k.size(12, () => { }, () => { }, 'fsize', [])), k.R(k.seg([{ label: 'B', on: true }, { icon: 'alignLeft', title: '左对齐' }]), k.btn('格式刷', 'painter', () => { })),
      k.grid(2, k.num('2.54 厘米', d => got.push(d), t => got.push(t), { lab: '上' }), k.chk('标题行', true, v => got.push(v))), null, k.R()),
      k.G('', k.R(k.swatch('底纹', '#FFFFFF', c => got.push(c)), k.sw(false, v => got.push(v)), k.link('删除表格', () => { })))] };
  const v = p.renderVals();
  assert.deepEqual(plain(v.tabs.map(t => [t.label, t.onA, t.ctxA])), [['开始', '0', '0'], ['表格', '1', '1']]);
  v.tabs[0].onClick(); assert.equal(got.pop(), 'home');
  assert.deepEqual(plain(v.groups.map(g => [g.title, g.rows.length])), [['字体', 3], ['', 1]], 'a group leaves out empty rows');
  const [row1, row2, row3] = v.groups[0].rows;
  assert.deepEqual([row1.items[0].isSelBtn, row1.items[1].isSize, row2.items[0].isSeg, row2.items[1].isBtn, row3.items[0].isNum, row3.items[1].isChk], [true, true, true, true, true, true]);
  assert.deepEqual(plain(row2.items[0].opts.map(o => [o.onA, o.plain, !!o.svg])), [['1', true, false], ['0', false, true]], 'a segment is text or one of the design\'s icons');
  assert.match(row2.items[1].svg.__html, /^<svg width="15"/);
  assert.deepEqual([row3.grid, row3.css, row3.items[0].fieldW], ['2', '--cols:2;', '56px']);
  row3.items[0].up(); row3.items[0].onKey({ key: 'ArrowDown', preventDefault() { } }); assert.deepEqual(got.splice(-2), [1, -1]);
  row3.items[0].onBlur({ target: { value: ' 3 ' } }); assert.equal(got.pop(), '3');
  row3.items[1].onChange({ target: { checked: false } }); assert.equal(got.pop(), false);
  const [sw, toggle, link] = v.groups[1].rows[0].items;
  assert.deepEqual([sw.isSwPick, sw.value, toggle.isSw, toggle.onA, link.isLink], [true, '#FFFFFF', true, '0', true]);
  sw.onChange({ target: { value: '#FF0000' } }); toggle.onClick(); assert.deepEqual(got.splice(-2), ['#FF0000', true]);
});
