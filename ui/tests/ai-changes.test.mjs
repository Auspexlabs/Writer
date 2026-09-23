// node --test ui/tests/   — the AI panel's change cards: what diffMark marks after an assistant turn, how the list reads,
// and the card (with 保留 / 撤销) staying in view.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');

// ---- Word: the engine numbers blocks per kind by position, so a paragraph added in the middle moves every later path ----
const P = text => ({ kind: 'paragraph', props: { text } });
const H = text => ({ kind: 'heading', props: { text, level: '2' } });
const T = rows => ({ kind: 'table', props: {}, children: rows.map(r => ({ kind: 'row', props: {}, children: r.map(text => ({ kind: 'cell', props: { text } })) })) });
/** A Word body as the editor reads it: engine paths, and for each block an element that takes the change mark. */
function body(specs) {
  const n = {};
  return EN.blocksOf(specs.map(s => Object.assign({ path: `/body/${s.kind}[${n[s.kind] = (n[s.kind] || 0) + 1}]` }, s)), 'a.docx')
    .map(b => Object.assign(b, { el: { setAttribute: (k, v) => { b.mark = k + '=' + v; } } }));
}
const marked = blocks => blocks.filter(b => b.mark === 'data-ai=1').map(b => b.props.html || b.kind);
const four = [P('一'), P('二'), P('三'), P('四')];

test('a paragraph added in the middle is the only one marked, though every paragraph after it has a new path', () => {
  const after = body([P('一'), P('二'), P('新的一段'), P('三'), P('四')]);
  assert.deepEqual(EN.diffBlocks(body(four), after), [['新增', '段落 · 新的一段']]);
  assert.deepEqual(marked(after), ['新的一段']);
});

test('a deleted paragraph is listed once and marks nothing; an edited one is 修改', () => {
  const shorter = body([P('一'), P('三'), P('四')]);
  assert.deepEqual(EN.diffBlocks(body(four), shorter), [['删除', '1 处内容']]);
  assert.deepEqual(marked(shorter), []);
  const edited = body([P('一'), P('二，改过'), P('三'), P('四')]);
  assert.deepEqual(EN.diffBlocks(body(four), edited), [['修改', '段落 · 二，改过']]);
  assert.deepEqual(marked(edited), ['二，改过']);
  const swapped = body([P('一'), P('二'), P('三'), T([['a'], ['b']])]);
  assert.deepEqual(EN.diffBlocks(body(four), swapped), [['新增', '表格 · 2 行'], ['删除', '1 处内容']], 'the last paragraph gone, a table added in its place: not an edit of it');
});

test('the weekly report turn: one paragraph rewritten into three, a table appended, the paragraph after them untouched', () => {
  const before = body([H('周报'), P('日期'), H('本周进展'), P('这周主要在忙预热'), H('下周计划'), P('待定')]);
  const after = body([H('周报'), P('日期'), H('本周进展'), P('海报已上墙'), P('种草笔记 36 篇'), P('短信覆盖 42 万人'), H('下周计划'), P('待定'),
    T([['日期', '事项'], ['9 月 28 日', '补齐物料'], ['9 月 30 日', '发布海报'], ['10 月 1 日', '新品首发']])]);
  assert.deepEqual(EN.diffBlocks(before, after), [['修改', '段落 · 海报已上墙'], ['新增', '段落 · 种草笔记 36 篇'], ['新增', '段落 · 短信覆盖 42 万人'], ['新增', '表格 · 4 行']]);
  assert.deepEqual(marked(after), ['海报已上墙', '种草笔记 36 篇', '短信覆盖 42 万人', 'table']);
});

test('pairing stays right with repeated paragraphs and at both ends, and is quick on a long document', () => {
  const blank = body([P('a'), P(''), P(''), P('b'), P('')]), more = body([P(''), P('a'), P(''), P(''), P(''), P('b'), P(''), P('c')]);
  assert.deepEqual(EN.diffBlocks(blank, more).map(i => i[0]), ['新增', '新增', '新增']);
  assert.equal(marked(more).length, 3);
  // 3000 paragraphs, one added near the top and the last one edited: everything in between is compared with itself
  const long = Array.from({ length: 3000 }, (_, i) => P('第 ' + i + ' 段'));
  const changed = [...long.slice(0, 2), P('插入'), ...long.slice(2, -1), P('最后一段改了')], t0 = Date.now();
  assert.deepEqual(EN.diffBlocks(body(long), body(changed)), [['新增', '段落 · 插入'], ['修改', '段落 · 最后一段改了']]);
  assert.ok(Date.now() - t0 < 2000, 'took ' + (Date.now() - t0) + ' ms');
});

test('the change list names each kind once: 表格 · 4 行, not 表格 · 表格 · 4 行', () => {
  const added = body([T([['a'], ['b'], ['c'], ['d']]), { kind: 'toc', props: { title: '目录' } }, { kind: 'toc', props: { title: 'Contents' } },
    { kind: 'image', props: {} }, { kind: 'code', props: { text: 'x = 1' } }, { kind: 'pagebreak', props: {} }, H('计划'), { kind: 'paragraph', props: { text: '要点', list: 'bullet' } }, P('')]);
  assert.deepEqual(EN.diffBlocks([], added).map(i => i[1]), ['表格 · 4 行', '目录', '目录 · Contents', '图片', '代码块', '分页符', '标题 · 计划', '列表项 · 要点', '段落 · （空）']);
});

test('same text, new look is 修改: heading level, paragraph style, alignment, list or plain; the rest stays paired', () => {
  const before = body([H('计划'), P('正文'), P('引言'), P('要点'), P('结尾'), P('新段落之前')]);
  const after = body([{ kind: 'heading', props: { text: '计划', level: '3' } }, { kind: 'paragraph', props: { text: '正文', align: 'center' } },
    { kind: 'paragraph', props: { text: '引言', style: 'Quote' } }, { kind: 'paragraph', props: { text: '要点', list: 'bullet' } }, P('结尾'), P('新的'), P('新段落之前')]);
  assert.deepEqual(EN.diffBlocks(before, after), [['修改', '标题 · 计划'], ['修改', '段落 · 正文'], ['修改', '段落 · 引言'], ['修改', '列表项 · 要点'], ['新增', '段落 · 新的']]);
  assert.deepEqual(marked(after), ['计划', '正文', '引言', '要点', '新的']);
  assert.deepEqual(EN.diffBlocks(body([{ kind: 'paragraph', props: { text: 'a', style: 'Normal' } }]), body([P('a')])), [], '正文 by name or by default is the same look');
});

test('a paragraph that became a heading is 修改, not 新增 + 删除: the same text pairs across kinds before the same kind does', () => {
  assert.deepEqual(EN.diffBlocks(body(four), body([P('一'), H('二'), P('三'), P('四')])), [['修改', '标题 · 二']]);
  // 二 became a heading and 三 was rewritten: the heading goes with 二 (same text), the new paragraph with 三 (same kind)
  assert.deepEqual(EN.diffBlocks(body(four), body([P('一'), P('新'), H('二'), P('四')])), [['修改', '段落 · 新'], ['修改', '标题 · 二']]);
});

// ---- Excel: sheet paths are positions too (/sheet[2]); a sheet added in front must not shift the comparison ----
const sheets = (...names) => ({ type: 'xlsx', sheets: names.map((name, i) => ({ name, path: `/sheet[${i + 1}]`, cells: { A1: { v: name } }, charts: [] })) });

test('a sheet added in front is new; the sheets after it are compared with themselves; a removed sheet is listed', () => {
  const after = sheets('汇总', '一月', '二月');
  assert.deepEqual(EN.diffMark(sheets('一月', '二月'), after), [['新增', '工作表 汇总']]);
  assert.ok(after.sheets.every(s => !s.cells.A1.ai), 'no cell marked');
  assert.deepEqual(EN.diffMark(sheets('一月', '二月'), sheets('二月')), [['删除', '1 个工作表']]);
});

// ---- PowerPoint: slides pair by the engine's slide id, which their path carries (/slide[@id=257]) ----
const deck = (...ids) => ({ type: 'pptx', slides: ids.map(id => ({ id: 's' + id, path: `/slide[@id=${id}]`, objs: [], bg: '#FFFFFF', notes: '', trans: 'none', duration: null, hidden: false })) });

test('a slide deleted while another is added is listed as deleted: slides go by id, not by how many there are', () => {
  assert.deepEqual(EN.diffMark(deck(256, 257, 258), deck(256, 258, 300)), [['新增', '第 3 页'], ['删除', '1 页']]);
  assert.deepEqual(EN.diffMark(deck(256, 257, 258), deck(256)), [['删除', '2 页']]);
  const made = deck(256, 300); made.slides[1].id = 'x7'; // added in the editor and saved: the editor's own id, the engine's path
  assert.deepEqual(EN.diffMark(made, deck(256, 300)), [], 'the same slide, not one deleted and one added');
});

// ---- the panel (index.dc.html): its logic in a vm, as storage.test.mjs runs it ----
function shell(docs, cur) {
  const code = readFileSync(new URL('../index.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, document: { querySelector: () => null },
    React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); } } };
  vm.runInNewContext(code + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell();
  Object.assign(c, { props: {}, EN });
  Object.assign(c.state, { docs, cur, view: 'doc', showAI: true });
  return c;
}
/** The chat list's scroller; like a browser's, scrollTop stays between 0 and the content's bottom. */
const scroller = (clientHeight, scrollHeight) => ({ clientHeight, scrollHeight, top: 0, get scrollTop() { return this.top = Math.max(0, Math.min(this.top, this.scrollHeight - this.clientHeight)); }, set scrollTop(v) { this.top = v; this.scrollTop; } });
const card = { items: [['新增', '段落 · 新的一段']], state: 'pending' };

test('the panel scrolls a new change card into view, 保留 / 撤销 included, unless the user has scrolled up', () => {
  const c = shell([{ id: 'a', title: 'a', type: 'docx', path: 'a.docx', loaded: true }], 'a');
  const el = scroller(400, 400); c.chatRef = { current: el };
  const render = (msgs, height) => { c.state.chats = { a: msgs }; el.scrollHeight = height; c.componentDidUpdate(); };
  const ask = { role: 'user', text: '加一段' }, reply = { role: 'ai', text: '已加上。', steps: [{ text: 'add a.docx /body --type paragraph', ok: true }] };
  render([ask, reply], 1000);
  assert.equal(el.scrollTop, 600, 'the reply is followed, as before');
  render([ask, { ...reply, change: card }], 1180);
  assert.equal(el.scrollTop, 780, 'the card arrived after the reply: down to its buttons');

  const again = [ask, { ...reply, change: { ...card, state: 'kept' } }, { role: 'user', text: '再加一段' }, reply];
  render(again, 1800);
  assert.equal(el.scrollTop, 1400);
  el.scrollTop = 900; c.componentDidUpdate(); // the user scrolls up to read; the shell renders for something else
  render([...again.slice(0, 3), { ...reply, change: card }], 1980);
  assert.equal(el.scrollTop, 900, 'scrolled up on purpose: left where the user is');
});
