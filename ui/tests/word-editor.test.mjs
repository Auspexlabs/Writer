// node --test ui/tests/ — the Word editor's own logic, run without a browser: the page view, the thumbnails and the status bar
// share one pagination, headers and footers show on every page with its number, and each comment shows its author.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');

const html = readFileSync(new URL('../WordEditor.dc.html', import.meta.url), 'utf8');
const script = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)?.[1];
assert.ok(script, 'word editor script exists');
const ctx = { React: { createRef: () => ({ current: null }) }, setTimeout, clearTimeout, getComputedStyle: () => ({ marginBottom: '28px' }), // a page break line's margin
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), $lang: () => 'zh', // i18n.js stand-in: Chinese passthrough, placeholders filled
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(script + '\nglobalThis.WordEditor = Component; globalThis.HF = { hfOut, hfEdit, hfFill, hfPreset }; globalThis.HF_PRESETS = HF_PRESETS;', ctx);

function editor(doc) {
  const component = new ctx.WordEditor();
  component.EN = EN;
  component.props = { doc, onChange() {}, toast() {} };
  component.pgCss = { textContent: '' };
  return component;
}
/** The editor body as laid out, in document order, the last entry ending the text: paragraphs as [top, height] in px, a page break
 *  as ['hr', top, height] (a paragraph of its own), a paragraph that starts a page as ['before', top, height], and one with a page break
 *  inside as ['p', top, height, top of the break]. */
function laidOut(blocks) {
  const ed = { innerText: 'text', querySelectorAll: () => [] };
  ed.children = blocks.map(([a, b, c, d]) => {
    const k = typeof a === 'number' ? { offsetTop: a, offsetHeight: b } : { tag: a, offsetTop: b, offsetHeight: c };
    const spans = d == null ? [] : [{ offsetTop: d, offsetHeight: 1, offsetParent: ed }];
    spans.forEach(s => Object.assign(s, { parentElement: k, children: [] }));
    return Object.assign(k, { offsetParent: ed, parentElement: ed, children: spans, querySelectorAll: () => spans, matches: s => k.tag === 'hr' && /^hr/.test(s),
      getAttribute: n => n !== 'data-pb' ? null : k.tag === 'before' ? 'before' : k.tag === 'hr' ? '1' : null });
  });
  return ed;
}
const plain = x => JSON.parse(JSON.stringify(x)); // the editor runs in another vm realm: compare values, not prototypes
const pages = (c, blocks) => { c.edRef.current = laidOut(blocks); c.refreshInfo(); const v = c.renderVals(); return plain({ text: v.pagesText, thumbs: v.pageThumbs.map(t => [t.off, t.ch]) }); };

test('page thumbnails split where the status bar counts pages: every page break starts the next page', () => {
  const c = editor({ id: 'd', html: '<p>one</p><hr data-pb="1"><p>two</p>' }); // A4, normal margins: 931px of text a page, thumbnails at 128 / 794 (the design's sidebar)
  assert.deepEqual(pages(c, [[0, 100], ['hr', 128, 1], [157, 100]]), { text: '共 2 页', thumbs: [['0px', '20px'], ['-157px', '16px']] },
    'page 1 ends at the break and page 2 begins with what follows it (page 2 was blank and page 1 showed both)');
  assert.deepEqual(pages(c, [[0, 900], ['hr', 928, 1], [957, 900]]), { text: '共 2 页', thumbs: [['0px', '149px'], ['-957px', '145px']] },
    'two nearly full pages are two pages, not three');
  assert.deepEqual(pages(c, [[0, 1500], ['hr', 1528, 1], [1557, 50]]), { text: '共 3 页', thumbs: [['0px', '150px'], ['-1143px', '62px'], ['-1557px', '8px']] },
    'a block longer than a page runs on over the next one, up to the break');
  assert.deepEqual(pages(c, [[0, 1500]]), { text: '共 2 页', thumbs: [['0px', '150px'], ['-1143px', '57px']] }, 'without breaks too');
  assert.deepEqual(pages(c, [[0, 900], [908, 100]]), { text: '共 2 页', thumbs: [['0px', '145px'], ['-908px', '16px']] }, 'a paragraph that does not fit starts the next page');
  assert.equal(c.pgCss.textContent, `[data-pg="${c.pgId}"]>:nth-child(2){margin-top:243px!important}`, 'the page view moves it to the top of page 2 (1123px + 20px gap)');
});

test('a page break inside a paragraph and a paragraph that starts a page split pages too, and never make an empty one', () => {
  const c = editor({ id: 'd', html: '' });
  assert.deepEqual(pages(c, [['p', 0, 200, 88]]), { text: '共 2 页', thumbs: [['0px', '14px'], ['-117px', '13px']] }, 'Ctrl+Enter after text');
  assert.equal(c.pgCss.textContent, `[data-pg="${c.pgId}"]>:nth-child(1)>:nth-child(1){margin-bottom:1054px!important}`, 'the rest of the paragraph moves on');
  assert.deepEqual(pages(c, [[0, 100], ['before', 128, 60]]), { text: '共 2 页', thumbs: [['0px', '16px'], ['-128px', '9px']] }, 'pageBreakBefore');
  assert.deepEqual(pages(c, [['before', 0, 60], [88, 60]]), { text: '共 1 页', thumbs: [['0px', '23px']] }, 'the first paragraph is on a new page already');
  assert.deepEqual(pages(c, [[0, 100], ['hr', 128, 1], ['before', 157, 60]]), { text: '共 2 页', thumbs: [['0px', '20px'], ['-157px', '9px']] }, 'and so is one right after a page break');
});

test('typing that moves no block keeps the pages: the page view\'s gaps are not taken out to measure the whole file again', () => {
  const c = editor({ id: 'd', html: '' }), ed = laidOut([[0, 900], [908, 100]]);
  c.edRef.current = ed; c.refreshInfo();
  const css = c.pgCss.textContent; let writes = 0;
  c.pgCss = { get textContent() { return css; }, set textContent(v) { writes++; } };
  c.refreshInfo(true);
  assert.equal(writes, 0, 'the same blocks, as tall as they were: nothing laid out again');
  assert.equal(c.state.info.pages, 2);
  ed.children[1].offsetHeight = 140; c.refreshInfo(true);
  assert.equal(writes, 2, 'a paragraph that grew a line: measured without the gaps, then the gaps put back');
  c.refreshInfo(true); assert.equal(writes, 2, 'and then remembered as it is now');
  c.refreshInfo(); assert.equal(writes, 4, 'without `typed` (a load, a new page size, a resize) it always measures');
  const onInput = c.renderVals().onInput; c.syncT = null;
  onInput({ nativeEvent: { inputType: 'insertText' } }); clearTimeout(c.syncT); assert.ok(!c.pgDirty, 'typing a character');
  onInput({ nativeEvent: { inputType: 'insertParagraph' } }); clearTimeout(c.syncT); assert.equal(c.pgDirty, true, 'Enter, a paste, formatting: counted from scratch at the next pause');
  c.refreshInfo(!c.pgDirty); assert.equal(writes, 6); assert.ok(!c.pgDirty);
});

test('each page thumbnail copies only its own page\'s blocks, under an empty block as tall as the text above them', () => {
  const c = editor({ id: 'd', html: '<p>0</p><p>1</p><p>2</p>' }), ed = laidOut([[0, 400], [410, 400], [820, 400]]);
  ed.children.forEach((k, i) => Object.assign(k, { isConnected: true, outerHTML: `<p>${i}</p>` }));
  c.edRef.current = ed; c.refreshInfo();
  assert.deepEqual(plain(c.state.info.cuts), [[0, 1, 0], [1, 2, 410]], 'page 2 starts at block 2; its copy also takes block 1, the one before');
  const th = c.renderVals().pageThumbs;
  assert.deepEqual(plain(th.map(t => [t.off, t.html.__html])), [['0px', '<p>0</p><p>1</p>'], ['-820px', '<div style="height:410px;margin:0;padding:0"></div><p>1</p><p>2</p>']],
    'the page\'s blocks sit where they do in the whole document, so the thumbnail\'s offset is the page\'s own');
  const d = editor({ id: 'e', html: '<p>whole</p>' }); pages(d, [['p', 0, 200, 88]]);
  assert.equal(d.state.info.cuts, null);
  assert.deepEqual(plain(d.renderVals().pageThumbs.map(t => t.html.__html)), ['<p>whole</p>', '<p>whole</p>'], 'a page break inside a paragraph: each thumbnail copies the whole document, as before');
});

test('headers and footers: every page shows its own number, the first page its own once 首页不同 is on', () => {
  const footer = '<p style="text-align:center">第 {page} 页 / 共 {pages} 页</p>';
  const c = editor({ id: 'd', html: '', header: 'H', footer, titlePg: true, firstHeader: '', firstFooter: '' });
  pages(c, [[0, 900], [928, 900], [1856, 900]]);
  const zones = plain(c.renderVals().hfZones);
  assert.deepEqual(zones.map(z => z.html.__html), ['', '', 'H', '<p style="text-align:center">第 2 页 / 共 3 页</p>', 'H', '<p style="text-align:center">第 3 页 / 共 3 页</p>']);
  assert.deepEqual(zones.map(z => z.top), ['0px', '1027px', '1143px', '2170px', '2286px', '3313px'], 'each at its page\'s top and bottom margin');
  assert.deepEqual(zones.slice(0, 3).map(z => z.label), ['首页页眉', '首页页脚', '页眉']);

  const { hfOut, hfEdit } = ctx.HF;
  const box = hfEdit(footer, 2, 3);
  assert.match(box, /第 <span class="wd-fld" data-f="page" contenteditable="false" title="页码">2<\/span> 页/, 'fields show as chips while editing');
  assert.equal(hfOut(box), footer, 'and go back as {page} / {pages}');
  assert.equal(hfOut('<p style="text-align: right;"><b>A <span class="wd-fld" data-f="page">4</span></b><br></p>x'), '<p style="text-align:right"><b>A {page}</b></p><p style="text-align:left">x</p>', 'bold, alignment and loose text');
  assert.equal(hfOut('<p>A&nbsp;<span data-f="page">1</span>\u200B<p style="text-align: center;">B</p></p>'), '<p style="text-align:left">A {page}</p><p style="text-align:center">B</p>', 'blocks the browser nests are paragraphs of their own');
  assert.equal(hfOut(hfEdit('<img data-keep="0" src="data:,">', 1, 1)), '<p style="text-align:left"><img contenteditable="false" data-keep="0" src="data:,"></p>', 'a logo the editor only shows is kept');
  assert.equal(hfOut('<p><br></p>'), '', 'an emptied header goes, so the engine removes it');
});

test('样式 / 页码 presets: a template plus an alignment becomes footer html, previewed and round-tripped through the edit box', () => {
  const { hfPreset, hfOut, hfEdit, hfFill } = ctx.HF;
  assert.equal(hfPreset('{page} / {pages}', 'left'), '<p style="text-align:left">{page} / {pages}</p>');
  assert.equal(hfPreset('{page}', ''), '<p style="text-align:center">{page}</p>', 'no alignment given: centered, as Word does for a page number');
  const html = hfPreset('第 {page} 页 / 共 {pages} 页', 'right');
  assert.equal(hfFill(html, 1, 5), '<p style="text-align:right">第 1 页 / 共 5 页</p>', 'the menu preview: page 1 of the current document');
  assert.equal(hfOut(hfEdit(html, 2, 5)), html, 'and it survives the edit box (fields as chips) unchanged');
});

test('页码 and 样式 offer the same six presets, previewed as page 1 of the current document', () => {
  const c = editor({ id: 'd', html: '', footer: '' });
  let saved = null; c.props.onChange = patch => { saved = patch; };
  pages(c, [[0, 900], [928, 900], [1856, 900]]); // three pages, as the headers/footers test above
  const n = c.state.info.pages;
  const previews = plain(ctx.HF_PRESETS.map(tpl => ctx.HF.hfFill(tpl, 1, n)));

  c.setState({ tab: 'insert' });
  let v = c.renderVals();
  assert.ok(v.ribbon.some(it => it.isMenu && it.label === '页码'), '插入 tab offers 页码 as a menu of presets, not one fixed insert');
  assert.deepEqual(plain(c.menus.pagenum.map(i => i.label)), previews);
  c.menus.pagenum[3].onClick(); // 第 {page} 页 / 共 {pages} 页 — the preset 插入 › 页码 used to hard-code
  assert.equal(saved.footer, '<p style="text-align:center">第 {page} 页 / 共 {pages} 页</p>', '插入 › 页码 still writes the footer centered');

  // 样式, in the header/footer floating bar, offers the same presets while one is open
  c.setState({ hf: { kind: 'footer', page: 0, key: 'footer' } });
  v = c.renderVals();
  assert.ok(v.hfBtns.some(b => b.label === '样式'), 'the floating bar has a 样式 button');
  assert.deepEqual(plain(c.menus.hfStyle.map(i => i.label)), previews, 'and its popup lists the same presets');
});

test('comment avatars: 我 only on my own comments; other authors show their initials in a colour of their own', () => {
  const c = editor({ id: 'd', html: '', comments: [
    { id: '1', author: '王律师', text: '' }, { id: '2', author: 'Ann Lee', initials: 'AL', text: '' }, { id: '3', author: 'mia zhou', text: '' },
    { id: '4', author: '王律师', text: '' }, { id: '5', author: '林晓', mine: true, text: '' }, { id: '6', mine: true, text: '' }] });
  const list = c.renderVals().comments;
  assert.deepEqual(list.map(x => x.av), ['王', 'AL', 'MZ', '王', '我', '我'], 'the file\'s initials, else the author\'s; 我 for mine, saved or new');
  assert.equal(list[0].avBg, list[3].avBg, 'one author, one colour');
  assert.equal(list[4].avBg, list[5].avBg, 'mine keep the designed colour');
  assert.equal(new Set([0, 1, 2, 4].map(i => list[i].avBg)).size, 4, 'different authors, different colours');
  assert.deepEqual(c.renderVals().comments.map(x => x.avBg), list.map(x => x.avBg), 'the colours stay put');
  const other = editor({ id: 'e', html: '', comments: [{ id: '9', author: 'Aaron', text: '' }, { id: '10', author: 'Ann Lee', text: '' }] }).renderVals().comments;
  assert.equal(other[1].avBg, list[1].avBg, 'an author keeps their colour in every document, whoever else commented');
  for (const x of other) assert.match(x.avBg, /^oklch\(0\.56 0\.1 (225|345|165|285|105)\)$/, 'one of five colours');
});
