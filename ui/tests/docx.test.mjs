// node --test ui/tests/   (Node 20) — the docx bridge's pure parts: tree → blocks → html, page setup props, header/footer text.
import { test } from 'node:test';
import assert from 'node:assert/strict';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');

test('page breaks: tree → block → html', () => {
  const blocks = EN.blocksOf([
    { kind: 'paragraph', path: '/body/paragraph[1]', props: { text: 'one' } },
    { kind: 'pagebreak', path: '/body/pagebreak[1]', props: {} },
    { kind: 'paragraph', path: '/body/paragraph[2]', props: { list: 'bullet', text: 'two' } }
  ], 'a.docx');
  assert.deepEqual(blocks[1], { kind: 'pagebreak', path: '/body/pagebreak[1]', props: {} });
  assert.equal(EN.blocksToHtml(blocks), '<p data-path="/body/paragraph[1]">one</p><hr data-pb="1" data-path="/body/pagebreak[1]"><ul><li data-path="/body/paragraph[2]">two</li></ul>');
});

test('pageOf maps engine props to the editor page, keeping unknown sizes in raw', () => {
  assert.deepEqual(EN.pageOf({ page: 'Letter', orientation: 'landscape', margin: 'moderate', columns: '2' }), { size: 'Letter', orient: 'landscape', margin: 'normal', cols: 2, raw: { size: 'Letter', margin: 'moderate' } });
  assert.deepEqual(EN.pageOf({ page: 'Legal', margin: '2cm 1cm 2cm 1cm' }), { size: 'A4', orient: 'portrait', margin: 'normal', cols: 1, raw: { size: 'Legal', margin: '2cm 1cm 2cm 1cm' } });
  assert.deepEqual(EN.pageOf(undefined), { size: 'A4', orient: 'portrait', margin: 'normal', cols: 1, raw: { size: '', margin: '' } });
});

test('plainOf strips header html but keeps page tokens and lines', () => {
  assert.equal(EN.plainOf('<b>Q3</b> report<br>page {page} of {pages}'), 'Q3 report\npage {page} of {pages}');
  assert.equal(EN.plainOf('a &lt;b&gt; &amp; c'), 'a <b> & c');
  assert.equal(EN.plainOf(undefined), '');
});

test('pageDiff emits only the changed set props, and a header or footer only once it is edited', () => {
  const logo = '<img data-keep="0" src="data:image/png;base64,AA"> Acme';
  const orig = { page: { page: 'A4', orientation: 'portrait', margin: 'normal', columns: '1' }, header: logo, footer: '', firstHeader: '', firstFooter: '', titlePg: false };
  assert.deepEqual(EN.pageDiff(orig, { page: { size: 'A4', orient: 'portrait', margin: 'normal', cols: 1 }, header: logo, footer: '' }), {}, 'untouched: the file keeps what the editor only shows');
  const footer = '<p style="text-align:center"><b>第 {page} 页</b></p>';
  assert.deepEqual(EN.pageDiff(orig, { page: { size: 'Letter', orient: 'landscape', margin: 'wide', cols: 2, color: '#fff', wm: '草稿', hf: false }, header: logo, footer, titlePg: true, firstFooter: '<p style="text-align:left">cover</p>' }),
    { page: 'Letter', orientation: 'landscape', margin: 'wide', columns: '2', footer, firstFooter: '<p style="text-align:left">cover</p>', titlePg: 'true' });
  assert.deepEqual(EN.pageDiff({ page: {}, header: 'x' }, { page: { size: 'A4' }, header: '' }), { header: '' }, 'an emptied header is sent as an empty string so the engine removes it');
});

test('pageDiff leaves sizes and margins the editor cannot show alone', () => {
  const orig = { page: { page: 'Legal', margin: 'moderate' } };
  assert.deepEqual(EN.pageDiff(orig, { page: EN.pageOf(orig.page) }), {});
  assert.deepEqual(EN.pageDiff(orig, { page: Object.assign(EN.pageOf(orig.page), { size: 'A5' }) }), { page: 'A5' });
});

test('Word tables and TOC keep their file paths and structural props in editor HTML', () => {
  const blocks = EN.blocksOf([
    { kind: 'table', path: '/body/table[1]', props: { style: 'TableGrid', borders: 'outside', borderColor: '00AA00', width: '100%', widths: '["3cm","5cm"]' }, children: [
      { kind: 'row', path: '/body/table[1]/row[1]', props: { header: 'true' }, children: [
        { kind: 'cell', path: '/body/table[1]/row[1]/cell[1]', props: { html: '<b>A</b>', colspan: '2', rowspan: '2', fill: 'FFF000', valign: 'middle' } }
      ] }
    ] },
    { kind: 'toc', path: '/body/toc[1]', props: { levels: '3', title: '目录', text: '第一章\n第二章' } }
  ], 'a.docx');
  assert.equal(blocks[0].rows[0].cells[0].props.colspan, '2');
  const html = EN.blocksToHtml(blocks);
  assert.match(html, /data-path="\/body\/table\[1\]"[^>]*data-w-bordercolor="00AA00"/);
  assert.match(html, /data-path="\/body\/table\[1\]\/row\[1\]\/cell\[1\]"[^>]*colspan="2" rowspan="2"/);
  assert.match(html, /data-toc="1" data-path="\/body\/toc\[1\]" data-levels="3" data-title="目录"/);
  assert.match(html, /第一章<\/span><\/div><div[^>]*><span[^>]*>第二章/);
  const fromTree = EN.blocksOf([{ kind: 'table', path: '/body/table[1]', props: { widths: ['3cm', '5cm'] }, children: [] }], 'a.docx');
  assert.equal(fromTree[0].props.widths, '["3cm","5cm"]', 'the /json tree gives widths parsed; the editor keeps the JSON the engine takes back');
});

test('unchanged Word table and TOC plan no writes, including centered blank cells and untitled TOC', async () => {
  const before = EN.blocksOf([
    { kind: 'table', path: '/body/table[1]', props: { borders: 'all' }, children: [
      { kind: 'row', path: '/body/table[1]/row[1]', props: {}, children: [
        { kind: 'cell', path: '/body/table[1]/row[1]/cell[1]', props: { html: '', align: 'center' } }
      ] }
    ] },
    { kind: 'toc', path: '/body/toc[1]', props: { levels: '3', text: 'Chapter' } }
  ], 'a.docx');
  const after = structuredClone(before), calls = [];
  assert.equal(await EN.planDocxBlocks('a.docx', before, after, async argv => { calls.push(argv); return {}; }), 0);
  assert.deepEqual(calls, []);
});

test('automatic text takes the colour Word gives it on the fill behind it: dark on a light fill, white on a dark one, in either theme', () => {
  assert.deepEqual(['D9D9D9', 'F5F5F7', 'FFFF00', 'DDEBF7', '1F3864', '000080', '000000'].map(f => EN.inkOn(f)), ['dark', 'dark', 'dark', 'dark', 'light', 'light', 'light']);
  // every element with a fill of its own is marked (cell shading, a highlight, the contents box); a fill taken away takes its mark along
  const el = (bg, ink) => { const attrs = ink ? { 'data-ink': ink } : {}; return { style: { backgroundColor: bg }, attrs, setAttribute(k, v) { attrs[k] = v; }, removeAttribute(k) { delete attrs[k]; } }; };
  const els = [el('rgb(217, 217, 217)'), el('rgb(31, 56, 100)', 'dark'), el('transparent', 'dark'), el('', 'light')];
  let asked = '';
  EN.inkFills({ querySelectorAll: sel => { asked = sel; return els; } });
  assert.deepEqual(els.map(e => e.attrs['data-ink'] || null), ['dark', 'light', null, null]);
  assert.match(asked, /\[style\*="background"\]/);
  assert.match(asked, /\[data-ink\]/, 'marks left by a fill that went are found too');
});
