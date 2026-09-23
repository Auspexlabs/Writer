// node --test ui/tests/ — Word paragraph formatting through the editor. Page breaks: a break inside a paragraph (<w:br w:type="page"/>)
// and a paragraph that starts a page (pageBreakBefore) are drawn as page breaks and, when the paragraph is edited, reach the file as
// page breaks again. Shading (w:pPr/w:shd): drawn, its own or its style's, and kept. The round trips run the real engine when the CLI
// is built.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');
const clone = x => structuredClone(x);
const PB = '<br style="page-break-before:always">';

test('the editor draws a page break inside a paragraph as a line it cannot type into, and sends it back as Word\'s', () => {
  const blocks = EN.blocksOf([
    { kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'end of page one' + PB + 'page two' } },
    { kind: 'paragraph', path: '/body/paragraph[2]', props: { html: 'own', pageBreakBefore: 'true' } },
    { kind: 'heading', path: '/body/heading[1]', props: { html: 'by its style', level: '1' }, computed: { pageBreakBefore: 'true' } },
    { kind: 'paragraph', path: '/body/paragraph[3]', props: { html: 'switched off', pageBreakBefore: 'false' }, computed: {} }
  ], 'a.docx');
  const html = EN.blocksToHtml(blocks);
  assert.match(html, /end of page one<span data-pb="1" contenteditable="false"><\/span>page two/);
  assert.match(html, /<p data-path="\/body\/paragraph\[2\]" data-w-pagebreakbefore="true" data-pb="before">own<\/p>/, 'its own break: kept for the save, and drawn');
  assert.match(html, /<h1 data-path="\/body\/heading\[1\]" data-pb="before">by its style<\/h1>/, 'the style\'s break: drawn, not written as its own');
  assert.match(html, /<p data-path="\/body\/paragraph\[3\]" data-w-pagebreakbefore="false">switched off<\/p>/);
  assert.equal(EN.pbOut(EN.pbIn('a' + PB + 'b' + PB)), 'a' + PB + 'b' + PB);
  assert.equal(EN.pbOut('a<span data-pb="1" contenteditable="false" style="color: red;"></span>b'), 'a' + PB + 'b', 'whatever the browser adds to it');
  assert.notEqual(EN.runsOf({ nodeType: 1, tagName: 'BODY', style: {}, childNodes: [{ nodeType: 1, tagName: 'BR', style: { pageBreakBefore: 'always' }, childNodes: [] }] }),
    EN.runsOf({ nodeType: 1, tagName: 'BODY', style: {}, childNodes: [{ nodeType: 1, tagName: 'BR', style: {}, childNodes: [] }] }), 'a page break is not a line break');
});

test('the editor draws a paragraph\'s shading, its own or its style\'s, and keeps its own for the save', () => {
  const html = EN.blocksToHtml(EN.blocksOf([
    { kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'own', fill: 'D9D9D9', align: 'center' } },
    { kind: 'paragraph', path: '/body/paragraph[2]', props: { html: 'by its style', style: 'Note' }, computed: { fill: 'FFF2CC' } },
    { kind: 'heading', path: '/body/heading[1]', props: { html: 'both', level: '2', fill: '1F3864', pageBreakBefore: 'true' } }
  ], 'a.docx'));
  assert.match(html, /<p data-path="\/body\/paragraph\[1\]" data-w-fill="D9D9D9" style="text-align:center;background:#D9D9D9">own<\/p>/);
  assert.match(html, /<p data-path="\/body\/paragraph\[2\]" data-style="Note" style="background:#FFF2CC">by its style<\/p>/, 'drawn, not written as its own');
  assert.match(html, /<h2 data-path="\/body\/heading\[1\]" data-w-pagebreakbefore="true" data-w-fill="1F3864" data-pb="before" style="background:#1F3864">both<\/h2>/);
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-paragraphs-'));
  server = spawn('dotnet', [cli, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  url = await new Promise((resolve, reject) => {
    let err = '';
    server.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1]); });
    server.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
  });
});
after(() => { server && server.kill(); dir && rmSync(dir, { recursive: true, force: true }); });
const run = async argv => {
  const r = await (await fetch(url + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) })).json();
  if (r.code !== 0) throw new Error(argv.join(' ') + ' → ' + JSON.stringify(r.error));
  return r.output && /^[{[]/.test(r.output.trim()) ? JSON.parse(r.output) : r.output;
};
const bodyBlocks = async file => EN.blocksOf((await run(['get', file, '/body', '--depth', '4'])).children, file);
// no DOM here: html parses as one text node, which is enough for a save to see which paragraphs changed
globalThis.DOMParser ??= class { parseFromString(s) { return { body: { nodeType: 1, tagName: 'BODY', style: {}, childNodes: [{ nodeType: 3, nodeValue: s }] } }; } };
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

test('engine: a page break inside a paragraph and a paragraph that starts a page stay page breaks when the paragraph is edited', { skip: skip() }, async () => {
  const file = join(dir, 'breaks.docx');
  await run(['create', file]);
  await run(['add', file, '/body', '--type', 'paragraph', '--prop', 'text=x']);
  await run(['set', file, '/body/paragraph[1]', '--raw', `<w:p ${W}><w:r><w:t>第一页末尾</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>第二页开头</w:t></w:r></w:p>`]);
  await run(['add', file, '/body', '--type', 'paragraph', '--prop', 'text=第三页', '--prop', 'pageBreakBefore=true']);
  const opened = await bodyBlocks(file);
  assert.match(EN.blocksToHtml(opened), /第一页末尾<span data-pb="1"[^>]*><\/span>第二页开头/);
  // the user types in both; the editor sends each paragraph's html as inlineHtml does (its page break line back to Word's)
  const edited = clone(opened);
  edited[0].props.html = EN.pbOut(EN.pbIn(edited[0].props.html).replace('第二页开头', '第二页开头，改过'));
  edited[1].props.html += '，也改过';
  const calls = [];
  await EN.planDocxBlocks(file, opened, edited, async argv => { calls.push(argv); return run(argv); });
  assert.equal(calls.length, 2, 'one set per edited paragraph');
  const [one, two] = [await run(['get', file, '/body/paragraph[1]', '--raw']), await run(['get', file, '/body/paragraph[2]', '--raw'])];
  assert.equal((await run(['get', file, '/body/paragraph[1]'])).props.text, '第一页末尾\f第二页开头，改过');
  assert.match(one, /<w:br w:type="page"\s*\/>/, 'still a page break');
  assert.doesNotMatch(one, /<w:br\s*\/>/, 'not a line break');
  assert.match(two, /<w:pageBreakBefore\s*\/>/);
  const reopened = await bodyBlocks(file);
  assert.equal(await EN.planDocxBlocks(file, reopened, clone(reopened), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
});

test('engine: a shaded paragraph keeps its shading when edited, and one Enter splits off it is shaded too', { skip: skip() }, async () => {
  const file = join(dir, 'shading.docx');
  await run(['create', file]);
  await run(['add', file, '/body', '--type', 'paragraph', '--prop', 'text=浅灰底纹', '--prop', 'fill=D9D9D9']);
  await run(['add', file, '/body', '--type', 'paragraph', '--prop', 'text=深蓝底纹', '--prop', 'fill=1F3864']);
  const opened = await bodyBlocks(file);
  assert.match(EN.blocksToHtml(opened), /data-w-fill="D9D9D9" style="background:#D9D9D9">浅灰底纹</);
  const edited = clone(opened);
  edited[0].props.html = '浅灰底纹，改过';
  edited.push({ kind: 'paragraph', path: null, props: { html: '新的一段', fill: '1F3864' } }); // the browser's copy of the paragraph Enter was pressed in
  await EN.planDocxBlocks(file, opened, edited, run);
  assert.deepEqual((await run(['get', file, '/body', '--depth', '2'])).children.map(c => [c.props.text, c.props.fill]),
    [['浅灰底纹，改过', 'D9D9D9'], ['深蓝底纹', '1F3864'], ['新的一段', '1F3864']]);
  const reopened = await bodyBlocks(file);
  assert.equal(await EN.planDocxBlocks(file, reopened, clone(reopened), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
});
