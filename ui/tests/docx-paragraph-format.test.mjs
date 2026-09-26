// node --test ui/tests/ — Word's 段落 settings in the editor: a paragraph keeps its own line spacing, spacing, indents, borders,
// keep-with-next and tab stops as data-w-* (what the save writes), draws them as CSS, and reads 分散对齐 back as distribute. The
// round trip runs the real engine when the CLI is built.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { install } from './dom-stub.mjs';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const parse = install();
const EN = await import('../engine.js');
const clone = x => structuredClone(x);
const FORMAT = { lineSpacing: '1.5', spaceBefore: '6pt', spaceAfter: '8pt', indentLeft: '1cm', indentFirst: '2ch', border: 'top bottom', keepNext: 'true', tabs: 'left 2cm, right 15cm' };

test('the editor draws a paragraph\'s own 段落 settings as CSS and keeps them on the element for the save', () => {
  const html = EN.blocksToHtml(EN.blocksOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: Object.assign({ html: 'text', align: 'distribute' }, FORMAT) },
    { kind: 'paragraph', path: '/body/paragraph[2]', props: { html: 'hang', indentFirst: '-2ch', lineSpacing: '18pt', border: 'box' } }], 'a.docx'));
  assert.match(html, /<p data-path="\/body\/paragraph\[1\]" data-w-linespacing="1.5" data-w-spacebefore="6pt" data-w-spaceafter="8pt" data-w-indentleft="1cm" data-w-indentfirst="2ch" data-w-border="top bottom" data-w-keepnext="true" data-w-tabs="left 2cm, right 15cm" style="text-align:justify;text-align-last:justify;line-height:2.25;margin-top:6pt;margin-bottom:8pt;margin-left:1cm;text-indent:2em;border-top:1px solid currentColor;border-bottom:1px solid currentColor;padding:1px 4px">text<\/p>/);
  assert.match(html, /style="line-height:18pt;text-indent:-2em;padding-left:2em;border-top:1px solid currentColor;border-bottom:1px solid currentColor;border-left:1px solid currentColor;border-right:1px solid currentColor;padding:1px 4px">hang</, 'a hanging indent pulls the first line back; box is all four sides');
  const back = EN.blocksFromHtml(parse(html));
  assert.deepEqual([back[0].align, back[1].align], ['distribute', null]);
  assert.deepEqual(Object.fromEntries(Object.keys(FORMAT).map(k => [k, back[0].props[k]])), FORMAT, 'every setting comes back as it went in');
  assert.deepEqual([back[1].props.indentFirst, back[1].props.lineSpacing, back[1].props.border], ['-2ch', '18pt', 'box']);
  // the ruler's tab stops and the ribbon's redraw after a change
  assert.deepEqual(EN.tabStopsOf('left 2cm, decimal 10.5cm'), [{ kind: 'left', cm: 2 }, { kind: 'decimal', cm: 10.5 }]);
  assert.equal(EN.tabsText([{ kind: 'center', cm: 8 }]), 'center 8cm');
  const el = parse('<p data-w-spaceafter="12pt" data-w-fill="D9D9D9" style="text-align:center;margin-top:6pt;line-height:1.5">x</p>').firstChild;
  EN.drawPara(el);
  assert.equal(el.getAttribute('style'), 'text-align: center; margin-bottom: 12pt; background: #D9D9D9;', 'the old spacing goes, the alignment stays');
});

test('段前 / 段后 in lines (行) draw 12pt a line; 孤行控制 off is the paragraph\'s own, and one its style turns off shows on data-widow', () => {
  const html = EN.blocksToHtml(EN.blocksOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'a', spaceBefore: '0.5lines', spaceAfter: '1lines', widowControl: 'false' } },
    { kind: 'paragraph', path: '/body/paragraph[2]', props: { html: 'b' }, computed: { widowControl: 'false' } }], 'a.docx'));
  assert.match(html, /data-w-spacebefore="0.5lines" data-w-spaceafter="1lines" data-w-widowcontrol="false" style="margin-top:6pt;margin-bottom:12pt">a</);
  assert.match(html, /<p data-path="\/body\/paragraph\[2\]" data-widow="off">b<\/p>/);
  const back = EN.blocksFromHtml(parse(html));
  assert.deepEqual([back[0].props.spaceBefore, back[0].props.widowControl, back[1].props.widowControl], ['0.5lines', 'false', undefined], 'what the style says is not written as the paragraph\'s');
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-parafmt-'));
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
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';

test('engine: 段落 settings typed in the editor survive save → reopen; a change sends one set, and taking one off says none', { skip: skip() }, async () => {
  const file = join(dir, 'format.docx');
  await run(['create', file]);
  const attrs = Object.entries(FORMAT).map(([k, v]) => `data-w-${k.toLowerCase()}="${v}"`).join(' ');
  const typed = EN.blocksFromHtml(parse(`<p ${attrs} style="text-align:justify;text-align-last:justify">正文</p><h1 data-w-keepnext="true" data-w-spacebefore="24pt">标题</h1><p data-w-indentfirst="-0.74cm" data-w-border="box">悬挂</p>`));
  await EN.planDocxBlocks(file, [], typed, run);
  const opened = await bodyBlocks(file);
  assert.deepEqual(Object.fromEntries(Object.keys(FORMAT).map(k => [k, opened[0].props[k]])), FORMAT);
  assert.equal(opened[0].props.align, 'distribute');
  assert.deepEqual([opened[1].kind, opened[1].props.keepNext, opened[1].props.spaceBefore], ['heading', 'true', '24pt']);
  assert.deepEqual([opened[2].props.indentFirst, opened[2].props.border], ['-0.74cm', 'box']);
  const shown = () => EN.blocksFromHtml(parse(EN.blocksToHtml(opened))); // what the editor holds once the file is on screen
  assert.equal(await EN.planDocxBlocks(file, opened, shown(), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
  assert.equal(await EN.planDocxBlocks(file, opened, clone(opened), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0, 'engine blocks compare with themselves too');
  const edited = shown(); edited[0].props.lineSpacing = '2'; delete edited[0].props.border; edited[2].props.keepLines = 'true';
  const calls = [];
  await EN.planDocxBlocks(file, opened, edited, async argv => { calls.push(argv.slice(2)); return run(argv); });
  assert.deepEqual(calls, [['/body/paragraph[1]', '--prop', 'lineSpacing=2', '--prop', 'border=none'], ['/body/paragraph[2]', '--prop', 'keepLines=true']]);
  const again = await bodyBlocks(file);
  assert.deepEqual([again[0].props.lineSpacing, again[0].props.border, again[2].props.keepLines], ['2', undefined, 'true']);
});

test('engine: spacing in lines and 孤行控制 survive save → reopen; turning widow control back on sends none', { skip: skip() }, async () => {
  const file = join(dir, 'lines.docx');
  await run(['create', file]);
  await EN.planDocxBlocks(file, [], EN.blocksFromHtml(parse('<p data-w-spacebefore="0.5lines" data-w-spaceafter="1lines" data-w-widowcontrol="false">一</p>')), run);
  const opened = await bodyBlocks(file);
  assert.deepEqual([opened[0].props.spaceBefore, opened[0].props.spaceAfter, opened[0].props.widowControl], ['0.5lines', '1lines', 'false']);
  const edited = EN.blocksFromHtml(parse(EN.blocksToHtml(opened))); delete edited[0].props.widowControl;
  const calls = [];
  await EN.planDocxBlocks(file, opened, edited, async argv => { calls.push(argv.slice(2)); return run(argv); });
  assert.deepEqual(calls, [['/body/paragraph[1]', '--prop', 'widowControl=none']]);
  assert.equal((await bodyBlocks(file))[0].props.widowControl, undefined);
});
