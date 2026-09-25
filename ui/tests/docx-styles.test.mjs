// node --test ui/tests/ — the style gallery in the editor: the document's styles come in as a list with their looks, are drawn on
// the page by a stylesheet of their own, a character style rides on a span the save compares like any other formatting, and
// 修改样式 / 新建样式 reach the file as style definitions. The round trip runs the real engine when the CLI is built.
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

test('styles list, looks as CSS, and a character style counts as formatting when the save compares runs', () => {
  const styles = EN.stylesOf('[{"id":"1","name":"heading 1","type":"paragraph","heading":1,"look":{"bold":"true","size":"16","font":"黑体"}},{"id":"Note","name":"备注","type":"paragraph","look":{"italic":"true","color":"595959","spaceAfter":"12pt","align":"center"}},{"id":"Emphasis","name":"Emphasis","type":"character","look":{"italic":"true"}}]');
  assert.equal(styles.length, 3);
  assert.deepEqual(EN.stylesOf('nonsense'), []);
  assert.equal(EN.lookCss(styles[1].look), 'font-style:italic;color:#595959;text-align:center;margin-bottom:12pt');
  assert.equal(EN.styleCss(styles, '[data-pg="x"]'),
    `[data-pg="x"] [data-style="1"],[data-pg="x"] h1:not([data-style]){font-family:'黑体','Noto Sans SC',sans-serif;font-size:16pt;font-weight:700}\n`
    + `[data-pg="x"] [data-style="Note"]{font-style:italic;color:#595959;text-align:center;margin-bottom:12pt}\n`
    + `[data-pg="x"] [data-style="Emphasis"]{font-style:italic}`, 'heading styles also dress the h1–h3 the editor shows');
  // the document's own look (the root's computed) dresses the page and plain paragraphs, and is what a paragraph style builds on
  const base = { font: 'Calibri', fontEa: '等线', size: '10.5', lineSpacing: '1', spaceAfter: '8pt', indentFirst: '2ch', pageSize: '21cm x 29.7cm' };
  assert.equal(EN.styleCss(styles.slice(0, 1).concat(styles.slice(2)), '[data-pg="x"]', base),
    `[data-pg="x"] > *{font-family:'Calibri','等线','Noto Sans SC',sans-serif;font-size:10.5pt}\n`
    + `[data-pg="x"] p:not([data-style]){line-height:1.5;margin-bottom:8pt;text-indent:2em}\n`
    + `[data-pg="x"] [data-style="1"],[data-pg="x"] h1:not([data-style]){font-family:'黑体','等线','Noto Sans SC',sans-serif;font-size:16pt;font-weight:700;line-height:1.5;margin-bottom:8pt;text-indent:2em}\n`
    + `[data-pg="x"] [data-style="Emphasis"]{font-style:italic}`, 'a character style keeps the text it sits in');
  assert.equal(EN.familyCss('宋体', 'Times New Roman'), "'宋体','Times New Roman','Noto Serif SC',serif");
  assert.notEqual(EN.runsOf('a <span data-style="Emphasis">b</span>'), EN.runsOf('a b'), 'applying 强调 is a change to save');
  assert.equal(EN.runsOf('a <span data-style="Emphasis">b</span>'), EN.runsOf('a <span data-style="Emphasis"><span>b</span></span>'));
  const blocks = EN.blocksFromHtml(parse('<p data-style="Note">n</p><h1 data-style="Title">t</h1><blockquote>q</blockquote><h2>h</h2>'));
  assert.deepEqual(blocks.map(b => [b.kind, b.props.style || '', b.props.level || '']), [['paragraph', 'Note', ''], ['paragraph', 'Title', ''], ['paragraph', 'Quote', ''], ['heading', '', '2']]);
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-styles-'));
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
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';

test('engine: a new style, a paragraph in it and a 强调 span survive save → reopen as pStyle and rStyle', { skip: skip() }, async () => {
  const file = join(dir, 'styles.docx');
  await run(['create', file]);
  const def = { id: 'Note', name: '备注', type: 'paragraph', basedOn: 'Normal', italic: 'true', color: '595959', spaceAfter: '12pt' }; // what defineStyle records
  await run(['set', file, '/', '--prop', 'style=' + JSON.stringify(def)]);
  const typed = EN.blocksFromHtml(parse('<p data-style="Note">一条<span data-style="Emphasis">强调</span>备注</p><h1>标题</h1>'));
  await EN.planDocxBlocks(file, [], typed, run);
  const styles = EN.stylesOf((await run(['get', file, '/'])).props.styles), note = styles.find(s => s.id === 'Note');
  assert.deepEqual([note.name, note.look.italic, note.look.color, note.look.spaceAfter], ['备注', 'true', '595959', '12pt']);
  const opened = EN.blocksOf((await run(['get', file, '/body', '--depth', '4'])).children, file);
  assert.deepEqual([opened[0].props.style, opened[0].props.html, opened[1].kind], ['Note', '一条<span data-style="Emphasis">强调</span>备注', 'heading']);
  assert.match(EN.blocksToHtml(opened), /<p data-path="\/body\/paragraph\[1\]" data-style="Note">一条<span data-style="Emphasis">强调<\/span>备注<\/p>/);
  const shown = EN.blocksFromHtml(parse(EN.blocksToHtml(opened)));
  assert.equal(await EN.planDocxBlocks(file, opened, shown, async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
});
