// node --test ui/tests/ — table formatting in the editor: column widths as a colgroup, row heights, a header row, a cell's own
// borders and alignment, the table's width and place, all kept as data-w-* and written back; sorted rows make the save write
// the table anew. The round trip runs the real engine when the CLI is built.
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

const table = (props, rows) => ({ kind: 'table', path: '/body/table[1]', props, rows: rows.map((cells, r) => ({ kind: 'row', path: `/body/table[1]/row[${r + 1}]`, props: cells.rowProps || {},
  cells: cells.map((c, i) => ({ kind: 'cell', path: `/body/table[1]/row[${r + 1}]/cell[${i + 1}]`, props: typeof c === 'string' ? { html: c } : c })) })) });

test('the editor draws widths, heights, a header row, cell borders and alignment, and reads them back as the same props', () => {
  const r1 = ['a', { html: 'b', borders: 'top bottom', align: 'center', valign: 'middle' }]; r1.rowProps = { header: 'true', height: '360000' };
  const html = EN.blocksToHtml([table({ style: 'PlainTable1', widths: '["3cm","5cm"]', width: '8cm', align: 'center' }, [r1, ['c', 'd']])]);
  assert.match(html, /<table data-path="\/body\/table\[1\]" data-w-style="PlainTable1" data-w-width="8cm" data-w-widths="\[&quot;3cm&quot;,&quot;5cm&quot;\]" data-w-align="center" style="border-collapse:collapse;width:302\.36\d*px;margin:8px auto;table-layout:fixed"><colgroup><col style="width:113\.38\d*px"><col style="width:188\.97\d*px"><\/colgroup>/);
  assert.match(html, /<tr data-path="\/body\/table\[1\]\/row\[1\]" data-w-header="true" data-w-height="360000" style="height:38px">/);
  assert.match(html, /<td data-path="\/body\/table\[1\]\/row\[1\]\/cell\[2\]" data-w-borders="top bottom" data-w-valign="middle" data-w-align="center" style="[^"]*border-top:1px solid #C7C7CC;border-right:1px dashed #E5E5EA;border-bottom:1px solid #C7C7CC;border-left:1px dashed #E5E5EA;vertical-align:middle;text-align:center">b</);
  const back = EN.blocksFromHtml(parse(html))[0];
  assert.deepEqual(back.props, { style: 'PlainTable1', width: '8cm', widths: '["3cm","5cm"]', align: 'center' });
  assert.deepEqual(back.rows[0].props, { header: 'true', height: '360000' });
  assert.deepEqual(back.rows[0].cells[1].props, { html: 'b', borders: 'top bottom', valign: 'middle', align: 'center' });
  assert.equal(EN.tableCss({}), 'border-collapse:collapse;width:100%;margin:8px 0');
  assert.equal(EN.ownLines('none'), 'border-top:1px dashed #E5E5EA;border-right:1px dashed #E5E5EA;border-bottom:1px dashed #E5E5EA;border-left:1px dashed #E5E5EA');
  // sorted rows are the same rows in another order: no recorded command leads there, so the table is written anew
  const before = table({}, [['b'], ['a']]), sorted = table({}, [['b'], ['a']]); sorted.rows.reverse();
  assert.deepEqual(EN.tablesToRewrite([before], [sorted]), [sorted]);
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-tablefmt-'));
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

test('engine: a table formatted in the editor survives save → reopen, and a second save sends nothing', { skip: skip() }, async () => {
  const file = join(dir, 'table.docx');
  await run(['create', file]);
  const typed = EN.blocksFromHtml(parse('<table data-w-style="GridTable4Accent1" data-w-widths=\'["3cm","5cm"]\' data-w-align="center" style="border-collapse:collapse"><colgroup><col><col></colgroup><tbody>'
    + '<tr data-w-header="true" data-w-height="1cm"><td>名称</td><td data-w-borders="top bottom" data-w-align="center" data-w-valign="middle">值</td></tr><tr><td>甲</td><td data-w-fill="D9E2F3">1</td></tr></tbody></table>'));
  await EN.planDocxBlocks(file, [], typed, run);
  const opened = await bodyBlocks(file), t = opened[0];
  assert.deepEqual([t.props.style, t.props.align, JSON.parse(t.props.widths).map(w => EN.cmOf(w))], ['GridTable4Accent1', 'center', [3, 5]]);
  assert.deepEqual([t.rows[0].props.header, EN.cmOf(t.rows[0].props.height + 'emu')], ['true', 1]);
  assert.deepEqual([t.rows[0].cells[1].props.borders, t.rows[0].cells[1].props.align, t.rows[0].cells[1].props.valign, t.rows[1].cells[1].props.fill], ['top bottom', 'center', 'middle', 'D9E2F3']);
  const shown = EN.blocksFromHtml(parse(EN.blocksToHtml(opened)));
  assert.equal(await EN.planDocxBlocks(file, opened, shown, async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
});
