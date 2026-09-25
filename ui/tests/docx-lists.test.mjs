// node --test ui/tests/ — Word's list galleries in the editor: the kind of a list (1. a. i., 1. 1.1 1.1.1, 一、（一）1.) rides on
// its ol as data-w-list, numbering that starts again is an ol of its own (data-w-restart), and both reach the file through
// numbering.xml. The round trips run the real engine when the CLI is built.
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
const para = (path, html, props) => ({ kind: 'paragraph', path, props: Object.assign({ html }, props) });

test('the editor draws each list kind and a restart as its own ol, and reads them back as list / level / restart props', () => {
  const html = EN.blocksToHtml(EN.blocksOf([
    para('/body/paragraph[1]', '一', { list: 'chinese', level: '0' }), para('/body/paragraph[2]', '（一）', { list: 'chinese', level: '1' }),
    para('/body/paragraph[3]', 'one', { list: 'number', level: '0' }), para('/body/paragraph[4]', 'one again', { list: 'number', level: '0', restart: 'true' }),
    para('/body/paragraph[5]', '1.', { list: 'outline', level: '0' }), para('/body/paragraph[6]', '1.1', { list: 'outline', level: '1' }), para('/body/paragraph[7]', 'dot', { list: 'bullet', level: '0' })], 'a.docx'));
  assert.equal(html, '<ol data-w-list="chinese"><li data-path="/body/paragraph[1]">一</li><ol data-w-list="chinese"><li data-path="/body/paragraph[2]">（一）</li></ol></ol>'
    + '<ol><li data-path="/body/paragraph[3]">one</li></ol><ol data-w-restart="1"><li data-path="/body/paragraph[4]">one again</li></ol>'
    + '<ol data-w-list="outline"><li data-path="/body/paragraph[5]">1.</li><ol data-w-list="outline"><li data-path="/body/paragraph[6]">1.1</li></ol></ol><ul><li data-path="/body/paragraph[7]">dot</li></ul>');
  const back = EN.blocksFromHtml(parse(html)).map(b => [b.props.list, b.props.level, b.props.restart || '']);
  assert.deepEqual(back, [['chinese', '0', ''], ['chinese', '1', ''], ['number', '0', ''], ['number', '0', 'true'], ['outline', '0', ''], ['outline', '1', ''], ['bullet', '0', '']]);
  // what the browser makes: a nested ol without the attribute belongs to the outer list's kind; a plain ol is 1. a. i.
  const typed = EN.blocksFromHtml(parse('<ol data-w-list="outline"><li>a</li><ol><li>b</li></ol></ol><ul><li>c<ol><li>d</li></ol></li></ul>'));
  assert.deepEqual(typed.map(b => [b.props.html, b.props.list, b.props.level]), [['a', 'outline', '0'], ['b', 'outline', '1'], ['c', 'bullet', '0'], ['d', 'number', '1']]);
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-lists-'));
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

test('engine: the galleries and a restart typed in the editor survive save → reopen, and a second save sends nothing', { skip: skip() }, async () => {
  const file = join(dir, 'lists.docx');
  await run(['create', file]);
  const typed = EN.blocksFromHtml(parse('<ol data-w-list="chinese"><li>第一章</li><ol><li>第一节</li></ol><li>第二章</li></ol><p>说明</p>'
    + '<ol><li>one</li><li>two</li></ol><ol data-w-restart="1"><li>one again</li><li>two again</li></ol><ol data-w-list="outline"><li>a</li><ol><li>a.a</li></ol></ol>'));
  await EN.planDocxBlocks(file, [], typed, run);
  const opened = await bodyBlocks(file);
  assert.deepEqual(opened.map(b => [b.props.html, b.props.list || '', b.props.level || '', b.props.restart || '']), [
    ['第一章', 'chinese', '0', ''], ['第一节', 'chinese', '1', ''], ['第二章', 'chinese', '0', ''], ['说明', '', '', ''],
    ['one', 'number', '0', ''], ['two', 'number', '0', ''], ['one again', 'number', '0', 'true'], ['two again', 'number', '0', ''], ['a', 'outline', '0', ''], ['a.a', 'outline', '1', '']]);
  assert.match(EN.blocksToHtml(opened), /<ol data-w-restart="1"><li data-path="[^"]+">one again<\/li><li data-path="[^"]+">two again<\/li><\/ol>/, 'the reopened restart is an ol of its own again');
  assert.equal(await EN.planDocxBlocks(file, opened, clone(opened), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
  // 继续编号: the restarted ol merged back into the one before it
  const merged = clone(opened); delete merged[6].props.restart;
  const calls = [];
  await EN.planDocxBlocks(file, opened, merged, async argv => { calls.push(argv.slice(2)); return run(argv); });
  assert.deepEqual(calls, [['/body/paragraph[7]', '--prop', 'restart=false']]);
  assert.equal((await bodyBlocks(file)).filter(b => b.props.restart).length, 0);
});
