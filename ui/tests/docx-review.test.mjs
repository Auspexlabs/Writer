// node --test ui/tests/ — 审阅 in the Word editor: comment threads (a reply shares its comment's place and is saved as Word's reply),
// 解决 / 重新打开, and the markup views (所有标记, 无标记, 原始版本, 批注 on or off).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import { install } from './dom-stub.mjs';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const parse = install();
const EN = await import('../engine.js');

test('planComments adds a reply beside its comment, by the id that comment has in the file now, without a quote', async () => {
  const calls = [];
  const exec = async argv => { calls.push(argv); return { path: argv[2] + '/comment[1]', props: { id: String(calls.length + 6) } }; };
  const r = await EN.planComments('a.docx', [], [
    { cid: 'c1', thread: '', text: '数据来源？', resolved: false, parent: '/body/paragraph[1]', origin: null, quote: '二十' },
    { cid: 'c2', thread: 'c1', text: '见附表', resolved: false, parent: '/body/paragraph[1]', origin: null, quote: '二十' }], null, exec);
  assert.deepEqual(calls.map(a => a.slice(2)), [
    ['/body/paragraph[1]', '--type', 'comment', '--prop', 'text=数据来源？', '--prop', 'quote=二十'],
    ['/body/paragraph[1]', '--type', 'comment', '--prop', 'text=见附表', '--prop', 'parent=7']]);
  assert.deepEqual(r.list.map(x => [x.cid, x.id]), [['c1', '7'], ['c2', '8']]);
});

const src = readFileSync(new URL('../WordEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = { React: { createRef: () => ({ current: null }) }, setTimeout, clearTimeout, getComputedStyle: () => ({ marginBottom: '0px' }),
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), $lang: () => 'zh',
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(src + '\nglobalThis.WordEditor = Component;', ctx);

test('the comments panel shows each comment with its replies, 解决 dims it, and the markup views reach the page', () => {
  let doc = { id: 'd', html: '', comments: [{ id: '1', author: 'Ann', text: '数据来源？', quote: '二十' }, { id: '2', author: 'Bo', parent: '1', text: '见附表' }] };
  const c = new ctx.WordEditor(); c.EN = EN; c.props = { get doc() { return doc; }, onChange(d) { doc = d; } }; c.pgCss = { textContent: '' };
  let v = c.renderVals();
  assert.equal(v.comments.length, 1, 'a reply is not a card of its own');
  assert.deepEqual(JSON.parse(JSON.stringify(v.comments[0].replies.map(r => r.text))), ['见附表']);
  v.comments[0].onResolve({ stopPropagation() { } });
  assert.equal(doc.comments[0].resolved, true);
  v = c.renderVals(); assert.deepEqual([v.comments[0].op, v.comments[0].resolveLabel], [0.55, '重新打开']);
  c.replyTo({ id: '1', quote: '二十' }, '好的');
  assert.deepEqual(doc.comments.map(x => [x.parent || '', x.text]), [['', '数据来源？'], ['1', '见附表'], ['1', '好的']]);
  c.edRef.current = { querySelectorAll: () => [], innerHTML: '' };
  c.delComment('1');
  assert.deepEqual(doc.comments, [], 'the replies go with their comment');
  assert.deepEqual([v.markupAttr, v.cmAttr], ['all', '1']);
  c.state.markup = 'none'; c.state.hideCm = true; doc = Object.assign({}, doc, { comments: [{ id: '3', text: 'x' }] });
  v = c.renderVals(); assert.deepEqual([v.markupAttr, v.cmAttr, v.hasComments], ['none', '0', false]);
});

// ----- round trip through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-review-'));
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

test('engine: a comment, its reply and 解决 survive save → reopen', { skip: !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx' }, async () => {
  const file = join(dir, 'review.docx');
  await run(['create', file]);
  await EN.planDocxBlocks(file, [], EN.blocksFromHtml(parse('<p>增长了二十</p>')), run);
  await EN.planComments(file, [], [
    { cid: 'c1', thread: '', text: '数据来源？', resolved: true, parent: '/body/paragraph[1]', origin: null, quote: '二十' },
    { cid: 'c2', thread: 'c1', text: '见附表', resolved: false, parent: '/body/paragraph[1]', origin: null, quote: '二十' }], null, run);
  const list = EN.commentsOf((await run(['get', file, '/body', '--depth', '4'])).children, 'Writer');
  assert.deepEqual(list.map(x => [x.text, x.quote, x.resolved, x.parent ? 'reply' : '']), [['数据来源？', '二十', true, ''], ['见附表', '二十', false, 'reply']]);
  assert.equal(list[1].parent, list[0].id);
});
