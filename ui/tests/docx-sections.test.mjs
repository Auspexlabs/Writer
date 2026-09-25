// node --test ui/tests/ — sections and notes in the Word editor: a paragraph that ends a section keeps its break and that section's own
// page setup as data-w-* (a new page after it in the page view), footnotes and endnotes go to the file as marks at character offsets
// with their text. The round trips run the real engine when the CLI is built.
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

test('notesOf reads footnotes and endnotes with the paragraph they hang on; planNotes adds, moves, edits and removes them by id', async () => {
  const list = EN.notesOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: {}, children: [
    { kind: 'footnote', path: '/body/paragraph[1]/footnote[1]', props: { id: '1', kind: 'footnote', text: 'a', at: '3' } },
    { kind: 'footnote', path: '/body/paragraph[1]/footnote[2]', props: { id: 'e1', kind: 'endnote', text: 'b', at: '5' } }] }]);
  assert.deepEqual(list.map(x => [x.nid, x.kind, x.path, x.at]), [['1', 'footnote', '/body/paragraph[1]', 3], ['e1', 'endnote', '/body/paragraph[1]', 5]]);
  const calls = [], exec = async argv => { calls.push(argv); return argv[0] === 'add' ? { path: argv[2] + '/footnote[1]', props: { id: '9' } } : {}; };
  const r = await EN.planNotes('a.docx', list, [
    { nid: '1', kind: 'footnote', text: 'a2', parent: '/body/paragraph[1]', at: 3 }, // text changed
    { nid: 'e1', kind: 'endnote', text: 'b', parent: '/body/paragraph[2]', at: 0 }, // moved to another paragraph
    { nid: 'nx', kind: 'footnote', text: 'new', parent: '/body/paragraph[1]', at: 1 },
    { nid: 'gone', kind: 'footnote', text: '', parent: null, at: 0 }], null, exec);
  assert.deepEqual(calls, [
    ['set', 'a.docx', '//footnote[@id=1]', '--prop', 'text=a2'],
    ['remove', 'a.docx', '//footnote[@id=e1]'], ['add', 'a.docx', '/body/paragraph[2]', '--type', 'footnote', '--prop', 'kind=endnote', '--prop', 'text=b', '--prop', 'at=0'],
    ['add', 'a.docx', '/body/paragraph[1]', '--type', 'footnote', '--prop', 'kind=footnote', '--prop', 'text=new', '--prop', 'at=1']]);
  assert.equal(r.count, 4);
  assert.deepEqual(r.list.map(x => [x.nid, x.id]), [['1', '1'], ['e1', '9'], ['nx', '9']]);
  const marks = parse('<p>a<sup data-fn="x" data-kind="endnote">?</sup>b<sup data-fn="y">?</sup><sup data-fn="z" data-kind="endnote">?</sup></p>');
  assert.deepEqual(EN.numberNotes(marks).map(m => m.textContent), ['i', '1', 'ii'], 'footnotes count 1, 2…, endnotes i, ii… as Word numbers them');
});

test('a paragraph\'s section break and its section\'s own page setup come back from the html; the setup is only sent with a break', () => {
  const [a, b] = EN.blocksFromHtml(parse('<p data-w-sectionbreak="nextPage" data-w-orientation="landscape" data-w-columns="2">one</p><p data-w-orientation="landscape">two</p>'));
  assert.deepEqual([a.props.sectionBreak, a.props.orientation, a.props.columns], ['nextPage', 'landscape', '2']);
  assert.equal(b.props.orientation, 'landscape');
  const html = EN.blocksToHtml(EN.blocksOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'x', sectionBreak: 'continuous', page: 'A4', orientation: 'portrait', margin: 'normal', columns: '1' } }], 'a.docx'));
  assert.match(html, /data-w-sectionbreak="continuous" data-w-page="A4" data-w-orientation="portrait" data-w-margin="normal" data-w-columns="1"/);
});

// ----- the editor: a section that starts on a new page starts one in the page view -----
const src = readFileSync(new URL('../WordEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = { React: { createRef: () => ({ current: null }) }, setTimeout, clearTimeout, getComputedStyle: () => ({ marginBottom: '0px' }),
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), $lang: () => 'zh',
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(src + '\nglobalThis.WordEditor = Component;', ctx);

test('the page view: a nextPage section break starts the next page, a continuous one does not', () => {
  const c = new ctx.WordEditor(); c.EN = EN; c.props = { doc: { id: 'd', html: '' }, onChange() {} }; c.pgCss = { textContent: '' };
  const ed = { innerText: 'text', querySelectorAll: () => [] };
  const block = (top, h, sb) => ({ offsetTop: top, offsetHeight: h, offsetParent: ed, parentElement: ed, children: [], querySelectorAll: () => [], matches: () => false, getAttribute: n => n === 'data-w-sectionbreak' ? sb || null : null });
  ed.children = [block(0, 100, 'nextPage'), block(100, 100, 'continuous'), block(200, 100)];
  c.edRef.current = ed; c.refreshInfo();
  assert.equal(c.state.info.pages, 2);
  assert.deepEqual(JSON.parse(JSON.stringify(c.state.info.at)), [[0, 100], [100, 300]]);
  const v = c.renderVals();
  assert.ok(v.tabs.some(t => t.label === '引用'), 'a 引用 tab holds 目录 and the notes');
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-sections-'));
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
const body = async file => (await run(['get', file, '/body', '--depth', '4'])).children;
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';

test('engine: a section break with its own landscape page and two columns, and footnotes and endnotes, survive save → reopen', { skip: skip() }, async () => {
  const file = join(dir, 'sections.docx');
  await run(['create', file]);
  const typed = EN.blocksFromHtml(parse('<p data-w-sectionbreak="nextPage" data-w-orientation="landscape" data-w-columns="2">横向两栏的一节</p><p>下一节</p>'));
  await EN.planDocxBlocks(file, [], typed, run);
  let blocks = EN.blocksOf(await body(file), file);
  assert.deepEqual([blocks[0].props.sectionBreak, blocks[0].props.orientation, blocks[0].props.columns], ['nextPage', 'landscape', '2']);
  assert.equal(blocks[1].props.sectionBreak, undefined);
  const shown = () => EN.blocksFromHtml(parse(EN.blocksToHtml(blocks)));
  assert.equal(await EN.planDocxBlocks(file, blocks, shown(), async argv => { throw new Error('a second save sent ' + argv.join(' ')); }), 0);
  const edited = shown(); edited[0].props.sectionBreak = 'continuous';
  await EN.planDocxBlocks(file, blocks, edited, run);
  blocks = EN.blocksOf(await body(file), file);
  assert.deepEqual([blocks[0].props.sectionBreak, blocks[0].props.orientation], ['continuous', 'landscape']);

  let r = await EN.planNotes(file, [], [{ nid: 'n1', kind: 'footnote', text: '脚注的文字', parent: '/body/paragraph[1]', at: 2 }, { nid: 'n2', kind: 'endnote', text: '尾注', parent: '/body/paragraph[2]', at: 3 }], null, run);
  let notes = EN.notesOf(await body(file));
  assert.deepEqual(notes.map(x => [x.kind, x.text, x.path, x.at]), [['footnote', '脚注的文字', '/body/paragraph[1]', 2], ['endnote', '尾注', '/body/paragraph[2]', 3]]);
  assert.equal((await body(file))[0].props.text, '横向两栏的一节', 'the mark is not text');
  r = await EN.planNotes(file, r.list, [{ nid: 'n1', kind: 'footnote', text: '改过', parent: '/body/paragraph[1]', at: 2 }, { nid: 'n2', kind: 'endnote', text: '尾注', parent: null, at: 0 }], null, run);
  notes = EN.notesOf(await body(file));
  assert.deepEqual(notes.map(x => [x.kind, x.text, x.at]), [['footnote', '改过', 2]], 'the text changed in place; the endnote whose mark went is gone');
});
