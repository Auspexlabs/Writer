// node --test ui/tests/ — Word's 插入 and 引用 in the editor: equations (LaTeX, drawn by KaTeX, saved as Word's equations), text boxes
// and shapes floating in a paragraph, bookmarks, numbered captions and drop caps. The round trips run the real engine when the CLI is built.
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
const plain = x => JSON.parse(JSON.stringify(x));

test('equations and shapes come out of the tree with the paragraph they sit in, and go into the text as elements the editor draws', () => {
  const tree = [{ kind: 'paragraph', path: '/body/paragraph[1]', props: {}, children: [
    { kind: 'equation', path: '/body/paragraph[1]/equation[1]', props: { latex: '\\frac{a}{b}', at: '3' } },
    { kind: 'shape', path: '/body/paragraph[1]/shape[1]', props: { id: '4', geometry: 'rect', width: '5cm', height: '2.5cm', fill: 'FFFFFF', line: '000000', wrap: 'square', x: '1cm', y: '0cm', xFrom: 'column', yFrom: 'paragraph', text: '第一行\n第二行' } }] }];
  assert.deepEqual(EN.eqsOf(tree), [{ path: '/body/paragraph[1]', at: 3, latex: '\\frac{a}{b}', display: false }]);
  const [shape] = EN.shapesOf(tree);
  assert.deepEqual(plain(shape), { id: '4', path: '/body/paragraph[1]', props: { geometry: 'rect', width: '5cm', height: '2.5cm', fill: 'FFFFFF', line: '000000', wrap: 'square', xFrom: 'column', yFrom: 'paragraph', x: '1cm', y: '0cm', text: '第一行\n第二行' } });
  assert.equal(EN.shapeHtml(shape), '<span data-shape="1" data-sid="4" data-w-geometry="rect" data-w-width="5cm" data-w-height="2.5cm" data-w-fill="FFFFFF" data-w-line="000000" data-w-wrap="square" data-w-xfrom="column" data-w-yfrom="paragraph" data-w-x="1cm" data-w-y="0cm" contenteditable="false"><span class="wd-shtext" contenteditable="true">第一行<br>第二行</span></span>');
  assert.equal(EN.eqHtml('a<b', true), '<span data-eq="1" data-latex="a&lt;b" data-display="1" contenteditable="false"></span>');
  assert.match(EN.shapeSvg('ellipse', '4472C4', 'none'), /<ellipse [^>]*fill="#4472C4" stroke="none"/);
  assert.match(EN.shapeSvg('star5', 'none', '2F528F'), /<polygon points="[\d., ]+" fill="none" stroke="#2F528F"/);
  // the save sends a paragraph's text without them
  const [b] = EN.blocksFromHtml(parse('<p>ab<span data-eq="1" data-latex="x">x</span>c' + EN.shapeHtml(shape) + '</p>'));
  assert.equal(b.props.html, 'abc');
});

test('bookmarks, captions and drop caps are a paragraph\'s own props, kept as data-w-* and sent as they change', () => {
  const [cap, bm, drop] = EN.blocksFromHtml(parse('<p data-style="Caption" data-w-caption="图">图 1 流程</p><p data-w-bookmark="Results">结果</p><p data-w-dropcap="drop">春</p>'));
  assert.deepEqual([cap.props.caption, cap.props.style, bm.props.bookmark, drop.props.dropCap], ['图', 'Caption', 'Results', 'drop']);
  const html = EN.blocksToHtml(EN.blocksOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: '图 2 结构', caption: '图', bookmark: '_Ref123', style: 'Caption' } }], 'a.docx'));
  assert.match(html, /data-w-bookmark="_Ref123" data-w-caption="图"/);
});

// ----- the editor's ribbons -----
const src = readFileSync(new URL('../WordEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = { React: { createRef: () => ({ current: null }) }, setTimeout, clearTimeout, getComputedStyle: () => ({ marginBottom: '0px' }),
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), $lang: () => 'zh',
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(src + '\nglobalThis.WordEditor = Component;', ctx);

test('插入 offers text boxes, shapes, equations, a cover page, drop caps and bookmarks; 引用 captions and cross-references', () => {
  const c = new ctx.WordEditor(); c.EN = EN; c.props = { doc: { id: 'd', html: '' }, onChange() {} }; c.pgCss = { textContent: '' };
  const labels = tab => { c.state.tab = tab; return c.renderVals().ribbon.map(x => x.label).filter(Boolean); };
  const ins = labels('insert');
  for (const l of ['文本框', '形状', '公式', '封面', '首字下沉', '书签']) assert.ok(ins.includes(l), l);
  const refs = labels('refs');
  for (const l of ['插入脚注', '插入尾注', '插入题注', '交叉引用']) assert.ok(refs.includes(l), l);
  c.state.eqDlg = { el: null, latex: '\\frac{1}{2}' };
  const v = c.renderVals();
  assert.deepEqual([v.hasEq, v.eqTitle, v.eqLatex], [true, '插入公式', '\\frac{1}{2}']);
  v.eqPresets[0].onClick();
  assert.equal(c.state.eqDlg.latex, '\\frac{1}{2} \\frac{a}{b}', 'a preset goes after what is typed');
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-insert-'));
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
const never = async argv => { throw new Error('a second save sent ' + argv.join(' ')); };

test('engine: captions, bookmarks and drop caps typed in the editor survive save → reopen, and a second save sends nothing', { skip: skip() }, async () => {
  const file = join(dir, 'marks.docx');
  await run(['create', file]);
  const typed = EN.blocksFromHtml(parse('<p data-style="Caption" data-w-caption="图" style="text-align:center">图 1 流程</p><p data-w-bookmark="Results">结果</p><p data-w-dropcap="drop">春</p><p>眠不觉晓</p><p>见 <a href="#Results">结果</a></p>'));
  await EN.planDocxBlocks(file, [], typed, run);
  const opened = EN.blocksOf(await body(file), file);
  assert.deepEqual([opened[0].props.caption, opened[0].props.style, opened[1].props.bookmark, opened[2].props.dropCap], ['图', 'Caption', 'Results', 'drop']);
  assert.match(opened[0].props.html, /^图 1 流程$/);
  assert.match(opened[4].props.html, /<a href="#Results">结果<\/a>/);
  assert.equal(await EN.planDocxBlocks(file, opened, EN.blocksFromHtml(parse(EN.blocksToHtml(opened))), never), 0);
  const renumbered = EN.blocksFromHtml(parse(EN.blocksToHtml(opened))); renumbered[0].props.html = '图 2 流程';
  await EN.planDocxBlocks(file, opened, renumbered, run);
  const again = EN.blocksOf(await body(file), file);
  assert.deepEqual([again[0].props.caption, again[0].props.html], ['图', '图 2 流程'], 'the number changes and the field stays');
});

test('engine: equations and text boxes placed by the editor survive save → reopen; changes are set, and gone ones are removed', { skip: skip() }, async () => {
  const file = join(dir, 'objects.docx');
  await run(['create', file]);
  await EN.planDocxBlocks(file, [], EN.blocksFromHtml(parse('<p>面积是 平方米</p><p>第二段</p>')), run);
  const p1 = '/body/paragraph[1]';
  let r = await EN.planEquations(file, [], [{ path: p1, was: p1, block: {}, eqs: [{ latex: '\\pi r^2', at: 3, display: false }] }], () => false, run);
  assert.equal(r.count, 1);
  let eqs = EN.eqsOf(await body(file));
  assert.deepEqual(eqs, [{ path: p1, at: 3, latex: '\\pi r^{2}', display: false }]);
  assert.equal((await EN.planEquations(file, eqs, [{ path: p1, was: p1, block: {}, eqs: eqs.map(({ latex, at, display }) => ({ latex, at, display })) }], () => false, never)).count, 0, 'unchanged: nothing sent');
  await EN.planEquations(file, eqs, [{ path: p1, was: p1, block: {}, eqs: [{ latex: 'E=mc^{2}', at: 0, display: false }, { latex: '\\sqrt{2}', at: 5, display: false }] }], () => false, run);
  eqs = EN.eqsOf(await body(file));
  assert.deepEqual(eqs.map(e => [e.latex, e.at]), [['E=mc^{2}', 0], ['\\sqrt{2}', 5]]);
  assert.equal((await body(file))[0].props.text, '面积是 平方米', 'the text is untouched');

  const box = { el: null, sid: null, props: { geometry: 'rect', width: '5cm', height: '2.5cm', fill: 'FFFFFF', line: '000000', wrap: 'square', xFrom: 'column', yFrom: 'paragraph', x: '1cm', y: '0cm', text: '框里的字' } };
  r = await EN.planShapes(file, [], [{ path: p1, was: p1, shapes: [box] }], run);
  let shapes = EN.shapesOf(await body(file));
  assert.deepEqual(plain(shapes.map(s => [s.path, s.props])), [[p1, box.props]]);
  assert.equal((await EN.planShapes(file, shapes, [{ path: p1, was: p1, shapes: [{ sid: shapes[0].id, props: box.props }] }], never)).count, 0, 'unchanged: nothing sent');
  const changed = Object.assign({}, box.props, { fill: 'FFF2CC', text: '改过的字', x: '2cm' });
  r = await EN.planShapes(file, shapes, [{ path: '/body/paragraph[2]', was: '/body/paragraph[2]', shapes: [] }, { path: p1, was: p1, shapes: [{ sid: shapes[0].id, props: changed }] }], run);
  shapes = EN.shapesOf(await body(file));
  assert.deepEqual(plain(shapes.map(s => [s.props.fill, s.props.text, s.props.x])), [['FFF2CC', '改过的字', '2cm']]);
  const moved = await EN.planShapes(file, shapes, [{ path: '/body/paragraph[2]', was: '/body/paragraph[2]', shapes: [{ sid: shapes[0].id, props: changed }] }], run);
  shapes = EN.shapesOf(await body(file));
  assert.deepEqual(plain(shapes.map(s => [s.path, s.props.text])), [['/body/paragraph[2]', '改过的字']], 'moved to the other paragraph');
  await EN.planShapes(file, moved.list, [], run);
  assert.deepEqual(EN.shapesOf(await body(file)), []);
});
