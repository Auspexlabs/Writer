import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

globalThis.window = globalThis;
const EN = await import('../engine.js'), K = await import('../office-io.js'), P = await import('../picture.js');
const { open, save, slideProps } = EN;

const slide = {
  kind: 'slide', path: '/slide[@id=256]', props: {
    id: '256', layout: 'Content', background: 'F7F8F9', notes: 'Speaker note',
    hidden: 'true', transition: 'wipe', duration: '700'
  }, children: [
    { kind: 'decor', path: '/slide[@id=256]/decor[1]', props: { id: '3', source: 'master', type: 'shape', text: 'Brand', x: '1cm', y: '1cm', w: '3cm', h: '2cm', fill: 'AABBCC' } },
    { kind: 'decor', path: '/slide[@id=256]/decor[2]', props: { source: 'layout', type: 'image', src: 'image1.png', x: '0cm', y: '0cm', w: '2cm', h: '2cm' } },
    { kind: 'image', path: '/slide[@id=256]/image[@id=4]', props: { id: '4', x: '2cm', y: '2cm', w: '8cm', h: '3cm' } }
  ]
};
const documentTree = { kind: 'document', props: { width: '33.867cm', height: '19.05cm' }, children: [slide] };
const response = json => ({ ok: true, json: async () => json });

test('PPTX open keeps decor separate and reads slide metadata', async () => {
  const prev = globalThis.fetch;
  globalThis.fetch = async url => {
    assert.match(String(url), /^\/json\?/);
    return response(documentTree);
  };
  try {
    const doc = await open({ id: 'p1', path: 'deck.pptx', type: 'pptx' });
    const s = doc.slides[0];
    assert.equal(s.notes, 'Speaker note');
    assert.equal(s.trans, 'wipe');
    assert.equal(s.duration, 700);
    assert.equal(s.hidden, true);
    assert.equal(s.decor.length, 2);
    assert.equal(s.decor[0].html, '<p>Brand</p>');
    assert.match(s.decor[1].src, /\/binary\?file=deck.pptx&path=/);
    assert.equal(s.objs.length, 1);
    assert.equal(s.objs[0].kind, 'image');
    assert.equal(await save(doc), 0, 'opening the file does not rewrite inherited elements or slide metadata');
  } finally { globalThis.fetch = prev; }
});

test('PPTX save writes changed notes, visibility and transition without editing decor', async () => {
  const prev = globalThis.fetch, commands = [];
  globalThis.fetch = async (url, opts) => {
    if (String(url).startsWith('/json?')) return response(documentTree);
    assert.equal(url, '/run');
    const { argv } = JSON.parse(opts.body); commands.push(argv);
    return response({ code: 0, output: '{}' });
  };
  try {
    const doc = await open({ id: 'p1', path: 'deck.pptx', type: 'pptx' });
    const s = doc.slides[0];
    s.notes = 'Updated'; s.hidden = false; s.trans = 'fade'; s.duration = 900;
    assert.equal(await save(doc), 1);
    assert.deepEqual(commands, [[
      'set', 'deck.pptx', '/slide[@id=256]',
      '--prop', 'notes=Updated', '--prop', 'hidden=false', '--prop', 'transition=fade', '--prop', 'duration=900'
    ]]);
    assert.equal(await save(doc), 0, 'saved snapshot includes slide properties');
  } finally { globalThis.fetch = prev; }
});

test('decor ids come from the path, so master and layout shapes sharing a cNvPr id stay distinct', async () => {
  const box = { x: '0cm', y: '0cm', w: '1cm', h: '1cm' };
  const twins = { ...slide, children: [
    { kind: 'decor', path: '/slide[1]/decor[1]', props: { id: '2', source: 'master', type: 'shape', text: 'A', ...box } },
    { kind: 'decor', path: '/slide[1]/decor[2]', props: { id: '2', source: 'layout', type: 'shape', text: 'B', ...box } },
    { kind: 'decor', path: '/slide[1]/decor[3]', props: { source: 'master', type: 'shape', geometry: 'line', line: 'FF5722', x: '0cm', y: '17cm', w: '33.867cm', h: '0cm' } },
    { kind: 'shape', path: '/slide[1]/shape[1]', props: { id: '2', ...box } }
  ] };
  const prev = globalThis.fetch;
  globalThis.fetch = async () => response({ ...documentTree, children: [twins] });
  try {
    const s = (await open({ id: 'p2', path: 'deck.pptx', type: 'pptx' })).slides[0];
    assert.deepEqual(s.decor.map(d => d.id), ['d/slide[1]/decor[1]', 'd/slide[1]/decor[2]', 'd/slide[1]/decor[3]']);
    assert.deepEqual(s.objs.map(o => o.id), ['e/slide[@id=256]/shape[@id=2]'], 'a slide\'s own shape too: cNvPr ids start over on every slide');
    assert.deepEqual([s.decor[2].h, s.decor[2].fill, s.decor[2].t], [2, '#FF5722', 'shape'], 'a zero-height connector draws as a thin bar');
  } finally { globalThis.fetch = prev; }
});

test('PPTX slide prop diff leaves unsupported transitions untouched', () => {
  const old = { notes: 'x', hidden: false, trans: 'other', duration: 500 };
  assert.deepEqual(slideProps(old, { ...old, notes: 'y' }), { notes: 'y' });
  assert.deepEqual(slideProps(old, { ...old, trans: 'none', duration: null }), { transition: 'none' });
});

test('new slide leaves inherited decor in place while clearing editable placeholders', async () => {
  const prev = globalThis.fetch, commands = [];
  globalThis.fetch = async (url, opts) => {
    if (String(url).startsWith('/json?')) return response(documentTree);
    const { argv } = JSON.parse(opts.body); commands.push(argv);
    const output = argv[0] === 'add' ? { props: { id: '300' } }
      : argv[0] === 'get' ? { children: [
        { kind: 'decor', path: '/slide[@id=300]/decor[1]' },
        { kind: 'shape', path: '/slide[@id=300]/shape[@id=5]' }
      ] } : {};
    return response({ code: 0, output: JSON.stringify(output) });
  };
  try {
    const doc = await open({ id: 'p1', path: 'deck.pptx', type: 'pptx' });
    doc.slides.push({ id: 'new', layout: 'Content', decor: [], objs: [], notes: '', trans: 'none', hidden: false, bg: '#FFFFFF' });
    await save(doc);
    assert.ok(commands.some(c => c[0] === 'add' && c.includes('layout=Content')));
    assert.ok(commands.some(c => c[0] === 'remove' && c[2] === '/slide[@id=300]/shape[@id=5]'));
    assert.ok(!commands.some(c => c[0] === 'remove' && c[2].includes('/decor[')));
  } finally { globalThis.fetch = prev; }
});

// ----- the real engine (when the CLI is built): open → duplicate in the slide editor → save → reopen -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';
let server, base, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-slides-'));
  server = spawn('dotnet', [cli, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  base = await new Promise((resolve, reject) => {
    let err = '';
    server.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1]); });
    server.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
  });
});
after(() => { server && server.kill(); dir && rmSync(dir, { recursive: true, force: true }); });
/** Runs fn with the bridge talking to the engine. */
async function engine(fn) { const prev = globalThis.fetch; globalThis.fetch = (u, o) => prev(new URL(u, base), o); try { return await fn(); } finally { globalThis.fetch = prev; } }
globalThis.DOMParser ??= class { parseFromString() { return { body: { children: [] } }; } }; // a save looks for list items in a new text box; these have none
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
/** A deck whose slide 1 holds a text box and a cropped grey picture, and slide 2 nothing; opened as the editor opens it. */
async function deck(file) {
  for (const argv of [['create', file], ['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank'], ['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank'],
    ['add', file, '/slide[1]', '--type', 'shape', '--prop', 'text=Hello', '--prop', 'x=2cm', '--prop', 'y=2cm', '--prop', 'w=12cm', '--prop', 'h=3cm'],
    ['add', file, '/slide[1]', '--type', 'image', '--prop', 'src=' + PNG, '--prop', 'x=4cm', '--prop', 'y=6cm', '--prop', 'w=8cm', '--prop', 'h=6cm'],
    ['set', file, '/slide[1]/image[1]', '--prop', 'crop=10,0,10,0', '--prop', 'grayscale=true', '--prop', 'alt=logo']]) await EN.run(argv);
  return EN.open({ id: file, path: file, type: 'pptx' });
}
/** The slide editor on the doc that get() returns; its changes go to set(d). */
function editor(get, set) {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const c = new ctx.SlideEditor(); c.props = { get doc() { return get(); }, onChange: set };
  return Object.assign(c, { K, P, EN });
}
const key = (c, k) => c.renderVals().onRootKey({ key: k, ctrlKey: true, target: { tagName: 'DIV' }, preventDefault() { } });
/** An object as the editor shows it, moved by d: kind, box, and its text and type or its picture's look. */
const seen = (o, d = 0) => [o.t, o.x + d, o.y + d, o.w, o.h, o.t === 'image' ? o.look : [o.html.replace(/<[^>]*>/g, ''), o.fs, o.color, o.font]];
const content = s => s.objs.map(o => seen(o));

test('engine: a duplicated slide is saved beside the original, picture and all', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'slides.pptx');
  let doc = await deck(file);
  const c = editor(() => doc, d => { doc = d; }), first = doc.slides[0].path;
  c.renderVals().ribbon.find(i => i.title === '复制幻灯片').onClick(); // 开始 › 复制
  c.setState({ pop: { id: 'thumb', i: 0, x: 0, y: 0 } });
  c.renderVals().popItems.find(i => i.label === '复制幻灯片').onClick(); // the thumbnail's menu
  await EN.save(doc);
  const back = await EN.open({ id: file, path: file, type: 'pptx' }), one = content(back.slides[0]);
  assert.equal(back.slides[0].path, first, 'the original stays first');
  assert.deepEqual(one.map(o => o[0]), ['text', 'image']);
  assert.deepEqual(back.slides.map(content), [one, one, one, []]);
}));

test('engine: a duplicated, copied or cut object is saved as a new object, and the original stays', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'objects.pptx');
  let doc = await deck(file);
  const c = editor(() => doc, d => { doc = d; }), [text, pic] = doc.slides[0].objs;
  c.setState({ sel: text.id }); key(c, 'd'); // Ctrl+D
  c.setState({ sel: pic.id, tab: 'picture' }); c.renderVals().ribbon.find(i => i.label === '复制').onClick(); // the 图片 tab's 复制
  const held = structuredClone(doc); // the doc the editor holds, had it changed while the save ran
  await EN.save(doc); EN.adopt(held, doc);
  const copy = held.slides[0].objs[3];
  assert.equal(copy.src, '/binary?file=' + encodeURIComponent(file) + '&path=' + encodeURIComponent(copy.path), 'the saved copy loads from its own place in the file');
  c.setState({ sel: copy.id }); key(c, 'x'); // cut the picture's copy…
  await EN.save(doc); // …which this save takes out of the file
  c.setState({ cur: 1, sel: null }); key(c, 'v'); // …and paste it on slide 2
  await EN.save(doc);
  const back = await EN.open({ id: file, path: file, type: 'pptx' }), [t, p] = back.slides[0].objs;
  assert.deepEqual([t.path, p.path], [text.path, pic.path], 'the originals stay');
  assert.deepEqual(content(back.slides[0]), [seen(t), seen(p), seen(t, 32)]);
  assert.deepEqual(content(back.slides[1]), [seen(p, 56)]);
  assert.equal((await EN.run(['get', file, back.slides[1].objs[0].path])).props.alt, 'logo', 'the pasted picture is the cut one exactly, alt text and all');
}));

test('engine: a copy is the original\'s own markup — a title copied stays a title placeholder, with what it inherits', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'title.pptx');
  await deck(file);
  await EN.run(['add', file, '/', '--type', 'slide', '--prop', 'layout=Title', '--prop', 'title=Q4']);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; }), title = doc.slides[2].objs[0];
  c.setState({ cur: 2, sel: title.id }); key(c, 'd');
  await EN.save(doc);
  const back = await EN.open({ id: file, path: file, type: 'pptx' }), [t, , copy] = back.slides[2].objs;
  assert.deepEqual([t.ph, t.bold], ['title', true]);
  assert.deepEqual([copy.ph, copy.bold, copy.fs, copy.html, copy.x - t.x], [t.ph, t.bold, t.fs, t.html, 32]);
}));

test('engine: a paste into another deck brings what was copied, not what sits at the same place in this one', { skip: skip() }, () => engine(async () => {
  const [fa, fb] = [join(dir, 'a.pptx'), join(dir, 'b.pptx')];
  await deck(fa);
  await EN.run(['set', fa, '/slide[1]/shape[1]', '--prop', 'text=From A']);
  const a = await EN.open({ id: fa, path: fa, type: 'pptx' });
  let doc = a;
  const c = editor(() => doc, d => { doc = d; });
  c.setState({ sel: a.slides[0].objs[0].id }); key(c, 'c');
  doc = await deck(fb); c.setState({ cur: 1, sel: null }); key(c, 'v'); // b's text box has the path a's has
  await EN.save(doc);
  assert.deepEqual((await EN.open({ id: fb, path: fb, type: 'pptx' })).slides[1].objs.map(o => o.html.replace(/<[^>]*>/g, '')), ['From A']);
}));

test('engine: undoing a deletion a save already made puts the picture back exactly', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'undo.pptx');
  let doc = await deck(file);
  const c = editor(() => doc, d => { doc = d; }), before = structuredClone(doc), pic = doc.slides[0].objs[1];
  c.setState({ sel: pic.id }); c.renderVals().onRootKey({ key: 'Delete', target: { tagName: 'DIV' }, preventDefault() { } });
  await EN.save(doc);
  assert.equal((await EN.open({ id: file, path: file, type: 'pptx' })).slides[0].objs.length, 1, 'the save took it out');
  doc = EN.adopt(structuredClone(before), doc); // what 撤销 does (index.dc.html undo)
  await EN.save(doc);
  const back = await EN.open({ id: file, path: file, type: 'pptx' }), p = back.slides[0].objs[1];
  assert.deepEqual(seen(p), seen(pic));
  assert.equal(p.path, pic.path, 'the same picture: its drawing id was free');
  assert.equal((await EN.run(['get', file, p.path])).props.alt, 'logo');
}));
