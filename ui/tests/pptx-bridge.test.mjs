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

test('a slide from the layout gallery binds its placeholders to the layout\'s own: typed text goes into them, empty ones stay empty and unwritten', async () => {
  const prev = globalThis.fetch, commands = [];
  const ph = (id, placeholder, x) => ({ kind: 'shape', path: `/slide[@id=300]/shape[@id=${id}]`, props: { id: String(id), placeholder, text: '', x, y: '4.995cm', w: '13.547cm', h: '12.361cm' }, computed: { size: '20pt', color: '1D1D1F' } });
  globalThis.fetch = async (url, opts) => {
    if (String(url).startsWith('/json?')) return response(documentTree);
    const { argv } = JSON.parse(opts.body); commands.push(argv);
    const output = argv[0] === 'add' ? { props: { id: '300' } }
      : argv[0] === 'get' ? { children: [{ kind: 'decor', path: '/slide[@id=300]/decor[1]' }, ph(2, 'title', '2.709cm'), ph(3, 'body', '2.709cm'), ph(4, 'body', '17.611cm')] } : {};
    return response({ code: 0, output: JSON.stringify(output) });
  };
  try {
    const doc = await open({ id: 'p1', path: 'deck.pptx', type: 'pptx' });
    const s = K.makeSlide('two', '16:9'); s.objs[0].html = '<p>Q3</p>'; doc.slides.push(s);
    await save(doc);
    assert.ok(commands.some(c => c[0] === 'add' && c.includes('layout=two')), 'the slide is made of the layout');
    assert.deepEqual(s.objs.map(o => o.path), ['/slide[@id=300]/shape[@id=2]', '/slide[@id=300]/shape[@id=3]', '/slide[@id=300]/shape[@id=4]'], 'the title to the title, the columns to the content placeholders, in order');
    assert.deepEqual(commands.filter(c => c[0] === 'set' && c[2].includes('/shape[')).map(c => c.slice(2)), [['/slide[@id=300]/shape[@id=2]', '--prop', 'html=<p>Q3</p>']], 'only the typed title is written: the empty columns keep their prompts, no box or size is pinned');
    assert.ok(!commands.some(c => c[0] === 'remove'), 'no placeholder is cleared away');
    assert.equal(await save(doc), 0);
  } finally { globalThis.fetch = prev; }
});

test('a layout slot of a new slide is drawn where the file puts it: LAYOUT_SPECS and PptxTemplate.Layouts agree on the placeholders', () => {
  const kinds = Object.fromEntries(Object.keys(K.LAYOUT_SPECS).map(k => [k, K.makeSlide(k, '16:9').objs.map(o => K.phFamily(o.ph))]));
  assert.deepEqual(kinds, { // PptxLayoutTests: ctrTitle/subTitle, title/obj, title/body, title/obj/obj, title/body/obj/body/obj, title, -, title/obj/body, title/pic/body, title/body
    title: ['title', 'body'], content: ['title', 'body'], section: ['title', 'body'], two: ['title', 'body', 'body'], comparison: ['title', 'body', 'body', 'body', 'body'],
    titleOnly: ['title'], blank: [], caption: ['title', 'body', 'body'], picture: ['title', 'pic', 'body'], quote: ['title', 'body'] });
  const two = K.makeSlide('two', '16:9');
  assert.deepEqual(two.objs.map(o => [o.x, o.y, o.w, o.h, o.html]), [[128, 56, 1344, 140, ''], [128, 236, 640, 584, ''], [832, 236, 640, 584, '']], 'empty, on the grid the file scales from');
  assert.deepEqual([K.objView(two.objs[0], K.THEMES.paper).hint, K.objView(two.objs[1], K.THEMES.paper).hint, K.objView({ ...two.objs[0], html: '<p>Q3</p>' }, K.THEMES.paper).hint], ['单击此处添加标题', '单击此处添加文本', '']);
  assert.equal(K.makeSlide('title', '4:3').objs[1].y, Math.round(508 * 1200 / 900), 'a 4:3 deck stretches the grid down');
  assert.deepEqual(K.makeSlide('title', '16:9').decor.map(d => [d.t, d.x, d.y, d.w, d.h]), [['shape', 128, 482, 96, 6]], 'the layout\'s accent bar is decor: drawn, never saved');
  assert.deepEqual(Object.keys(K.THEMES), ['ink', 'paper', 'sea', 'clay', 'mist', 'sand', 'rose', 'night'], 'the engine\'s palette= keys');
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
// a save looks for list items in a new text box (these have none) and compares a placeholder's text with the file's (plain text is enough here)
globalThis.DOMParser ??= class { parseFromString(s) { return { body: { nodeType: 1, tagName: 'BODY', style: {}, children: [], childNodes: [{ nodeType: 3, nodeValue: s.replace(/<[^>]*>/g, '') }] } }; } };
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
function editor(get, set, win) { // win: a window two editors share, as two decks open in one window do
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: win || { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone,
    // a stand-in for ui/i18n.js: no dictionary loaded, so every call falls back to the Chinese (the '@@' context dropped), as in Chinese mode
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const c = new ctx.SlideEditor(); c.props = { get doc() { return get(); }, onChange: set };
  return Object.assign(c, { K, P, EN });
}
const key = (c, k) => c.renderVals().onRootKey({ key: k, ctrlKey: true, target: { tagName: 'DIV' }, preventDefault() { } });
/** The layout gallery as opened by 新建幻灯片 / the thumbnails' ＋ (mode new, a slide at index at) or 版式 (mode relayout); picks the layout called label. */
const pick = (c, mode, label, at) => { c.setState({ pop: { id: 'g', mode, at, x: 0, y: 0 } }); c.renderVals().gallery.find(g => g.label === label).onClick(); };

test('版式 moves placeholders as the engine will: text kept in the first free slot of its family, an empty leftover dropped, missing slots added empty', () => {
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [K.makeSlide('title', '16:9')] };
  doc.slides[0].objs[0].html = '<p>Q3</p>'; doc.slides[0].decor.push(Object.assign(K.shape({ html: '' }), { source: 'master' }));
  const c = editor(() => doc, d => { doc = d; });
  pick(c, 'relayout', '两栏内容');
  let s = doc.slides[0];
  const seen = objs => Array.from(objs, o => [o.ph, o.html, o.x, o.y]); // the editor's arrays come from its vm
  assert.equal(s.layout, 'two');
  assert.deepEqual(seen(s.objs), [['title', '<p>Q3</p>', 128, 56], ['body', '', 128, 236], ['body', '', 832, 236]], 'the empty subtitle became the first column');
  assert.deepEqual(Array.from(s.decor, d => d.source), ['master'], 'the title layout\'s bar went with it; the master\'s decor stays');
  s.objs[1].html = '<p>左</p>';
  pick(c, 'relayout', '仅标题');
  s = doc.slides[0];
  assert.deepEqual(seen(s.objs), [['title', '<p>Q3</p>', 128, 56], ['body', '<p>左</p>', 128, 236]], 'the column with text keeps its place, the empty one goes');
  pick(c, 'new', '引用', 1);
  assert.deepEqual([doc.slides.length, doc.slides[1].layout, Array.from(doc.slides[1].objs, o => o.ph), c.state.cur], [2, 'quote', ['title', 'sub'], 1]);
});

test('✦ 美化 asks the assistant, through the shell, for the current slide or the whole deck', () => {
  const asked = [];
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [K.makeSlide('title', '16:9'), K.makeSlide('content', '16:9')] };
  const c = editor(() => doc, d => { doc = d; }); c.props.onAsk = t => asked.push(t);
  c.setState({ cur: 1 }); c.renderVals();
  c.menus.beautify[0].onClick(); c.menus.beautify[1].onClick();
  assert.match(asked[0], /^美化第 2 页幻灯片/);
  assert.match(asked[1], /^美化整份幻灯片/);
});
/** An object as the editor shows it, moved by d: kind, box, and its text and type or its picture's look. */
const seen = (o, d = 0) => [o.t, o.x + d, o.y + d, o.w, o.h, o.t === 'image' ? o.look : [o.html.replace(/<[^>]*>/g, ''), o.fs, o.color, o.font]];
const content = s => s.objs.map(o => seen(o));

test('engine: a slide from the gallery is saved with the layout\'s real placeholders, which come back as such; 版式 moves them in the file too', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'layouts.pptx');
  await EN.run(['create', file]);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; });
  pick(c, 'new', '两栏内容', 0);
  doc.slides[0].objs[0].html = '<p>Q3 回顾</p>'; // typed into the title; both columns stay empty
  await EN.save(doc);
  let back = await EN.open({ id: file, path: file, type: 'pptx' });
  const text = s => s.objs.map(o => [o.ph, o.html.replace(/<[^>]*>/g, '')]);
  assert.equal(back.slides[0].layout, 'Two Content');
  assert.deepEqual(text(back.slides[0]), [['title', 'Q3 回顾'], ['body', ''], ['body', '']]);
  const column = await EN.run(['get', file, back.slides[0].objs[2].path]);
  assert.deepEqual([column.props.placeholder, column.props.text, column.props.x], ['body', '', '17.611cm'], 'a real content placeholder, placed by the layout');
  assert.equal(await EN.save(back), 0, 'reopening writes nothing');

  doc = back; const c2 = editor(() => doc, d => { doc = d; });
  pick(c2, 'relayout', '仅标题');
  await EN.save(doc);
  back = await EN.open({ id: file, path: file, type: 'pptx' });
  assert.deepEqual([back.slides[0].layout, text(back.slides[0])], ['Title Only', [['title', 'Q3 回顾']]], 'the empty columns went with the layout');
}));

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

test('engine: a table round-trips — style flags, column widths, a merge and a cell\'s own fill and line', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'table.pptx');
  await EN.run(['create', file]); await EN.run(['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank']);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; });
  c.insertTable(3, 3); const t = doc.slides[0].objs[0];
  c.setState({ tab: 'table', bubble: true, cell: { r: 1, c: 0 }, cellSel: { r: 2, c: 1 } });
  c.renderVals().ribbon.find(i => i.label === '合并单元格').onClick();
  c.setState({ cell: { r: 0, c: 2 }, cellSel: null });
  c.renderVals().ribbon.find(i => i.isColor && i.label === '填充').onChange({ target: { value: '#D9E2F3' } });
  c.renderVals().ribbon.find(i => i.isColor && i.label === '边框').onChange({ target: { value: '#1F2937' } });
  c.menus.tstyle.find(i => i.label === '浅色样式 2').onClick(); c.renderVals().ribbon.find(i => i.label === '第一列').onClick();
  c.patchSel({ colW: [280, 500, 500], w: 1280 });
  await EN.save(doc);
  let back = await EN.open({ id: file, path: file, type: 'pptx' }), b = back.slides[0].objs[0];
  assert.deepEqual([b.t, b.rows, b.merges, b.cells, b.tstyle, b.header, b.banded, b.firstCol], ['table', [['标题 1', '标题 2', '标题 3'], ['', '', ''], ['', '', '']], [{ r: 1, c: 0, rs: 2, cs: 2 }], { '0:2': { fill: '#D9E2F3', line: '#1F2937' } }, 'LightStyle2Accent1', true, true, true]);
  assert.deepEqual(b.colW.map(w => Math.round(w / 10)), [28, 50, 50]);
  assert.equal(await EN.save(back), 0, 'reopening writes nothing');
  doc = back; const c2 = editor(() => doc, d => { doc = d; });
  c2.setState({ tab: 'table', bubble: true, sel: b.id, sels: [b.id], cell: { r: 1, c: 0 } });
  c2.renderVals().ribbon.find(i => i.label === '拆分单元格').onClick(); c2.renderVals().ribbon.find(i => i.label === '上方插入行').onClick();
  await EN.save(doc);
  back = await EN.open({ id: file, path: file, type: 'pptx' }); b = back.slides[0].objs[0];
  assert.deepEqual([b.rows.length, b.merges, b.cells], [4, [], { '0:2': { fill: '#D9E2F3', line: '#1F2937' } }], 'the merge is gone, the row is in, the styled cell stayed where it was');
  void t;
}));

test('engine: animations and a morph round-trip — PowerPoint\'s presets in click order; deleting an object takes its effects along', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'anims.pptx');
  await EN.run(['create', file]); await EN.run(['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank']);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; });
  c.insertText(); c.insertShape('star5');
  const [a, b] = doc.slides[0].objs, pickObj = o => c.setState({ sel: o.id, sels: [o.id] });
  pickObj(a); c.addFx('fade'); pickObj(b); c.addFx('fly'); c.patchFx({ start: 'with', delay: 200 }); pickObj(a); c.addFx('spin'); c.patchFx({ start: 'after' });
  c.commit((d, s) => { s.trans = 'morph'; s.duration = 1200; });
  await EN.save(doc);
  let back = await EN.open({ id: file, path: file, type: 'pptx' }), s = back.slides[0];
  assert.deepEqual(s.anims.map(x => [x.fx, x.start, x.delay, x.id]), [['fade', 'click', 0, s.objs[0].id], ['fly', 'with', 200, s.objs[1].id], ['spin', 'after', 0, s.objs[0].id]]);
  assert.deepEqual([s.trans, s.duration], ['morph', 1200]);
  assert.equal(await EN.save(back), 0, 'reopening writes nothing');
  doc = back; const c2 = editor(() => doc, d => { doc = d; });
  c2.setState({ sel: s.objs[1].id, sels: [s.objs[1].id] }); c2.deleteSels();
  await EN.save(doc);
  back = await EN.open({ id: file, path: file, type: 'pptx' });
  assert.deepEqual(back.slides[0].anims.map(x => x.fx), ['fade', 'spin']);
}));

test('engine: notes typed under the slide go to its notes page and come back; the divider drags the pane taller', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'notes.pptx');
  await EN.run(['create', file]); await EN.run(['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank']);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; });
  c.renderVals().onNotes({ target: { value: 'First point\nSecond point' } });
  await EN.save(doc);
  assert.equal((await EN.run(['get', file, '/slide[1]'])).props.notes, 'First point\nSecond point');
  assert.equal((await EN.open({ id: file, path: file, type: 'pptx' })).slides[0].notes, 'First point\nSecond point');
  c.setState({ box: { w: 1200, h: 700 } }); c.renderVals().notesGrab({ preventDefault() { }, clientY: 600 }); c.onWM({ clientX: 0, clientY: 500 }); c.onWU();
  assert.equal(c.renderVals().notesH, '184px');
}));

test('engine: sections are saved and read back; a slide copied into another deck is made there, on a blank layout when that deck lacks its own', { skip: skip() }, () => engine(async () => {
  const fa = join(dir, 'secA.pptx'), fb = join(dir, 'secB.pptx'), win = { innerWidth: 1360, innerHeight: 860 };
  for (const f of [fa, fb]) { await EN.run(['create', f]); for (let i = 0; i < 3; i++) await EN.run(['add', f, '/', '--type', 'slide', '--prop', 'layout=Blank']); }
  let A = await EN.open({ id: fa, path: fa, type: 'pptx' }), B = await EN.open({ id: fb, path: fb, type: 'pptx' });
  const ca = editor(() => A, d => { A = d; }, win), cb = editor(() => B, d => { B = d; }, win);
  ca.addSection(2); ca.commit(d => { d.slides[2].sec = 'Results'; });
  ca.setState({ cur: 1 }); ca.insertText(); ca.setState({ sel: null, sels: [] });
  ca.renderVals().thumbs[1].onClick(); key(ca, 'c');
  win.__writerSlideClip.slide.layout = 'Fancy'; // a layout deck B does not have
  cb.renderVals().thumbs[0].onClick(); key(cb, 'v');
  assert.deepEqual([B.slides.length, B.slides[1].from, B.slides[1].objs[0].from], [4, undefined, undefined], 'into another deck: made from the model');
  await EN.save(A); await EN.save(B);
  const A2 = await EN.open({ id: fa, path: fa, type: 'pptx' }), B2 = await EN.open({ id: fb, path: fb, type: 'pptx' });
  assert.deepEqual(A2.slides.map(s => s.sec), ['默认节', '', 'Results']);
  assert.deepEqual([B2.slides.length, B2.slides[1].layout, K.textOf(B2.slides[1].objs[0].html)], [4, 'Blank', '单击输入文本']);
  assert.equal(await EN.save(A2), 0, 'reopening writes nothing');
}));

test('engine: the shape gallery round-trips — a dashed gradient star with a shadow, a connector stuck to it, and a group come back as drawn', { skip: skip() }, () => engine(async () => {
  const file = join(dir, 'shapes.pptx');
  await EN.run(['create', file]); await EN.run(['add', file, '/', '--type', 'slide', '--prop', 'layout=Blank']);
  let doc = await EN.open({ id: file, path: file, type: 'pptx' });
  const c = editor(() => doc, d => { doc = d; });
  c.insertShape('star5'); const star = doc.slides[0].objs[0];
  c.patchSel({ fill: 'grad:#E3B25A,#B5563A,90', stroke: '#1D1D1F', sw: 4, dash: 'dash', shadow: true, rot: 15, lockAspect: true });
  c.insertShape('arrowLine'); const ln = doc.slides[0].objs[1];
  c.patchObj(ln.id, { end: { id: star.id, idx: 1 }, x: 100, y: 100, w: 200, h: 0 });
  c.insertShape('rect'); c.insertShape('ellipse');
  const [r, e] = doc.slides[0].objs.slice(2); c.setState({ sel: e.id, sels: [r.id, e.id] }); c.groupSels();
  await EN.save(doc);
  const back = await EN.open({ id: file, path: file, type: 'pptx' }), [s, l, g] = back.slides[0].objs;
  assert.deepEqual([s.t, s.shape, s.fill, s.dash, s.shadow, s.rot, s.lockAspect, s.stroke], ['shape', 'star5', 'grad:#E3B25A,#B5563A,90', 'dash', true, 15, true, '#1D1D1F']);
  assert.equal(s.sw, 4);
  assert.deepEqual([l.t, l.tail, l.end, [l.x + l.w, l.y + l.h]], ['line', 'triangle', { id: s.id, idx: 1 }, [s.x, s.y + s.h / 2]], 'the connector\'s end names the star and sits on its left side');
  assert.deepEqual([g.t, g.kids.map(k => k.shape), g.kids.map(k => k.path.startsWith(g.path + '/shape['))], ['group', ['rect', 'ellipse'], [true, true]]);
  assert.deepEqual([g.x, g.y, g.w, g.h], [Math.min(r.x, e.x), Math.min(r.y, e.y), Math.max(r.x + r.w, e.x + e.w) - Math.min(r.x, e.x), Math.max(r.y + r.h, e.y + e.h) - Math.min(r.y, e.y)]);
  assert.equal(await EN.save(back), 0, 'reopening writes nothing');
  // the group moves as one, then dissolves in the file
  doc = back; const c2 = editor(() => doc, d => { doc = d; });
  c2.setState({ sel: g.id, sels: [g.id] }); c2.patchSel(c2.moveKids(g, { x: g.x + 100, y: g.y + 50 }));
  await EN.save(doc);
  const moved = (await EN.open({ id: file, path: file, type: 'pptx' })).slides[0].objs[2];
  assert.deepEqual([moved.x, moved.y, moved.kids[0].x - g.kids[0].x], [g.x + 100, g.y + 50, 100]);
  c2.ungroup(); await EN.save(doc);
  const flat = await EN.open({ id: file, path: file, type: 'pptx' });
  assert.deepEqual(flat.slides[0].objs.map(o => o.t), ['shape', 'line', 'shape', 'shape']);
  assert.equal(flat.slides[0].objs[2].x, g.kids[0].x + 100, 'the members stay where the group showed them');
}));
