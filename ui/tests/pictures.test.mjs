import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync, mkdtempSync, rmSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import * as E from '../sheet-engine.js';
globalThis.window = globalThis; // pdf-kit is imported by the bridge and expects a browser global
const EN = await import('../engine.js');
const P = await import('../picture.js');
const K = await import('../office-io.js');

const plain = x => JSON.parse(JSON.stringify(x)); // the editors run in another vm realm: compare values, not prototypes
const response = json => ({ ok: true, json: async () => json });
/** Replaces fetch for one test: /run bodies go to onRun(argv), /json to the tree. */
async function withEngine(onRun, tree, fn) {
  const prev = globalThis.fetch;
  globalThis.fetch = async (url, opts) => {
    if (String(url).startsWith('/json?')) return response(tree);
    const r = await onRun(JSON.parse(opts.body).argv);
    return response({ code: 0, output: JSON.stringify(r || {}) });
  };
  try { return await fn(); } finally { globalThis.fetch = prev; }
}

test('a look keeps only what differs from the picture as it is; a diff sends neutral values for what went away', () => {
  assert.deepEqual(P.lookFrom({ src: 'x', crop: '0,0,0,0', rotation: '0', brightness: '20', contrast: '0', line: 'none', lineWidth: '2pt', geometry: 'rect', grayscale: 'true', flipH: 'false', transparency: '0' }),
    { brightness: '20', grayscale: 'true' });
  assert.deepEqual(P.lookFrom({ line: '000000', lineWidth: '1.5pt', crop: '10,0,0,0', geometry: 'ellipse' }), { crop: '10,0,0,0', line: '000000', lineWidth: '1.5pt', geometry: 'ellipse' }, 'a black border is a border');
  assert.deepEqual(EN.lookDiff({ brightness: '20', grayscale: 'true', line: 'C00000', lineWidth: '3pt', crop: '5,5,5,5' }, { brightness: '30', rotation: '90' }),
    { crop: '0,0,0,0', rotation: '90', brightness: '30', grayscale: 'false', line: 'none' });
  assert.deepEqual(EN.lookDiff({ geometry: 'roundRect' }, { geometry: 'roundRect' }), {});
  assert.deepEqual(EN.mergeLook({ brightness: '20', grayscale: 'true' }, { grayscale: null, shadow: 'true' }), { brightness: '20', shadow: 'true' });
});

test('crop to a ratio takes the largest centred part of the whole picture', () => {
  assert.equal(P.cropToRatio({}, 400, 300, '1:1'), '12.5,0,12.5,0');
  assert.equal(P.cropToRatio({}, 400, 300, '16:9'), '0,12.5,0,12.5');
  assert.equal(P.cropToRatio({ crop: '25,0,0,0' }, 300, 300, '4:3'), '0,0,0,0', 'the whole picture is 4:3 already');
  assert.deepEqual(P.cropOf({ crop: '10,20,30,5' }).map(v => +v.toFixed(3)), [0.1, 0.2, 0.3, 0.05]);
});

test('a look is drawn with CSS: the crop as the image\'s place and size, mirrors about the frame, colour as filters', () => {
  const v = P.pictureView({ crop: '25,0,0,20', flipH: 'true', transparency: '30', grayscale: 'true', brightness: '20', line: 'C00000', lineWidth: '2pt', shadow: 'true', geometry: 'roundRect' }, 300, 200, 1);
  assert.deepEqual([v.left, v.top, v.width, v.height], ['-33.333%', '0%', '133.333%', '125%']);
  assert.equal(v.flip, 'scale(-1,1)');
  assert.equal(v.origin, '62.5% 40%', 'the centre of the part shown');
  assert.equal(v.opacity, '0.7');
  assert.equal(v.filter, 'url(#wlum20_0) grayscale(1)');
  assert.equal(v.radius, '33.33px', 'a sixth of the shorter side, as Office\'s roundRect');
  assert.equal(v.shadow, '0 0 0 2px #C00000,2.12px 2.12px 4px 2px rgba(0,0,0,0.4)');
  assert.match(v.frame, /overflow:hidden;border-radius:33.33px/);
  const tri = P.pictureView({ geometry: 'triangle', flipV: 'true' }, 100, 100, 1);
  assert.equal(tri.clip, 'polygon(50% 100%,100% 0%,0% 0%)', 'a triangle turned upside down points down');
  assert.equal(P.pictureView({}, 100, 100, 1).filter, 'none');
});

test('the 图片 tab appears with a selected picture, takes over an open ribbon, and closes with the picture', () => {
  assert.deepEqual(P.tabAfterSelect({ tab: 'home', bubble: true }, 'picture', true), { tab: 'picture' });
  assert.equal(P.tabAfterSelect({ tab: 'home', bubble: false }, 'picture', true), null, 'a closed ribbon stays closed');
  assert.deepEqual(P.tabAfterSelect({ tab: 'picture', bubble: true }, 'picture', false), { tab: 'home', bubble: false });
  assert.equal(P.tabAfterSelect({ tab: 'insert', bubble: true }, 'picture', false), null);
});

/** A host that records what the tools ask of the editor. */
function host(look, answer) {
  const h = { shown: [], sent: [], committed: [], busy: [], toasts: [], look: () => Object.assign({}, look) };
  Object.assign(h, {
    show: l => h.shown.push(l), busyFn: null, toast: m => h.toasts.push(m),
    send: async props => { h.sent.push(props); if (answer instanceof Error) throw answer; return { path: '/p', props: Object.assign({}, look, props) }; },
    commit: (r, binary) => h.committed.push([r.props, binary]),
  });
  h.busy = []; const busy = op => h.busy.push(op);
  return Object.assign(h, { tools: P.pictureTools(Object.assign({}, h, { busy, frame: () => ({ cx: 0, cy: 0, w: 400, h: 300 }), src: () => 'u' })) });
}

test('a tool shows at once and sends one command; a slider sends once it rests', async () => {
  const h = host({ grayscale: 'true' });
  await h.tools.set({ shadow: 'true' });
  assert.deepEqual(plain(h.shown), [{ grayscale: 'true', shadow: 'true' }]);
  assert.deepEqual(h.sent, [{ shadow: 'true' }]);
  assert.deepEqual(h.committed, [[{ grayscale: 'true', shadow: 'true' }, false]]);
  h.tools.set({ brightness: '10' }, true); h.tools.set({ brightness: '25' }, true); h.tools.set({ contrast: '-5' }, true);
  assert.equal(h.shown.length, 4, 'every step of the drag is drawn');
  assert.equal(h.sent.length, 1, 'nothing is sent while the slider moves');
  await new Promise(r => setTimeout(r, 420));
  assert.deepEqual(h.sent[1], { brightness: '25', contrast: '-5' }, 'one command with where it came to rest');
  h.tools.ratio('1:1');
  await new Promise(r => setTimeout(r, 0));
  assert.deepEqual(h.sent[2], { crop: '12.5,0,12.5,0' });
});

test('a slow tool shows busy, a failed one puts the look back and says why', async () => {
  const ok = host({});
  await ok.tools.cutout();
  assert.deepEqual(ok.busy, ['cutout', null]);
  assert.deepEqual(ok.sent, [{ background: 'remove' }]);
  assert.equal(ok.committed[0][1], true, 'the image itself changed: the editor loads it again');
  await ok.tools.compress('web');
  assert.deepEqual(ok.sent[1], { compress: 'web' });
  const toasts = [], zip = P.pictureTools({ look: () => ({}), show() { }, busy() { }, commit() { }, toast: m => toasts.push(m), send: async () => ({ path: '/p', props: { bytes: '52000' }, before: 2400000 }) });
  await zip.compress('print');
  assert.deepEqual(toasts, ['已压缩：2.3 MB → 51 KB，节省 2.2 MB']);
  assert.equal(P.savedText(40000, 40000), '这张图片已经不大于显示所需，没有再压缩');
  const bad = host({ grayscale: 'true' }, Object.assign(new Error('抠图需要 macOS 14 以上的 Apple Vision'), { hint: 'build the helper' }));
  await bad.tools.cutout();
  assert.deepEqual(plain(bad.shown), [{ grayscale: 'true' }], 'the look as the model has it is drawn again');
  assert.deepEqual(bad.toasts, ['抠图需要 macOS 14 以上的 Apple Vision（build the helper）']);
  assert.deepEqual(bad.busy, ['cutout', null]);
});

test('the 图片 ribbon: sliders carry the look\'s values, a running tool says so', () => {
  const made = [], k = { B: (label, onClick, o) => ({ isBtn: true, label, o }), M: (id, label, items) => { made.push(id); return { isMenu: true, id, label, items }; }, I: (label, onClick, o) => ({ isItem: true, label, o }), C: (label, swatch) => ({ isColor: true, label, swatch }), SEP: { isSep: true } };
  const items = P.pictureRibbon(k, host({}).tools, { brightness: '20', transparency: '35', line: 'C00000', lineWidth: '3pt', geometry: 'roundRect' }, 'cutout');
  const ranges = items.filter(i => i.isRange).map(i => [i.label, i.value, i.min, i.max]);
  assert.deepEqual(ranges, [['亮度', 20, -100, 100], ['对比度', 0, -100, 100], ['透明度', 35, 0, 100]]);
  assert.equal(items[0].label, '抠图中…');
  assert.equal(items.find(i => i.label === '边框').swatch, '#C00000');
  assert.equal(items.find(i => i.label === '圆角').o.on, true);
  assert.deepEqual(made, ['pic-crop', 'pic-shape', 'pic-width', 'pic-compress']);
  assert.ok(items.find(i => i.id === 'pic-width').items.find(i => i.label === '3 磅').o.on);
});

test('setPicture runs one set per tool, after a save of the same file that is under way, reading the path when its turn comes', async () => {
  const order = [];
  let release;
  const slow = new Promise(r => { release = r; });
  const doc = { id: 'm', type: 'md', path: 'a.md', loaded: true, text: 'new', _orig: 'old' };
  const prev = globalThis.fetch;
  globalThis.fetch = async (url, opts) => {
    if (String(url).startsWith('/file?')) { order.push('save starts'); await slow; order.push('save ends'); return response({}); }
    order.push(JSON.parse(opts.body).argv.join(' '));
    return response({ code: 0, output: JSON.stringify({ kind: 'image', path: '/body/image[2]', props: { grayscale: 'true' } }) });
  };
  try {
    let path = '/body/image[1]';
    const saving = EN.save(doc);
    const picture = EN.setPicture('a.md', () => path, { grayscale: 'true' });
    await new Promise(r => setTimeout(r, 10));
    assert.deepEqual(order, ['save starts'], 'the picture waits for the save');
    path = '/body/image[2]'; // the save renumbered the pictures
    release();
    await saving;
    assert.deepEqual(await picture, { path: '/body/image[2]', props: { grayscale: 'true' } });
    assert.deepEqual(order, ['save starts', 'save ends', 'set a.md /body/image[2] --prop grayscale=true']);
    await assert.rejects(EN.setPicture('a.md', () => null, { grayscale: 'true' }), /还没有存进文件/);
    order.length = 0;
    const zipped = await EN.setPicture('a.md', '/body/image[1]', { compress: 'web' });
    assert.deepEqual(order, ['get a.md /body/image[1]', 'set a.md /body/image[1] --prop compress=web'], 'a compress reads the size it starts from');
    assert.equal(zipped.before, 0);
  } finally { globalThis.fetch = prev; }
});

test('Word pictures: the look rides on the img, a figure frames it, and a save sends only an undone look', async () => {
  const nodes = [{ kind: 'image', path: '/body/image[1]', props: { src: '/word/media/image1.png', width: '8cm', height: '6cm', grayscale: 'true', crop: '10,0,0,0', rotation: '90' } },
    { kind: 'image', path: '/body/image[2]', props: { src: '/word/media/image2.png', width: '4cm', height: '4cm' } }];
  const blocks = EN.blocksOf(nodes, 'a.docx');
  const html = EN.blocksToHtml(blocks);
  assert.match(html, /<figure data-pic="1" style="[^"]*width:302px;[^"]*aspect-ratio:302 \/ 227;transform:rotate\(90deg\)/);
  assert.match(html, /<img data-path="\/body\/image\[1\]" data-fw="302" data-fh="227" data-w-crop="10,0,0,0" data-w-rotation="90" data-w-grayscale="true"/);
  assert.match(html, /<img data-path="\/body\/image\[2\]" data-fw="151" data-fh="151" src="[^"]+" style="max-width:100%;display:block;margin:8px auto;width:151px">/, 'a picture without a look stays a plain img');
  // the file has the look the editor shows: nothing to save; the look undone in the editor: its way back
  const cmds = [];
  await EN.planDocxBlocks('a.docx', blocks, structuredClone(blocks), async argv => { cmds.push(argv); return {}; });
  assert.deepEqual(cmds, []);
  const undone = structuredClone(blocks);
  delete undone[0].props.grayscale;
  await EN.planDocxBlocks('a.docx', blocks, undone, async argv => { cmds.push(argv); return {}; });
  assert.deepEqual(cmds, [['set', 'a.docx', '/body/image[1]', '--prop', 'grayscale=false']]);
});

test('pictureSaved takes a command\'s answer into the snapshot and gives the model the change, not the result', async () => {
  const slide = { kind: 'slide', path: '/slide[1]', props: { id: '256', layout: 'Blank' }, children: [
    { kind: 'image', path: '/slide[1]/image[1]', props: { id: '4', x: '2cm', y: '2cm', w: '8cm', h: '6cm', rotation: '90', shadow: 'true' } }] };
  const tree = { kind: 'document', props: { width: '33.867cm', height: '19.05cm' }, children: [slide] };
  const doc = await withEngine(() => ({}), tree, () => EN.open({ id: 'p', path: 'deck.pptx', type: 'pptx' }));
  const o = doc.slides[0].objs[0];
  assert.deepEqual(plain([o.path, o.look, o.rot]), ['/slide[@id=256]/image[@id=4]', { shadow: 'true' }, 90]);
  o.x += 50; // moved in the editor, not saved yet
  const change = EN.pictureSaved(doc, o.path, { x: '2cm', y: '2.75cm', w: '8cm', h: '4.5cm', rotation: '90', shadow: 'true', crop: '0,12.5,0,12.5' });
  assert.deepEqual(plain(change.look), { crop: '0,12.5,0,12.5' });
  assert.deepEqual([change.dx, change.dy, change.dw, change.dh, change.drot], [0, 36, 0, -70, 0]);
  const snap = doc._orig.slides[0].objs[0];
  assert.deepEqual(plain([snap.look, snap.h]), [{ shadow: 'true', crop: '0,12.5,0,12.5' }, 213]);
  Object.assign(o, { look: EN.mergeLook(o.look, change.look), y: o.y + change.dy, h: o.h + change.dh });
  const cmds = [];
  await withEngine(argv => { cmds.push(argv); return {}; }, tree, () => EN.save(doc));
  assert.equal(cmds.length, 1);
  assert.deepEqual(cmds[0].slice(0, 4), ['set', 'deck.pptx', '/slide[@id=256]/image[@id=4]', '--prop']);
  assert.match(cmds[0][4], /^x=3\.04\dcm$/, 'only the move the editor had not saved');
  assert.equal(cmds[0].length, 5);
  // an undo takes the deck back: the picture's look and turn go back to the file
  o.look = { shadow: 'true' }; o.rot = 0;
  cmds.length = 0;
  await withEngine(argv => { cmds.push(argv); return {}; }, tree, () => EN.save(doc));
  assert.deepEqual(cmds[0].slice(3), ['--prop', 'crop=0,0,0,0', '--prop', 'rotation=0']);
});

test('Excel pictures come with their sheet, and a save sends an undone look back', async () => {
  const s = EN.sheetModel({ kind: 'sheet', path: '/sheet[1]', props: { name: 'S' }, children: [
    { kind: 'image', path: '/sheet[1]/image[1]', props: { id: '3', x: '2cm', y: '1cm', w: '4cm', h: '3cm', flipH: 'true', alt: 'logo' } }] }, 'book.xlsx');
  assert.deepEqual(plain(s.images), [{ id: '/sheet[1]/image[@id=3]', path: '/sheet[1]/image[@id=3]', src: '/binary?file=book.xlsx&path=%2Fsheet%5B1%5D%2Fimage%5B%40id%3D3%5D', x: 76, y: 38, w: 151, h: 113, alt: 'logo', look: { flipH: 'true' } }]);
  const before = structuredClone([s]), after = structuredClone([s]);
  after[0].images[0].look = {};
  const cmds = [];
  await EN.planXlsx('book.xlsx', before, after, async argv => { cmds.push(argv); return {}; });
  assert.deepEqual(cmds, [['set', 'book.xlsx', '/sheet[1]/image[@id=3]', '--prop', 'flipH=false']]);
});

// ---- the tab in the editors themselves ----
const script = name => readFileSync(new URL(`../${name}.dc.html`, import.meta.url), 'utf8');
function mount(name, props) {
  const html = script(name), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: { documentElement: { dataset: {} } }, React: { createRef: () => ({ current: null }) }, structuredClone,
    // a stand-in for ui/i18n.js: no dictionary loaded, so every call falls back to the Chinese (the '@@' context dropped), as in Chinese mode
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + `\nglobalThis.${name} = Component;`, ctx);
  const c = new ctx[name]();
  c.props = props;
  return { c, names: [...new Set([...html.split('<script type="text/x-dc"')[0].matchAll(/\{\{\s*([A-Za-z_$][\w$]*)\s*\}\}/g)].map(m => m[1]))].filter(n => n !== 'true' && n !== 'false') };
}
const menusRender = (c, render) => Object.keys(c.menus).every(id => { c.state.pop = { id, x: 0, y: 0 }; const n = render().popItems.length; c.state.pop = null; return n > 0; });

test('slide editor: a selected picture brings the 图片 tab; its tools and menus render; deselecting closes it', () => {
  let doc = { id: 'd', path: 'deck.pptx', type: 'pptx', theme: 'paper', ratio: '16:9', slides: [{ id: 's1', path: '/slide[@id=256]', objs: [
    K.txt({ id: 'e4', t: 'image', path: '/slide[@id=256]/image[@id=4]', src: '/binary?x', x: 100, y: 100, w: 400, h: 300, look: { grayscale: 'true' }, rot: 0 }),
    K.txt({ id: 'e5', t: 'text', html: '<p>t</p>' })] }] };
  const { c, names } = mount('SlideEditor', { get doc() { return doc; }, onChange(d) { doc = d; } });
  Object.assign(c, { K, P, EN });
  const tabs = () => c.renderVals().tabs.map(t => t.label);
  assert.ok(!tabs().includes('图片'));
  c.setState({ sel: 'e5' });
  assert.ok(tabs().includes('形状格式') && !tabs().includes('图片'), 'a text box keeps 形状格式');
  c.setState({ tab: 'home', bubble: true, sel: 'e4' }); c.componentDidUpdate({});
  assert.equal(c.state.tab, 'picture', 'the open ribbon turns to 图片');
  assert.ok(tabs().includes('图片') && !tabs().includes('形状格式'));
  const v = c.renderVals();
  assert.deepEqual(names.filter(n => !(n in v)), []);
  assert.ok(v.ribbon.some(i => i.isRange && i.label === '亮度') && v.ribbon.some(i => i.label === '排列'));
  assert.ok(menusRender(c, () => c.renderVals()), 'every menu of the tab has items');
  assert.match(v.objs.find(o => o.id === 'e4').picImage, /filter:grayscale\(1\)/);
  c.setState({ sel: null }); c.componentDidUpdate({});
  assert.deepEqual([c.state.tab, c.state.bubble], ['home', false]);
});

test('slide editor: the picture tools act on the picture on its own slide, though every slide numbers its shapes anew', async () => {
  // as the engine prints it: slide 1's text box and slide 2's picture both have cNvPr id 2
  const box = { x: '2cm', y: '2cm', w: '10cm', h: '6cm' };
  const tree = { kind: 'document', props: { width: '33.867cm', height: '19.05cm' }, children: [
    { kind: 'slide', path: '/slide[1]', props: { id: '256', layout: 'Blank' }, children: [{ kind: 'shape', path: '/slide[1]/shape[1]', props: { id: '2', text: 'Title', ...box } }] },
    { kind: 'slide', path: '/slide[2]', props: { id: '257', layout: 'Blank' }, children: [{ kind: 'image', path: '/slide[2]/image[1]', props: { id: '2', ...box } }] }] };
  let doc = await withEngine(() => ({}), tree, () => EN.open({ id: 'p', path: 'deck.pptx', type: 'pptx' }));
  const toasts = [], { c } = mount('SlideEditor', { get doc() { return doc; }, onChange(d) { doc = d; }, toast: m => toasts.push(m) });
  Object.assign(c, { K, P, EN });
  const ev = { stopPropagation() { }, preventDefault() { }, clientX: 0, clientY: 0, currentTarget: { closest: () => null } };
  c.renderVals().thumbs[1].onClick();
  c.renderVals().objs.find(o => o.isImg).onMD(ev); c.componentDidUpdate({});
  const sent = [], pic = Object.assign({ id: '2' }, box);
  await withEngine(argv => {
    sent.push(argv.slice(2).join(' '));
    if (argv[2] !== '/slide[@id=257]/image[@id=2]') throw new Error("shape has no property 'background' in pptx");
    const [k, v] = argv[4].split('='); if (k !== 'background') pic[k] = v;
    return { kind: 'image', props: Object.assign({}, pic) };
  }, tree, async () => { await c.picTools().cutout(); await c.picTools().set({ grayscale: 'true' }); });
  assert.deepEqual(sent, ['/slide[@id=257]/image[@id=2] --prop background=remove', '/slide[@id=257]/image[@id=2] --prop grayscale=true']);
  assert.deepEqual(toasts, []);
  assert.deepEqual(plain(doc.slides[1].objs[0].look), { grayscale: 'true' });
  assert.equal(doc.slides[0].objs[0].look, undefined, 'the text box on slide 1 is left alone');
  assert.equal(new URLSearchParams(doc.slides[1].objs[0].src.split('?')[1]).get('path'), '/slide[@id=257]/image[@id=2]', 'the picture loads by its id, which moving or deleting slides leaves alone');
  // slide 1 opened from its thumbnail's menu: the selection does not pass to the shape there with the same number
  c.renderVals().thumbs[0].onCtx(ev);
  assert.equal(c.obj, undefined);
});

test('sheet editor: a picture on the sheet shows and selects; the 图片 tab comes and goes with it', () => {
  let doc = { id: 'x', path: 'book.xlsx', type: 'xlsx', active: 0, sheets: [{ name: 'S', path: '/sheet[1]', cells: {}, images: [{ id: '3', path: '/sheet[1]/image[@id=3]', src: '/binary?y', x: 64, y: 20, w: 200, h: 150, alt: 'logo', look: { flipH: 'true' } }] }] };
  const { c, names } = mount('SheetEditor', { get doc() { return doc; }, onChange(d) { doc = d; } });
  Object.assign(c, { E, P, EN });
  let v = c.renderVals();
  assert.deepEqual(plain(v.pics.map(p => [p.id, p.w, p.h, p.line])), [['3', '200px', '150px', 'none']]);
  assert.ok(!v.tabs.some(t => t.label === '图片'));
  c.setState({ picSel: '3', tab: 'picture', bubble: true }); c.componentDidUpdate({});
  v = c.renderVals();
  assert.ok(v.tabs.some(t => t.label === '图片'));
  assert.deepEqual(names.filter(n => !(n in v)), []);
  assert.match(v.pics[0].image, /transform:scale\(-1,1\)/);
  assert.equal(v.pics[0].line, '2px solid var(--k59, #3F7D5C)');
  assert.ok(menusRender(c, () => c.renderVals()));
  c.setState({ picSel: null }); c.componentDidUpdate({});
  assert.deepEqual([c.state.tab, c.state.bubble], ['home', false]);
});

test('sheet editor: another sheet starts with nothing selected, and ids name one object in the workbook, though every sheet numbers its own', async () => {
  // as the engine prints it: each sheet's first picture and first chart are both id 2
  const sheet = n => ({ kind: 'sheet', path: `/sheet[${n}]`, props: { name: 'S' + n }, children: [
    { kind: 'image', path: `/sheet[${n}]/image[1]`, props: { id: '2', x: '2cm', y: '1cm', w: '4cm', h: '3cm' } },
    { kind: 'chart', path: `/sheet[${n}]/chart[1]`, props: { id: '2', type: 'column', title: 'T' + n, categories: 'A1:A1', series: '[{"values":"B1:B1"}]', x: '8cm', y: '1cm' } }] });
  const tree = { kind: 'document', children: [sheet(1), sheet(2)] };
  let doc = await withEngine(() => ({}), tree, () => EN.open({ id: 'x', path: 'book.xlsx', type: 'xlsx' }));
  const ids = doc.sheets.flatMap(s => [...s.images, ...s.charts].map(o => o.id));
  assert.equal(new Set(ids).size, 4, 'one id per object: ' + ids.join(' '));
  const { c } = mount('SheetEditor', { get doc() { return doc; }, onChange(d) { doc = d; } });
  Object.assign(c, { E, P, EN });
  const ev = { stopPropagation() { }, preventDefault() { } };
  const toSheet = i => { const pp = { doc }; c.renderVals().sheetTabs[i].onClick(); c.componentDidUpdate(pp); };
  toSheet(1);
  c.renderVals().pics[0].onSel(ev); c.renderVals().charts[0].onSel(ev); c.componentDidUpdate({ doc });
  assert.ok(c.pic && c.state.chartSel, 'the picture and the chart on S2 are selected');
  toSheet(0);
  const v = c.renderVals();
  assert.deepEqual([c.pic, c.state.picSel, c.state.chartSel], [null, null, null], 'the selection stayed behind on S2');
  assert.deepEqual([v.pics[0].line, v.charts[0].border], ['none', 'var(--k6, #E5E5EA)'], 'nothing on S1 shows as selected');
  // S1's own picture, selected: its tools act on it, not on S2's picture with the same number
  v.pics[0].onSel(ev); c.componentDidUpdate({ doc });
  const sent = [];
  await withEngine(argv => { sent.push(argv[2]); return { kind: 'image', props: { id: '2', x: '2cm', y: '1cm', w: '4cm', h: '3cm', grayscale: 'true' } }; }, tree, () => c.picTools().set({ grayscale: 'true' }));
  assert.deepEqual(sent, ['/sheet[1]/image[@id=2]']);
  assert.deepEqual(plain(doc.sheets.map(s => s.images[0].look)), [{ grayscale: 'true' }, {}]);
});

// ---- the real engine (skipped until the CLI is built): the bridge's commands must come back when the file is opened again ----
const cli = fileURLToPath(new URL('../../src/Writer.Cli/bin/Debug/net10.0/writer.dll', import.meta.url));
/** Runs fn(dir) against `writer serve` on a temporary folder, the bridge's relative URLs sent to it. */
async function withRealEngine(fn) {
  const dir = mkdtempSync(join(tmpdir(), 'writer-pictures-')), prev = globalThis.fetch;
  const server = spawn('dotnet', [cli, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  try {
    const url = await new Promise((resolve, reject) => {
      let err = '';
      server.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1].replace(/\/$/, '')); });
      server.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
    });
    globalThis.fetch = (u, o) => prev(String(u).startsWith('/') ? url + u : u, o);
    return await fn(dir);
  } finally { globalThis.fetch = prev; server.kill(); rmSync(dir, { recursive: true, force: true }); }
}
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

test('engine: a copied sheet\'s pictures and charts are new objects, saved with the copy; the copy\'s picture tools leave the original alone', { skip: !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx' }, async () => {
  await withRealEngine(async dir => {
    const file = join(dir, 'book.xlsx');
    for (const argv of [['create', file], ['set', file, '/sheet[1]', '--prop', 'name=S'], ['set', file, '/sheet[1]/range[A1:B3]', '--prop', 'values=[["k","v"],["a",3],["b",5]]'],
      ['add', file, '/sheet[1]', '--type', 'image', '--prop', 'src=' + PNG, '--prop', 'x=1cm', '--prop', 'y=2cm', '--prop', 'w=2cm', '--prop', 'h=2cm', '--prop', 'alt=logo'],
      ['set', file, '/sheet[1]/image[@id=2]', '--prop', 'crop=10,0,0,0', '--prop', 'x=1cm', '--prop', 'y=2cm', '--prop', 'w=2cm', '--prop', 'h=2cm'],
      ['add', file, '/sheet[1]', '--type', 'chart', '--prop', 'type=column', '--prop', 'title=T', '--prop', 'categories=A2:A3', '--prop', 'series=[{"values":"B2:B3"}]', '--prop', 'x=5cm', '--prop', 'y=2cm']]) await EN.run(argv);
    let doc = await EN.open({ id: 'b', path: file, type: 'xlsx' });
    const { c } = mount('SheetEditor', { get doc() { return doc; }, onChange(d) { doc = d; } });
    Object.assign(c, { E, P, EN });
    const copy = () => { c.state.pop = { id: 'tabctx', i: 0 }; c.renderVals().popItems.find(i => i.label === '复制工作表').onClick(); c.state.pop = null; };
    copy(); copy(); // twice: the second copy needs a name of its own, or the file refuses the sheet
    assert.ok(await EN.save(doc) > 0);
    const look = s => [s.name, s.images.map(im => [im.x, im.y, im.w, im.h, im.alt, plain(im.look)]), s.charts.map(ch => [ch.title, ch.cat, ch.x, ch.y])];
    const again = await EN.open({ id: 'b2', path: file, type: 'xlsx' }), [orig, ...copies] = again.sheets.map(look);
    assert.deepEqual(copies.map(s => s[0]), ['S (3)', 'S (2)']);
    for (const s of copies) assert.deepEqual(s.slice(1), orig.slice(1), 'the copy has the original\'s picture and chart');
    // S (2)'s picture, selected on its sheet: a picture tool changes it, and only it
    const pp = { doc }; c.renderVals().sheetTabs[2].onClick(); c.componentDidUpdate(pp);
    c.renderVals().pics[0].onSel({ stopPropagation() { } }); c.componentDidUpdate({ doc });
    const toasts = []; c.props.toast = m => toasts.push(m);
    await c.picTools().set({ grayscale: 'true' });
    assert.deepEqual(toasts, []);
    const last = await EN.open({ id: 'b3', path: file, type: 'xlsx' });
    assert.deepEqual(plain(last.sheets.map(s => [s.name, s.images[0].look])), [['S', { crop: '10,0,0,0' }], ['S (3)', { crop: '10,0,0,0' }], ['S (2)', { crop: '10,0,0,0', grayscale: 'true' }]]);
  });
});

test('engine: sheets are saved in the editor\'s order: a copy beside its original, a new sheet at the end, 左移 and 右移', { skip: !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx' }, async () => {
  await withRealEngine(async dir => {
    const file = join(dir, 'order.xlsx');
    for (const argv of [['create', file], ['set', file, '/sheet[1]', '--prop', 'name=A'], ['add', file, '/', '--type', 'sheet', '--prop', 'name=B'], ['add', file, '/', '--type', 'sheet', '--prop', 'name=C'],
      ['set', file, '/sheet[1]/cell[A1]', '--prop', 'value=a'], ['set', file, '/sheet[2]/cell[A1]', '--prop', 'value=b'], ['set', file, '/sheet[3]/cell[A1]', '--prop', 'value=c']]) await EN.run(argv);
    let doc = await EN.open({ id: 'o', path: file, type: 'xlsx' });
    const { c } = mount('SheetEditor', { get doc() { return doc; }, onChange(d) { doc = d; } });
    Object.assign(c, { E, P, EN });
    const tab = (i, label) => { c.state.pop = { id: 'tabctx', i }; c.renderVals().popItems.find(x => x.label === label).onClick(); c.state.pop = null; };
    const sheets = d => d.sheets.map(s => s.name + ':' + ((s.cells.A1 || {}).v || ''));
    const reopened = async () => sheets(await EN.open({ id: 'o2', path: file, type: 'xlsx' }));
    tab(0, '复制工作表'); c.addSheet(); tab(1, '右移'); tab(4, '左移'); tab(0, '右移');
    assert.deepEqual(sheets(doc), ['B:b', 'A:a', 'A (2):a', 'Sheet5:', 'C:c']);
    await EN.save(doc);
    assert.deepEqual(await reopened(), sheets(doc));
    // again on the model the save updated: its sheets are still addressed right after the file's order changed
    tab(4, '左移'); tab(0, '右移');
    assert.deepEqual(sheets(doc), ['A:a', 'B:b', 'A (2):a', 'C:c', 'Sheet5:']);
    await EN.save(doc);
    assert.deepEqual(await reopened(), sheets(doc));
  });
});
