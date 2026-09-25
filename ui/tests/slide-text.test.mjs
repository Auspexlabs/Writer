// node --test ui/tests/ — text in a box: bullet levels (Tab / ⇧Tab as data-lvl), paragraph and character spacing, columns,
// direction, autofit and the WordArt effects are drawn by textStyle, kept in the model and written to the file as the shape's
// box-wide settings and each paragraph's list and level.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
// a DOMParser for the simple markup a text box holds (ul / ol / li / p / br, attributes, text): what the bridge's list pass and run
// comparison read of an element
globalThis.DOMParser ??= class {
  parseFromString(html) {
    const mk = tag => ({ nodeType: 1, tagName: tag, attrs: {}, childNodes: [], style: {}, get children() { return this.childNodes.filter(c => c.nodeType === 1); }, getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; } });
    const body = mk('BODY'), stack = [body], re = /<\/?([a-z0-9]+)((?:\s+[\w-]+="[^"]*")*)\s*\/?>|([^<]+)/gi; let m;
    while ((m = re.exec(html))) {
      const top = stack[stack.length - 1];
      if (m[3]) { top.childNodes.push({ nodeType: 3, nodeValue: m[3].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&') }); continue; }
      const tag = m[1].toUpperCase();
      if (tag === 'BODY' || tag === 'HTML') continue; // the bridge wraps the markup in a body of its own
      if (m[0][1] === '/') { if (top.tagName === tag) stack.pop(); continue; }
      const el = mk(tag); for (const a of m[2].matchAll(/([\w-]+)="([^"]*)"/g)) el.attrs[a[1]] = a[2];
      top.childNodes.push(el); if (tag !== 'BR' && !m[0].endsWith('/>')) stack.push(el);
    }
    return { body };
  }
};
const K = await import('../office-io.js'), EN = await import('../engine.js');
const th = K.THEMES.paper;

test('textStyle: spacing, columns, direction and the WordArt effects as CSS; autofit scales the font', () => {
  const o = K.txt({ sb: 10, sa: 5, cs: 2, cols: 2, vert: 'eaVert', tOutline: 'fg', tShadow: true, tGrad: 'acc,sub,90', fs: 40 });
  const tx = K.textStyle(o, 40, th);
  assert.match(tx, /--sb:10px;--sa:5px;letter-spacing:2px;column-count:2;column-gap:48px;writing-mode:vertical-rl;text-orientation:mixed;/);
  assert.match(tx, /-webkit-text-stroke:1\.0px #1D1D1F;paint-order:stroke fill;text-shadow:2\.5px 2\.5px 5px rgba\(0,0,0,0\.4\);background:linear-gradient\(180deg,#1D1D1F,#6E6E73\);-webkit-background-clip:text/);
  assert.equal(K.textStyle(K.txt({}), 40, th), '', 'a plain box adds nothing');
  const v = K.objView(K.txt({ fs: 40, autofit: 'shrink', fit: 0.75 }), th);
  assert.equal(v.fs, '30px');
  assert.equal(K.objView(K.txt({ fs: 40, autofit: 'none', fit: 0.75 }), th).fs, '40px', 'the scale counts only while autofit is on');
  assert.match(K.LVL_CSS, /\[data-lvl="2"\]\{margin-left:3\.2em\}/);
});

test('indent: Tab and ⇧Tab move the paragraph at the caret between levels 0 and 4; with the box only selected, every paragraph', () => {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const mkEl = (tag, attrs = {}) => ({ tagName: tag, nodeType: 1, attrs, getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; }, setAttribute(k, v) { this.attrs[k] = v; }, removeAttribute(k) { delete this.attrs[k]; }, parentNode: null });
  const box = mkEl('DIV'), li = mkEl('LI', { 'data-lvl': '1' }); li.parentNode = box; const textNode = { nodeType: 3, parentNode: li };
  const ctx = { window: { innerWidth: 1360, innerHeight: 860, getSelection: () => ({ rangeCount: 1, anchorNode: textNode }) }, React: { createRef: () => ({ current: null }) }, structuredClone,
    document: { createElement: () => ({ set innerHTML(h) { this._h = h; }, get innerHTML() { return this._h.replace(/<li>/g, '<li data-lvl="1">').replace(/<li data-lvl="2">/g, '<li data-lvl="3">'); }, querySelectorAll: () => [] }) },
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const a = K.txt({ id: 'a', html: '<ul><li>one</li><li data-lvl="2">two</li></ul>' });
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [{ id: 's1', layout: 'blank', decor: [], objs: [a], notes: '', trans: 'none', hidden: false, bg: null }] };
  const c = new ctx.SlideEditor(); c.props = { get doc() { return doc; }, onChange: d => { doc = d; }, toast() { } }; c.K = K;
  c.stageRef.current = { querySelector: sel => sel === '[data-edit="a"]' ? box : null };
  c.setState({ sel: 'a', sels: ['a'], editing: 'a' });
  const key = (k, shift) => c.renderVals().onRootKey({ key: k, shiftKey: shift, target: { isContentEditable: true, tagName: 'DIV' }, preventDefault() { } });
  key('Tab'); assert.equal(li.attrs['data-lvl'], '2');
  key('Tab'); key('Tab'); key('Tab'); assert.equal(li.attrs['data-lvl'], '4', 'four is the deepest');
  key('Tab', true); key('Tab', true); key('Tab', true); key('Tab', true); assert.equal(li.attrs['data-lvl'], undefined, 'back at the top level the attribute goes');
  c.setState({ editing: null });
  c.renderVals().ribbon.find(i => i.title === '增加缩进 Tab').onClick();
  assert.equal(doc.slides[0].objs[0].html, '<ul><li data-lvl="1">one</li><li data-lvl="3">two</li></ul>', 'every paragraph moves one level (the stand-in document does the DOM part)');
  // the 开始 tab's menus set the box-wide values in slide units (16:9: 1 pt = 5/3 units)
  c.menus.para.find(i => i.label === '6 磅').onClick(); assert.equal(doc.slides[0].objs[0].sb, 10);
  c.menus.cs.find(i => i.label === '宽松').onClick(); assert.equal(doc.slides[0].objs[0].cs, 2.5);
  c.menus.cols[1].onClick(); c.menus.fit[1].onClick(); c.menus.vert[2].onClick(); c.menus.art.find(i => i.label === '渐变 + 阴影').onClick();
  const o = doc.slides[0].objs[0];
  assert.deepEqual([o.cols, o.autofit, o.vert, o.tGrad, o.tShadow, o.tOutline], [2, 'shrink', 'eaVert', 'acc,fg,90', true, '']);
});

test('the bridge: box settings and list levels go to the file and come back', async () => {
  const slide = { kind: 'slide', path: '/slide[@id=256]', props: { id: '256', layout: 'Blank' }, children: [
    { kind: 'shape', path: '/slide[@id=256]/shape[@id=2]', props: { id: '2', text: 'a\nb', x: '2cm', y: '2cm', w: '8cm', h: '3cm', lineSpacing: '1.5', spaceBefore: '6pt', charSpacing: '1.5', columns: '2', autofit: 'shrink:80', direction: 'eaVert', textOutline: '1F2937', textShadow: 'true', textGradient: '4472C4,ED7D31,45' }, children: [
      { kind: 'paragraph', path: '/slide[@id=256]/shape[@id=2]/paragraph[1]', props: { text: 'a', html: 'a', list: 'number' } },
      { kind: 'paragraph', path: '/slide[@id=256]/shape[@id=2]/paragraph[2]', props: { text: 'b', html: 'b', list: 'number', level: '1' } }] }] };
  const tree = { kind: 'document', props: { width: '33.867cm', height: '19.05cm' }, children: [slide] };
  const prev = globalThis.fetch, commands = [];
  globalThis.fetch = async (url, opts) => { if (String(url).startsWith('/json?')) return { ok: true, json: async () => tree }; const { argv } = JSON.parse(opts.body); commands.push(argv); return { ok: true, json: async () => ({ code: 0, output: '{}' }) }; };
  try {
    const doc = await EN.open({ id: 'p1', path: 'deck.pptx', type: 'pptx' }), o = doc.slides[0].objs[0];
    assert.deepEqual([o.lh, o.sb, o.cs, o.cols, o.autofit, o.fit, o.vert, o.tOutline, o.tShadow, o.tGrad], [1.5, 10, 2.5, 2, 'shrink', 0.8, 'eaVert', '#1F2937', true, '#4472C4,#ED7D31,45']);
    assert.equal(o.html, '<ol><li>a</li><li data-lvl="1">b</li></ol>', 'a numbered list with its levels');
    assert.equal(await EN.save(doc), 0, 'nothing to write yet');
    Object.assign(o, { sb: 0, cs: 0, cols: 1, autofit: 'none', vert: 'horz', tGrad: '', tShadow: false, lh: 1.35, html: '<ul><li>a</li><li data-lvl="2">b</li></ul>' });
    await EN.save(doc);
    const set = commands.find(c => c[0] === 'set' && c[2] === '/slide[@id=256]/shape[@id=2]');
    const props = Object.fromEntries(set.slice(3).filter(x => x !== '--prop').map(x => x.split(/=(.*)/s).slice(0, 2)));
    assert.deepEqual(props, { lineSpacing: '1.35', spaceBefore: '0pt', charSpacing: '0', columns: '1', direction: 'horz', autofit: 'none', textShadow: 'false', textGradient: '' }, 'the runs are the same, so the html is not rewritten: the list pass carries the markers');
    assert.deepEqual(commands.filter(c => c[2].includes('/paragraph[')).map(c => c.slice(2)), [['/slide[@id=256]/shape[@id=2]/paragraph[1]', '--prop', 'list=bullet', '--prop', 'level=0'], ['/slide[@id=256]/shape[@id=2]/paragraph[2]', '--prop', 'list=bullet', '--prop', 'level=2']]);
  } finally { globalThis.fetch = prev; }
});
