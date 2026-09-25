// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { DEF, load, applyTo, instructions, chatOptions, pairQuote, launch, newFileArgs, loadLayout, saveLayout } from '../prefs.js';

/** A stand-in for localStorage that remembers what was last set, so save then load round-trips. */
function fakeStorage(initial) {
  let v = initial;
  return { getItem: () => v, setItem: (k, val) => { v = val; } };
}

/** A stand-in for <html>: attributes and inline custom properties. */
function fakeEl() {
  const a = {}, st = {};
  return { a, st, setAttribute: (k, v) => { a[k] = v; }, removeAttribute: k => { delete a[k]; }, style: { setProperty: (k, v) => { st[k] = v; }, removeProperty: k => { delete st[k]; } } };
}

test('applyTo: the defaults add nothing but the theme; every setting maps to its attribute', () => {
  const el = fakeEl(), none = () => false;
  applyTo(el, DEF, none);
  assert.deepEqual(el.a, { 'data-theme': 'light' });
  assert.deepEqual(el.st, {});
  applyTo(el, { ...DEF, theme: 'auto', density: 'compact', motion: true, ai: false, darkPages: true, accent: '#2F5D8A', glass: 100 }, q => q.includes('dark'));
  assert.deepEqual(el.a, { 'data-theme': 'dark', 'data-density': 'compact', 'data-motion': 'reduce', 'data-ai': 'off', 'data-dark-pages': '1', 'data-accent': '1', 'data-glass': '1' });
  assert.deepEqual(el.st, { '--accent': '#2F5D8A', '--gt': '1.818' });
  applyTo(el, { ...DEF, glass: 0 }, q => q.includes('reduced-motion'));
  assert.deepEqual(el.a, { 'data-theme': 'light', 'data-motion': 'reduce', 'data-glass': '1' }, 'the system asking for less motion counts too');
  assert.deepEqual(el.st, { '--gt': '0' });
  applyTo(el, { ...DEF, accent: 'red; x' }, none);
  assert.deepEqual(el.st, {}, 'only #rrggbb accents');
});

test('DEF has the same keys and defaults as the Settings window', () => {
  const html = readFileSync(new URL('../MacSettings.dc.html', import.meta.url), 'utf8');
  const lit = html.match(/const DEF = (\{[^;]*\});/)[1];
  assert.deepEqual(DEF, new Function('return ' + lit)());
});

test('load merges what is stored over the defaults and survives junk', () => {
  const store = v => ({ getItem: () => v });
  assert.equal(load(store('{"theme":"dark","paper":"B5"}')).paper, 'B5');
  assert.equal(load(store('{"theme":"dark"}')).newType, 'docx');
  assert.deepEqual(load(store('not json')), DEF);
  assert.deepEqual(load(null), DEF);
});

test('loadLayout: the sidebar and the assistant panel default closed, same as the shell\'s own defaults; junk and no storage fall back too', () => {
  assert.deepEqual(loadLayout(fakeStorage(null)), { showThumbs: false, showAI: false });
  assert.deepEqual(loadLayout(fakeStorage('{"thumbs":false,"ai":true}')), { showThumbs: false, showAI: true });
  assert.deepEqual(loadLayout(fakeStorage('not json')), { showThumbs: false, showAI: false });
  assert.deepEqual(loadLayout(null), { showThumbs: false, showAI: false });
});

test('saveLayout merges into \'writer-mac\' instead of replacing it: mac.dc.html/win.dc.html keep the collapsed toolbar in the same key', () => {
  const s = fakeStorage('{"collapsed":true}');
  saveLayout(false, true, s);
  assert.deepEqual(JSON.parse(s.getItem()), { collapsed: true, thumbs: false, ai: true });
  saveLayout(true, false, s);
  assert.deepEqual(JSON.parse(s.getItem()), { collapsed: true, thumbs: true, ai: false });
  assert.deepEqual(loadLayout(s), { showThumbs: true, showAI: false });
});

test('instructions: reply style plus the user\'s own text; chatOptions adds the selection only without whole-document context', () => {
  assert.equal(instructions(DEF), '');
  assert.equal(instructions({ ...DEF, tone: 'brief', instr: '  用正式书面语。 ' }), 'Keep every reply to one or two short sentences.\n用正式书面语。');
  assert.match(instructions({ ...DEF, tone: 'detail' }), /more detail/);
  assert.deepEqual(chatOptions(DEF, 'picked'), {});
  assert.deepEqual(chatOptions({ ...DEF, ctx: false, instr: 'x' }, 'picked'), { instructions: 'x', selection: 'picked' });
  assert.deepEqual(chatOptions({ ...DEF, ctx: false }, null), { selection: '' });
});

test('pairQuote pairs CJK punctuation, skips over a closing mark, leaves words alone', () => {
  assert.deepEqual(pairQuote('「', ''), { close: '」' });
  assert.deepEqual(pairQuote('《', ' '), { close: '》' });
  assert.deepEqual(pairQuote('“', '，'), { close: '”' });
  assert.deepEqual(pairQuote('你好「', '」'), { close: '」' }, 'an IME commit ending with an opening mark');
  assert.equal(pairQuote('「', '文'), null, 'not before a word');
  assert.deepEqual(pairQuote('」', '」'), { skip: true });
  assert.equal(pairQuote('」', ''), null);
  assert.equal(pairQuote('a', ''), null);
  assert.equal(pairQuote('', ''), null);
});

test('launch: restore keeps the open set and its order, startup picks what opens', () => {
  const files = ['new.md', 'a.docx', 'b.xlsx', 'c.pptx'].map(path => ({ path }));
  const session = { order: ['c.pptx', 'a.docx', 'b.xlsx'], closed: ['b.xlsx'], last: 'a.docx' };
  const paths = r => r.docs.map(d => d.path);
  const r = launch(files, DEF, session);
  assert.deepEqual(paths(r), ['new.md', 'c.pptx', 'a.docx'], 'closed stays closed, known keep their order, new files first');
  assert.equal(r.open.path, 'a.docx');
  assert.deepEqual(paths(launch(files, { ...DEF, restore: false }, session)), ['new.md', 'a.docx', 'b.xlsx', 'c.pptx']);
  assert.equal(launch(files, { ...DEF, restore: false }, session).open.path, 'new.md', 'without restore the newest file opens');
  assert.deepEqual(launch(files, { ...DEF, startup: 'home' }, session), { docs: r.docs, open: null, blank: null });
  assert.equal(launch(files, { ...DEF, startup: 'blank', newType: 'xlsx' }, session).blank, 'xlsx');
  assert.equal(launch([], DEF, {}).open, null);
  assert.equal(launch(files, DEF, { last: 'gone.docx' }).open.path, 'new.md');
});

test('newFileArgs: paper size and tracking for new Word documents only', () => {
  assert.deepEqual(newFileArgs('a.docx', 'docx', DEF, {}), []);
  assert.deepEqual(newFileArgs('a.docx', 'docx', { ...DEF, paper: 'Letter' }, {}), [['set', 'a.docx', '/', '--prop', 'page=Letter']]);
  assert.deepEqual(newFileArgs('a.docx', 'docx', { ...DEF, track: true }, {}), [], 'no track property in this engine: skipped');
  assert.deepEqual(newFileArgs('a.docx', 'docx', { ...DEF, paper: 'B5', track: true }, { track: true }), [['set', 'a.docx', '/', '--prop', 'page=B5', '--prop', 'track=true']]);
  assert.deepEqual(newFileArgs('a.xlsx', 'xlsx', { ...DEF, paper: 'A5' }, {}), []);
});

test('every window page loads theme-page.css, so 文稿页面也用深色 (and glass transparency) works in the Mac and Windows apps too', () => {
  for (const page of ['index.dc.html', 'mac.dc.html', 'win.dc.html'])
    assert.match(readFileSync(new URL('../' + page, import.meta.url), 'utf8'), /<link rel="stylesheet" href="\.\/theme-page\.css">/, page);
});
