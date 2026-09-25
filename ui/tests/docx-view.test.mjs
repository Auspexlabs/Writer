// node --test ui/tests/ — 视图 in the Word editor: 页面移动 › 并排 lays the pages side by side (each a column of the text), and 缩放
// offers Word's presets with 页宽, 整页 and 多页.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');
const src = readFileSync(new URL('../WordEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = { React: { createRef: () => ({ current: null }) }, setTimeout, clearTimeout, getComputedStyle: () => ({ marginBottom: '0px' }),
  $t: (s, v) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => v && k in v ? v[k] : m), $lang: () => 'zh',
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() {} } };
vm.runInNewContext(src + '\nglobalThis.WordEditor = Component;', ctx);
const plain = x => JSON.parse(JSON.stringify(x));

test('side by side: the text flows into page-wide columns, one sheet per column left to right, without headers or thumbnails', () => {
  const c = new ctx.WordEditor(); c.EN = EN; c.props = { doc: { id: 'd', html: '', page: { hf: true } }, onChange() {}, showThumbs: true }; c.pgCss = { textContent: '' };
  const ed = { innerText: 'text', querySelectorAll: () => [], children: [], scrollWidth: 3 * 602 + 2 * 212 }; // A4, normal margins: 602px of text, 212px between
  c.edRef.current = ed; c.state.view = 'side'; c.refreshInfo();
  assert.equal(c.state.info.pages, 3);
  const v = c.renderVals();
  assert.deepEqual(plain(v.sheets.map(s => [s.left, s.w, s.top])), [['0px', '794px', '0px'], ['814px', '794px', '0px'], ['1628px', '794px', '0px']]);
  assert.deepEqual([v.pageW, v.sideAttr, v.edColW, v.edGap, v.edFill, v.hfZones.length, v.pageThumbs.length, v.showRuler], ['2422px', '1', '602px', '212px', 'auto', 0, 0, false]);
  assert.equal(c.pgCss.textContent, '', 'no page gaps pushed in: the columns break the pages');
  c.state.tab = 'view';
  const menus = c.renderVals(); assert.ok(menus.ribbon.some(x => x.label === '页面移动') && menus.ribbon.some(x => x.label === '缩放：100%'));
  c.state.pop = { id: 'zoom' };
  assert.deepEqual(plain(c.renderVals().popItems.map(i => i.label)), ['50%', '75%', '100%', '125%', '150%', '200%', '页宽', '整页', '多页']);
});
