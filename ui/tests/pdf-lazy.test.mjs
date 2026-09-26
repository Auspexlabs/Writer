// node --test ui/tests/ — the PDF editor draws only the pages near the view and lets go of the others (a page's canvas holds its
// area × zoom² × pixel ratio²), draws the thumbnail of a page away from the view small on its own, and its sidebar lists the pages.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const src = readFileSync(new URL('../PdfEditor.dc.html', import.meta.url), 'utf8');
const script = src.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout() { }, requestAnimationFrame: () => 0, window: { devicePixelRatio: 1 },
  document: { querySelector: () => null, querySelectorAll: () => [] }, $t: s => s, React: { createRef: () => ({ current: null }) },
  DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } };
vm.runInNewContext(script + '\nglobalThis.C = Component;', ctx);

/** n pages 1000px tall, 20px apart, in a view 800px tall; `drawn` records each page drawn and at what scale. */
function editor(n) {
  const c = new ctx.C(), drawn = [], pages = Array.from({ length: n }, (_, i) => ({ id: 'p' + i, src: i }));
  c.props = { doc: { id: 'd', pages } }; Object.assign(c.state, { scale: 1, fields: {} });
  c.K = { keyOf: () => 'k', renderPage: async (id, pg, canvas, scale) => { drawn.push([pg.id, Math.round(scale * 1000) / 1000]); canvas.width = 100; canvas.height = 140; }, textLayer: async () => { }, markHits() { } };
  const els = pages.map((p, i) => ({ offsetTop: i * 1020, offsetHeight: 1000, getAttribute: () => p.id }));
  c.scRef.current = { clientHeight: 800, scrollTop: 0, querySelectorAll: () => els, querySelector: () => null };
  const canvas = () => ({ width: 0, height: 0, getContext: () => ({ drawImage() { } }) });
  pages.forEach(p => { c.cv[p.id] = canvas(); c.tl[p.id] = {}; });
  return { c, drawn, canvas, pages };
}
const plain = x => JSON.parse(JSON.stringify(x));

test('only the pages near the view are drawn, the nearest first; scrolling away lets go of them', async () => {
  const { c, drawn } = editor(30);
  await c.renderAll();
  assert.deepEqual(plain(drawn), [['p0', 1], ['p1', 1]], 'the top of a 30-page file: two pages, not thirty');
  drawn.length = 0; c.scRef.current.scrollTop = 10200; await c.renderAll();
  assert.deepEqual(plain(drawn).map(d => d[0]), ['p10', 'p9', 'p11', 'p8'], 'page 11 in view: it first, then the ones around it');
  assert.deepEqual([c.cv.p0.width, c.cv.p1.width, c.cv.p10.width], [0, 0, 100], 'the pages scrolled away from hold no pixels');
  drawn.length = 0; await c.renderAll(); assert.deepEqual(plain(drawn), [], 'nothing drawn twice');
});

test('with the thumbnails open, a page away from the view gets its thumbnail drawn small; a drawn page lends it its own', async () => {
  const { c, drawn, canvas, pages } = editor(6);
  pages.forEach(p => { c.tcv[p.id] = canvas(); });
  await c.renderAll();
  assert.deepEqual(plain(drawn), [['p0', 1], ['p1', 1], ['p2', 0.202], ['p3', 0.202], ['p4', 0.202], ['p5', 0.202]], 'the two in view at full size (their thumbnails copied from them), the others at thumbnail size');
  drawn.length = 0; c.state.scale = 1.5; await c.renderAll();
  assert.deepEqual(plain(drawn), [['p0', 1.5], ['p1', 1.5]], 'zooming draws the pages again, not the thumbnails');
});

test('the sidebar lists the pages unless its outline tab is on (it showed nothing for a file without an outline)', () => {
  assert.match(src, /<sc-if value="\{\{ thumbsOn \}\}"[^>]*>\s*<div[^>]*>\s*<sc-for list="\{\{ pages \}\}"/);
  assert.match(script, /thumbsOn: !outlineOn/);
});
