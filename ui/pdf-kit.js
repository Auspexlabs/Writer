// PDF open / render / text layer / search / annotations / saving: an incremental update appended to the file, or a rewrite
// when the pages themselves changed.
const V = '4.4.168', CDN = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${V}/`;
let lib = null, pdflib = null;
export async function pdfjs() { if (!lib) { lib = await import(CDN + 'build/pdf.min.mjs'); lib.GlobalWorkerOptions.workerSrc = CDN + 'build/pdf.worker.min.mjs'; } return lib; }
async function pl() { if (!pdflib) pdflib = await import('https://esm.sh/pdf-lib@1.17.1'); return pdflib; }
/** Tests and scripts hand in the libraries themselves (node has no CDN). */
export function useLibs(js, lb) { if (js) lib = js; if (lb) pdflib = lb; }

// ----- bytes -----
// store: key → bytes. A document's pages carry `from`, the key of the bytes they come from: `<id>@<generation>` for the
// file as opened and after every save, another id for pages merged in from a second PDF. Old generations stay around
// so that undoing past a save still finds its pages; gens[id] is the generation the file on disk has.
export const store = window.__pdfStore || (window.__pdfStore = {});
const cache = window.__pdfDocs || (window.__pdfDocs = {});
const gens = window.__pdfGens || (window.__pdfGens = {});
export const keyOf = (id, pg) => (pg && pg.from) || id;
export const genKey = (id, gen) => id + '@' + gen;
export const gen = id => gens[id] || 0;
export function setGen(id, n, bytes) { gens[id] = n; store[genKey(id, n)] = store[id] = bytes; reset(genKey(id, n)); }
export function reset(key) { const p = cache[key]; delete cache[key]; if (p) p.then(d => d.destroy(), () => { }); }
/** Drops generations up to `upto` (memory); a page of a dropped generation falls back to the current bytes. */
export function forget(id, upto) { for (let g = upto; g > 0 && (store[genKey(id, g)] || cache[genKey(id, g)]); g--) { delete store[genKey(id, g)]; reset(genKey(id, g)); } } // ponytail: keeps the last KEEP saves for undo; older ones degrade to the current file
export const KEEP = 10;
export async function openPdf(key) {
  if (cache[key]) return cache[key];
  const L = await pdfjs(), bytes = store[key] || (key.includes('@') && store[key.slice(0, key.lastIndexOf('@'))]);
  if (!bytes) throw new Error('no bytes for ' + key);
  cache[key] = L.getDocument({ data: bytes.slice(), cMapUrl: CDN + 'cmaps/', cMapPacked: true, standardFontDataUrl: CDN + 'standard_fonts/' }).promise;
  return cache[key];
}
export async function importPdf(file, id) {
  const bytes = new Uint8Array(await file.arrayBuffer()); store[id] = bytes;
  const pdf = await openPdf(id);
  return Array.from({ length: pdf.numPages }, (_, i) => ({ id: 'p' + i + '_' + Math.random().toString(36).slice(2, 6), src: i, rot: 0, from: id }));
}
/** w/h as displayed (rotation applied), w0/h0 unrotated, rot the total rotation the page is shown with. */
export async function pageInfo(id, pg) {
  if (pg.src < 0) { const r = (pg.rot || 0) % 180; return { w: r ? 842 : 595, h: r ? 595 : 842, w0: 595, h0: 842, rot: (pg.rot || 0) % 360 }; }
  const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1);
  const rot = (p.rotate + (pg.rot || 0)) % 360, vp = p.getViewport({ scale: 1, rotation: rot }), v0 = p.getViewport({ scale: 1, rotation: 0 });
  return { w: vp.width, h: vp.height, w0: v0.width, h0: v0.height, rot };
}
/** Draws the page at scale × the device pixel ratio; mode is pdf.js's annotationMode (default: the file's annotations, form fields as they are). */
export async function renderPage(id, pg, canvas, scale, mode) {
  const dpr = Math.min(2, window.devicePixelRatio || 1), ctx = canvas.getContext('2d');
  if (pg.src < 0) { const inf = await pageInfo(id, pg); canvas.width = inf.w * scale * dpr; canvas.height = inf.h * scale * dpr; ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height); return; }
  const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1);
  const vp = p.getViewport({ scale: scale * dpr, rotation: (p.rotate + (pg.rot || 0)) % 360 });
  canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
  if (canvas.__task) try { canvas.__task.cancel(); } catch (e) { }
  const task = p.render(Object.assign({ canvasContext: ctx, viewport: vp }, mode != null ? { annotationMode: mode } : {})); canvas.__task = task;
  try { await task.promise; } catch (e) { if (e && e.name !== 'RenderingCancelledException') throw e; }
}
/** Prints through the system dialog: every page drawn at 150 dpi (form values and pending annotations included) into a
 *  hidden frame sized to the first page, one image per sheet, then window.print() there. */
export async function printPdf(id, doc) {
  const dpr = Math.min(2, window.devicePixelRatio || 1), c = document.createElement('canvas'), imgs = []; let size = null;
  for (const pg of doc.pages) {
    const inf = await pageInfo(id, pg); size = size || inf;
    await renderPage(id, pg, c, 150 / 72 / dpr, 3); drawAnnots(c.getContext('2d'), (doc.annots || []).filter(a => a.page === pg.id), c.width / inf.w);
    imgs.push(c.toDataURL('image/jpeg', 0.92));
  }
  if (!size) return;
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${String(doc.title || '').replace(/[<&]/g, '')}</title><style>@page{size:${size.w}pt ${size.h}pt;margin:0}html,body{margin:0}.p{width:${size.w}pt;height:${size.h}pt;display:flex;align-items:center;justify-content:center;overflow:hidden;break-after:page}img{max-width:100%;max-height:100%}</style></head><body>${imgs.map(s => `<div class="p"><img src="${s}"></div>`).join('')}</body></html>`;
  const f = document.createElement('iframe'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0'; document.body.appendChild(f);
  f.contentDocument.open(); f.contentDocument.write(html); f.contentDocument.close();
  await Promise.all(Array.from(f.contentDocument.images, i => i.decode().catch(() => { })));
  try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { }
  setTimeout(() => f.remove(), 60000);
}
/** The annotations already in the file (not links, pop-ups or form fields), with their box as displayed. */
export async function fileAnnots(id, pg) {
  if (pg.src < 0) return [];
  const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1), vp = p.getViewport({ scale: 1, rotation: (p.rotate + (pg.rot || 0)) % 360 });
  return (await p.getAnnotations()).filter(a => a.rect && !/^(Link|Popup|Widget)$/.test(a.subtype) && !(a.annotationFlags & 2)).map(a => {
    const [x1, y1, x2, y2] = vp.convertToViewportRectangle(a.rect);
    return { id: a.id, subtype: a.subtype, x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), text: (a.contentsObj && a.contentsObj.str) || '', author: (a.titleObj && a.titleObj.str) || '' };
  });
}

// ----- forms: the page's fillable fields, drawn by the editor as inputs; values live in pdf.js's annotation storage, which
// saveBytes writes into the file (pdf.js's own incremental update, /V and appearances) -----
/** Text, checkbox, radio and choice fields with their box as displayed and current value. */
export async function widgets(id, pg) {
  if (pg.src < 0) return [];
  const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1), vp = p.getViewport({ scale: 1, rotation: (p.rotate + (pg.rot || 0)) % 360 }), out = [];
  for (const a of await p.getAnnotations()) {
    if (a.subtype !== 'Widget' || a.hidden || (a.annotationFlags & 2)) continue;
    const kind = a.fieldType === 'Tx' ? 'text' : a.fieldType === 'Btn' ? (a.checkBox ? 'check' : a.radioButton ? 'radio' : null) : a.fieldType === 'Ch' ? 'choice' : null;
    if (!kind) continue;
    const [x1, y1, x2, y2] = vp.convertToViewportRectangle(a.rect), st = pdf.annotationStorage.getRawValue(a.id), has = st && st.value !== undefined;
    const w = { id: a.id, kind, name: a.fieldName || '', x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), ro: !!a.readOnly, fs: (a.defaultAppearanceData && a.defaultAppearanceData.fontSize) || 0 };
    if (kind === 'text') { w.value = has ? String(st.value) : a.fieldValue == null ? '' : String(a.fieldValue); w.multi = !!a.multiLine; w.maxLen = a.maxLen || 0; }
    else if (kind === 'check') w.checked = has ? !!st.value : a.fieldValue === a.exportValue;
    else if (kind === 'radio') w.checked = has ? !!st.value : a.fieldValue === a.buttonValue;
    else { w.options = (a.options || []).map(o => ({ v: o.exportValue, t: o.displayValue })); const v = has ? st.value : Array.isArray(a.fieldValue) ? a.fieldValue[0] : a.fieldValue; w.value = v == null ? '' : String(v); }
    out.push(w);
  }
  return out;
}
/** Records a field's value; a radio turns the other buttons of its group (same field name on the page) off. */
export async function setField(id, pg, wid, value, group) {
  const s = (await openPdf(keyOf(id, pg))).annotationStorage;
  s.setValue(wid, { value });
  for (const o of group || []) if (o !== wid) s.setValue(o, { value: false });
}

// ----- text layer: one transparent span per text item, laid out in unrotated page space at scale 1 -----
// The editor sizes the container w0 × h0 and rotates and scales it with layerTransform, so a zoom or a rotation never
// rebuilds the spans. Each span is scaled horizontally so its measured width matches the width the PDF gives the run.
const ROT = { 0: '', 90: 'rotate(90deg) translateY(-100%)', 180: 'rotate(180deg) translate(-100%,-100%)', 270: 'rotate(270deg) translateX(-100%)' };
export const layerTransform = (rot, scale) => `scale(${scale}) ${ROT[((rot % 360) + 360) % 360] || ''}`.trim();
let mctx = null; const ascents = {};
function measurer() { if (!mctx) mctx = document.createElement('canvas').getContext('2d'); return mctx; }
function ascent(ctx, fam) { if (ascents[fam] == null) { ctx.font = '100px ' + fam; const m = ctx.measureText('Hg'); ascents[fam] = (m.fontBoundingBoxAscent || 80) / 100; } return ascents[fam]; }
/** The page's text as one string: items in reading order, a newline where the PDF ends a line. */
export function textOf(items) { let s = ''; for (const it of items) if (it.str !== undefined) s += it.str + (it.hasEOL ? '\n' : ''); return s; }
/** Fills `div` with the page's selectable text; div.__tl = { items: [{ el, str, start, len }], text } for search marks. */
export async function textLayer(id, pg, div) {
  div.replaceChildren(); div.__tl = null; if (pg.src < 0) return;
  const L = await pdfjs(), pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1);
  const vp = p.getViewport({ scale: 1, rotation: 0 }), tc = await p.getTextContent(), ctx = measurer();
  const items = [], frag = document.createDocumentFragment(); let text = '';
  for (const it of tc.items) {
    if (it.str === undefined) continue;
    const start = text.length; text += it.str + (it.hasEOL ? '\n' : '');
    if (it.str) {
      const st = tc.styles[it.fontName] || {}, tx = L.Util.transform(vp.transform, it.transform);
      let angle = Math.atan2(tx[1], tx[0]); if (st.vertical) angle += Math.PI / 2;
      const fh = Math.hypot(tx[2], tx[3]), fam = st.fontFamily || 'sans-serif', asc = fh * ascent(ctx, fam);
      const left = angle ? tx[4] + asc * Math.sin(angle) : tx[4], top = angle ? tx[5] - asc * Math.cos(angle) : tx[5] - asc;
      const el = document.createElement('span'); el.textContent = it.str; if (it.dir) el.dir = it.dir;
      el.style.cssText = `left:${left.toFixed(2)}px;top:${top.toFixed(2)}px;font-size:${fh.toFixed(2)}px;font-family:${fam}`;
      const cw = st.vertical ? it.height : it.width; let tf = '';
      if (cw > 0 && fh > 0) { ctx.font = `${fh * 4}px ${fam}`; const mw = ctx.measureText(it.str).width / 4; if (mw > 0) tf = `scaleX(${(cw / mw).toFixed(4)})`; }
      if (angle) tf = `rotate(${(angle * 180 / Math.PI).toFixed(2)}deg) ` + tf;
      if (tf) el.style.transform = tf;
      frag.appendChild(el); items.push({ el, str: it.str, start, len: it.str.length });
    }
    if (it.hasEOL) frag.appendChild(document.createElement('br'));
  }
  div.appendChild(frag); div.__tl = { items, text };
}
/** Client rects of a text selection (page units) merged into one box per line. */
export function mergeRects(rs) {
  const out = [];
  for (const r of rs.filter(r => r.w > 0.5 && r.h > 0.5).sort((a, b) => a.y - b.y || a.x - b.x)) {
    const l = out.find(o => Math.min(o.y + o.h, r.y + r.h) - Math.max(o.y, r.y) > 0.5 * Math.min(o.h, r.h) && r.x <= o.x + o.w + o.h && r.x + r.w >= o.x - o.h);
    if (!l) { out.push(Object.assign({}, r)); continue; }
    const x = Math.min(l.x, r.x), y = Math.min(l.y, r.y); l.w = Math.max(l.x + l.w, r.x + r.w) - x; l.h = Math.max(l.y + l.h, r.y + r.h) - y; l.x = x; l.y = y;
  }
  return out;
}

/** The PDF's outline (bookmarks) flattened: [{ title, level, index }], index the page it points to (null when unresolved). */
export async function outline(key) {
  const pdf = await openPdf(key), out = [];
  const walk = async (items, level) => {
    for (const it of items || []) {
      let index = null;
      try { let dest = it.dest; if (typeof dest === 'string') dest = await pdf.getDestination(dest); if (Array.isArray(dest) && dest[0] != null) index = typeof dest[0] === 'object' ? await pdf.getPageIndex(dest[0]) : dest[0]; } catch (e) { }
      out.push({ title: it.title, level, index });
      await walk(it.items, level + 1);
    }
  };
  await walk(await pdf.getOutline(), 0);
  return out;
}

// ----- search: every page's text is fetched once per source; hits are offsets into textOf(), marked inside the layer spans -----
const texts = {};
export async function pageText(id, pg) {
  if (pg.src < 0) return '';
  const k = keyOf(id, pg) + ':' + pg.src;
  if (texts[k] == null) { const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1); texts[k] = textOf((await p.getTextContent()).items); }
  return texts[k];
}
/** Case-insensitive matches of q in each page's text (any run of spaces matches a line end too): [{ p, start, end, i }]. */
export function findIn(pageTexts, q) {
  q = (q || '').trim(); if (!q) return [];
  const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+'), 'gi'), hits = [];
  pageTexts.forEach((t, p) => { for (const m of (t || '').matchAll(re)) if (m[0]) hits.push({ p, start: m.index, end: m.index + m[0].length, i: hits.length }); });
  return hits;
}
/** Wraps the parts of the layer's spans inside hits in <span class="pt-hit">, the hit numbered cur also .cur; no hits clears the marks. */
export function markHits(div, hits, cur) {
  const tl = div.__tl; if (!tl || (!hits.length && !div.__marked)) return;
  div.__marked = hits.length > 0;
  for (const it of tl.items) {
    const end = it.start + it.len, rs = hits.filter(h => h.start < end && h.end > it.start).sort((a, b) => a.start - b.start);
    if (!rs.length) { if (it.marked) { it.el.textContent = it.str; it.marked = false; } continue; }
    const frag = document.createDocumentFragment(); let pos = 0;
    for (const h of rs) {
      const s = Math.max(pos, h.start - it.start), e = Math.min(it.len, h.end - it.start); if (e <= s) continue;
      if (s > pos) frag.appendChild(document.createTextNode(it.str.slice(pos, s)));
      const m = document.createElement('span'); m.className = 'pt-hit' + (h.i === cur ? ' cur' : ''); m.textContent = it.str.slice(s, e); frag.appendChild(m); pos = e;
    }
    if (pos < it.len) frag.appendChild(document.createTextNode(it.str.slice(pos)));
    it.el.replaceChildren(frag); it.marked = true;
  }
}

export async function textItems(id, pg) {
  if (pg.src < 0) return [];
  const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1);
  const vp = p.getViewport({ scale: 1, rotation: (p.rotate + (pg.rot || 0)) % 360 });
  const tc = await p.getTextContent(); const out = [];
  tc.items.forEach(it => {
    if (!it.str || !it.str.trim()) return;
    const [a, b, c, d, e, f] = it.transform; const h = Math.hypot(c, d) || Math.hypot(a, b); const w = it.width;
    const [x1, y1, x2, y2] = vp.convertToViewportRectangle([e, f - h * 0.22, e + w, f + h * 0.9]);
    out.push({ str: it.str, x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), fs: h });
  });
  return out;
}
export async function extractText(id, pages) {
  const out = [];
  for (const pg of pages) {
    if (pg.src < 0) continue;
    const pdf = await openPdf(keyOf(id, pg)), p = await pdf.getPage(pg.src + 1), tc = await p.getTextContent();
    const lines = []; let cur = null, lastY = null;
    tc.items.forEach(it => { const y = Math.round(it.transform[5]); if (lastY === null || Math.abs(y - lastY) > 2) { cur = { y, h: Math.hypot(it.transform[2], it.transform[3]), t: '' }; lines.push(cur); lastY = y; } cur.t += it.str; });
    const paras = []; let buf = '', prevY = null, prevH = 12;
    lines.forEach(l => { const gap = prevY === null ? 0 : prevY - l.y; if (buf && gap > prevH * 1.7) { paras.push(buf); buf = ''; } buf += (buf && /[A-Za-z0-9]$/.test(buf) ? ' ' : '') + l.t.trim(); prevY = l.y; prevH = l.h || prevH; });
    if (buf) paras.push(buf);
    out.push(paras.filter(Boolean));
  }
  return out;
}

// ----- drawing the editor's own annotations (print, PNG export, appearance images) -----
export const QUAD = { hl: 'Highlight', ul: 'Underline', so: 'StrikeOut' };
export function drawAnnots(ctx, annots, k) {
  annots.forEach(a => {
    ctx.save();
    if (QUAD[a.t]) {
      ctx.fillStyle = a.color || '#FFE066'; if (a.t === 'hl') ctx.globalCompositeOperation = 'multiply';
      for (const [x, y, w, h] of a.quads || [[a.x, a.y, a.w, a.h]]) {
        if (a.t === 'hl') ctx.fillRect(x * k, y * k, w * k, h * k);
        else ctx.fillRect(x * k, (a.t === 'ul' ? y + h - 1.5 : y + h * 0.5) * k, w * k, 1.5 * k);
      }
    }
    if (a.t === 'rect') { ctx.strokeStyle = a.color || '#B5563A'; ctx.lineWidth = (a.sw || 2) * k; ctx.strokeRect(a.x * k, a.y * k, a.w * k, a.h * k); }
    if (a.t === 'white' || a.t === 'replace') { ctx.fillStyle = a.bg || '#FFFFFF'; ctx.fillRect(a.x * k, a.y * k, a.w * k, a.h * k); }
    if (a.t === 'text' || a.t === 'replace') {
      ctx.fillStyle = a.color || '#23211D'; const fs = (a.fs || 14) * k; ctx.font = `${a.bold ? '700 ' : ''}${fs}px 'Noto Sans SC','IBM Plex Sans',sans-serif`; ctx.textBaseline = 'top';
      const maxW = a.w * k - (a.t === 'text' ? 8 * k : 0); let y = a.y * k + (a.t === 'text' ? 4 * k : Math.max(0, (a.h * k - fs * 1.15) / 2));
      String(a.text || '').split('\n').forEach(par => { let line = ''; for (const ch of par) { if (ctx.measureText(line + ch).width > maxW && line) { ctx.fillText(line, a.x * k + (a.t === 'text' ? 4 * k : 0), y); y += fs * 1.3; line = ch; } else line += ch; } ctx.fillText(line, a.x * k + (a.t === 'text' ? 4 * k : 0), y); y += fs * 1.3; });
    }
    if (a.t === 'ink' && a.pts && a.pts.length > 1) { ctx.strokeStyle = a.color || '#2F5D8A'; ctx.lineWidth = (a.sw || 2.5) * k; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.beginPath(); a.pts.forEach(([x, y], i) => i ? ctx.lineTo(x * k, y * k) : ctx.moveTo(x * k, y * k)); ctx.stroke(); }
    if (a.t === 'note') { ctx.fillStyle = a.color || '#FFD54A'; ctx.beginPath(); ctx.roundRect(a.x * k, a.y * k, 20 * k, 20 * k, 4 * k); ctx.fill(); ctx.strokeStyle = '#FFFFFF'; ctx.lineWidth = k; [7, 10.5, 14].forEach((dy, i) => { ctx.beginPath(); ctx.moveTo((a.x + 4) * k, (a.y + dy) * k); ctx.lineTo((a.x + (i === 2 ? 12 : 16)) * k, (a.y + dy) * k); ctx.stroke(); }); }
    ctx.restore();
  });
}

// ----- writing annotations into the PDF -----
// Each annotation becomes a standard /Annot object (Highlight, Underline, StrikeOut, Square, Ink, Text; FreeText with an
// image appearance so any font, CJK included, shows everywhere). They are appended as an incremental update: the original
// bytes stay as they are, followed by the new objects, the touched page objects (their /Annots), a cross-reference section
// of the same kind as the file's last one and a trailer chaining to it. Other viewers see and can edit them.
const f2 = n => String(Math.round(n * 100) / 100);
const arr = a => '[' + a.map(f2).join(' ') + ']';
const latin1 = s => Uint8Array.from(s, c => c.charCodeAt(0) & 255);
const rgb = hex => { const m = /^#?([0-9a-f]{6})$/i.exec(hex || ''); if (!m) return '0 0 0'; const v = parseInt(m[1], 16); return [(v >> 16) & 255, (v >> 8) & 255, v & 255].map(c => f2(c / 255)).join(' '); };
const hex16 = s => '<FEFF' + Array.from(String(s || ''), ch => ch.charCodeAt(0).toString(16).padStart(4, '0')).join('').toUpperCase() + '>';
const pdfDate = () => { const d = new Date(), p = n => String(n).padStart(2, '0'); return `(D:${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())})`; };
/** A point as displayed (y down, page rotated by rot) → PDF user space, for the page box [x0 y0 x1 y1]. */
export function toPdf(x, y, rot, box) {
  const [x0, y0, x1, y1] = box;
  switch (((rot % 360) + 360) % 360) { case 90: return [x0 + y, y0 + x]; case 180: return [x1 - x, y0 + y]; case 270: return [x1 - y, y1 - x]; default: return [x0 + x, y1 - y]; }
}
function pdfRect(x, y, w, h, rot, box) { const [ax, ay] = toPdf(x, y, rot, box), [bx, by] = toPdf(x + w, y + h, rot, box); return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)]; }
/** The objects for one annotation, numbered from o.num: [{ num, body } | { num, dict, stream }], the first one the /Annot.
 *  o: { num, pageRef: 'n g R', rot, box, ap?: { w, h, rgb, alpha } (deflated image planes for text annotations) }. */
export function annotObjects(a, o) {
  const objs = [], R = o.rot || 0, B = o.box; let n = o.num;
  const base = (sub, rect, extra, flags = 4) => `<< /Type /Annot /Subtype /${sub} /Rect ${arr(rect)} /P ${o.pageRef} /F ${flags} /M ${pdfDate()} /NM (${a.id}) /CA 1 ${extra} >>`;
  if (QUAD[a.t]) {
    const quads = (a.quads || [[a.x, a.y, a.w, a.h]]).map(([x, y, w, h]) => pdfRect(x, y, w, h, R, B));
    const rect = [Math.min(...quads.map(q => q[0])), Math.min(...quads.map(q => q[1])), Math.max(...quads.map(q => q[2])), Math.max(...quads.map(q => q[3]))];
    objs.push({ num: n++, body: base(QUAD[a.t], rect, `/QuadPoints ${arr(quads.flatMap(([x1, y1, x2, y2]) => [x1, y2, x2, y2, x1, y1, x2, y1]))} /C [${rgb(a.color || '#FFE066')}]`) });
  } else if (a.t === 'rect' || a.t === 'white') {
    const white = a.t === 'white';
    objs.push({ num: n++, body: base('Square', pdfRect(a.x, a.y, a.w, a.h, R, B), `/C [${white ? '1 1 1' : rgb(a.color)}] /BS << /W ${f2(white ? 0 : a.sw || 2)} /S /S >>${white ? ' /IC [1 1 1]' : ''}`) });
  } else if (a.t === 'ink' && a.pts && a.pts.length > 1) {
    const pts = a.pts.map(([x, y]) => toPdf(x, y, R, B)), sw = a.sw || 2.5, xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
    objs.push({ num: n++, body: base('Ink', [Math.min(...xs) - sw, Math.min(...ys) - sw, Math.max(...xs) + sw, Math.max(...ys) + sw], `/InkList [${arr(pts.flat())}] /C [${rgb(a.color)}] /BS << /W ${f2(sw)} >>`) });
  } else if (a.t === 'note') {
    const icon = `q ${rgb(a.color || '#FFD54A')} rg 1 1 18 18 re f 1 1 1 RG 1 w 4 13 m 16 13 l S 4 9.5 m 16 9.5 l S 4 6 m 12 6 l S Q`;
    objs.push({ num: n, body: base('Text', pdfRect(a.x, a.y, 20, 20, R, B), `/Contents ${hex16(a.text)} /Name /Comment /C [${rgb(a.color || '#FFD54A')}] /Open false /AP << /N ${n + 1} 0 R >>`, 28) });
    objs.push({ num: n + 1, dict: `<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Length ${icon.length} >>`, stream: latin1(icon) }); n += 2;
  } else if ((a.t === 'text' || a.t === 'replace') && o.ap) {
    const { w, h, rgb: px, alpha } = o.ap, rw = R % 180 ? a.h : a.w, rh = R % 180 ? a.w : a.h;
    const M = { 0: '1 0 0 1 0 0', 90: `0 1 -1 0 ${f2(a.h)} 0`, 180: `-1 0 0 -1 ${f2(a.w)} ${f2(a.h)}`, 270: `0 -1 1 0 0 ${f2(a.w)}` }[((R % 360) + 360) % 360];
    const cs = `q ${M} cm ${f2(a.w)} 0 0 ${f2(a.h)} 0 0 cm /Im0 Do Q`;
    objs.push({ num: n, body: base('FreeText', pdfRect(a.x, a.y, a.w, a.h, R, B), `/Contents ${hex16(a.text)} /DA (${rgb(a.color)} rg /Helv ${f2(a.fs || 14)} Tf) /AP << /N ${n + 3} 0 R >>`) });
    objs.push({ num: n + 1, dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /SMask ${n + 2} 0 R /Length ${px.length} >>`, stream: px });
    objs.push({ num: n + 2, dict: `<< /Type /XObject /Subtype /Image /Width ${w} /Height ${h} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode /Length ${alpha.length} >>`, stream: alpha });
    objs.push({ num: n + 3, dict: `<< /Type /XObject /Subtype /Form /BBox [0 0 ${f2(rw)} ${f2(rh)}] /Resources << /XObject << /Im0 ${n + 1} 0 R >> >> /Length ${cs.length} >>`, stream: latin1(cs) }); n += 4;
  }
  return { objs, next: n };
}
/** The first object number an update may use: past every `n g obj` and every trailer /Size in the file (pdf-lib does not
 *  count the object streams and cross-reference streams it unpacked, and an update must not reuse their numbers). */
export function firstFree(bytes) {
  const s = new TextDecoder('latin1').decode(bytes); let n = 0;
  for (const m of s.matchAll(/(?:^|[^0-9])(\d+)\s+\d+\s+obj\b/g)) n = Math.max(n, +m[1] + 1);
  for (const m of s.matchAll(/\/Size\s+(\d+)/g)) n = Math.max(n, +m[1]);
  return n;
}
/** Appends an incremental update to bytes: the objects, a cross-reference table or stream (whichever the file ends with)
 *  and a trailer with /Prev. t: { size, root, info?, id? } in PDF syntax. Throws when the file has no usable startxref. */
export function appendUpdate(bytes, objs, t) {
  const tail = String.fromCharCode(...bytes.slice(-2048)), m = /startxref\s+(\d+)\s*%%EOF\s*$/.exec(tail);
  if (!m || +m[1] >= bytes.length) throw new Error('no xref');
  const prev = +m[1]; let i = prev; while (i < bytes.length && /\s/.test(String.fromCharCode(bytes[i]))) i++;
  const table = String.fromCharCode(...bytes.slice(i, i + 4)) === 'xref';
  const chunks = [bytes, latin1('\n')], entries = []; let off = bytes.length + 1;
  const push = u8 => { chunks.push(u8); off += u8.length; };
  for (const o of [...objs].sort((a, b) => a.num - b.num)) {
    entries.push({ num: o.num, gen: o.gen || 0, off });
    push(latin1(`${o.num} ${o.gen || 0} obj\n`));
    if (o.stream) { push(latin1(o.dict + '\nstream\n')); push(o.stream); push(latin1('\nendstream')); } else push(latin1(o.body));
    push(latin1('\nendobj\n'));
  }
  const xref = off, trailer = `/Root ${t.root} /Prev ${prev}${t.info ? ' /Info ' + t.info : ''}${t.id ? ' /ID ' + t.id : ''}`;
  let size = t.size;
  if (table) {
    const runs = [];
    for (const e of entries) { const r = runs[runs.length - 1]; if (r && r.start + r.list.length === e.num) r.list.push(e); else runs.push({ start: e.num, list: [e] }); }
    let s = 'xref\n';
    for (const r of runs) { s += `${r.start} ${r.list.length}\n`; for (const e of r.list) s += `${String(e.off).padStart(10, '0')} ${String(e.gen).padStart(5, '0')} n \n`; }
    push(latin1(s + `trailer\n<< /Size ${size} ${trailer} >>\nstartxref\n${xref}\n%%EOF\n`));
  } else {
    const num = size++; entries.push({ num, gen: 0, off: xref });
    const rows = new Uint8Array(entries.length * 7), idx = [];
    entries.forEach((e, k) => { rows.set([1, (e.off >>> 24) & 255, (e.off >>> 16) & 255, (e.off >>> 8) & 255, e.off & 255, (e.gen >> 8) & 255, e.gen & 255], k * 7); const r = idx[idx.length - 1]; if (r && r[0] + r[1] === e.num) r[1]++; else idx.push([e.num, 1]); });
    push(latin1(`${num} 0 obj\n<< /Type /XRef /Size ${size} /W [1 4 2] /Index [${idx.flat().join(' ')}] ${trailer} /Length ${rows.length} >>\nstream\n`)); push(rows); push(latin1(`\nendstream\nendobj\nstartxref\n${xref}\n%%EOF\n`));
  }
  const out = new Uint8Array(off); let p = 0; for (const c of chunks) { out.set(c, p); p += c.length; }
  return out;
}
async function deflate(u8) { return new Uint8Array(await new Response(new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate'))).arrayBuffer()); }
/** A text annotation drawn at 2× as RGB + alpha planes, deflated, for its appearance stream. */
async function apPlanes(a) {
  const K = 2, w = Math.max(1, Math.ceil(a.w * K)), h = Math.max(1, Math.ceil(a.h * K)), c = document.createElement('canvas'); c.width = w; c.height = h;
  const ctx = c.getContext('2d'); ctx.translate(-a.x * K, -a.y * K); drawAnnots(ctx, [a], K);
  const d = ctx.getImageData(0, 0, w, h).data, px = new Uint8Array(w * h * 3), al = new Uint8Array(w * h);
  for (let i = 0, j = 0; i < al.length; i++, j += 4) { px[i * 3] = d[j]; px[i * 3 + 1] = d[j + 1]; px[i * 3 + 2] = d[j + 2]; al[i] = d[j + 3]; }
  return { w, h, rgb: await deflate(px), alpha: await deflate(al) };
}
const refId = r => r.objectNumber + 'R' + (r.generationNumber || ''); // pdf.js's annotation ids
/** Appends the annotations (Map page index → [annot]), drops the removed ones (pdf.js ids) and carries edited form values
 *  (pdf.js ids of the widgets) up to their field dictionaries, as one incremental update. */
async function writeAnnots(bytes, byIndex, removed, planes, fields) {
  const { PDFDocument, PDFName, PDFArray, PDFRef, PDFDict } = await pl();
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false }), ctx = doc.context, ti = ctx.trailerInfo;
  if (ti.Encrypt) throw new Error('encrypted');
  const pages = doc.getPages(), objs = []; let next = Math.max(ctx.largestObjectNumber + 1, firstFree(bytes));
  // pdf.js puts a new value on the widget; when the widget is a kid of its field, the field must carry /V too (Acrobat reads it there)
  const N = s => PDFName.of(s);
  for (const fid of fields || []) {
    const m = /^(\d+)R(\d*)$/.exec(fid); if (!m) continue;
    const w = ctx.lookup(PDFRef.of(+m[1], m[2] ? +m[2] : 0)), v = w instanceof PDFDict && w.get(N('V')), pr = w instanceof PDFDict && w.get(N('Parent'));
    if (!v || !(pr instanceof PDFRef) || w.has(N('T'))) continue;
    const parent = ctx.lookup(pr); if (!(parent instanceof PDFDict) || !parent.has(N('T')) || String(parent.get(N('V'))) === String(v)) continue;
    parent.set(N('V'), v); objs.push({ num: pr.objectNumber, gen: pr.generationNumber, body: parent.toString() });
  }
  for (const [i, page] of pages.entries()) {
    const list = byIndex.get(i) || [], old = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray), was = old ? old.asArray() : [];
    const keep = was.filter(r => !(r instanceof PDFRef && removed.has(refId(r))));
    if (!list.length && keep.length === was.length) continue;
    const mb = page.getMediaBox(), cb = page.getCropBox ? page.getCropBox() : mb;
    const x0 = Math.max(mb.x, cb.x), y0 = Math.max(mb.y, cb.y), box = [x0, y0, Math.min(mb.x + mb.width, cb.x + cb.width), Math.min(mb.y + mb.height, cb.y + cb.height)];
    const o = { pageRef: `${page.ref.objectNumber} ${page.ref.generationNumber} R`, rot: page.getRotation().angle, box }, refs = [];
    for (const a of list) { const r = annotObjects(a, Object.assign({ num: next, ap: planes[a.id] }, o)); if (!r.objs.length) continue; objs.push(...r.objs); refs.push(PDFRef.of(r.objs[0].num)); next = r.next; }
    page.node.set(PDFName.of('Annots'), ctx.obj([...keep, ...refs]));
    objs.push({ num: page.ref.objectNumber, gen: page.ref.generationNumber, body: page.node.toString() });
  }
  if (!objs.length) return bytes;
  if (!ti.Root) throw new Error('no root');
  return appendUpdate(bytes, objs, { size: next, root: ti.Root.toString(), info: ti.Info && ti.Info.toString(), id: ti.ID && ti.ID.toString() });
}
/** The document rebuilt with the pages in the editor's order: its own pages moved, copies of pages from other sources
 *  (merged files, older generations), blank pages, rotations set; the catalog (outline, forms, metadata) stays. */
async function restructure(bytes, id, key, pages) {
  const { PDFDocument, degrees } = await pl(), opts = { ignoreEncryption: true, updateMetadata: false };
  const doc = await PDFDocument.load(bytes, opts), own = doc.getPages(), others = {};
  for (let i = own.length - 1; i >= 0; i--) doc.removePage(i);
  for (const [i, pg] of pages.entries()) {
    let page;
    if (pg.src >= 0 && keyOf(id, pg) === key && own[pg.src]) { page = own[pg.src]; doc.insertPage(i, page); }
    else if (pg.src >= 0 && store[keyOf(id, pg)]) { const k = keyOf(id, pg), o = others[k] || (others[k] = await PDFDocument.load(store[k], opts)); [page] = await doc.copyPages(o, [pg.src]); doc.insertPage(i, page); }
    else page = doc.insertPage(i, [595, 842]);
    if (pg.rot) page.setRotation(degrees((((page.getRotation().angle + pg.rot) % 360) + 360) % 360));
  }
  return await doc.save({ useObjectStreams: false });
}
/** The key of the bytes the editor's pages are relative to. */
export const baseKey = (id, doc) => doc._gen ? genKey(id, doc._gen) : id;
/** Whether the editor's state differs from the file on disk. */
export function dirty(id, doc) {
  const key = baseKey(id, doc);
  return !!((doc.annots && doc.annots.length) || (doc.removed && doc.removed.length) || (doc.form || 0) !== (doc._formSaved || 0) || (doc._gen || 0) !== gen(id)
    || doc.pages.length !== doc._n || doc.pages.some((p, i) => p.src !== i || (p.rot || 0) % 360 || keyOf(id, p) !== key));
}
/** The file as the editor shows it: form values, then the page structure, then the annotations as an incremental update. */
export async function saveBytes(id, doc) {
  const key = baseKey(id, doc), pdf = await openPdf(key); let bytes = store[key] || store[id], fields = [];
  const stored = pdf.annotationStorage && pdf.annotationStorage.size ? pdf.annotationStorage.getAll() : null;
  if (stored) { bytes = await pdf.saveDocument(); fields = Object.keys(stored); }
  const same = doc.pages.length === pdf.numPages && doc.pages.every((p, i) => p.src === i && !((p.rot || 0) % 360) && keyOf(id, p) === key);
  if (!same) bytes = await restructure(bytes, id, key, doc.pages);
  const byIndex = new Map(); (doc.annots || []).forEach(a => { const i = doc.pages.findIndex(p => p.id === a.page); if (i >= 0) byIndex.set(i, [...(byIndex.get(i) || []), a]); });
  const removed = new Set(doc.removed || []);
  if (!byIndex.size && !removed.size && !fields.length) return bytes;
  const planes = {}; for (const a of doc.annots || []) if (a.t === 'text' || a.t === 'replace') planes[a.id] = await apPlanes(a);
  try { return await writeAnnots(bytes, byIndex, removed, planes, fields); }
  catch (e) { // no usable cross-reference to chain to (a damaged file pdf-lib repaired on load): rewrite it first, then append
    if (!/xref/.test(e.message)) throw e;
    const { PDFDocument } = await pl(); bytes = await (await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })).save({ useObjectStreams: false });
    return await writeAnnots(bytes, byIndex, removed, planes, fields);
  }
}
export const exportPdf = saveBytes;
/** The pages at these indices (into doc.pages) as a PDF of their own, annotations included. */
export async function extractPdf(id, doc, indices) {
  const { PDFDocument } = await pl(), src = await PDFDocument.load(await saveBytes(id, doc), { ignoreEncryption: true, updateMetadata: false }), out = await PDFDocument.create();
  for (const p of await out.copyPages(src, indices)) out.addPage(p);
  return await out.save({ useObjectStreams: false });
}
/** "1-3,5" → [0, 1, 2, 4] within n pages; null when it is not a page list. */
export function parseRange(s, n) {
  const out = new Set();
  for (const part of String(s || '').split(/[,，;\s]+/).filter(Boolean)) {
    const m = /^(\d+)(?:[-–~](\d+))?$/.exec(part); if (!m) return null;
    const a = +m[1], b = m[2] ? +m[2] : a; for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (i >= 1 && i <= n) out.add(i - 1);
  }
  return [...out].sort((a, b) => a - b);
}
/** The page at 2× as a PNG blob, the editor's pending annotations drawn on top. */
export async function pagePng(id, pg, annots) {
  const c = document.createElement('canvas'); await renderPage(id, pg, c, 2 / Math.min(2, window.devicePixelRatio || 1));
  drawAnnots(c.getContext('2d'), annots || [], c.width / (await pageInfo(id, pg)).w);
  return new Promise(res => c.toBlob(res, 'image/png'));
}
export function download(data, name, type) { const a = document.createElement('a'); a.href = URL.createObjectURL(data instanceof Blob ? data : new Blob([data], { type: type || 'application/pdf' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }
