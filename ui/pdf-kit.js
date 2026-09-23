// PDF open / render / text extraction / export with annotations.
const V = '4.4.168', CDN = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${V}/`;
let lib = null, pdflib = null;
export async function pdfjs() { if (!lib) { lib = await import(CDN + 'build/pdf.min.mjs'); lib.GlobalWorkerOptions.workerSrc = CDN + 'build/pdf.worker.min.mjs'; } return lib; }
async function pl() { if (!pdflib) pdflib = await import('https://esm.sh/pdf-lib@1.17.1'); return pdflib; }
export const store = window.__pdfStore || (window.__pdfStore = {});
const cache = window.__pdfDocs || (window.__pdfDocs = {});
export async function openPdf(id) {
  if (cache[id]) return cache[id];
  const L = await pdfjs();
  cache[id] = L.getDocument({ data: store[id].slice(), cMapUrl: CDN + 'cmaps/', cMapPacked: true, standardFontDataUrl: CDN + 'standard_fonts/' }).promise;
  return cache[id];
}
export async function importPdf(file, id) {
  const bytes = new Uint8Array(await file.arrayBuffer()); store[id] = bytes;
  const pdf = await openPdf(id);
  return Array.from({ length: pdf.numPages }, (_, i) => ({ id: 'p' + i + '_' + Math.random().toString(36).slice(2, 6), src: i, rot: 0 }));
}
export function copyStore(from, to) { if (store[from]) store[to] = store[from]; }
export async function pageInfo(id, pg) {
  if (pg.src < 0) { const r = (pg.rot || 0) % 180; return { w: r ? 842 : 595, h: r ? 595 : 842, base: 0 }; }
  const pdf = await openPdf(id), p = await pdf.getPage(pg.src + 1);
  const vp = p.getViewport({ scale: 1, rotation: (p.rotate + (pg.rot || 0)) % 360 });
  return { w: vp.width, h: vp.height };
}
export async function renderPage(id, pg, canvas, scale) {
  const dpr = Math.min(2, window.devicePixelRatio || 1), ctx = canvas.getContext('2d');
  if (pg.src < 0) { const inf = await pageInfo(id, pg); canvas.width = inf.w * scale * dpr; canvas.height = inf.h * scale * dpr; ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height); return; }
  const pdf = await openPdf(id), p = await pdf.getPage(pg.src + 1);
  const vp = p.getViewport({ scale: scale * dpr, rotation: (p.rotate + (pg.rot || 0)) % 360 });
  canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
  if (canvas.__task) try { canvas.__task.cancel(); } catch (e) { }
  const task = p.render({ canvasContext: ctx, viewport: vp }); canvas.__task = task;
  try { await task.promise; } catch (e) { if (e && e.name !== 'RenderingCancelledException') throw e; }
}
export async function textItems(id, pg) {
  if (pg.src < 0) return [];
  const pdf = await openPdf(id), p = await pdf.getPage(pg.src + 1);
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
  const pdf = await openPdf(id); const out = [];
  for (const pg of pages) {
    if (pg.src < 0) continue;
    const p = await pdf.getPage(pg.src + 1), tc = await p.getTextContent();
    const lines = []; let cur = null, lastY = null;
    tc.items.forEach(it => { const y = Math.round(it.transform[5]); if (lastY === null || Math.abs(y - lastY) > 2) { cur = { y, h: Math.hypot(it.transform[2], it.transform[3]), t: '' }; lines.push(cur); lastY = y; } cur.t += it.str; });
    const paras = []; let buf = '', prevY = null, prevH = 12;
    lines.forEach(l => { const gap = prevY === null ? 0 : prevY - l.y; if (buf && gap > prevH * 1.7) { paras.push(buf); buf = ''; } buf += (buf && /[A-Za-z0-9]$/.test(buf) ? ' ' : '') + l.t.trim(); prevY = l.y; prevH = l.h || prevH; });
    if (buf) paras.push(buf);
    out.push(paras.filter(Boolean));
  }
  return out;
}
export function drawAnnots(ctx, annots, k) {
  annots.forEach(a => {
    ctx.save();
    if (a.t === 'hl') { ctx.fillStyle = a.color || '#FFE066'; ctx.globalAlpha = 0.42; ctx.fillRect(a.x * k, a.y * k, a.w * k, a.h * k); }
    if (a.t === 'rect') { ctx.strokeStyle = a.color || '#B5563A'; ctx.lineWidth = (a.sw || 2) * k; ctx.strokeRect(a.x * k, a.y * k, a.w * k, a.h * k); }
    if (a.t === 'white' || a.t === 'replace') { ctx.fillStyle = a.bg || '#FFFFFF'; ctx.fillRect(a.x * k, a.y * k, a.w * k, a.h * k); }
    if (a.t === 'text' || a.t === 'replace') {
      ctx.fillStyle = a.color || '#23211D'; const fs = (a.fs || 14) * k; ctx.font = `${a.bold ? '700 ' : ''}${fs}px 'Noto Sans SC','IBM Plex Sans',sans-serif`; ctx.textBaseline = 'top';
      const maxW = a.w * k - (a.t === 'text' ? 8 * k : 0); let y = a.y * k + (a.t === 'text' ? 4 * k : Math.max(0, (a.h * k - fs * 1.15) / 2));
      String(a.text || '').split('\n').forEach(par => { let line = ''; for (const ch of par) { if (ctx.measureText(line + ch).width > maxW && line) { ctx.fillText(line, a.x * k + (a.t === 'text' ? 4 * k : 0), y); y += fs * 1.3; line = ch; } else line += ch; } ctx.fillText(line, a.x * k + (a.t === 'text' ? 4 * k : 0), y); y += fs * 1.3; });
    }
    if (a.t === 'ink' && a.pts && a.pts.length > 1) { ctx.strokeStyle = a.color || '#2F5D8A'; ctx.lineWidth = (a.sw || 2.5) * k; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.beginPath(); a.pts.forEach(([x, y], i) => i ? ctx.lineTo(x * k, y * k) : ctx.moveTo(x * k, y * k)); ctx.stroke(); }
    ctx.restore();
  });
}
const toBytes = c => new Promise(res => c.toBlob(b => b.arrayBuffer().then(ab => res(new Uint8Array(ab))), 'image/png'));
export async function exportPdf(id, doc) {
  const { PDFDocument, degrees } = await pl();
  const src = store[id] ? await PDFDocument.load(store[id], { ignoreEncryption: true }) : null;
  const out = await PDFDocument.create();
  for (const pg of doc.pages) {
    let page;
    if (pg.src < 0 || !src) page = out.addPage([595, 842]);
    else { const [cp] = await out.copyPages(src, [pg.src]); page = out.addPage(cp); }
    const total = (((page.getRotation().angle || 0) + (pg.rot || 0)) % 360 + 360) % 360;
    page.setRotation(degrees(total));
    const an = (doc.annots || []).filter(a => a.page === pg.id);
    if (!an.length) continue;
    const { width: ow, height: oh } = page.getSize(); const sw = total % 180 ? oh : ow, shh = total % 180 ? ow : oh, K = 2;
    const c1 = document.createElement('canvas'); c1.width = sw * K; c1.height = shh * K; drawAnnots(c1.getContext('2d'), an, K);
    const c2 = document.createElement('canvas'); c2.width = ow * K; c2.height = oh * K; const x = c2.getContext('2d');
    x.translate(c2.width / 2, c2.height / 2); x.rotate(-total * Math.PI / 180); x.drawImage(c1, -c1.width / 2, -c1.height / 2);
    const img = await out.embedPng(await toBytes(c2));
    const box = page.getMediaBox ? page.getMediaBox() : { x: 0, y: 0 };
    page.drawImage(img, { x: box.x || 0, y: box.y || 0, width: ow, height: oh });
  }
  return await out.save();
}
export function download(bytes, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }
