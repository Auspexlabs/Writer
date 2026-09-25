// Shared slide kit + file import/export/print for 素笺 Office.
import { pictureView, picSrc } from './picture.js';
// $t under node (this module is node-tested): falls back to the Chinese, vars filled the same way. Only for text a new
// slide/table is created with — never for existing content, which engine.js reads from the file as it is.
const T = (s, v) => globalThis.$t ? globalThis.$t(s, v) : v ? String(s).replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : s;
export const THEMES = {
  ink: { name: '墨色', bg: '#1D1D1F', fg: '#FFFFFF', sub: '#BDB7AA', acc: '#E3B25A', card: '#2C2C2E', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  paper: { name: '素白', bg: '#FFFFFF', fg: '#1D1D1F', sub: '#6E6E73', acc: '#1D1D1F', card: '#F5F5F7', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  sea: { name: '海蓝', bg: '#1F3550', fg: '#FFFFFF', sub: '#C5D2E0', acc: '#6CC6D9', card: '#2A4666', hf: 'Noto Sans SC', bf: 'Noto Sans SC' },
  clay: { name: '陶土', bg: '#F3E6DA', fg: '#3A2618', sub: '#7A5A45', acc: '#B5563A', card: '#EAD5C3', hf: 'Noto Serif SC', bf: 'Noto Sans SC' }
};
export const SW = 1600;
export const slideH = ratio => ratio === '4:3' ? 1200 : 900;
let uid = Date.now() % 100000;
export const oid = () => 'o' + (uid++).toString(36);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const escHtml = esc;
const P = t => `<p>${esc(t)}</p>`;
const UL = items => '<ul>' + items.map(t => `<li>${esc(t)}</li>`).join('') + '</ul>';
export function txt(o) { return Object.assign({ id: oid(), t: 'text', x: 128, y: 100, w: 1344, h: 120, rot: 0, html: '<p></p>', fs: 32, color: null, font: null, bold: false, italic: false, underline: false, align: 'left', va: 'top', lh: 1.35, fill: '', stroke: '', sw: 0, op: 1, anim: 'none' }, o); }
export function shape(o) { return Object.assign(txt({ t: 'shape', shape: 'rect', fill: null, html: '', align: 'center', va: 'middle', fs: 28 }), o); }
export function makeSlide(layout, c, ratio) {
  c = c || {}; const H = slideH(ratio), objs = [];
  const title = (y, h, fs, extra) => txt(Object.assign({ ph: 'title', x: 128, y, w: 1344, h, fs, bold: true, html: P(c.title || T('单击添加标题')), va: 'bottom' }, extra || {}));
  if (layout === 'title') {
    objs.push(shape({ x: 128, y: H * 0.36 - 20, w: 110, h: 8, fill: null, html: '' }));
    objs.push(title(H * 0.36, 160, 76, { va: 'top' }));
    objs.push(txt({ ph: 'sub', x: 128, y: H * 0.36 + 180, w: 1344, h: 90, fs: 32, html: P(c.subtitle || T('单击添加副标题')) }));
  } else if (layout === 'content') {
    objs.push(title(70, 130, 54));
    objs.push(txt({ ph: 'body', x: 128, y: 250, w: 1344, h: H - 350, fs: 32, lh: 1.6, html: UL(c.bullets || [T('单击添加文本')]) }));
  } else if (layout === 'two') {
    objs.push(title(70, 130, 54));
    objs.push(txt({ ph: 'body', x: 128, y: 250, w: 640, h: H - 350, fs: 30, lh: 1.6, html: UL(c.left || [T('左栏要点')]) }));
    objs.push(txt({ ph: 'body', x: 832, y: 250, w: 640, h: H - 350, fs: 30, lh: 1.6, html: UL(c.right || [T('右栏要点')]) }));
  } else if (layout === 'stats') {
    objs.push(title(70, 130, 54));
    (c.stats || [{ v: '0', k: T('指标') }, { v: '0', k: T('指标') }, { v: '0', k: T('指标') }]).slice(0, 3).forEach((s, i) => {
      const x = 128 + i * 459;
      objs.push(shape({ x, y: 280, w: 427, h: 340, fill: 'card', html: '', va: 'top' }));
      objs.push(txt({ x: x + 36, y: 320, w: 355, h: 140, fs: 84, bold: true, color: 'acc', font: 'head', html: P(s.v) }));
      objs.push(txt({ ph: 'sub', x: x + 36, y: 480, w: 355, h: 100, fs: 28, html: P(s.k) }));
    });
  } else if (layout === 'titleOnly') objs.push(title(70, 130, 54));
  else if (layout === 'section') {
    objs.push(title(H / 2 - 110, 150, 68));
    objs.push(shape({ x: 128, y: H / 2 + 60, w: 1344, h: 4, fill: null, html: '' }));
  }
  return { id: oid(), objs, notes: c.notes || '', trans: 'fade', hidden: false, bg: null };
}
export const LAYOUTS = [['title', '标题幻灯片'], ['content', '标题和内容'], ['two', '两栏内容'], ['stats', '数据卡片'], ['section', '节标题'], ['titleOnly', '仅标题'], ['blank', '空白']];

export function resolveColor(v, th, fallback) { if (v === 'acc') return th.acc; if (v === 'card') return th.card; if (v === 'sub') return th.sub; if (v === 'fg') return th.fg; if (v == null) return fallback; return v; }
/** How a slide object is drawn. ptPx: slide units per point (16:9 decks are 960 pt wide), for picture borders and shadows. */
export function objView(o, th, ptPx) {
  const isShape = o.t === 'shape';
  const pic = o.t === 'image' ? pictureView(o.look, o.w, o.h, ptPx || SW / 960) : null;
  const fill = isShape ? resolveColor(o.fill, th, th.acc) : resolveColor(o.fill, th, '');
  const color = resolveColor(o.color, th, o.ph === 'sub' ? th.sub : (isShape && (o.fill == null) ? '#FFFFFF' : th.fg));
  const font = o.font === 'head' || (!o.font && o.ph === 'title') ? th.hf : (o.font || th.bf);
  let radius = '0', clip = 'none';
  if (o.shape === 'round') radius = Math.min(o.w, o.h) * 0.18 + 'px';
  if (o.shape === 'ellipse') radius = '50%';
  if (o.shape === 'triangle') clip = 'polygon(50% 0,100% 100%,0 100%)';
  if (o.shape === 'diamond') clip = 'polygon(50% 0,100% 50%,50% 100%,0 50%)';
  if (o.shape === 'arrow') clip = 'polygon(0 30%,65% 30%,65% 0,100% 50%,65% 100%,65% 70%,0 70%)';
  if (o.shape === 'pill') radius = Math.min(o.w, o.h) / 2 + 'px';
  return {
    left: o.x + 'px', top: o.y + 'px', width: o.w + 'px', height: o.h + 'px', tf: `rotate(${o.rot || 0}deg)`, op: o.op ?? 1,
    bg: fill || 'transparent', radius, clip, stroke: o.sw && o.stroke ? `inset 0 0 0 ${o.sw}px ${o.stroke}` : 'none',
    color, fs: o.fs + 'px', font: `'${font}','Noto Sans SC',sans-serif`, fw: o.bold ? 700 : 400, fst: o.italic ? 'italic' : 'normal', td: o.underline ? 'underline' : 'none',
    align: o.align || 'left', va: o.va === 'middle' ? 'center' : o.va === 'bottom' ? 'flex-end' : 'flex-start', lh: o.lh || 1.35,
    pad: o.t === 'text' ? '8px 12px' : '16px 24px', isImg: o.t === 'image', src: picSrc(o.src || ''), picFrame: pic ? pic.frame : '', picImage: pic ? pic.image : '', isTable: o.t === 'table', hasText: o.t === 'text' || o.t === 'shape',
    rows: (o.rows || []).map((r, ri) => ({ cells: r.map((c, ci) => ({ text: c, bg: ri === 0 ? th.acc : (ri % 2 ? th.card : 'transparent'), color: ri === 0 ? '#FFFFFF' : th.fg, fw: ri === 0 ? 700 : 400, ri, ci })) })),
    inner: { __html: o.html || '' }
  };
}
function slideHtml(s, th, H, scale) {
  const bg = s.bg || th.bg;
  let h = `<div class="sl" style="width:${SW}px;height:${H}px;position:relative;overflow:hidden;background:${bg};zoom:${scale}">`;
  (s.decor || []).concat(s.objs).forEach(o => {
    const v = objView(o, th);
    h += `<div style="position:absolute;left:${v.left};top:${v.top};width:${v.width};height:${v.height};transform:${v.tf};opacity:${v.op}">`;
    h += `<div style="position:absolute;inset:0;background:${v.bg};border-radius:${v.radius};clip-path:${v.clip};box-shadow:${v.stroke}"></div>`;
    if (v.isImg) h += `<div style="position:absolute;inset:0;${v.picFrame}"><img src="${esc(v.src)}" style="${v.picImage}"></div>`;
    if (v.hasText) h += `<div style="position:absolute;inset:0;padding:${v.pad};display:flex;flex-direction:column;justify-content:${v.va};color:${v.color};font-size:${v.fs};font-family:${v.font};font-weight:${v.fw};font-style:${v.fst};text-decoration:${v.td};text-align:${v.align};line-height:${v.lh}"><div class="t">${o.html || ''}</div></div>`;
    if (v.isTable) h += `<table style="width:100%;height:100%;border-collapse:collapse;font-size:${v.fs};font-family:${v.font}">` + v.rows.map(r => '<tr>' + r.cells.map(c => `<td style="background:${c.bg};color:${c.color};font-weight:${c.fw};padding:10px 16px;border-bottom:1px solid rgba(128,128,128,.25)">${esc(c.text)}</td>`).join('') + '</tr>').join('') + '</table>';
    h += '</div>';
  });
  return h + '</div>';
}

export const DOC_CSS = `h1{font-size:26px;font-weight:700;margin:18px 0 8px;line-height:1.4}h2{font-size:20px;font-weight:600;margin:16px 0 6px}h3{font-size:16px;font-weight:600;margin:14px 0 4px}p{margin:0 0 8px}blockquote{margin:8px 0;padding:4px 14px;border-left:3px solid #D1D1D6;color:#6E6E73}pre{font-family:'IBM Plex Mono',monospace;background:#F5F5F7;padding:10px 12px;font-size:13px;white-space:pre-wrap}ul,ol{margin:0 0 8px;padding-left:1.6em}img{max-width:100%}a{color:#2F5D8A}ins{color:#3F7D5C;text-decoration:underline}del{color:#B5563A}[data-cid]{background:#F3E6C4}hr[data-pb]{border:none;border-top:1px dashed #C7C7CC;margin:24px 0;break-after:page}`;
const PAGES = { A4: ['210mm', '297mm'], Letter: ['8.5in', '11in'], A5: ['148mm', '210mm'] };
const MARG = { narrow: '12.7mm', normal: '25.4mm', wide: '38mm' };
/** The Noto aliases of assets/fonts/fonts.css (local() fonts only, nothing to fetch), for the pages the app writes out: prints and HTML exports. */
export const fontFaces = () => [...document.styleSheets].filter(s => /\/fonts\.css$/.test(s.href)).flatMap(s => [...s.cssRules].map(r => r.cssText)).filter(t => t.includes('local(')).join('');
export function printDoc(doc, ctx) {
  let css = '', body = '';
  const fonts = `<style>${fontFaces()}</style>`;
  if (doc.type === 'docx') {
    const pg = doc.page || {}; const sz = PAGES[pg.size || 'A4']; const [w, h] = pg.orient === 'landscape' ? [sz[1], sz[0]] : sz;
    css = `@page{size:${w} ${h};margin:${MARG[pg.margin || 'normal']}}body{margin:0;font-family:'Noto Serif SC',serif;font-size:11pt;line-height:1.8;color:#1D1D1F}.hd,.ft{font-size:9pt;color:#8E8E93}.ed{column-count:${pg.cols || 1};column-gap:32px}` + DOC_CSS;
    body = (doc.header ? `<div class="hd">${doc.header}</div>` : '') + `<div class="ed">${doc.html}</div>` + (doc.footer ? `<div class="ft">${doc.footer}</div>` : ''); // header html as the engine gives it
  } else if (doc.type === 'xlsx') {
    const E = ctx.E, calc = new E.Calc(doc);
    css = `@page{size:A4 landscape;margin:12mm}body{font-family:'Noto Sans SC',sans-serif;font-size:10pt}table{border-collapse:collapse;margin-bottom:24px}td{border:1px solid #D1D1D6;padding:4px 8px;white-space:nowrap}h2{font-size:12pt}`;
    doc.sheets.forEach((sh, si) => {
      const u = E.usedRange(sh); if (!u) return;
      body += `<h2>${esc(sh.name)}</h2><table>`;
      for (let r = u.r1; r <= u.r2; r++) {
        body += '<tr>';
        for (let c = u.c1; c <= u.c2; c++) { const cell = sh.cells[E.A(r, c)], s = (cell && cell.s) || {}, v = calc.value(si, r, c); body += `<td style="${s.b ? 'font-weight:700;' : ''}${s.fill ? 'background:' + s.fill + ';' : ''}${s.color ? 'color:' + s.color + ';' : ''}text-align:${s.align || (typeof v === 'number' ? 'right' : 'left')}">${esc(E.fmt(v, s))}</td>`; }
        body += '</tr>';
      }
      body += '</table>';
    });
  } else {
    const th = THEMES[doc.theme] || THEMES.paper, H = slideH(doc.ratio);
    css = `@page{size:${SW * 0.6}px ${H * 0.6}px;margin:0}body{margin:0}.sl{break-after:page}.t p{margin:0}.t ul,.t ol{margin:0;padding-left:1.2em}`;
    body = doc.slides.filter(s => !s.hidden).map(s => slideHtml(s, th, H, 0.6)).join('');
  }
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${esc(doc.title)}</title>${fonts}<style>${css}</style></head><body>${body}</body></html>`;
  const f = document.createElement('iframe');
  f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.appendChild(f);
  f.contentDocument.open(); f.contentDocument.write(html); f.contentDocument.close();
  setTimeout(() => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { } setTimeout(() => f.remove(), 2000); }, 700);
}

let JSZ = null;
async function zipLib() { if (!JSZ) JSZ = (await import('https://esm.sh/jszip@3.10.1')).default; return JSZ; }
const X = s => new DOMParser().parseFromString(s, 'application/xml');
const kids = (el, name) => Array.from(el ? el.children : []).filter(c => c.localName === name);
const kid = (el, name) => kids(el, name)[0];
const desc = (el, name) => el ? Array.from(el.getElementsByTagName('*')).filter(c => c.localName === name) : [];
const at = (el, name) => el ? (el.getAttribute(name) || el.getAttribute(name.split(':').pop())) : null;
async function rels(z, path) {
  const dir = path.slice(0, path.lastIndexOf('/') + 1), file = path.slice(path.lastIndexOf('/') + 1);
  const f = z.file(dir + '_rels/' + file + '.rels'); const m = {};
  if (!f) return m;
  desc(X(await f.async('string')), 'Relationship').forEach(r => { let t = at(r, 'Target'); if (!t.startsWith('/')) { const parts = (dir + t).split('/'); const out = []; parts.forEach(p => { if (p === '..') out.pop(); else if (p !== '.') out.push(p); }); t = out.join('/'); } else t = t.slice(1); m[at(r, 'Id')] = t; });
  return m;
}
async function dataUrl(z, path) { const f = z.file(path); if (!f) return ''; const ext = path.split('.').pop().toLowerCase(); const b64 = await f.async('base64'); return `data:image/${ext === 'jpg' ? 'jpeg' : ext === 'svg' ? 'svg+xml' : ext};base64,${b64}`; }
function csvCells(t, sep) {
  const cells = {}; const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < t.length; i++) { const ch = t[i]; if (q) { if (ch === '"' && t[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; } else if (ch === '"') q = true; else if (ch === sep) { row.push(cur); cur = ''; } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; } else cur += ch; }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  rows.forEach((r, ri) => r.forEach((v, ci) => { if (v !== '') cells[colN(ci) + (ri + 1)] = { v, s: ri === 0 ? { b: true } : undefined }; }));
  return cells;
}
const colN = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };

async function docxHtml(z) {
  const doc = X(await z.file('word/document.xml').async('string'));
  const r = await rels(z, 'word/document.xml');
  const smap = {};
  const sf = z.file('word/styles.xml');
  if (sf) desc(X(await sf.async('string')), 'style').forEach(s => { const n = kid(s, 'name'); smap[at(s, 'w:styleId')] = (n ? at(n, 'w:val') : '').toLowerCase(); });
  const body = desc(doc, 'body')[0];
  let out = '', list = null;
  const flushList = () => { if (list) { out += `<${list.t}>${list.items.join('')}</${list.t}>`; list = null; } };
  const runs = async p => {
    let h = '';
    for (const node of Array.from(p.children)) {
      if (node.localName === 'hyperlink') { h += `<a href="${r[at(node, 'r:id')] || '#'}">${await runs(node)}</a>`; continue; }
      if (node.localName !== 'r') continue;
      const pr = kid(node, 'rPr'); const st = [];
      let o = '', c = '';
      if (pr) {
        if (kid(pr, 'b') && at(kid(pr, 'b'), 'w:val') !== '0') { o += '<b>'; c = '</b>' + c; }
        if (kid(pr, 'i') && at(kid(pr, 'i'), 'w:val') !== '0') { o += '<i>'; c = '</i>' + c; }
        if (kid(pr, 'u') && at(kid(pr, 'u'), 'w:val') !== 'none') { o += '<u>'; c = '</u>' + c; }
        if (kid(pr, 'strike')) { o += '<s>'; c = '</s>' + c; }
        const col = kid(pr, 'color'); if (col && at(col, 'w:val') && at(col, 'w:val') !== 'auto') st.push('color:#' + at(col, 'w:val'));
        const sz = kid(pr, 'sz'); if (sz) st.push('font-size:' + (+at(sz, 'w:val') / 2) + 'pt');
        const hl = kid(pr, 'highlight'); if (hl) st.push('background-color:' + at(hl, 'w:val'));
      }
      let t = '';
      for (const ch of Array.from(node.children)) {
        if (ch.localName === 't') t += esc(ch.textContent);
        else if (ch.localName === 'tab') t += '&emsp;';
        else if (ch.localName === 'br') t += '<br>';
        else if (ch.localName === 'drawing') { const b = desc(ch, 'blip')[0]; const id = b && at(b, 'r:embed'); if (id && r[id]) t += `<img src="${await dataUrl(z, r[id])}" style="max-width:100%">`; }
      }
      if (t) h += (st.length ? `<span style="${st.join(';')}">` : '') + o + t + c + (st.length ? '</span>' : '');
    }
    return h;
  };
  for (const el of Array.from(body.children)) {
    if (el.localName === 'p') {
      const pPr = kid(el, 'pPr'); const ps = pPr && kid(pPr, 'pStyle'); const sname = ps ? (smap[at(ps, 'w:val')] || at(ps, 'w:val').toLowerCase()) : '';
      const jc = pPr && kid(pPr, 'jc'); const al = jc ? { center: 'center', right: 'right', both: 'justify', end: 'right' }[at(jc, 'w:val')] : null;
      const style = al ? ` style="text-align:${al}"` : '';
      const inner = (await runs(el)) || '<br>';
      if (pPr && kid(pPr, 'numPr')) { if (!list) list = { t: /number|decimal|编号/.test(sname) ? 'ol' : 'ul', items: [] }; list.items.push(`<li${style}>${inner}</li>`); continue; }
      flushList();
      let tag = 'p';
      if (/^title|标题$/.test(sname) || sname === 'title') tag = 'h1';
      else if (/heading 1|标题 1/.test(sname) || sname === '1') tag = 'h1';
      else if (/heading 2|标题 2/.test(sname) || sname === '2') tag = 'h2';
      else if (/heading [3-9]|标题 [3-9]/.test(sname) || sname === '3') tag = 'h3';
      else if (/quote|引用/.test(sname)) tag = 'blockquote';
      out += `<${tag}${style}>${inner}</${tag}>`;
    } else if (el.localName === 'tbl') {
      flushList();
      out += '<table style="border-collapse:collapse;width:100%;margin:8px 0"><tbody>';
      for (const tr of kids(el, 'tr')) { out += '<tr>'; for (const tc of kids(tr, 'tc')) { const ps = kids(tc, 'p'); let t = ''; for (const p of ps) t += (t ? '<br>' : '') + await runs(p); const gs = desc(tc, 'gridSpan')[0]; out += `<td${gs ? ` colspan="${at(gs, 'w:val')}"` : ''} style="border:1px solid #C7C7CC;padding:6px 8px">${t || '&nbsp;'}</td>`; } out += '</tr>'; }
      out += '</tbody></table>';
    }
  }
  flushList();
  return out || '<p><br></p>';
}
async function xlsxSheets(z) {
  const wb = X(await z.file('xl/workbook.xml').async('string'));
  const r = await rels(z, 'xl/workbook.xml');
  const ss = [];
  const sf = z.file('xl/sharedStrings.xml');
  if (sf) desc(X(await sf.async('string')), 'si').forEach(si => ss.push(desc(si, 't').map(t => t.textContent).join('')));
  const sheets = [];
  for (const s of desc(wb, 'sheet')) {
    const path = r[at(s, 'r:id')]; const f = path && z.file(path); if (!f) continue;
    const x = X(await f.async('string')); const cells = {}; const colW = {};
    desc(x, 'col').forEach(c => { const w = Math.round(+at(c, 'width') * 7 + 5); for (let i = +at(c, 'min'); i <= Math.min(+at(c, 'max'), 40); i++) colW[colN(i - 1)] = w; });
    desc(x, 'c').forEach(c => {
      const ref = at(c, 'r'), t = at(c, 't'), fe = kid(c, 'f'), ve = kid(c, 'v');
      let v = '';
      if (fe && fe.textContent) v = '=' + fe.textContent;
      else if (t === 's') v = ss[+(ve && ve.textContent)] || '';
      else if (t === 'inlineStr') v = desc(c, 't').map(n => n.textContent).join('');
      else if (t === 'b') v = ve && ve.textContent === '1' ? 'TRUE' : 'FALSE';
      else v = ve ? ve.textContent : '';
      if (v !== '') cells[ref] = { v };
    });
    const merges = desc(x, 'mergeCell').map(m => { const [a, b] = at(m, 'ref').split(':'); const pa = /([A-Z]+)(\d+)/.exec(a), pb = /([A-Z]+)(\d+)/.exec(b); const ci = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }; return { r: +pa[2] - 1, c: ci(pa[1]), rs: +pb[2] - +pa[2] + 1, cs: ci(pb[1]) - ci(pa[1]) + 1 }; });
    sheets.push({ name: at(s, 'name'), cells, colW, merges });
  }
  return sheets.length ? sheets : [{ name: 'Sheet1', cells: {} }];
}
async function pptxSlides(z) {
  const pres = X(await z.file('ppt/presentation.xml').async('string'));
  const r = await rels(z, 'ppt/presentation.xml');
  const sz = desc(pres, 'sldSz')[0]; const cx = +at(sz, 'cx') || 12192000, cy = +at(sz, 'cy') || 6858000;
  const k = SW / cx, ratio = cy / cx > 0.7 ? '4:3' : '16:9', H = slideH(ratio), ky = H / cy;
  const slides = [];
  for (const sid of desc(pres, 'sldId')) {
    const path = r[at(sid, 'r:id')]; const f = path && z.file(path); if (!f) continue;
    const x = X(await f.async('string')); const sr = await rels(z, path); const objs = [];
    const bgc = desc(kid(desc(x, 'bg')[0], 'bgPr'), 'srgbClr')[0];
    const pos = (el, fb) => { const off = desc(el, 'off')[0], ext = desc(el, 'ext')[0]; if (!off || !ext) return fb; return { x: Math.round(+at(off, 'x') * k), y: Math.round(+at(off, 'y') * ky), w: Math.round(+at(ext, 'cx') * k), h: Math.round(+at(ext, 'cy') * ky) }; };
    for (const sp of desc(x, 'sp')) {
      const ph = desc(sp, 'ph')[0], phT = ph ? (at(ph, 'type') || 'body') : null;
      const fb = phT === 'title' || phT === 'ctrTitle' ? { x: 128, y: 70, w: 1344, h: 150 } : { x: 128, y: 250, w: 1344, h: H - 350 };
      const p = pos(desc(sp, 'spPr')[0], fb);
      const spPr = desc(sp, 'spPr')[0]; const geom = desc(spPr, 'prstGeom')[0]; const fillC = desc(kid(spPr, 'solidFill'), 'srgbClr')[0];
      let fsz = null, html = '', bold = false;
      for (const ap of desc(sp, 'p').filter(n => n.namespaceURI && n.namespaceURI.includes('drawingml'))) {
        let t = ''; for (const run of kids(ap, 'r')) { const rp = kid(run, 'rPr'); if (rp && at(rp, 'sz')) fsz = fsz || +at(rp, 'sz') / 100; if (rp && at(rp, 'b') === '1') bold = true; const col = desc(rp, 'srgbClr')[0]; const tt = esc(desc(run, 't').map(n => n.textContent).join('')); t += col ? `<span style="color:#${at(col, 'val')}">${tt}</span>` : tt; }
        const bu = desc(kid(ap, 'pPr'), 'buChar')[0];
        html += bu ? `<ul><li>${t}</li></ul>` : `<p>${t || '<br>'}</p>`;
      }
      html = html.replace(/<\/ul><ul>/g, '');
      const isTitle = phT === 'title' || phT === 'ctrTitle';
      const base = { x: p.x, y: p.y, w: p.w, h: p.h, html, fs: Math.round((fsz || (isTitle ? 40 : 20)) * SW / 960), bold: bold || isTitle, ph: isTitle ? 'title' : phT === 'subTitle' ? 'sub' : null };
      const prst = geom ? at(geom, 'prst') : 'rect';
      if (fillC || (geom && prst !== 'rect')) objs.push(shape(Object.assign(base, { fill: fillC ? '#' + at(fillC, 'val') : null, shape: { roundRect: 'round', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'arrow' }[prst] || 'rect', align: 'center', va: 'middle' })));
      else objs.push(txt(base));
    }
    for (const pic of desc(x, 'pic')) { const b = desc(pic, 'blip')[0]; const id = b && at(b, 'r:embed'); const p = pos(desc(pic, 'spPr')[0], { x: 400, y: 200, w: 800, h: 450 }); if (id && sr[id]) objs.push(txt({ t: 'image', src: await dataUrl(z, sr[id]), x: p.x, y: p.y, w: p.w, h: p.h, html: '' })); }
    for (const gf of desc(x, 'graphicFrame')) { const tbl = desc(gf, 'tbl')[0]; if (!tbl) continue; const p = pos(desc(gf, 'xfrm')[0], { x: 128, y: 250, w: 1344, h: 400 }); objs.push(txt({ t: 'table', x: p.x, y: p.y, w: p.w, h: p.h, fs: 24, html: '', rows: kids(tbl, 'tr').map(tr => kids(tr, 'tc').map(tc => desc(tc, 't').map(n => n.textContent).join(''))) })); }
    slides.push({ id: oid(), objs, notes: '', trans: 'fade', hidden: false, bg: bgc ? '#' + at(bgc, 'val') : '#FFFFFF' });
  }
  return { theme: 'paper', ratio, slides: slides.length ? slides : [makeSlide('title', {}, ratio)] };
}
export async function importFile(file) {
  const name = file.name, ext = name.split('.').pop().toLowerCase(), title = name.replace(/\.[^.]+$/, '');
  if (ext === 'txt' || ext === 'md') { const t = await file.text(); return { type: 'docx', title, html: t.split(/\r?\n/).map(l => l.startsWith('### ') ? `<h3>${esc(l.slice(4))}</h3>` : l.startsWith('## ') ? `<h2>${esc(l.slice(3))}</h2>` : l.startsWith('# ') ? `<h1>${esc(l.slice(2))}</h1>` : /^[-*] /.test(l) ? `<ul><li>${esc(l.slice(2))}</li></ul>` : `<p>${esc(l) || '<br>'}</p>`).join('').replace(/<\/ul><ul>/g, '') }; }
  if (ext === 'html' || ext === 'htm') { const d = new DOMParser().parseFromString(await file.text(), 'text/html'); d.querySelectorAll('script,style').forEach(n => n.remove()); return { type: 'docx', title, html: d.body.innerHTML }; }
  if (ext === 'csv' || ext === 'tsv') return { type: 'xlsx', title, sheets: [{ name: 'Sheet1', cells: csvCells(await file.text(), ext === 'tsv' ? '\t' : ',') }] };
  const Z = await zipLib(); const z = await Z.loadAsync(file);
  if (ext === 'docx') return { type: 'docx', title, html: await docxHtml(z) };
  if (ext === 'xlsx') return { type: 'xlsx', title, sheets: await xlsxSheets(z) };
  if (ext === 'pptx') return Object.assign({ type: 'pptx', title }, await pptxSlides(z));
  throw new Error(T('暂不支持 .{ext} 文件', { ext }));
}

const xe = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function rgbHex(c) { if (!c) return null; if (c[0] === '#') return c.slice(1, 7).toUpperCase(); const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); return m ? [m[1], m[2], m[3]].map(n => (+n).toString(16).padStart(2, '0')).join('').toUpperCase() : null; }
function docxBody(html) {
  const root = new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html').body.firstChild;
  const runsOf = (node, f) => {
    let out = '';
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { if (!n.nodeValue) return; let pr = ''; if (f.b) pr += '<w:b/>'; if (f.i) pr += '<w:i/>'; if (f.u) pr += '<w:u w:val="single"/>'; if (f.s) pr += '<w:strike/>'; if (f.color) pr += `<w:color w:val="${f.color}"/>`; if (f.sz) pr += `<w:sz w:val="${Math.round(f.sz * 2)}"/>`; if (f.hl) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${f.hl}"/>`; if (f.va) pr += `<w:vertAlign w:val="${f.va}"/>`; out += `<w:r>${pr ? '<w:rPr>' + pr + '</w:rPr>' : ''}<w:t xml:space="preserve">${xe(n.nodeValue.replace(/\u200B/g, ''))}</w:t></w:r>`; return; }
      if (n.nodeType !== 1) return;
      const t = n.tagName.toLowerCase(); if (t === 'br') { out += '<w:r><w:br/></w:r>'; return; }
      if (t === 'del') return;
      const g = Object.assign({}, f), st = n.style;
      if (t === 'b' || t === 'strong' || st.fontWeight === 'bold' || +st.fontWeight >= 600) g.b = true;
      if (t === 'i' || t === 'em' || st.fontStyle === 'italic') g.i = true;
      if (t === 'u' || t === 'ins' || (st.textDecoration || '').includes('underline')) g.u = true;
      if (t === 's' || t === 'strike' || (st.textDecoration || '').includes('line-through')) g.s = true;
      if (t === 'sup') g.va = 'superscript'; if (t === 'sub') g.va = 'subscript';
      if (st.color) g.color = rgbHex(st.color); if (st.backgroundColor) g.hl = rgbHex(st.backgroundColor);
      if (st.fontSize) { const m = /([\d.]+)(pt|px)/.exec(st.fontSize); if (m) g.sz = m[2] === 'pt' ? +m[1] : +m[1] * 0.75; }
      out += runsOf(n, g);
    });
    return out;
  };
  const para = (el, style, extra) => { const al = { center: 'center', right: 'right', justify: 'both' }[el.style && el.style.textAlign]; let pPr = ''; if (style) pPr += `<w:pStyle w:val="${style}"/>`; if (al) pPr += `<w:jc w:val="${al}"/>`; return `<w:p>${pPr ? '<w:pPr>' + pPr + '</w:pPr>' : ''}${extra || ''}${runsOf(el, {})}</w:p>`; };
  let out = '';
  root.childNodes.forEach(n => {
    if (n.nodeType === 3) { if (n.nodeValue.trim()) out += `<w:p><w:r><w:t xml:space="preserve">${xe(n.nodeValue)}</w:t></w:r></w:p>`; return; }
    if (n.nodeType !== 1) return;
    const t = n.tagName.toLowerCase();
    if (t === 'h1') out += para(n, 'Heading1'); else if (t === 'h2') out += para(n, 'Heading2'); else if (t === 'h3') out += para(n, 'Heading3');
    else if (t === 'ul' || t === 'ol') Array.from(n.children).forEach((li, i) => { out += para(li, null, `<w:r><w:t xml:space="preserve">${t === 'ol' ? (i + 1) + '. ' : '• '}</w:t></w:r>`); });
    else if (t === 'table') { out += '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'; n.querySelectorAll('tr').forEach(tr => { out += '<w:tr>'; tr.querySelectorAll('td,th').forEach(td => { out += `<w:tc><w:tcPr>${td.colSpan > 1 ? `<w:gridSpan w:val="${td.colSpan}"/>` : ''}</w:tcPr>${para(td)}</w:tc>`; }); out += '</w:tr>'; }); out += '</w:tbl>'; }
    else if (t === 'hr' && n.hasAttribute('data-pb')) out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    else if (t === 'hr') out += '<w:p/>';
    else out += para(n, t === 'blockquote' ? 'Quote' : null);
  });
  return out;
}
function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); }
const CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export async function exportDocx(doc) {
  const Z = await zipLib(); const z = new Z();
  const pg = doc.page || {}; const sizes = { A4: [11906, 16838], Letter: [12240, 15840], A5: [8391, 11906] }; let [w, h] = sizes[pg.size || 'A4']; if (pg.orient === 'landscape') [w, h] = [h, w];
  const m = { narrow: 720, normal: 1440, wide: 2160 }[pg.margin || 'normal'];
  z.file('[Content_Types].xml', CT + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file('_rels/.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  const hs = (id, name, sz) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${id.slice(-1) - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
  z.file('word/styles.xml', CT + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:eastAsia="SimSun" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${hs('Heading1', 'heading 1', 40)}${hs('Heading2', 'heading 2', 32)}${hs('Heading3', 'heading 3', 26)}<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="5C5850"/></w:rPr></w:style></w:styles>`);
  z.file('word/document.xml', CT + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${docxBody(doc.html || '')}<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${pg.orient === 'landscape' ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${m}" w:right="${m}" w:bottom="${m}" w:left="${m}" w:header="708" w:footer="708" w:gutter="0"/>${(pg.cols || 1) > 1 ? `<w:cols w:num="${pg.cols}" w:space="720"/>` : ''}</w:sectPr></w:body></w:document>`);
  download(await z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), doc.title + '.docx');
}
export async function exportXlsx(doc, E) {
  const Z = await zipLib(); const z = new Z(); const calc = new E.Calc(doc);
  const n = doc.sheets.length;
  z.file('[Content_Types].xml', CT + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + doc.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '</Types>');
  z.file('_rels/.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  z.file('xl/_rels/workbook.xml.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + doc.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  z.file('xl/workbook.xml', CT + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + doc.sheets.map((s, i) => `<sheet name="${xe(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>');
  z.file('xl/styles.xml', CT + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;¥&quot;#,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="6"><xf/><xf fontId="1" applyFont="1"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="164" fontId="1" applyNumberFormat="1" applyFont="1"/><xf numFmtId="9" applyNumberFormat="1"/><xf numFmtId="9" fontId="1" applyNumberFormat="1" applyFont="1"/></cellXfs></styleSheet>');
  doc.sheets.forEach((sh, si) => {
    const rows = {};
    Object.keys(sh.cells).forEach(a => { const p = E.parseA(a); if (p) (rows[p.r] = rows[p.r] || []).push([p.c, a]); });
    let data = '';
    Object.keys(rows).map(Number).sort((a, b) => a - b).forEach(r => {
      data += `<row r="${r + 1}">`;
      rows[r].sort((a, b) => a[0] - b[0]).forEach(([c, a]) => {
        const cell = sh.cells[a], s = cell.s || {}; const raw = cell.v == null ? '' : String(cell.v); if (raw === '') return;
        const st = (s.fmt === 'money' ? 2 : s.fmt === 'pct' ? 4 : 0) + (s.b ? 1 : 0); const sa = st ? ` s="${st}"` : '';
        if (raw[0] === '=') { const v = calc.value(si, r, c); const vv = typeof v === 'number' ? `<v>${v}</v>` : ''; data += `<c r="${a}"${sa}${typeof v === 'string' ? ' t="str"' : ''}><f>${xe(raw.slice(1))}</f>${vv || (typeof v === 'string' ? `<v>${xe(v)}</v>` : '')}</c>`; }
        else if (raw.trim() !== '' && !isNaN(Number(raw)) && s.fmt !== 'text') data += `<c r="${a}"${sa}><v>${Number(raw)}</v></c>`;
        else data += `<c r="${a}"${sa} t="inlineStr"><is><t xml:space="preserve">${xe(raw)}</t></is></c>`;
      });
      data += '</row>';
    });
    const cols = Object.keys(sh.colW || {}).map(k => { const i = E.colIdx(k) + 1; return `<col min="${i}" max="${i}" width="${((sh.colW[k] - 5) / 7).toFixed(2)}" customWidth="1"/>`; }).join('');
    const merges = (sh.merges || []).length ? `<mergeCells count="${sh.merges.length}">` + sh.merges.map(m => `<mergeCell ref="${E.A(m.r, m.c)}:${E.A(m.r + m.rs - 1, m.c + m.cs - 1)}"/>`).join('') + '</mergeCells>' : '';
    z.file(`xl/worksheets/sheet${si + 1}.xml`, CT + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols ? '<cols>' + cols + '</cols>' : ''}<sheetData>${data}</sheetData>${merges}</worksheet>`);
  });
  download(await z.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), doc.title + '.xlsx');
}
export function exportCsv(doc, E) {
  const sh = doc.sheets[doc.active || 0], u = E.usedRange(sh), calc = new E.Calc(doc); let out = '';
  if (u) for (let r = 0; r <= u.r2; r++) { const row = []; for (let c = 0; c <= u.c2; c++) { const v = E.fmt(calc.value(doc.active || 0, r, c), (sh.cells[E.A(r, c)] || {}).s); row.push(/[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v); } out += row.join(',') + '\n'; }
  download(new Blob(['\ufeff' + out], { type: 'text/csv' }), doc.title + '.csv');
}
