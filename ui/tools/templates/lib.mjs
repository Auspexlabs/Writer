// What the template sources (docx.mjs, xlsx.mjs, pptx.mjs, md.mjs, mm.mjs) are written with. A source exports its gallery
// categories, cats = [[zh, en], …] in display order, and its templates, default = [{ id, cat, name: [zh, en], build(b, t) }].
// build gets b, the builder for its format bound to the file being made, and t(zh, en), which picks the language being built:
// one source makes both files. Every builder call becomes the engine's own commands (add / set), run in order.
// A Markdown template's build returns the file's text instead (b is null).

/** Text for the inline HTML that paragraphs, cells and shapes take. */
export const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
/** A run with a look of its own: span('合计', { color: '1F3A5F', size: 10, bold: true }). size is in points. */
export function span(text, o = {}) {
  const css = [o.color && 'color:#' + o.color, o.size && 'font-size:' + o.size + 'pt', o.bg && 'background-color:#' + o.bg, o.font && 'font-family:' + o.font].filter(Boolean).join(';');
  let h = esc(text);
  if (css) h = `<span style="${css}">${h}</span>`;
  if (o.bold) h = `<b>${h}</b>`;
  if (o.italic) h = `<i>${h}</i>`;
  return h;
}
export const bold = text => `<b>${esc(text)}</b>`;
/** Lines of inline HTML in one paragraph or cell. */
export const lines = (...ls) => ls.filter(l => l != null).join('<br>');

/** Colours the templates share: an accent, a soft fill of it, a rule and a secondary text colour per palette. */
export const INK = { acc: '1D1D1F', soft: 'F2F2F4', line: 'D2D2D7', sub: '6E6E73' };
export const NAVY = { acc: '1F3A5F', soft: 'EEF2F7', line: 'C9D3E0', sub: '5E6B7A' };
export const CLAY = { acc: 'A4492F', soft: 'F6EEE8', line: 'E4D2C5', sub: '7A5F50' };
export const SLATE = { acc: '2F5D8A', soft: 'EAF1F8', line: 'C5D6E8', sub: '5B6B7C' };
export const PLUM = { acc: '4B3B63', soft: 'F1EEF5', line: 'D7D0E1', sub: '6D6478' };
export const OCHRE = { acc: '8A6A2F', soft: 'F7F2E8', line: 'E3D7C0', sub: '7A6E5A' };

const P = props => Object.entries(props).filter(([, v]) => v != null && v !== '').flatMap(([k, v]) => ['--prop', k + '=' + (typeof v === 'object' ? JSON.stringify(v) : v)]);
const plain = h => String(h).replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const cm = v => typeof v === 'number' ? Math.round(v * 1000) / 1000 + 'cm' : v;

// ---------- Word ----------
/** Paragraph content is inline HTML (see span); page(props) takes the document's page setup (page, margin, header, footer…).
 *  The same calls bound to a table cell come from into(cellPath): a two-column résumé or a form writes paragraphs into cells. */
export function docx(run, file) {
  const add = (type, props, parent) => run(['add', file, parent, '--type', type, ...P(props)]);
  const set = (path, props) => run(['set', file, path, ...P(props)]);
  const api = parent => {
    const para = (html, props) => add('paragraph', Object.assign(html ? { html } : { text: '' }, props), parent);
    return {
      title: (html, props) => para(html, Object.assign({ style: 'Title' }, props)),
      h1: (html, props) => add('heading', Object.assign({ level: 1, html }, props), parent),
      h2: (html, props) => add('heading', Object.assign({ level: 2, html }, props), parent),
      h3: (html, props) => add('heading', Object.assign({ level: 3, html }, props), parent),
      p: para,
      /** A thin bar across the text width in colour c: the rule under a letterhead or a résumé section. */
      rule: (c, size = 1) => para(span(' ', { size, color: c }), { fill: c }),
      /** One list paragraph per item; props.list = number for a numbered list. */
      list: async (items, props) => { for (const html of items) await para(html, Object.assign({ list: 'bullet' }, props)); },
      table: (rows, opts) => docxTable(add, set, rows, opts || {}, parent),
      pagebreak: () => add('pagebreak', {}, parent),
      toc: props => add('toc', props, parent),
    };
  };
  return Object.assign({ file, run, set, page: props => set('/', props), into: api }, api('/body'));
}

/** rows: the full grid, row by row. A cell is inline HTML, or { html, fill, align, valign, borders, colspan, rowspan }, or null where
 *  a merged cell covers it. opts: the table's own props (style, width, widths, borders, borderColor, align) and heights: [row height…]. */
async function docxTable(add, set, rows, o, parent) {
  const cellOf = c => c == null ? {} : typeof c === 'object' ? c : { html: String(c) };
  const t = await add('table', { data: rows.map(r => r.map(c => plain(cellOf(c).html || ''))), style: o.style, width: o.width, widths: o.widths, borders: o.borders, borderColor: o.borderColor, align: o.align }, parent);
  const at = (r, c) => `${t.path}/row[${r + 1}]/cell[${c + 1}]`;
  for (const [r, h] of (o.heights || []).entries()) if (h) await set(`${t.path}/row[${r + 1}]`, { height: cm(h) });
  for (let r = 0; r < rows.length; r++) for (let c = 0; c < rows[r].length; c++) {
    const x = cellOf(rows[r][c]), props = {};
    if (x.html && /[<&]/.test(x.html)) props.html = x.html;
    for (const k of ['fill', 'align', 'valign', 'borders']) if (x[k] != null) props[k] = x[k];
    if (Object.keys(props).length) await set(at(r, c), props);
  }
  // merges from the last row up, right to left: the cells still to merge keep their places in the grid
  for (let r = rows.length - 1; r >= 0; r--) for (let c = rows[r].length - 1; c >= 0; c--) {
    const x = cellOf(rows[r][c]);
    if ((x.colspan || 1) > 1 || (x.rowspan || 1) > 1) await set(at(r, c), { colspan: x.colspan, rowspan: x.rowspan });
  }
  return t.path;
}

// ---------- Excel ----------
const colIdx = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const colName = c => { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; };
export { colName };
/** Works on one sheet at a time: sheet(name) names the first, then adds the next. values() takes rows where a string starting
 *  with = is a formula. Refs are A1 or A1:C3; style() takes the cell look (bold, fill, color, format, align, border…);
 *  merge() adds merged ranges to the sheet's list; layout() takes the other sheet props (widths, heights, freeze, filter). */
export function xlsx(run, file) {
  let at = 0;
  const merges = [];
  const sp = () => `/sheet[${at}]`;
  const set = (path, props) => run(['set', file, path, ...P(props)]);
  const where = ref => ref.includes(':') ? `${sp()}/range[${ref}]` : `${sp()}/cell[${ref}]`;
  return {
    file, run,
    sheet: async name => { if (at) await run(['add', file, '/', '--type', 'sheet', ...P({ name })]); else await set('/sheet[1]', { name }); at++; },
    values: async (ref, rows) => {
      const m = /^([A-Z]+)(\d+)/.exec(ref), c0 = colIdx(m[1]), r0 = +m[2];
      const isF = v => typeof v === 'string' && v.startsWith('=');
      await set(`${sp()}/range[${ref}]`, { values: rows.map(r => r.map(v => v == null || isF(v) ? '' : v)) });
      for (const [i, r] of rows.entries()) for (const [j, v] of r.entries()) if (isF(v)) await set(`${sp()}/cell[${colName(c0 + j)}${r0 + i}]`, { formula: v.slice(1) });
    },
    cell: (ref, props) => set(where(ref), typeof props === 'object' ? props : { value: props }),
    fx: (ref, formula, props) => set(where(ref), Object.assign({ formula }, props)),
    style: (ref, props) => set(where(ref), props),
    layout: props => set(sp(), props),
    merge: (...refs) => set(sp(), { merges: merges[at] = (merges[at] || []).concat(refs) }),
    chart: props => run(['add', file, sp(), '--type', 'chart', ...P(Object.assign({}, props, { x: cm(props.x), y: cm(props.y), w: cm(props.w), h: cm(props.h) }))]),
  };
}

// ---------- PowerPoint (16:9, 33.867 × 19.05 cm) ----------
export const SW = 33.867, SH = 19.05;
/** slide(background) adds a blank slide and returns its drawing calls; lengths are in cm. */
export function pptx(run, file) {
  let n = 0;
  return {
    file, run,
    slide: async (bg = 'FFFFFF', props) => {
      await run(['add', file, '/', '--type', 'slide', ...P(Object.assign({ layout: 'Blank', background: bg }, props))]);
      return slideApi(run, file, `/slide[${++n}]`);
    },
  };
}
function slideApi(run, file, sp) {
  const add = (type, props, parent = sp) => run(['add', file, parent, '--type', type, ...P(props)]);
  const set = (path, props) => run(['set', file, path, ...P(props)]);
  const box = (x, y, w, h) => ({ x: cm(x), y: cm(y), w: cm(w), h: cm(h) });
  return {
    path: sp,
    notes: text => set(sp, { notes: text }),
    /** A drawn shape with no text: a band, a rule, a card, a dot. o.geometry: rect (default) | roundRect | ellipse | triangle | diamond. */
    shape: (x, y, w, h, fill, o = {}) => add('shape', Object.assign({ geometry: o.geometry || 'rect', fill, line: o.line || 'none' }, box(x, y, w, h))),
    /** Text: one paragraph per item, each inline HTML (bold as <b>) or { html, list, align }. o: size (pt), color and font for all
     *  of it, align, and fill (+ geometry) for text on a shape of its own, which the app centres both ways. */
    text: async (x, y, w, h, content, o = {}) => {
      const items = (Array.isArray(content) ? content : [content]).map(p => typeof p === 'object' ? p : { html: p });
      const s = await add('shape', Object.assign({ geometry: o.fill ? o.geometry || 'rect' : undefined, html: items[0].html || ' ', fill: o.fill || 'none', line: o.line || 'none' }, box(x, y, w, h)));
      for (const [i, p] of items.entries()) {
        if (i) await add('paragraph', { html: p.html || ' ', list: p.list, align: p.align || o.align }, s.path);
        else if (p.list || p.align || o.align) await set(`${s.path}/paragraph[1]`, { list: p.list, align: p.align || o.align });
      }
      if (o.size || o.color || o.font) await set(s.path, { size: o.size ? o.size + 'pt' : undefined, color: o.color, font: o.font });
      return s.path;
    },
    /** A table: rows of text; o.header is the first row's fill (its text white and bold), o.size the text size in points,
     *  o.zebra a fill for every other body row, o.x/y/w/h its frame. */
    table: async (rows, o) => {
      const t = await add('table', Object.assign({ data: rows }, box(o.x, o.y, o.w, o.h)));
      for (let r = 0; r < rows.length; r++) for (let c = 0; c < rows[r].length; c++) {
        const head = r === 0 && o.header, props = {};
        if (head) props.fill = o.header; else if (o.zebra && r % 2 === 0) props.fill = o.zebra;
        if (head || o.size) props.html = span(rows[r][c], { size: o.size, bold: !!head, color: head ? 'FFFFFF' : o.color });
        if (Object.keys(props).length) await set(`${t.path}/row[${r + 1}]/cell[${c + 1}]`, props);
      }
      return t.path;
    },
  };
}

// ---------- mind map ----------
/** tree(node): a node is [text, …children] or plain text for a leaf; text may be { text, note, icon, side, fill, color }. */
export function mm(run, file) {
  const props = v => typeof v === 'object' ? v : { text: v };
  const grow = async (path, kids) => {
    for (const k of kids) {
      const [head, ...rest] = Array.isArray(k) ? k : [k];
      const n = await run(['add', file, path, '--type', 'topic', ...P(props(head))]);
      if (rest.length) await grow(n.path, rest);
    }
  };
  return {
    file, run,
    tree: async ([head, ...kids]) => { await run(['set', file, '/topic[1]', ...P(props(head))]); await grow('/topic[1]', kids); },
  };
}
