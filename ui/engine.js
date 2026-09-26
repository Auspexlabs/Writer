// engine.js — the bridge between the editors and the writer engine serving this page.
// Open: the engine's JSON tree becomes an editor model, every block remembering its path.
// Save: the model is diffed against what was opened and the difference becomes writer commands.
// The file on disk is the only source of truth; the engine keeps everything the editor does not model.
import { txt, shape as mkShape, line as mkLine, SW, slideH, THEMES, phFamily, resolveColor, POLY, union, flatObjs } from './office-io.js';
import { lookFrom, pictureView, picSrc, placeStyle, PLACE } from './picture.js';
import * as PK from './pdf-kit.js';
import * as MM from './mindmap.js';

export const state = { workspace: '', chat: false, model: '', version: '', drafts: '', open: null };
const enc = encodeURIComponent;
// i18n: this module runs both in the page (window.$t from ui/i18n.js, loaded before it) and under node tests (no $t at
// all) — _t mirrors i18n.js's own {name} fill for that second case, so a node test sees the same Chinese it always has.
const _t = (zh, v) => globalThis.$t ? globalThis.$t(zh, v) : (v ? String(zh).replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : zh);
// unique for the session: a document keeps its id when 存储 or a rename moves its file, and a new draft may then reuse the old name
let docSeq = 0;
const idOf = path => 'f' + Array.from(path).reduce((h, c) => (h * 31 + c.charCodeAt(0)) >>> 0, 7).toString(36) + '.' + (++docSeq).toString(36);
const norm = p => String(p || '').replace(/\\/g, '/');
const isAbs = p => /^([a-zA-Z]:)?\//.test(norm(p));
const titleOf = path => norm(path).split('/').pop().replace(/\.[^.]+$/, '');
const dirOf = path => { const n = norm(path); return n.includes('/') ? n.slice(0, n.lastIndexOf('/')) : ''; };
const inside = (full, root) => !!root && norm(full).toLowerCase().startsWith(norm(root).replace(/\/+$/, '').toLowerCase() + '/');
const fullOf = path => isAbs(path) ? norm(path) : norm(state.workspace).replace(/\/+$/, '') + '/' + norm(path);
/** A document in the drafts folder: new documents stay there until their first 存储 (desktop apps only). */
export const isDraft = path => !!state.drafts && !!path && inside(fullOf(path), state.drafts);
/** Inside the workspace a path is kept relative (as /files lists it), elsewhere absolute. */
export const relOf = path => inside(fullOf(path), state.workspace) ? fullOf(path).slice(norm(state.workspace).replace(/\/+$/, '').length + 1) : norm(path);
/** Where new files go: the drafts folder in the desktop apps ('' = the workspace itself, which it is in the drafts window). */
const newDir = () => !state.drafts ? '' : norm(state.drafts).toLowerCase() === norm(state.workspace).toLowerCase() ? '' : norm(state.drafts);
// the editor each extension opens in, before the engine has said (files() replaces it with the engine's own list, which
// also holds the compatibility formats: .doc .xls .ppt .wps .et .dps .docm .odt .rtf .csv .html ... and .txt as text)
const FORMATS = { docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', md: 'md', markdown: 'md', txt: 'md', pdf: 'pdf', mm: 'mm', xmind: 'xmind' }; // an .xmind opens as a fresh .mm beside it (see openXmind)
const extOf = path => norm(path).split('/').pop().split('.').pop().toLowerCase();
/** The editor a file opens in (docx, xlsx, pptx, md, pdf, mm), or null for a type the engine does not read. */
export const editorOf = path => (state.open || FORMATS)[extOf(path)] || null;
/** Every readable extension as a file input's accept list. */
export const accept = () => Object.keys(state.open || FORMATS).map(e => '.' + e).join(',');
/** A compatibility format: not the editor's own file type, so it is read into a Word, Excel or PowerPoint draft on
 *  open and never written back — Office's compatibility mode. Its first 存储 offers a .docx/.xlsx/.pptx beside it. */
export const isCompat = path => { const t = editorOf(path); return (t === 'docx' || t === 'xlsx' || t === 'pptx') && extOf(path) !== t; };

export class EngineError extends Error {
  constructor(message, code, hint) { super(message); this.code = code; this.hint = hint; }
}

async function http(path, opts) {
  const res = await fetch(path, Object.assign({ credentials: 'same-origin' }, opts));
  if (res.ok) return res;
  let body = null;
  try { body = await res.json(); } catch (e) { }
  const err = body && body.error;
  throw new EngineError(err ? err.message : res.status + ' ' + res.statusText, err ? err.code : 'HTTP_' + res.status, err ? err.hint : '');
}

/** Runs one writer command; argv is the command line without the program name. Rejects with EngineError when the command fails. */
export async function run(argv) {
  const res = await http('/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) });
  const r = await res.json();
  if (r.code !== 0) throw new EngineError(r.error.message || 'command failed', r.error.code, r.error.hint);
  const out = r.output || '';
  try { return out.trim().startsWith('{') || out.trim().startsWith('[') ? JSON.parse(out) : out; } catch (e) { return out; }
}

/** The documents in the workspace, newest first, as lazy docs (loaded on open). */
export async function files() {
  const r = await (await http('/files')).json();
  state.workspace = r.workspace; state.chat = !!r.chat; state.model = r.model || ''; state.version = r.version || ''; state.drafts = r.drafts || '';
  if (r.open && typeof r.open === 'object') state.open = r.open;
  return r.files.filter(f => editorOf(f.path)).map(f => lazy(f.path));
}

export function lazy(path, format) {
  return { id: idOf(path), type: editorOf(path) || format || 'md', title: titleOf(path), path, loaded: false };
}

export const fileUrl = path => '/file?file=' + enc(path);
const binaryUrl = (path, node) => '/binary?file=' + enc(path) + '&path=' + enc(node);

/** { path, mtime, size } of a file on disk — cheap enough to poll, to notice another program (the MCP server, an agent's
 *  CLI calls) changing it outside this window. mtime is milliseconds since the epoch, the same unit PUT /file returns. */
export async function stat(path) { return (await http('/stat?file=' + enc(path))).json(); }

async function takenIn(dir) {
  const list = state.drafts && dir && norm(dir).toLowerCase() === norm(state.drafts).toLowerCase()
    ? (await (await http('/files?drafts=1')).json()).files.map(f => f.path) : (await files()).map(f => f.path);
  return new Set(list.map(p => norm(p).toLowerCase()));
}
/** A fresh name that does not clash with an existing file: title.ext, title 2.ext, … */
export async function freeName(title, ext, dir) {
  const taken = await takenIn(dir);
  const clean = (title || _t('未命名')).replace(/[\\/:*?"<>|]/g, ' ').trim().slice(0, 60) || _t('未命名');
  for (let n = 1; ; n++) { const p = (dir ? norm(dir).replace(/\/+$/, '') + '/' : '') + clean + (n > 1 ? ' ' + n : '') + '.' + ext; if (!taken.has(p.toLowerCase())) return p; }
}

export async function create(type, title) {
  const path = await freeName(title || _t('未命名'), type, newDir());
  await run(['create', path]);
  if (type === 'mm') await run(['set', path, '/topic[1]', '--prop', 'text=' + titleOf(path)]);
  if (type === 'pptx') await run(['add', path, '/', '--type', 'slide', '--prop', 'layout=title']); // a new deck opens on its title slide, as PowerPoint's does
  return lazy(path, type);
}

export async function upload(file) {
  const ext = extOf(file.name);
  if (!editorOf(file.name)) throw new EngineError(_t('不支持的文件类型 .{ext}', { ext }), 'UNKNOWN_FORMAT', '');
  const path = await freeName(file.name.replace(/\.[^.]+$/, ''), ext);
  await http('/file?file=' + enc(path), { method: 'PUT', body: file });
  return lazy(path);
}

export async function copy(doc) {
  const ext = doc.path.split('.').pop();
  const path = await freeName(doc.title + _t(' 副本'), ext, newDir());
  const bytes = await (await http(fileUrl(doc.path))).arrayBuffer();
  await http('/file?file=' + enc(path), { method: 'PUT', body: bytes });
  return lazy(path, doc.type);
}

/** A new document that starts as a copy of the file at url (a gallery template under /app/templates/), named title, as a draft. */
export async function createFrom(url, title, type) {
  const path = await freeName(title, type, newDir());
  const bytes = await (await http(url)).arrayBuffer();
  await http('/file?file=' + enc(path), { method: 'PUT', body: bytes });
  return lazy(path, type);
}

export async function rename(doc, title) {
  const ext = doc.path.split('.').pop();
  const dir = dirOf(doc.path);
  const path = await freeName(title, ext, dir);
  await http('/file?file=' + enc(path) + '&from=' + enc(doc.path), { method: 'PUT', body: new Uint8Array(0) });
  return path;
}

/** 存储 / 另存为: moves the document's file to target (a path the desktop shell granted after its Save dialog), or copies it
 *  with keep. Returns the path the document continues in. */
export async function saveAs(doc, target, keep) {
  // in the file's lane, after a save already under way: a save must not land on the path the file just left
  await inLane(doc.path, () => http('/file?file=' + enc(norm(target)) + '&from=' + enc(doc.path) + (keep ? '&keep=1' : ''), { method: 'PUT', body: new Uint8Array(0) }));
  return relOf(target);
}
/** 不存储: deletes a draft, after any save already under way, so that save cannot bring it back. */
export async function discard(doc) { await inLane(doc.path, () => http('/file?file=' + enc(doc.path), { method: 'DELETE' })); }

/** Converts through the engine: export <file> --to <name>.<ext>. Returns the new lazy doc. */
export async function exportTo(doc, ext) {
  const path = await freeName(doc.title, ext, newDir());
  await run(['export', doc.path, '--to', path]);
  return lazy(path, ext);
}

export function download(path, name) {
  const a = document.createElement('a'); a.href = fileUrl(path); a.download = name || path.split('/').pop(); document.body.appendChild(a); a.click(); setTimeout(() => a.remove(), 500);
}

/** Server-sent events for one file; cb() runs whenever the file changes on disk. Returns a function that stops watching. */
export function watch(path, cb) {
  const es = new EventSource('/events?file=' + enc(path));
  es.addEventListener('change', cb);
  return () => es.close();
}

/** Streams a chat turn; onEvent(name, data) gets text / tool / done / error. Resolves when the stream ends.
 *  opts: { instructions?, selection?, signal? } — the settings' reply style and custom text, (when the whole document is not to
 *  be sent) the selection, and an AbortSignal (停止) that ends the request; signal is not sent, the rest goes in the body. */
export async function chat(doc, messages, onEvent, opts) {
  const { signal, ...body } = opts || {};
  const res = await http('/chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(Object.assign({ file: doc ? doc.path : null, messages }, body)), signal });
  const reader = res.body.getReader(), dec = new TextDecoder(); let buf = '';
  for (; ;) {
    const { done, value } = await reader.read(); if (done) break;
    buf += dec.decode(value, { stream: true });
    let i; while ((i = buf.indexOf('\n\n')) >= 0) {
      const chunk = buf.slice(0, i); buf = buf.slice(i + 2);
      let name = 'message', data = '';
      chunk.split('\n').forEach(l => { if (l.startsWith('event:')) name = l.slice(6).trim(); else if (l.startsWith('data:')) data += l.slice(5).trim(); });
      if (!data) continue;
      let parsed; try { parsed = JSON.parse(data); } catch (e) { parsed = { text: data }; }
      onEvent(name, parsed);
    }
  }
}

// ---------- 设置 › AI: the model settings live in the engine, so a settings window of its own (same engine) sees the same ----------
const JSON_BODY = { 'Content-Type': 'application/json' };

/** The model settings: { provider, baseUrl, model, hasKey, source } (source: env | file | none). The key itself never comes back. */
export async function aiConfig() { return (await http('/ai')).json(); }

/** Saves { provider, baseUrl, model, apiKey? }: apiKey left out keeps the stored key, '' clears it. The engine switches its assistant at
 *  once; this window hears 'writer-ai' and the others a storage event on the key 'writer-ai'. Resolves with the settings that apply now. */
export async function saveAiConfig(settings) {
  // keepalive: a save on leaving a field still lands when that was the settings window closing
  const r = await (await http('/ai', { method: 'PUT', headers: JSON_BODY, body: JSON.stringify(settings), keepalive: true })).json();
  try { globalThis.localStorage.setItem('writer-ai', String(Date.now())); } catch (e) { }
  try { globalThis.dispatchEvent(new Event('writer-ai')); } catch (e) { }
  return r;
}

/** One small request with these settings, or the saved ones when left out: { ok, error? } with a short reason such as Key 无效. */
export async function testAi(settings) {
  return (await http('/ai/test', { method: 'POST', headers: JSON_BODY, body: JSON.stringify(settings || {}) })).json();
}

// ---------- units ----------
const UNIT = { cm: 1, mm: 0.1, in: 2.54, pt: 2.54 / 72, px: 2.54 / 96, emu: 2.54 / 914400 };
export function cmOf(s) { const m = /^(-?[\d.]+)\s*([a-z]+)?$/i.exec(String(s || '').trim()); if (!m) return 0; return +m[1] * (UNIT[(m[2] || 'cm').toLowerCase()] || 1); }
const cmStr = cm => (Math.round(cm * 1000) / 1000) + 'cm';
const hex = c => c ? '#' + c : null;
const unhex = c => c && c[0] === '#' ? c.slice(1).toUpperCase() : c ? c.toUpperCase() : 'none';

// Run nodes are left out: every editor renders text from the html property of its block, and runs are most of a document's nodes.
async function tree(path) { return await (await http('/json?file=' + enc(path) + '&skip=run')).json(); }

// ---------- open ----------
export async function open(doc) {
  if (isCompat(doc.path)) {
    // compatibility mode: the engine converts the file into a draft of the editor's own type; the original is never
    // written. The draft's first 存储 offers a .docx/.xlsx/.pptx beside the original (from remembers which).
    const type = editorOf(doc.path);
    const target = await freeName(doc.title, type, newDir());
    await run(['export', doc.path, '--to', target]);
    doc = Object.assign({}, doc, { path: target, type, from: doc.path });
  }
  const t = doc.type;
  let model;
  if (t === 'docx') model = await openDocx(doc);
  else if (t === 'xlsx') model = await openXlsx(doc);
  else if (t === 'pptx') model = await openPptx(doc);
  else if (t === 'pdf') model = await openPdf(doc);
  else if (t === 'mm') model = await openMm(doc);
  else if (t === 'xmind') return openXmind(doc);
  else model = await openMd(doc);
  const st = await stat(doc.path).catch(() => null); // best-effort: a missing stat just turns off the external-change check for this doc
  return Object.assign({}, doc, model, { loaded: true, dirty: false, _mtime: st ? st.mtime : null });
}

async function openMd(doc) {
  const text = await (await http(fileUrl(doc.path))).text();
  return { text, _orig: text };
}

// A PDF's bytes live in pdf-kit's store under `<id>@<generation>`; each open and each save is a new generation, and the
// pages say which one they come from, so undo past a save still shows the pages it had (see pdf-kit.js).
async function openPdf(doc) {
  const bytes = new Uint8Array(await (await http(fileUrl(doc.path))).arrayBuffer());
  const gen = Math.max(doc._gen || 0, PK.gen(doc.id)) + 1, key = PK.genKey(doc.id, gen); PK.setGen(doc.id, gen, bytes);
  const pdf = await PK.openPdf(key);
  return { pages: Array.from({ length: pdf.numPages }, (_, i) => ({ id: 'p' + i + '_' + doc.id, src: i, rot: 0, from: key })), annots: [], removed: [], form: 0, _formSaved: 0, _gen: gen, _n: pdf.numPages };
}
/** Writes the PDF as the editor shows it (pdf-kit.saveBytes: form values, page order, annotations) over the file. */
async function savePdf(doc) {
  if (!PK.dirty(doc.id, doc)) return 0;
  const bytes = await PK.saveBytes(doc.id, doc);
  const r = await http('/file?file=' + enc(doc.path) + (doc._mtime != null ? '&ifMtime=' + doc._mtime : ''), { method: 'PUT', body: bytes });
  doc._mtime = (await r.json()).mtime;
  const gen = PK.gen(doc.id) + 1, key = PK.genKey(doc.id, gen); PK.setGen(doc.id, gen, bytes); PK.forget(doc.id, gen - PK.KEEP);
  doc._saved = { before: doc.pages, annots: (doc.annots || []).map(a => a.id), removed: (doc.removed || []).slice() };
  doc.pages = doc.pages.map((p, i) => ({ id: p.id, src: i, rot: 0, from: key })); doc.annots = []; doc.removed = []; doc._formSaved = doc.form || 0; doc._gen = gen; doc._n = doc.pages.length;
  return 1;
}

// ----- docx: tree → blocks → html -----
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const isQuote = style => /quote/i.test(style || '');

// ----- docx tables: one grid model for the editor's table commands and for saving them -----
// A model is { rows: [{ id, ref, cells: [{ id, ref, c, cs, rs }] }] }: a cell lives in its top row, c is its first grid column,
// cs and rs its spans; each row keeps its cells in column order, which is also the engine's order of visible cells.
// The editor runs a command on its table through the model and records it on the table (data-ops); saving replays the
// same command on the model of the table as it was opened, which says exactly which engine commands to run, and where.
// New rows and cells get ids ('~…') that the replay reproduces, so they match the editor's elements.

/** Places rows of cells ({ id, ref, cs, rs }) on the grid, as a browser lays out colspan and rowspan. */
export function gridModel(rows) {
  const covered = [];
  return { rows: rows.map((row, r) => {
    const used = covered[r] || new Set(); let c = 0;
    const cells = row.cells.map(x => {
      while (used.has(c)) c++;
      const cell = { id: x.id, ref: x.ref, c, cs: Math.max(1, +x.cs || 1), rs: Math.max(1, +x.rs || 1) };
      for (let i = 1; i < cell.rs; i++) for (let j = 0; j < cell.cs; j++) (covered[r + i] ||= new Set()).add(c + j);
      c += cell.cs;
      return cell;
    });
    return { id: row.id, ref: row.ref, cells };
  }) };
}
const copyModel = m => ({ rows: m.rows.map(row => Object.assign({}, row, { cells: row.cells.map(x => Object.assign({}, x)) })) });
/** grid[r][c] = the cell covering that position. */
function coverage(m) {
  const g = m.rows.map(() => []);
  m.rows.forEach((row, r) => row.cells.forEach(x => { for (let i = 0; i < x.rs && r + i < g.length; i++) for (let j = 0; j < x.cs; j++) g[r + i][x.c + j] = x; }));
  return g;
}
function locate(m, id) { for (let r = 0; r < m.rows.length; r++) { const x = m.rows[r].cells.find(x => x.id === id); if (x) return { r, x }; } throw new Error(_t('单元格不在表格中')); }
const rowOf = (m, x) => m.rows.findIndex(row => row.cells.includes(x));
const sortCells = row => row.cells.sort((a, b) => a.c - b.c);
/** Same rows, same cells in the same places: the check that a replay reached what the editor shows. */
export const sameGrid = (a, b) => JSON.stringify(a.rows.map(r => [r.id, r.cells.map(x => [x.id, x.c, x.cs, x.rs])])) === JSON.stringify(b.rows.map(r => [r.id, r.cells.map(x => [x.id, x.c, x.cs, x.rs])]));

/** Table commands. Each takes a model and returns a new one (plus what the save needs to know); refusals throw an Error whose
 * message is shown to the user. New rows and cells take ids from newId() in a fixed order and remember the element they copy (from). */
export const tableOps = {
  /** A row above or below the cell's rows. Like the engine it copies the row above (the first row when inserting at the top),
   * and a merged cell that runs across the new row grows over it instead of getting a cell. */
  insertRow(m0, id, below, newId) {
    const m = copyModel(m0), hit = locate(m, id), at = below ? hit.r + hit.x.rs : hit.r, tpl = Math.max(0, at - 1);
    const row = { id: newId(), from: m.rows[tpl].id, fresh: true, cells: [] };
    for (const x of [...new Set(coverage(m)[tpl].filter(Boolean))]) {
      if (at > 0 && at < m.rows.length && rowOf(m, x) + x.rs - 1 >= at) x.rs++;
      else row.cells.push({ id: newId(), c: x.c, cs: x.cs, rs: 1, from: x.id, fresh: true });
    }
    m.rows.splice(at, 0, row);
    return { m, at };
  },
  /** Removes one row. A merged cell that starts in it moves down with its content, as the engine does. */
  deleteRow(m0, rowId) {
    const m = copyModel(m0), r = m.rows.findIndex(row => row.id === rowId);
    if (r < 0) throw new Error(_t('行不在表格中'));
    for (const x of [...new Set(coverage(m)[r].filter(Boolean))]) {
      if (rowOf(m, x) < r) x.rs--;
      else if (x.rs > 1) { x.rs--; m.rows[r + 1].cells.push(x); sortCells(m.rows[r + 1]); }
    }
    m.rows.splice(r, 1);
    return { m, at: r };
  },
  /** A column left or right of the cell. A merged cell that spans the new column's place grows over it. */
  insertCol(m0, id, right, newId) {
    const m = copyModel(m0), hit = locate(m, id), c = right ? hit.x.c + hit.x.cs : hit.x.c;
    insertColAt(m, c, newId);
    return { m, c };
  },
  /** Removes the grid column c: cells in it go, merged cells across it get narrower. */
  deleteCol(m0, c) {
    const m = copyModel(m0), g = coverage(m), seen = new Set();
    m.rows.forEach((row, r) => { const x = g[r][c]; if (!x || seen.has(x)) return; seen.add(x); if (x.cs > 1) x.cs--; else x.gone = true; });
    const emptied = m.rows.filter(row => row.cells.length && row.cells.every(x => x.gone)).length;
    for (const row of m.rows) { row.cells = row.cells.filter(x => !x.gone); for (const x of row.cells) if (x.c > c) x.c--; }
    const empty = m.rows.every(row => !row.cells.length);
    if (emptied && !empty) throw new Error(_t('删除这一列会让某些行没有单元格，请先拆分合并单元格'));
    return { m, empty };
  },
  /** Joins the cell with its right neighbour, which must cover the same rows. */
  mergeRight(m0, id) {
    const m = copyModel(m0), { r, x } = locate(m, id), y = (coverage(m)[r] || [])[x.c + x.cs];
    if (!y || rowOf(m, y) !== r || y.rs !== x.rs) throw new Error(_t(y ? '右侧单元格的行数不同，不能合并' : '右侧没有可合并的单元格'));
    x.cs += y.cs; m.rows[r].cells.splice(m.rows[r].cells.indexOf(y), 1);
    return { m, into: x.id, from: y.id };
  },
  /** Joins the cell with the one below, which must cover the same columns. */
  mergeDown(m0, id) {
    const m = copyModel(m0), { r, x } = locate(m, id), below = r + x.rs, y = below < m.rows.length ? coverage(m)[below][x.c] : null;
    if (!y || rowOf(m, y) !== below || y.c !== x.c || y.cs !== x.cs) throw new Error(_t(y ? '下方单元格的列数不同，不能合并' : '下方没有可合并的单元格'));
    x.rs += y.rs; m.rows[below].cells.splice(m.rows[below].cells.indexOf(y), 1);
    return { m, into: x.id, from: y.id };
  },
  /** Splits a merged cell back into one cell per grid position; the new ones are empty and look like it. */
  split(m0, id, newId) {
    const m = copyModel(m0), { r, x } = locate(m, id);
    if (x.cs === 1 && x.rs === 1) throw new Error(_t('这个单元格没有合并'));
    for (let i = 0; i < x.rs; i++) {
      for (let j = 0; j < x.cs; j++) if (i || j) m.rows[r + i].cells.push({ id: newId(), c: x.c + j, cs: 1, rs: 1, from: x.id, fresh: true });
      sortCells(m.rows[r + i]);
    }
    const was = { cs: x.cs, rs: x.rs }; x.cs = 1; x.rs = 1;
    return { m, was };
  }
};
/** Inserts grid column c in place: a cell spanning across it grows by one column, every other row gets a new cell there
 * (it copies the cell on its left, else the one on its right). Returns what happened per row, top to bottom. */
function insertColAt(m, c, newId) {
  const g = coverage(m), made = [];
  m.rows.forEach((row, r) => {
    const x = g[r][c], left = c > 0 ? g[r][c - 1] : null;
    if (x && x === left) { if (rowOf(m, x) === r) made.push({ r, widen: x }); return; }
    made.push({ r, row, cell: { id: newId(), c, cs: 1, rs: 1, from: (left || x || {}).id, fresh: true } });
  });
  for (const row of m.rows) for (const x of row.cells) if (x.c >= c) x.c++;
  for (const d of made) { if (d.widen) d.widen.cs++; else { d.row.cells.push(d.cell); sortCells(d.row); } }
  return made;
}

/** The engine commands that replay the editor's table commands (ops) on a table as it was opened (o: path, rows of cells with
 * props), and the rows as the engine then has them: the cells keep their entries, new ones are marked fresh.
 * Throws when an op does not fit the table; the caller then writes the table anew instead. */
export function replayTable(o, ops, file) {
  let m = gridModel(o.rows.map(row => ({ id: row.path, ref: row, cells: row.cells.map(x => ({ id: x.path, ref: x, cs: x.props && x.props.colspan, rs: x.props && x.props.rowspan })) })));
  const T = o.cur || o.path, cmds = [], html = new Map(), rowProps = new Map(o.rows.map(r => [r.path, r.props || {}]));
  const rowPath = r => `${T}/row[${r + 1}]`;
  const cellPath = x => { const r = rowOf(m, x); return `${rowPath(r)}/cell[${m.rows[r].cells.indexOf(x) + 1}]`; };
  const htmlOf = id => html.has(id) ? html.get(id) : (m.rows.flatMap(r => r.cells).find(x => x.id === id)?.ref?.props?.html || '');
  const cmd = (...argv) => cmds.push([argv[0], file, ...argv.slice(1)]);
  for (const op of ops || []) {
    const ids = (op.ids || []).slice(), newId = () => { if (!ids.length) throw new Error('replay needs more new ids than were recorded'); return ids.shift(); };
    if (op.op === 'insertRow') {
      const r = tableOps.insertRow(m, op.ref, op.below, newId);
      cmd('add', T, '--type', 'row', '--index', String(r.at + 1));
      const row = r.m.rows[r.at]; rowProps.set(row.id, Object.fromEntries(Object.entries(rowProps.get(row.from) || {}).filter(([k]) => k !== 'header')));
      m = r.m;
    } else if (op.op === 'deleteRows') {
      for (const id of op.rows) { cmd('remove', rowPath(m.rows.findIndex(row => row.id === id))); m = tableOps.deleteRow(m, id).m; }
    } else if (op.op === 'mergeRight' || op.op === 'mergeDown') {
      const { x } = locate(m, op.ref), r = tableOps[op.op](m, op.ref);
      const joined = locate(r.m, op.ref).x;
      cmd('set', cellPath(x), '--prop', op.op === 'mergeRight' ? 'colspan=' + joined.cs : 'rowspan=' + joined.rs);
      html.set(op.ref, joinCells(htmlOf(op.ref), htmlOf(r.from)));
      m = r.m;
    } else if (op.op === 'split') {
      const { x } = locate(m, op.ref), path = cellPath(x);
      if (x.cs > 1) cmd('set', path, '--prop', 'colspan=1');
      if (x.rs > 1) cmd('set', path, '--prop', 'rowspan=1');
      m = tableOps.split(m, op.ref, newId).m;
    } else if (op.op === 'deleteCols') {
      const { x } = locate(m, op.ref), first = x.c;
      for (let k = x.cs - 1; k >= 0; k--) { planDeleteCol(m, first + k, cmd, cellPath, rowPath); m = tableOps.deleteCol(m, first + k).m; }
    } else if (op.op === 'insertCol') {
      m = planInsertCol(m, op, newId, cmd, rowPath);
    } else throw new Error('unknown table command ' + op.op);
    if (ids.length) throw new Error('replay made fewer new ids than were recorded');
  }
  const spans = x => Object.assign({}, x.cs > 1 ? { colspan: String(x.cs) } : {}, x.rs > 1 ? { rowspan: String(x.rs) } : {});
  const rows = m.rows.map(row => {
    const cells = row.cells.map(x => {
      if (!x.ref) return { kind: 'cell', path: x.id, props: Object.assign({ html: html.get(x.id) || '' }, spans(x)), fresh: true };
      const props = Object.assign({}, x.ref.props); delete props.colspan; delete props.rowspan;
      return Object.assign({}, x.ref, { props: Object.assign(props, spans(x), html.has(x.id) ? { html: html.get(x.id) } : {}) });
    });
    return row.ref ? Object.assign({}, row.ref, { cells }) : { kind: 'row', path: row.id, props: rowProps.get(row.id) || {}, cells, fresh: true };
  });
  return { m, cmds, rows };
}

/** Deleting grid column c: a one-column cell goes (with the rows it spans); a wider one gets narrower, which in the engine is
 * lowering its colspan and removing the empty cell that splits off in each of its rows. Lower cells first, so paths hold. */
function planDeleteCol(m, c, cmd, cellPath, rowPath) {
  const g = coverage(m), owners = [...new Set(g.map(row => row[c]).filter(Boolean))];
  for (const x of owners.reverse()) {
    if (x.cs === 1) { cmd('remove', cellPath(x)); continue; }
    const top = rowOf(m, x), at = m.rows[top].cells.indexOf(x);
    cmd('set', cellPath(x), '--prop', 'colspan=' + (x.cs - 1));
    for (let i = x.rs - 1; i >= 1; i--) cmd('remove', `${rowPath(top + i)}/cell[${m.rows[top + i].cells.filter(y => y.c < x.c).length + 1}]`);
    cmd('remove', `${rowPath(top)}/cell[${at + 2}]`);
  }
}

/** Inserting a grid column. The engine adds a cell before the next visible cell of its row, which misplaces it when a merged
 * cell from a row above hides a cell in between, and it cannot widen that hidden part. Those merges are split for the
 * moment (rowspan=1), the column goes in row by row, and they are merged again. */
function planInsertCol(m0, op, newId, cmd, rowPath) {
  const { m: result } = tableOps.insertCol(m0, op.ref, op.right, (ids => () => ids.shift())((op.ids || []).slice()));
  const m = copyModel(m0), hit = locate(m, op.ref), c = op.right ? hit.x.c + hit.x.cs : hit.x.c, g = coverage(m), flat = new Set();
  m.rows.forEach((row, r) => {
    const x = g[r][c], left = c > 0 ? g[r][c - 1] : null, straddles = x && x === left;
    if (straddles && x.rs > 1) flat.add(x);
    const p = straddles ? x.c + x.cs : c, next = row.cells.find(y => y.c >= p);
    for (let col = p; col < (next ? next.c : g[r].length); col++) if (g[r][col] && rowOf(m, g[r][col]) < r) flat.add(g[r][col]);
  });
  const order = [...flat].sort((a, b) => rowOf(m, b) - rowOf(m, a) || b.c - a.c);
  const pathOf = x => { const r = rowOf(m, x); return `${rowPath(r)}/cell[${m.rows[r].cells.indexOf(x) + 1}]`; };
  for (const x of order) {
    cmd('set', pathOf(x), '--prop', 'rowspan=1');
    const r = rowOf(m, x);
    for (let i = 1; i < x.rs; i++) { m.rows[r + i].cells.push({ id: '#' + x.id + '#' + i, c: x.c, cs: x.cs, rs: 1, part: x }); sortCells(m.rows[r + i]); }
    x.was = x.rs; x.rs = 1;
  }
  const gf = coverage(m);
  m.rows.forEach((row, r) => {
    const x = gf[r][c], left = c > 0 ? gf[r][c - 1] : null;
    if (x && x === left) {
      const i = row.cells.indexOf(x);
      cmd('add', rowPath(r), '--type', 'cell', '--index', String(i + 2));
      cmd('set', `${rowPath(r)}/cell[${i + 1}]`, '--prop', 'colspan=' + (x.cs + 1));
    } else {
      const k = row.cells.filter(y => y.c < c).length + 1;
      cmd('add', rowPath(r), '--type', 'cell', ...(k <= row.cells.length ? ['--index', String(k)] : []));
    }
  });
  insertColAt(m, c, newId);
  for (const x of order) {
    cmd('set', pathOf(x), '--prop', 'rowspan=' + x.was);
    for (const row of m.rows) row.cells = row.cells.filter(y => y.part !== x);
    x.rs = x.was; delete x.was;
  }
  if (!sameGrid(m, result)) throw new Error('column insert did not replay');
  return result;
}

// ----- docx tables and contents in the editor: what they look like -----
/** The lines a table draws, from its props as the engine reports them: its own borders, else what its style draws. A table
 * without style or borders draws none in Word. */
export function tableLook(p) {
  p = p || {};
  if (p.borders) return p.borders === 'all' ? 'grid' : p.borders; // none | outside | inside | horizontal
  const style = p.style || '';
  if (!style || /^TableNormal$/i.test(style)) return 'none';
  if (/^ThreeLineTable$/i.test(style)) return p.header === 'false' ? 'three0' : 'three'; // three0: no 标题行, so no rule under the first row
  if (/^PlainTable1$/i.test(style)) return 'horizontal';
  return 'grid';
}
const LINE = '1px solid #C7C7CC', GUIDE = '1px dashed #E5E5EA', RULE = '1.5px solid #1D1D1F', HAIR = '1px solid #1D1D1F';
/** Border CSS of a cell at grid row r, place x ({ c, cs, rs }) in an R×C table. Lines Word would not draw show as faint guides. */
export function cellLines(look, r, x, R, C) {
  const top = r === 0, bottom = r + x.rs >= R, left = x.c === 0, right = x.c + x.cs >= C;
  const edge = (outer, isOuter) => look === 'grid' ? LINE : look === 'outside' ? (isOuter ? LINE : GUIDE) : look === 'inside' ? (isOuter ? GUIDE : LINE) : outer;
  const three = look === 'three' || look === 'three0';
  const t = three ? (top ? RULE : GUIDE) : look === 'horizontal' ? LINE : edge(GUIDE, top);
  const b = three ? (bottom ? RULE : r === 0 && look === 'three' ? HAIR : GUIDE) : look === 'horizontal' ? LINE : edge(GUIDE, bottom);
  const l = three || look === 'horizontal' ? GUIDE : edge(GUIDE, left), rt = three || look === 'horizontal' ? GUIDE : edge(GUIDE, right);
  return `border-top:${t};border-right:${rt};border-bottom:${b};border-left:${l}`;
}
/** A cell's own lines over the table's (its borders prop): none, all, or the sides listed; '' leaves the table's. */
export function ownLines(borders) {
  if (!borders) return '';
  const all = ['top', 'right', 'bottom', 'left'], on = borders === 'none' ? [] : borders === 'all' || borders === 'outside' || borders === 'box' ? all : borders.split(/[\s,]+/);
  return all.map(s => `border-${s}:${on.includes(s) ? LINE : GUIDE}`).join(';');
}
const cellStyle = (p, look, r, x, R, C) => `padding:6px 8px;${cellLines(look, r, x, R, C)}${p.borders ? ';' + ownLines(p.borders) : ''}${p.fill && p.fill !== 'none' ? ';background:#' + esc(p.fill) : ''}${p.valign && p.valign !== 'top' ? ';vertical-align:' + esc(p.valign) : ''}${p.align && p.align !== 'left' ? ';text-align:' + esc(p.align) : ''}`;
/** Column widths (the table's widths prop, a JSON list of lengths) as a colgroup, so the editor lays the columns out as Word does. */
const colgroupOf = widths => { let list = []; try { list = JSON.parse(widths || '[]'); } catch (e) { list = []; } return Array.isArray(list) && list.length ? '<colgroup>' + list.map(w => `<col style="width:${cmOf(w) * CM_PX}px">`).join('') + '</colgroup>' : ''; };
const rowStyle = props => props.height ? ` style="height:${Math.round(cmOf(props.height + 'emu') * CM_PX)}px"` : '';

const TOC_STYLE = "border:1px solid #E5E5EA;border-radius:6px;padding:14px 18px;margin:12px 0;font-family:'Noto Sans SC',sans-serif;font-size:14px;line-height:1.9;background:#FFFFFF;color:#1D1D1F";
/** A table of contents in the editor: read-only, its entries indented by level, page numbers when the file has them. Its style
 *  (DocxToc.StyleOf) says how an entry ends: classic runs dots to the number, simple leaves a gap, plain has no number. */
export function tocHtml({ path, levels, title, entries, style }) {
  style = style === 'simple' || style === 'plain' ? style : 'classic';
  const head = title ? `<div style="font-weight:600;margin-bottom:4px">${esc(title)}</div>` : '';
  const dots = style === 'classic' ? '<span style="flex:1;min-width:12px;margin:0 4px;border-bottom:1.5px dotted #AEAEB2;transform:translateY(-5px)"></span>' : '<span style="flex:1;min-width:12px"></span>';
  const body = entries.length ? entries.map(e => `<div style="display:flex;align-items:baseline;padding-left:${(Math.max(1, e.level) - 1) * 18}px"><span>${e.href ? `<a href="#${esc(e.href)}">${esc(e.text)}</a>` : esc(e.text)}</span>${style !== 'plain' && e.page ? dots + `<span style="color:#8E8E93">${esc(e.page)}</span>` : ''}</div>`).join('')
    : `<div style="color:#8E8E93">${_t('添加标题后，目录会在保存时生成')}</div>`;
  return `<nav data-toc="1"${path ? ` data-path="${esc(path)}"` : ''} data-levels="${esc(levels || '3')}" data-title="${esc(title || '')}"${style !== 'classic' ? ` data-toc-style="${style}"` : ''} contenteditable="false" style="${TOC_STYLE}">${head}${body}</nav>`;
}
/** Headings of the live editor that a contents of `levels` levels lists, each with an id to jump to. */
export function editorHeadings(root, levels) {
  return Array.from(root.querySelectorAll('h1,h2,h3')).filter(h => !h.closest('[data-toc]') && !h.hasAttribute('data-style'))
    .map((h, i) => { if (!h.id) h.id = 'h' + Date.now().toString(36) + i; return { level: +(h.getAttribute('data-level') || h.tagName[1]), text: h.innerText.replace(/\s+/g, ' ').trim(), href: h.id }; })
    .filter(e => e.text && e.level <= (+levels || 3));
}
/** Draws every table of contents in the editor again from its headings, as the engine builds it on save. */
export function refreshTocs(root) {
  for (const nav of Array.from(root.querySelectorAll('[data-toc]'))) {
    const levels = nav.getAttribute('data-levels') || '3', title = nav.getAttribute('data-title') || '';
    const tmp = parseHtml(tocHtml({ path: nav.getAttribute('data-path'), levels, title, entries: editorHeadings(root, levels), style: nav.getAttribute('data-toc-style') })).firstChild;
    nav.replaceWith(tmp);
  }
}

// ----- docx tables in the editor: its commands on the live table -----
let tableIds = 0;
/** An id for a row or cell the editor makes, until the save gives it its path. */
export const newTableId = () => '~' + Date.now().toString(36) + (++tableIds).toString(36);
/** The editor's table as a model; rows and cells keep their elements. Elements without a path get an id. */
export function tableModelOf(table) {
  const id = el => el.getAttribute('data-path') || (el.setAttribute('data-path', newTableId()), el.getAttribute('data-path'));
  return gridModel(Array.from(table.rows).map(tr => ({ id: id(tr), ref: tr, cells: Array.from(tr.cells).map(td => ({ id: id(td), ref: td, cs: td.colSpan, rs: td.rowSpan })) })));
}
/** One command of the editor's table ribbon on a model, by the id of the cell it starts from: the new model, and the record a
 * save replays. insertRow(below?), insertCol(right?), deleteRows and deleteCols (every row or column the cell covers, as in
 * Word), mergeRight, mergeDown, split. Refusals throw with the message to show. */
export function tableCommand(before, op, cell, flag, newId) {
  const ids = [], take = () => { const id = newId(); ids.push(id); return id; };
  let r, record = { op, ref: cell };
  if (op === 'insertRow') { r = tableOps.insertRow(before, cell, flag, take); record.below = !!flag; }
  else if (op === 'insertCol') { r = tableOps.insertCol(before, cell, flag, take); record.right = !!flag; }
  else if (op === 'deleteRows') {
    const { r: top, x } = locate(before, cell), rows = before.rows.slice(top, top + x.rs).map(row => row.id).reverse();
    r = { m: rows.reduce((m, id) => tableOps.deleteRow(m, id).m, before) }; record = { op, rows };
  }
  else if (op === 'deleteCols') { const { x } = locate(before, cell); let m = before, empty = false; for (let k = x.cs - 1; k >= 0; k--) ({ m, empty } = tableOps.deleteCol(m, x.c + k)); r = { m, empty }; }
  else if (op === 'mergeRight' || op === 'mergeDown') r = tableOps[op](before, cell);
  else if (op === 'split') r = tableOps.split(before, cell, take);
  else throw new Error('unknown table command ' + op);
  if (ids.length) record.ids = ids;
  return Object.assign(r, { record, empty: r.empty || !r.m.rows.length });
}

/** Runs a table command on the live table: the DOM follows the new model (rows and cells in order, spans, copies for new ones)
 * and the command is recorded on the table for the save. Returns the new model, or null when the table went away. */
export function runTableOp(table, op, cell, flag) {
  const before = tableModelOf(table), r = tableCommand(before, op, cell, flag, newTableId), record = r.record;
  const ref = id => { for (const row of before.rows) { if (row.id === id) return row.ref; for (const x of row.cells) if (x.id === id) return x.ref; } return null; };
  if (r.empty) { table.remove(); return null; }
  if (r.into) { const into = ref(r.into), from = ref(r.from); into.innerHTML = joinCells(into.innerHTML, from.innerHTML); }
  const body = table.tBodies[0] || table, trs = r.m.rows.map(row => {
    let tr = row.ref;
    if (!tr) { tr = (ref(row.from) || document.createElement('tr')).cloneNode(false); tr.setAttribute('data-path', row.id); tr.removeAttribute('data-w-header'); }
    tr.replaceChildren(...row.cells.map(x => {
      let td = x.ref;
      if (!td) { td = (ref(x.from) || document.createElement('td')).cloneNode(false); td.setAttribute('data-path', x.id); td.innerHTML = '<br>'; }
      if (x.cs > 1) td.colSpan = x.cs; else td.removeAttribute('colspan');
      if (x.rs > 1) td.rowSpan = x.rs; else td.removeAttribute('rowspan');
      return td;
    }));
    return tr;
  });
  body.replaceChildren(...trs);
  let ops = []; try { ops = JSON.parse(table.getAttribute('data-ops') || '[]'); } catch (e) { ops = []; }
  table.setAttribute('data-ops', JSON.stringify(ops.concat([record])));
  if (op === 'insertCol' || op === 'deleteCols') { table.querySelector('colgroup')?.remove(); table.removeAttribute('data-w-widths'); table.style.tableLayout = ''; } // the columns changed: Word lays them out anew
  drawTableLines(table);
  return r.m;
}
/** Draws the live table's lines from its style and borders attributes, after a command or a new look. */
export function drawTableLines(table) {
  const m = tableModelOf(table), look = tableLook({ style: table.getAttribute('data-w-style'), borders: table.getAttribute('data-w-borders'), header: table.getAttribute('data-w-header') });
  const R = m.rows.length, C = Math.max(1, ...coverage(m).map(r => r.length));
  m.rows.forEach((row, r) => row.cells.forEach(x => { x.ref.style.border = ''; x.ref.style.cssText += ';' + cellLines(look, r, x, R, C) + (x.ref.getAttribute('data-w-borders') ? ';' + ownLines(x.ref.getAttribute('data-w-borders')) : ''); }));
}

/** Cell contents after a merge: both, one under the other, as the engine joins their paragraphs (empty ones dropped). */
const blankHtml = h => !/<img/i.test(h || '') && !String(h || '').replace(/<[^>]*>/g, '').replace(/&nbsp;| |​/g, '').trim();
export const joinCells = (a, b) => [a, b].filter(h => !blankHtml(h)).join('<br>') || '<br>';

/** A page break inside a paragraph: the engine writes it as Word's HTML does; the editor draws it as a line across the page that
 *  cannot be typed into (a span, which a paragraph can hold), and sends it back as the engine's. */
const PB_WORD = '<br style="page-break-before:always">', PB_LINE = '<span data-pb="1" contenteditable="false"></span>';
export const pbIn = html => String(html).split(PB_WORD).join(PB_LINE);
export const pbOut = html => String(html).replace(/<span data-pb="1"[^>]*><\/span>/g, PB_WORD);
/** A heading or paragraph's own props (kept for the save), and how it shows, its own or its style's: a page break above it, a fill;
 *  widow control its style turns off rides on data-widow, for the panel's 孤行控制. */
const paraOf = (b, p, n) => { const c = n.computed || {}; for (const k of PARA_OWN) if (p[k]) b.props[k] = p[k]; b.pbb = (p.pageBreakBefore || c.pageBreakBefore) === 'true'; b.shade = p.fill || c.fill; b.widowOff = !p.widowControl && c.widowControl === 'false'; };

/** Word pictures are addressed by id (//image[@id=n], their wp:docPr) when the file gives every picture one of its own: the address
 *  survives every move and renumbering; else by their place (/body/image[n]). */
export function uniquePictureIds(nodes) {
  const ids = []; const walk = list => { for (const n of list || []) { if (n.kind === 'image') ids.push(n.props && n.props.id); walk(n.children); } }; walk(nodes);
  return ids.every(Boolean) && new Set(ids).size === ids.length;
}
const picPath = (n, byId) => byId && n.props && n.props.id ? `//image[@id=${n.props.id}]` : n.path;
const placeOf = p => Object.fromEntries(PLACE.filter(k => p[k] != null && p[k] !== '').map(k => [k, String(p[k])]));
/** A picture as the tree gives it: its look and its place (picture.js PLACE), the frame in px. */
const picOf = (n, file, byId) => { const p = n.props || {}, path = picPath(n, byId); return { path, props: Object.assign({ src: binaryUrl(file, path) }, lookFrom(p)), place: placeOf(p), width: p.width ? cmOf(p.width) / 2.54 * 96 : 0, height: p.height ? cmOf(p.height) / 2.54 * 96 : 0 }; };

export function blocksOf(nodes, file) {
  const byId = uniquePictureIds(nodes);
  return (nodes || []).map(n => {
    const p = n.props || {}, b = { kind: n.kind, path: n.path, props: {} };
    const pics = (n.children || []).filter(c => c.kind === 'image'); // a paragraph's own pictures: floating in it, or in its line of text
    if (pics.length && (n.kind === 'heading' || n.kind === 'paragraph')) b.pics = pics.map(c => picOf(c, file, byId));
    if (n.kind === 'heading') { b.props.html = p.html || esc(p.text); b.props.level = p.level || '1'; if (p.align) b.props.align = p.align; paraOf(b, p, n); }
    else if (n.kind === 'paragraph') { b.props.html = p.html || esc(p.text); if (p.list && p.list !== 'none') { b.props.list = p.list; b.props.level = p.level || '0'; if (p.restart === 'true') b.props.restart = 'true'; } if (p.align) b.props.align = p.align; if (p.style) b.props.style = p.style; paraOf(b, p, n); }
    else if (n.kind === 'code') b.props.text = p.text || '';
    else if (n.kind === 'table') {
      // the tree gives json props (widths) parsed; the editor keeps them as the JSON text the engine takes back
      b.props = Object.fromEntries(['style', 'header', 'borders', 'borderColor', 'width', 'widths', 'align'].filter(k => p[k] != null).map(k => [k, typeof p[k] === 'object' ? JSON.stringify(p[k]) : p[k]]));
      b.rows = (n.children || []).filter(r => r.kind === 'row').map(r => ({ kind: 'row', path: r.path,
        props: Object.fromEntries(['header', 'height'].filter(k => r.props?.[k] != null).map(k => [k, r.props[k]])),
        cells: (r.children || []).filter(c => c.kind === 'cell').map(c => ({ kind: 'cell', path: c.path,
          props: Object.assign({ html: c.props.html != null ? c.props.html : esc(c.props.text) },
            Object.fromEntries(['fill', 'colspan', 'rowspan', 'borders', 'valign', 'width', 'align'].filter(k => c.props?.[k] != null).map(k => [k, c.props[k]]))) })) }));
    }
    else if (n.kind === 'toc') b.props = { levels: p.levels || '3', title: p.title || '', text: p.text || '', style: p.style || 'classic' };
    else if (n.kind === 'image') Object.assign(b, picOf(n, file, byId));
    else if (n.kind === 'pagebreak') { }
    else b.props.html = esc(p.text || '');
    return b;
  });
}

/** A table's own CSS: its width (a percentage or a length, else the text width; fixed columns once widths are set) and its place between the margins. */
export function tableCss(p) {
  const w = p.width && p.width !== 'auto' ? (/%$/.test(p.width) ? p.width : cmOf(p.width) * CM_PX + 'px') : '100%';
  return `border-collapse:collapse;width:${w};margin:8px ${p.align === 'center' ? 'auto' : p.align === 'right' ? '0 8px auto' : '0'}${p.widths ? ';table-layout:fixed' : ''}`;
}
/** The list kinds beyond the html tags' own: the ol carries them as data-w-list (number is a plain ol, bullet a plain ul). */
export const LIST_KINDS = ['number', 'outline', 'chinese'];
export function blocksToHtml(blocks) {
  let out = ''; const stack = []; // open lists: {type, kind}
  const closeLists = n => { while (stack.length > n) { out += '</' + stack.pop().type + '>'; } };
  const pa = (b, tag, extra) => {
    const css = [alignCss(b.props.align), b.shade ? 'background:#' + esc(b.shade) : '', paraCss(b.props)].filter(Boolean).join(';');
    return `<${tag} data-path="${esc(b.path)}"${extra || ''}${attrs(b.props, PARA_OWN)}${b.pbb ? ' data-pb="before"' : ''}${b.widowOff ? ' data-widow="off"' : ''}${css ? ` style="${css}"` : ''}>${(b.pics || []).map(x => picHtml(x, true)).join('')}${pbIn(b.props.html || '<br>')}</${tag}>`;
  };
  const attrs = (p, keys) => keys.map(k => p?.[k] != null ? ` data-w-${k.toLowerCase()}="${esc(p[k])}"` : '').join('');
  for (const b of blocks) {
    if (b.kind === 'paragraph' && b.props.list) {
      // a list of another kind, or numbering that starts again (data-w-restart), is an ol of its own
      const level = +b.props.level || 0, kind = b.props.list, type = kind === 'bullet' ? 'ul' : 'ol';
      while (stack.length > level + 1) out += '</' + stack.pop().type + '>';
      if (stack.length === level + 1 && (stack[level].kind !== kind || b.props.restart)) { out += '</' + stack.pop().type + '>'; }
      while (stack.length < level + 1) {
        const own = stack.length === level, t = own ? type : 'ul';
        out += own ? `<${t}${kind !== 'bullet' && kind !== 'number' ? ` data-w-list="${kind}"` : ''}${b.props.restart ? ' data-w-restart="1"' : ''}>` : '<ul>';
        stack.push({ type: t, kind: own ? kind : 'bullet' });
      }
      out += pa(b, 'li');
      continue;
    }
    closeLists(0);
    if (b.kind === 'heading') out += pa(b, 'h' + Math.min(3, +b.props.level || 1), +b.props.level > 3 ? ` data-level="${b.props.level}"` : '');
    else if (b.kind === 'paragraph') {
      const st = b.props.style || '';
      if (/^title$/i.test(st)) out += pa(b, 'h1', ' data-style="Title"');
      else if (/^subtitle$/i.test(st)) out += pa(b, 'h2', ' data-style="Subtitle"');
      else if (isQuote(st)) out += pa(b, 'blockquote', ` data-style="${esc(st)}"`);
      else out += pa(b, 'p', st && st !== 'Normal' ? ` data-style="${esc(st)}"` : '');
    }
    else if (b.kind === 'code') out += `<pre data-path="${esc(b.path)}">${esc(b.props.text)}</pre>`;
    else if (b.kind === 'table') {
      const look = tableLook(b.props), m = blocksModel(b.rows), R = m.rows.length, C = Math.max(1, ...coverage(m).map(r => r.length));
      out += `<table data-path="${esc(b.path)}"${attrs(b.props, ['style', 'header', 'borders', 'borderColor', 'width', 'widths', 'align'])} style="${tableCss(b.props)}">${colgroupOf(b.props.widths)}<tbody>` + m.rows.map((row, r) => `<tr data-path="${esc(row.ref.path)}"${attrs(row.ref.props, ['header', 'height'])}${rowStyle(row.ref.props)}>` + row.cells.map(x => { const c = x.ref;
        return `<td data-path="${esc(c.path)}"${attrs(c.props, ['fill', 'borders', 'valign', 'width', 'align'])}${x.cs > 1 ? ` colspan="${x.cs}"` : ''}${x.rs > 1 ? ` rowspan="${x.rs}"` : ''} style="${cellStyle(c.props, look, r, x, R, C)}">${pbIn(c.props.html || '<br>')}</td>`; }).join('') + '</tr>').join('') + '</tbody></table>';
    }
    else if (b.kind === 'toc') {
      const oneLine = h => plainOf(h).replace(/\s+/g, ' ').trim(), levels = new Map(blocks.filter(h => h.kind === 'heading').map(h => [oneLine(h.props.html), +h.props.level || 1]));
      const entries = (b.props.text || '').split('\n').filter(Boolean).map(line => { const tab = line.lastIndexOf('\t'); const text = tab < 0 ? line : line.slice(0, tab); return { text, page: tab < 0 ? '' : line.slice(tab + 1), level: levels.get(text.replace(/\s+/g, ' ').trim()) || 1 }; });
      out += tocHtml({ path: b.path, levels: b.props.levels, title: b.props.title, entries, style: b.props.style });
    }
    else if (b.kind === 'image') out += picHtml(b);
    else if (b.kind === 'pagebreak') out += `<hr data-pb="1" data-path="${esc(b.path)}">`;
    else out += pa(b, 'p');
  }
  closeLists(0);
  return out || '<p><br></p>';
}

// ----- docx: page setup, header and footer (document props) -----
const PAPERS = ['A4', 'Letter', 'A5', 'B5', 'A3', 'Legal'], MARGINS = ['narrow', 'normal', 'moderate', 'wide'];
/** The editor's page model from the document props: margins are a preset (DocxSection.Margins) or the four lengths the engine
 *  prints (2.54cm 3.18cm 2.54cm 3.18cm); a paper size the editor cannot show falls back to A4 but stays in `raw`. */
export function pageOf(p) {
  p = p || {};
  return { size: PAPERS.includes(p.page) ? p.page : 'A4', orient: p.orientation === 'landscape' ? 'landscape' : 'portrait', margin: MARGINS.includes(p.margin) || /\d/.test(p.margin || '') ? p.margin : 'normal', cols: Math.max(1, +p.columns || 1), raw: { size: p.page || '', margin: p.margin || '' } };
}
/** Header or footer text for the editor: inline html → plain lines, {page}/{pages} kept. */
export function plainOf(html) {
  return String(html || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&').trim();
}
/** Header and footer html as the engine reads and writes it (document props); the editor keeps it as it came until edited. */
const HF = ['header', 'footer', 'firstHeader', 'firstFooter'];
/** The `set <file> /` props that take the file from `orig` (raw engine values) to the model: only what differs, so a header
 * or footer the user did not edit (a logo, a table in it) is never written. */
export function pageDiff(orig, doc) {
  orig = orig || {};
  const o = pageOf(orig.page), n = Object.assign(pageOf(null), doc.page || {}), p = {};
  if (n.size !== o.size) p.page = n.size;
  if (n.orient !== o.orient) p.orientation = n.orient;
  if (n.margin !== o.margin) p.margin = n.margin;
  if (+n.cols !== +o.cols) p.columns = String(+n.cols || 1);
  for (const k of HF) if ((doc[k] || '') !== (orig[k] || '')) p[k] = doc[k] || '';
  if (!!doc.titlePg !== !!orig.titlePg) p.titlePg = doc.titlePg ? 'true' : 'false';
  for (const k of ['lineNumbers', 'hyphenation']) if (!!doc[k] !== !!orig[k]) p[k] = doc[k] ? 'true' : 'false';
  if ((doc.noteFormat || '') !== (orig.noteFormat || '')) p.noteFormat = doc.noteFormat || 'none';
  return p;
}

const hfOf = p => Object.assign(Object.fromEntries(HF.map(k => [k, p[k] || ''])), { titlePg: p.titlePg === 'true' });
async function openDocx(doc) {
  const t = await tree(doc.path);
  const p = t.props || {};
  const body = (t.children || []).find(c => c.kind === 'body') || { children: [] };
  const blocks = blocksOf(body.children, doc.path);
  const page = Object.assign({ hf: true }, doc.page || {}, pageOf(p));
  const comments = commentsOf(body.children, p.author || 'Writer'), track = p.track === 'true'; // the engine writes comments as the document's author, else Writer
  const notes = notesOf(body.children), eqs = eqsOf(body.children), shapes = shapesOf(body.children);
  const html = inkFills(parseHtml(anchorObjects(anchorNotes(anchorComments(trackHtml(blocksToHtml(blocks)), comments), notes, p.noteFormat), eqs, shapes))).innerHTML;
  const flags = { lineNumbers: p.lineNumbers === 'true', hyphenation: p.hyphenation === 'true', noteFormat: p.noteFormat || '' };
  return { html, rev: (doc.rev || 0) + 1, track, comments: comments.map(c => ({ id: c.cid, author: c.author, initials: c.initials, mine: c.mine, time: c.time, text: c.text, quote: c.quote, path: c.path, resolved: c.resolved, parent: c.parent })),
    notes: notes.map(x => ({ id: x.nid, kind: x.kind, text: x.text })), styles: stylesOf(p.styles), base: t.computed || {}, styleEdits: [], page, ...hfOf(p), ...flags,
    _orig: { blocks, ids: uniquePictureIds(body.children), page: { page: p.page, orientation: p.orientation, margin: p.margin, columns: p.columns }, ...hfOf(p), ...flags, track, comments, notes, eqs, shapes } };
}

// ----- docx: footnotes and endnotes -----
/** Every note in the tree: its engine id (also the editor id, `nid`), kind, text, the paragraph it hangs on and the character offset of
 *  its mark there. */
export function notesOf(nodes) {
  const out = [];
  const walk = list => { for (const n of list || []) { if (n.kind === 'footnote') { const q = n.props || {}; out.push({ nid: String(q.id), id: String(q.id), kind: q.kind || 'footnote', text: q.text || '', path: n.path.replace(/\/footnote\[[^\]]*\]$/, ''), at: +q.at || 0 }); } walk(n.children); } };
  walk(nodes);
  return out;
}
/** A note's mark in the text: a superscript number the caret steps over; numbered by order of its kind. */
const NOTE_MARK = 'sup[data-fn]';
/** What sits in a paragraph beside its text: deleted text, notes' marks, equations and shapes. Character offsets leave them out. */
const MARKS = 'del,sup[data-fn],span[data-eq],span[data-shape]';
/** The character offset of `node` in block `el`: the visible text before it (a br and a page break one character each). */
export function offsetIn(el, node) {
  let at = 0; const w = el.ownerDocument.createTreeWalker(el, 5); let n;
  while ((n = w.nextNode()) && n !== node) {
    if (n.nodeType === 1) { if (n.tagName === 'BR' || n.hasAttribute('data-pb')) at++; continue; }
    if (!n.parentElement.closest(MARKS)) at += n.nodeValue.replace(/\u200B/g, '').length;
  }
  return at;
}
/** Puts `mark` at character offset `at` of block `el` (after the last text when the paragraph is shorter). */
function placeAt(el, at, mark) {
  const d = el.ownerDocument, w = d.createTreeWalker(el, 5); let n, pos = 0;
  while ((n = w.nextNode())) {
    if (n.nodeType === 1) { if (n.tagName === 'BR' || n.hasAttribute('data-pb')) pos++; continue; }
    if (n.parentElement.closest(MARKS)) continue;
    const len = n.nodeValue.replace(/\u200B/g, '').length;
    if (at <= pos + len) { const r = d.createRange(); r.setStart(n, Math.min(n.nodeValue.length, at - pos)); r.collapse(true); r.insertNode(mark); return; }
    pos += len;
  }
  el.appendChild(mark);
}
const noteMark = (d, x) => { const s = d.createElement('sup'); s.setAttribute('data-fn', x.nid); s.setAttribute('data-kind', x.kind); s.setAttribute('contenteditable', 'false'); s.textContent = '?'; return s; };
/** Puts each note's mark at its character offset in its paragraph (deleted text and other marks not counted, a br one character). */
export function anchorNotes(html, notes, format) {
  if (!notes.length) return html;
  const root = parseHtml(html), d = root.ownerDocument;
  for (const x of notes) { const el = root.querySelector(`[data-path="${x.path}"]`); if (el) placeAt(el, x.at, noteMark(d, x)); }
  numberNotes(root, format);
  return root.innerHTML;
}
/** Numbers the marks 1… per kind in reading order, as Word does: in the document's noteFormat (DocxSection.NoteFormats) for both
 *  kinds, else footnotes 1, 2, 3 and endnotes i, ii, iii. Returns the marks. */
export function numberNotes(root, format) {
  const count = {}, marks = Array.from(root.querySelectorAll(NOTE_MARK)), num = NOTE_NUMBERS[format];
  for (const s of marks) { const k = s.getAttribute('data-kind') || 'footnote'; count[k] = (count[k] || 0) + 1; const n = num ? num(count[k]) : k === 'endnote' ? toRoman(count[k]) : String(count[k]); if (s.textContent !== n) s.textContent = n; }
  return marks;
}
const toRoman = n => [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']].reduce((s, [v, r]) => { while (n >= v) { s += r; n -= v; } return s; }, '');
const CN_DIGITS = '〇一二三四五六七八九'; // i18n-ok — Word's 一二三 numbering itself
/** Word's numFmt as it draws the numbers: letters run a…z, then aa, bb…; circled numbers stop at ⑳; 一二三 count to 九十九. */
const NOTE_NUMBERS = { decimal: String, lowerRoman: toRoman, upperRoman: n => toRoman(n).toUpperCase(),
  lowerLetter: n => String.fromCharCode(97 + (n - 1) % 26).repeat(Math.floor((n - 1) / 26) + 1), upperLetter: n => NOTE_NUMBERS.lowerLetter(n).toUpperCase(),
  decimalEnclosedCircleChinese: n => n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n),
  chineseCounting: n => n < 10 ? CN_DIGITS[n] : n < 100 ? (n < 20 ? '' : CN_DIGITS[Math.floor(n / 10)]) + '十' + (n % 10 ? CN_DIGITS[n % 10] : '') : String(n) }; // i18n-ok
/** The editor's notes as the engine sees them: the paragraph each mark sits in now and its character offset there. */
function notesIn(doc, el, blocks) {
  const cells = blocks.filter(b => b.rows).flatMap(b => b.rows.flatMap(r => r.cells));
  return (doc.notes || []).map(x => {
    const mark = el.querySelector(`${NOTE_MARK.slice(0, 3)}[data-fn="${x.id}"]`);
    const b = mark && (blocks.find(y => (y.kind === 'paragraph' || y.kind === 'heading') && y.el && y.el.contains(mark)) || cells.find(y => y.el && y.el.contains(mark)));
    if (!b) return { nid: String(x.id), kind: x.kind, text: x.text || '', parent: null, at: 0 };
    return { nid: String(x.id), kind: x.kind, text: x.text || '', parent: b.kind === 'cell' ? b.path + '/paragraph[1]' : b.path, at: offsetIn(b.el, mark) };
  });
}
/** Commands that take the file's notes (`orig`) to the editor's (`current`): a new mark adds a note where it sits, a moved one is
 *  added again there, changed text is set, a mark gone takes its note away. Returns the count and the list to remember. */
export async function planNotes(file, orig, current, log, exec = run) {
  let n = 0; const list = [], byId = new Map((orig || []).map(o => [o.nid, o]));
  const add = async x => { const r = await exec(['add', file, x.parent, '--type', 'footnote', '--prop', 'kind=' + x.kind, '--prop', 'text=' + x.text, '--prop', 'at=' + x.at]); n++; log && log('add', r.path, 'footnote'); const p = r.props || {}; return { nid: x.nid, id: String(p.id), kind: x.kind, text: x.text, path: x.parent, at: x.at }; };
  for (const x of current) {
    const o = byId.get(x.nid);
    if (!x.parent) continue; // its mark is gone from the text: removed below
    try {
      if (!o) list.push(await add(x));
      else if (o.path !== x.parent || o.at !== x.at) { await exec(['remove', file, `//footnote[@id=${o.id}]`]); n++; list.push(await add(x)); }
      else { if ((o.text || '') !== x.text) { await exec(['set', file, `//footnote[@id=${o.id}]`, '--prop', 'text=' + x.text]); n++; log && log('set', o.path, 'footnote'); } list.push(Object.assign({}, o, { text: x.text })); }
    } catch (e) { log && log('skip', x.parent, e.message); }
  }
  const keep = new Set(list.map(x => x.nid));
  for (const o of orig || []) if (!keep.has(o.nid)) { try { await exec(['remove', file, `//footnote[@id=${o.id}]`]); n++; log && log('remove', o.path, 'footnote'); } catch (e) { log && log('skip', o.path, e.message); } }
  return { count: n, list };
}

// ----- docx: equations and shapes in paragraphs -----
/** Every equation in the tree: the paragraph it sits in, its offset there, its LaTeX and whether it has a line of its own. */
export function eqsOf(nodes) {
  const out = [];
  const walk = list => { for (const n of list || []) { if (n.kind === 'equation') { const q = n.props || {}; out.push({ path: n.path.replace(/\/equation\[[^\]]*\]$/, ''), at: +q.at || 0, latex: q.latex || '', display: q.display === 'true' }); } walk(n.children); } };
  walk(nodes);
  return out;
}
/** A shape's props as the editor keeps them on its element (data-w-*), in the order the engine takes them. */
export const SHAPE_PROPS = ['geometry', 'width', 'height', 'fill', 'line', 'wrap', 'xFrom', 'yFrom', 'x', 'y', 'xAlign', 'yAlign'];
/** Every Word shape in the tree: its id, the paragraph it floats in, its props and text. */
export function shapesOf(nodes) {
  const out = [];
  const walk = list => { for (const n of list || []) { if (n.kind === 'shape') { const q = n.props || {}; out.push({ id: String(q.id), path: n.path.replace(/\/shape\[[^\]]*\]$/, ''), props: Object.assign(Object.fromEntries(SHAPE_PROPS.filter(k => q[k] != null && q[k] !== '').map(k => [k, String(q[k])])), { text: q.text || '' }) }); } walk(n.children); } };
  walk(nodes);
  return out;
}
/** A shape's look as an SVG over its box: the preset's outline (the pptx editor's polygons), its fill and outline colours. */
export function shapeSvg(geometry, fill, line) {
  const a = `fill="${!fill || fill === 'none' ? 'none' : '#' + esc(fill)}" stroke="${!line || line === 'none' ? 'none' : '#' + esc(line)}" stroke-width="1.5" vector-effect="non-scaling-stroke"`;
  const body = geometry === 'ellipse' ? `<ellipse cx="50" cy="50" rx="50" ry="50" ${a}/>` : geometry === 'roundRect' ? `<rect width="100" height="100" rx="12" ry="12" ${a}/>`
    : POLY[geometry] ? `<polygon points="${POLY[geometry].map(q => q.join(',')).join(' ')}" ${a}/>` : `<rect width="100" height="100" ${a}/>`;
  return `<svg viewBox="0 0 100 100" preserveAspectRatio="none">${body}</svg>`;
}
/** An equation in the text: its LaTeX on the element, drawn by the editor (KaTeX); the caret steps over it. */
export const eqHtml = (latex, display) => `<span data-eq="1" data-latex="${esc(latex)}"${display ? ' data-display="1"' : ''} contenteditable="false"></span>`;
/** A shape in the text: its props as data-w-*, its text in a box of its own that takes typing; drawn and placed by the editor. */
export function shapeHtml(x) {
  const p = x.props || {}, attrs = SHAPE_PROPS.filter(k => p[k]).map(k => ` data-w-${k.toLowerCase()}="${esc(p[k])}"`).join('');
  return `<span data-shape="1"${x.id ? ` data-sid="${esc(x.id)}"` : ''}${attrs} contenteditable="false"><span class="wd-shtext" contenteditable="true">${esc(p.text || '').replace(/\n/g, '<br>')}</span></span>`;
}
/** Puts equations at their offsets and shapes at the start of their paragraphs. */
export function anchorObjects(html, eqs, shapes) {
  if (!eqs.length && !shapes.length) return html;
  const root = parseHtml(html), d = root.ownerDocument, box = d.createElement('div');
  const make = h => { box.innerHTML = h; return box.firstChild; };
  for (const x of eqs) { const el = root.querySelector(`[data-path="${x.path}"]`); if (el) placeAt(el, x.at, make(eqHtml(x.latex, x.display))); }
  for (const x of shapes.slice().reverse()) { const el = root.querySelector(`[data-path="${x.path}"]`); if (el) el.insertBefore(make(shapeHtml(x)), el.firstChild); }
  return root.innerHTML;
}
/** The paragraphs of the editor's blocks (a table cell counts as its first paragraph) with what sits in them: equations with their
 *  offsets, shapes with their props. `was` gives each block's path when last saved. */
function objectsIn(blocks, was) {
  const cells = blocks.filter(b => b.rows).flatMap(b => b.rows.flatMap(r => r.cells));
  return blocks.filter(b => (b.kind === 'paragraph' || b.kind === 'heading') && b.el).concat(cells.filter(c => c.el)).map(b => {
    const own = sel => Array.from(b.el.querySelectorAll(sel)).filter(x => x.parentElement.closest('p,h1,h2,h3,h4,h5,h6,li,blockquote,td,th') === b.el || x.parentElement === b.el);
    const at = path => path ? (b.kind === 'cell' ? path + '/paragraph[1]' : path) : null;
    return { block: b, path: at(b.path), was: at(was.get(b)),
      eqs: own('span[data-eq]').map(x => ({ latex: x.getAttribute('data-latex') || '', at: offsetIn(b.el, x), display: x.hasAttribute('data-display') })),
      shapes: own('span[data-shape]').map(x => ({ el: x, sid: x.getAttribute('data-sid'), props: Object.assign(Object.fromEntries(SHAPE_PROPS.filter(k => x.getAttribute('data-w-' + k.toLowerCase())).map(k => [k, x.getAttribute('data-w-' + k.toLowerCase())])), { text: shapeText(x) }) })) };
  });
}
/** A shape's text: its box's lines. */
const shapeText = el => { const t = el.querySelector('.wd-shtext'); if (!t) return ''; const c = t.cloneNode(true); c.querySelectorAll('br').forEach(b => b.replaceWith('\n')); c.querySelectorAll('div,p').forEach(b => { b.prepend('\n'); }); return c.textContent.replace(/\u200B/g, '').replace(/^\n/, '').replace(/\u00a0/g, ' '); };
/** Commands that give each paragraph the equations the editor shows in it: a paragraph whose equations or text changed has them
 *  taken out and put back where they are now (a rewritten text may have moved them); the others are left alone. */
export async function planEquations(file, orig, paras, changed, exec = run, log) {
  let n = 0; const list = [];
  for (const p of paras) {
    const had = p.was ? (orig || []).filter(e => e.path === p.was) : [], key = l => JSON.stringify(l.map(e => [e.latex, e.at, !!e.display]));
    list.push(...p.eqs.map(e => Object.assign({ path: p.path }, e)));
    if (!had.length && !p.eqs.length) continue;
    if (key(had) === key(p.eqs) && !changed(p.block)) continue;
    for (let k = had.length; k >= 1; k--) { try { await exec(['remove', file, `${p.path}/equation[${k}]`]); n++; } catch (e) { log && log('skip', p.path, e.message); } }
    for (const e of p.eqs) { try { await exec(['add', file, p.path, '--type', 'equation', '--prop', 'latex=' + e.latex, '--prop', 'at=' + e.at, ...(e.display ? ['--prop', 'display=true'] : [])]); n++; log && log('add', p.path, 'equation'); } catch (err) { log && log('skip', p.path, err.message); } }
  }
  return { count: n, list };
}
/** Commands that take the file's shapes to the editor's, by id: a new one is added in its paragraph (its element learns the id), one that
 *  moved to another paragraph is added again there, changed props are set, one gone is removed. */
export async function planShapes(file, orig, paras, exec = run, log) {
  let n = 0; const list = [], seen = new Set(), byId = new Map((orig || []).map(o => [o.id, o]));
  const add = async (path, x) => { const r = await exec(['add', file, path, '--type', 'shape', ...propsArgs(x.props)]); n++; log && log('add', path, 'shape'); const id = String((r && r.props || {}).id); if (x.el) x.el.setAttribute('data-sid', id); return { id, path, props: x.props }; };
  for (const p of paras) for (const x of p.shapes) {
    const o = x.sid && byId.get(x.sid);
    try {
      if (!o || seen.has(o.id)) { list.push(await add(p.path, x)); continue; }
      seen.add(o.id);
      if (o.path !== p.was) { try { await exec(['remove', file, `//shape[@id=${o.id}]`]); n++; } catch (e) { } list.push(await add(p.path, x)); continue; }
      const diff = {}; for (const k of SHAPE_PROPS.concat(['text'])) if ((o.props[k] || '') !== (x.props[k] || '') && (x.props[k] || k === 'text')) diff[k] = x.props[k] || '';
      if (Object.keys(diff).length) { await exec(['set', file, `//shape[@id=${o.id}]`, ...propsArgs(diff)]); n++; log && log('set', p.path, 'shape'); }
      list.push({ id: o.id, path: p.path, props: x.props });
    } catch (e) { log && log('skip', p.path, e.message); }
  }
  for (const o of orig || []) if (!seen.has(o.id)) { try { await exec(['remove', file, `//shape[@id=${o.id}]`]); n++; log && log('remove', o.path, 'shape'); } catch (e) { log && log('skip', o.path, e.message); } }
  return { count: n, list };
}

// ----- docx: the style gallery -----
/** The document's styles as the engine lists them (its styles prop), [] when it has none. */
export function stylesOf(json) { try { const v = typeof json === 'string' ? JSON.parse(json) : json; return Array.isArray(v) ? v : []; } catch (e) { return []; } }
/** A document's fonts as a CSS family: the Latin one, the East Asian one (Word picks per character), then the bundled face of the
 *  same kind for a machine that has neither (思源宋体 after 宋体, Times, Georgia…, else 思源黑体). */
export function familyCss(font, fontEa) {
  const fs = [...new Set([font, fontEa].filter(Boolean).map(f => String(f).replace(/'/g, '')))];
  if (!fs.length) return '';
  const serif = /song|宋|明|ming|kai|楷|fang|仿|times|georgia|cambria|garamond|palatino|century|serif/i.test(fs[0]) && !/sans|hei|黑/i.test(fs[0]);
  return fs.map(f => `'${f}'`).concat(serif ? ["'Noto Serif SC'", 'serif'] : ["'Noto Sans SC'", 'sans-serif']).join(',');
}
/** The CSS a style's look draws: font, size, weight, slant, colour, alignment, spacing, first-line indent. */
export function lookCss(look) {
  const l = look || {}, css = [];
  if (l.font || l.fontEa) css.push('font-family:' + familyCss(l.font, l.fontEa));
  if (l.size) css.push('font-size:' + l.size + 'pt');
  if (l.bold) css.push('font-weight:' + (l.bold === 'true' ? '700' : '400'));
  if (l.italic) css.push('font-style:' + (l.italic === 'true' ? 'italic' : 'normal'));
  if (l.color) css.push('color:#' + l.color);
  const a = alignCss(l.align); if (a) css.push(a);
  const para = paraCss(l); if (para) css.push(para);
  return css.join(';');
}
/** Stylesheet rules that draw the document's own styles in the editor under `scope`: a paragraph style on [data-style] blocks and,
 *  for the heading styles, on h1–h3 (so a document's headings look as Word shows them); a character style on its spans. `base` is the
 *  document's own look (the engine's computed on the root: its default paragraph style over docDefaults, theme fonts resolved): the
 *  text of the page and of plain paragraphs, and what a paragraph style's look builds on. */
export function styleCss(styles, scope, base) {
  const b = base || {}, rules = [];
  if (base) {
    const text = lookCss({ font: b.font, fontEa: b.fontEa, size: b.size, color: b.color }), para = paraCss(b);
    if (text) rules.push(`${scope} > *{${text}}`);
    if (para) rules.push(`${scope} p:not([data-style]){${para}}`);
  }
  return rules.concat((styles || []).map(s => {
    const css = lookCss(base && s.type !== 'character' ? Object.assign({}, b, s.look) : s.look); if (!css) return '';
    const sel = [`${scope} [data-style="${s.id}"]`]; if (s.heading >= 1 && s.heading <= 3) sel.push(`${scope} h${s.heading}:not([data-style])`);
    return `${sel.join(',')}{${css}}`;
  })).filter(Boolean).join('\n');
}

// ----- docx: tracked changes and comments -----
/** The engine's <ins>/<del> become the editor's tracked-change markup (ins[data-t] / del[data-t]). */
export const trackHtml = html => String(html || '').replace(/<(ins|del)\b(?![^>]*\sdata-t=)/gi, '<$1 data-t="1"');
const fmtTime = iso => {
  const d = new Date(iso || ''); if (isNaN(d)) return '';
  const hm = String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  if ((globalThis.$lang ? globalThis.$lang() : 'zh') === 'en') return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) + ' ' + hm;
  return `${d.getMonth() + 1}月${d.getDate()}日 ${hm}`; // i18n-ok: Chinese date format; English uses the branch above
};
/** Every comment in the tree: its engine id (also the editor id, `cid`), the paragraph path it hangs on, author, initials, time,
 * text, quote, and whether it is mine (written as `me`, the author this app writes comments as). */
export function commentsOf(nodes, me) {
  const out = [];
  const walk = list => { for (const n of list || []) { if (n.kind === 'comment') { const q = n.props || {}; out.push({ cid: String(q.id), id: String(q.id), path: n.path.replace(/\/comment\[[^\]]*\]$/, ''), author: q.author || '', initials: q.initials || '', mine: !!me && q.author === me, date: q.date || '', time: [q.author, fmtTime(q.date)].filter(Boolean).join(' · '), text: q.text || '', quote: q.quote || '', resolved: q.resolved === 'true', parent: q.parent ? String(q.parent) : '' }); } walk(n.children); } };
  walk(nodes);
  return out;
}
/** Wraps each comment's quoted text in its paragraph with <span data-cid>, so the side panel can find and highlight it. */
export function anchorComments(html, comments) {
  if (!comments.some(c => c.quote)) return html;
  const root = parseHtml(html), d = root.ownerDocument;
  for (const c of comments) {
    const el = c.quote && root.querySelector(`[data-path="${c.path}"]`); if (!el) continue;
    const nodes = []; const w = d.createTreeWalker(el, 4); let n; while ((n = w.nextNode())) if (!n.parentElement.closest('del')) nodes.push(n); // the engine's quote leaves deleted text out
    const at = nodes.map(x => x.nodeValue).join('').indexOf(c.quote); if (at < 0) continue;
    const r = d.createRange(); let pos = 0, open = false;
    for (const node of nodes) {
      const end = pos + node.nodeValue.length;
      if (!open && at < end) { r.setStart(node, at - pos); open = true; }
      if (open && at + c.quote.length <= end) { r.setEnd(node, at + c.quote.length - pos); break; }
      pos = end;
    }
    const span = d.createElement('span'); span.setAttribute('data-cid', c.cid);
    span.appendChild(r.extractContents()); r.insertNode(span);
  }
  return root.innerHTML;
}
/** The editor's comments as the engine sees them: the paragraph each anchor span sits in now (`parent`, a table cell counts as
 * its first paragraph) and sat in when last saved (`origin`, from `was`), and the quoted text without deleted text. */
function commentsIn(doc, el, blocks, was) {
  const cells = blocks.filter(b => b.rows).flatMap(b => b.rows.flatMap(r => r.cells));
  return (doc.comments || []).map(c => {
    const span = el.querySelector(`[data-cid="${c.parent || c.id}"]`); // a reply sits where its comment does
    const b = span && (blocks.find(x => (x.kind === 'paragraph' || x.kind === 'heading') && x.el && x.el.contains(span)) || cells.find(x => x.el && x.el.contains(span)));
    const at = path => path ? (b.kind === 'cell' ? path + '/paragraph[1]' : path) : null;
    const q = span && span.cloneNode(true); if (q) q.querySelectorAll('del').forEach(x => x.remove());
    return { cid: String(c.id), thread: c.parent ? String(c.parent) : '', text: c.text || '', resolved: !!c.resolved, parent: b ? at(b.path) : null, origin: b ? at(was.get(b)) : null,
      quote: q ? q.textContent.replace(/\u00a0/g, ' ').replace(/\u200B/g, '') : '' };
  });
}
/** Commands that take the file's comments (`orig`, as saved last) to the editor's (`current`): add / set / remove by engine id.
 * A comment whose anchor moved to another paragraph (paragraphs merged, text cut and pasted) is anchored again there, keeping its
 * author, initials and date. A quote the engine cannot find anchors the whole paragraph. A comment already gone with its paragraph is
 * dropped. Returns the count of commands and the list to remember as the new orig; a new comment that hangs on nothing waits. */
export async function planComments(file, orig, current, log, exec = run) {
  let n = 0; const list = [];
  const byCid = new Map((orig || []).map(o => [o.cid, o]));
  const tolerant = async argv => { try { await exec(argv); n++; log && log(argv[0], argv[2]); return true; } catch (e) { log && log('skip', argv[2], e.message); return false; } };
  const add = async (c, keep) => {
    const up = c.thread && (list.find(x => x.cid === c.thread) || byCid.get(c.thread)); // a reply goes beside its comment, by that comment's id now
    const props = Object.assign({ text: c.text }, keep || {}, up ? { parent: up.id } : c.quote ? { quote: c.quote } : {}, c.resolved ? { resolved: 'true' } : {});
    let r;
    try { r = await exec(['add', file, c.parent, '--type', 'comment', ...propsArgs(props)]); }
    catch (e) { if (!props.quote) throw e; delete props.quote; r = await exec(['add', file, c.parent, '--type', 'comment', ...propsArgs(props)]); }
    n++; log && log('add', r.path, 'comment');
    const p = r.props || {};
    return { cid: c.cid, id: String(p.id), path: c.parent, author: p.author || '', initials: p.initials || '', date: p.date || '', text: c.text, resolved: c.resolved };
  };
  for (const c of current) {
    const o = byCid.get(c.cid);
    try {
      if (o && c.parent && c.origin !== o.path) {
        await tolerant(['remove', file, `//comment[@id=${o.id}]`]);
        list.push(await add(c, Object.assign({}, o.author ? { author: o.author } : {}, o.initials ? { initials: o.initials } : {}, o.date ? { date: o.date } : {})));
      } else if (o) {
        const diff = {};
        if ((o.text || '') !== c.text) diff.text = c.text;
        if (!!o.resolved !== c.resolved) diff.resolved = c.resolved ? 'true' : 'false';
        if (Object.keys(diff).length && !await tolerant(['set', file, `//comment[@id=${o.id}]`, ...propsArgs(diff)])) continue;
        list.push(Object.assign({}, o, { text: c.text, resolved: c.resolved }, c.parent ? { path: c.parent } : {}));
      } else if (c.parent) list.push(await add(c));
    } catch (e) { log && log('skip', c.parent, e.message); }
  }
  const keep = new Set(current.map(c => c.cid));
  for (const o of orig || []) if (!keep.has(o.cid)) await tolerant(['remove', file, `//comment[@id=${o.id}]`]);
  return { count: n, list };
}

// ----- docx: html → blocks -----
const BLOCK = /^(P|DIV|H1|H2|H3|H4|H5|H6|UL|OL|LI|PRE|BLOCKQUOTE|TABLE|IMG|HR|FIGURE|SECTION|ARTICLE)$/;
function parseHtml(html) { return new DOMParser().parseFromString('<body>' + html + '</body>', 'text/html').body; }
const docxAttrs = (el, keys) => Object.fromEntries(keys.filter(k => el.hasAttribute('data-w-' + k.toLowerCase())).map(k => [k, el.getAttribute('data-w-' + k.toLowerCase())]));

/** Blocks in document order from an editor root (element or html). Each block keeps its element so paths can be written back. */
export function blocksFromHtml(root) {
  const el = typeof root === 'string' ? parseHtml(root) : root;
  const out = [];
  const walk = (node, listType, level) => {
    for (const c of Array.from(node.childNodes)) {
      if (c.nodeType === 3) { if (c.nodeValue.trim()) out.push({ kind: 'paragraph', path: null, props: { html: esc(c.nodeValue.trim()) }, el: null }); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.tagName;
      if (tag === 'UL' || tag === 'OL') { walk(c, listKind(c, listType), level + 1); continue; }
      if (tag === 'LI') {
        const inner = c.cloneNode(true); Array.from(inner.querySelectorAll('ul,ol')).forEach(x => x.remove());
        const restart = !c.previousElementSibling && node.getAttribute && node.hasAttribute('data-w-restart') ? { restart: 'true' } : {};
        out.push(withPics({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(inner), list: listType || 'bullet', level: String(Math.max(0, level)) }, restart, paraAttrs(c)), el: c, align: alignOf(c) }, picsIn(c)));
        Array.from(c.children).filter(x => /^(UL|OL)$/.test(x.tagName)).forEach(x => walk(x, listKind(x, listType), level + 1));
        continue;
      }
      if (/^H[1-6]$/.test(tag)) {
        const st = c.getAttribute('data-style');
        if (st) out.push(withPics({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), style: st }, paraAttrs(c)), el: c, align: alignOf(c) }, picsIn(c)));
        else out.push(withPics({ kind: 'heading', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), level: c.getAttribute('data-level') || tag[1] }, paraAttrs(c)), el: c, align: alignOf(c) }, picsIn(c)));
        continue;
      }
      if (tag === 'PRE') { out.push({ kind: 'code', path: pathOf(c), props: { text: c.innerText.replace(/\n$/, '') }, el: c }); continue; }
      if (tag === 'BLOCKQUOTE') { out.push(withPics({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), style: c.getAttribute('data-style') || 'Quote' }, paraAttrs(c)), el: c, align: alignOf(c) }, picsIn(c))); continue; }
      if (c.hasAttribute('data-toc')) { out.push({ kind: 'toc', path: pathOf(c), props: { levels: c.getAttribute('data-levels') ?? '3', title: c.getAttribute('data-title') ?? '目录', style: c.getAttribute('data-toc-style') || 'classic' }, refresh: c.hasAttribute('data-refresh'), el: c }); continue; }
      if (tag === 'TABLE') {
        let ops = []; try { ops = JSON.parse(c.getAttribute('data-ops') || '[]'); } catch (e) { ops = null; } // unreadable: the save writes the table anew
        out.push({ kind: 'table', path: pathOf(c), el: c, ops, props: docxAttrs(c, ['style', 'header', 'borders', 'borderColor', 'width', 'widths', 'align']),
          rows: Array.from(c.querySelectorAll('tr')).filter(tr => tr.closest('table') === c).map(tr => ({ kind: 'row', path: pathOf(tr), el: tr,
            props: docxAttrs(tr, ['header', 'height']), cells: Array.from(tr.children).filter(td => /^(TD|TH)$/.test(td.tagName)).map(td => ({ kind: 'cell', path: pathOf(td), el: td,
              props: Object.assign({ html: inlineHtml(td) }, docxAttrs(td, ['fill', 'borders', 'valign', 'width', 'align']), td.colSpan > 1 ? { colspan: String(td.colSpan) } : {}, td.rowSpan > 1 ? { rowspan: String(td.rowSpan) } : {}) })) })) });
        continue;
      }
      if (tag === 'IMG') { if (isFloat(c)) out.push(withPics({ kind: 'paragraph', path: null, props: { html: '' }, el: null }, [picOfEl(c)])); else out.push(imgBlock(c)); continue; } // a floating picture on its own: a paragraph holds it
      if (tag === 'HR') { if (c.getAttribute('data-pb')) out.push({ kind: 'pagebreak', path: pathOf(c), props: {}, el: c }); continue; }
      if (tag === 'BR' && node === el) continue;
      if (tag === 'P' || tag === 'DIV' || tag === 'FIGURE' || tag === 'SECTION' || tag === 'ARTICLE') {
        if (Array.from(c.children).some(x => BLOCK.test(x.tagName) && !/^(IMG|BR|FIGURE)$/.test(x.tagName))) { walk(c, listType, level); continue; }
        const pics = picsIn(c);
        if (pics.length === 1 && !c.textContent.trim() && !isFloat(pics[0].el)) { out.push(imgBlock(pics[0].el)); continue; } // a picture of its own: a block
        const st = c.getAttribute('data-style');
        out.push(withPics({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c) }, st ? { style: st } : {}, paraAttrs(c)), el: c, align: alignOf(c) }, pics));
        continue;
      }
      // inline content at block level (a stray span or text)
      out.push({ kind: 'paragraph', path: null, props: { html: inlineHtml(c) }, el: null });
    }
  };
  walk(el, null, -1);
  return out;
}
const pathOf = el => el.getAttribute && el.getAttribute('data-path') || null;
/** The kind of list an ol or ul is: a ul bullets; an ol what its data-w-list says, else its parent list's numbering (a nested
 *  level of the same list), else plain numbers. */
const listKind = (el, parent) => el.tagName === 'UL' ? 'bullet' : el.getAttribute('data-w-list') || (parent && parent !== 'bullet' ? parent : 'number');
/** A heading or paragraph's own props that its element keeps as data-w-* (the file's; not what its style gives), and the value
 *  that turns each off. */
const SECTION_PROPS = ['page', 'orientation', 'margin', 'columns']; // of the section a paragraph ends: only with its sectionBreak
const PARA_OWN = ['pageBreakBefore', 'fill', 'lineSpacing', 'spaceBefore', 'spaceAfter', 'indentLeft', 'indentRight', 'indentFirst', 'border', 'keepNext', 'keepLines', 'widowControl', 'tabs', 'bookmark', 'caption', 'dropCap', 'sectionBreak', ...SECTION_PROPS];
const PARA_OFF = Object.fromEntries(PARA_OWN.map(k => [k, k === 'pageBreakBefore' || k === 'keepNext' || k === 'keepLines' ? 'false' : 'none']));
const paraAttrs = el => docxAttrs(el, PARA_OWN);
/** A Word length as CSS: characters (2ch) as em, lines (0.5lines, Word's 行, 12pt each) as pt, cm and pt as they are. */
const lenCss = v => /(ch|em)$/i.test(v) ? parseFloat(v) + 'em' : /lines?$/i.test(v) ? Math.round(parseFloat(v) * 1200) / 100 + 'pt' : v;
/** text-align for a paragraph alignment; distribute (分散对齐) also spreads the last line. */
const alignCss = a => !a || a === 'left' ? '' : a === 'distribute' ? 'text-align:justify;text-align-last:justify' : 'text-align:' + a;
/** The CSS a paragraph's own 段落 settings draw: line spacing (a multiple of Word's single, 1.5 here, so the template's 1.15 is the
 *  editor's usual 1.8; an exact height as it is), space before and after, indents (a hanging indent pulls the first line back),
 *  borders. Tab stops and keep-with-next draw nothing; the ruler shows the stops. */
export function paraCss(p) {
  const css = [];
  if (p.lineSpacing && p.lineSpacing !== 'none') { const m = /^min\s+(.+)$/i.exec(p.lineSpacing), v = m ? m[1] : p.lineSpacing; css.push(isFinite(v) ? 'line-height:' + Math.round(+v * 150) / 100 : m ? 'min-height:' + v : 'line-height:' + v); }
  if (p.spaceBefore) css.push('margin-top:' + lenCss(p.spaceBefore));
  if (p.spaceAfter) css.push('margin-bottom:' + lenCss(p.spaceAfter));
  if (p.indentLeft) css.push('margin-left:' + lenCss(p.indentLeft));
  if (p.indentRight) css.push('margin-right:' + lenCss(p.indentRight));
  if (p.indentFirst) { const v = lenCss(p.indentFirst.replace(/^-/, '')); css.push(p.indentFirst.startsWith('-') ? `text-indent:-${v};padding-left:${v}` : 'text-indent:' + v); }
  if (p.border && p.border !== 'none') { for (const side of p.border === 'box' ? ['top', 'bottom', 'left', 'right'] : p.border.split(/[\s,]+/)) css.push(`border-${side}:1px solid currentColor`); css.push('padding:1px 4px'); }
  return css.join(';');
}
const PARA_STYLE = ['line-height', 'min-height', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right', 'text-indent', 'padding-left', 'padding', 'border-top', 'border-bottom', 'border-left', 'border-right'];
/** Draws a live paragraph again from its data-w-* settings (after the editor changed one): the CSS of paraCss and its own shading,
 *  leaving its alignment and everything else alone. */
export function drawPara(el) {
  const p = docxAttrs(el, PARA_OWN);
  for (const k of PARA_STYLE) el.style.removeProperty(k);
  for (const d of paraCss(p).split(';')) { const i = d.indexOf(':'); if (i > 0) el.style.setProperty(d.slice(0, i), d.slice(i + 1)); }
  if ('fill' in p) el.style.background = p.fill === 'none' ? '' : '#' + p.fill;
  return el;
}
/** The tab stops of a paragraph's tabs setting: [{ kind, cm }]. */
export const tabStopsOf = tabs => String(tabs || '').split(/[,;]/).map(s => s.trim().split(/\s+/)).filter(x => x.length === 2 && isFinite(parseFloat(x[1]))).map(([kind, pos]) => ({ kind, cm: Math.round(cmOf(pos) * 100) / 100 }));
export const tabsText = stops => stops.map(t => `${t.kind} ${t.cm}cm`).join(', ');
/** A picture in the editor: its path (by id, once saved), look, place (picture.js PLACE; inline unless it says otherwise), frame in px,
 *  and `first` when it stands before its paragraph's text. */
const picOfEl = img => ({ path: pathOf(img), el: img, props: Object.assign({ src: img.getAttribute('src') || '' }, docxAttrs(img, LOOK)), place: Object.assign({ wrap: 'inline' }, docxAttrs(img, PLACE)), first: img.hasAttribute('data-w-first'),
  width: +img.getAttribute('data-fw') || (img.style.width ? parseFloat(img.style.width) : 0), height: +img.getAttribute('data-fh') || 0 });
const imgBlock = img => Object.assign({ kind: 'image' }, picOfEl(img));
const isFloat = img => (img.getAttribute('data-w-wrap') || 'inline') !== 'inline';
const PIC_BLOCKS = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,td,th';
/** The pictures of one block (not those of a block nested in it: a list's sub-items, a cell's paragraphs). */
const picsIn = el => { const nested = Array.from(el.querySelectorAll(PIC_BLOCKS)); return Array.from(el.querySelectorAll('img')).filter(i => !nested.some(b => b.contains(i))).map(picOfEl); };
const withPics = (b, pics) => { if (pics.length) b.pics = pics; return b; };
const alignOf = el => { const st = el.style || {}, a = st.textAlign; if (!a || a === 'start' || a === 'left') return null; return a === 'justify' && st.textAlignLast === 'justify' ? 'distribute' : a === 'end' ? 'right' : a; };
/** A block's inline html for the engine: without the AI change marks, comment anchor spans (the file keeps anchors itself), the
 *  editor's zero-width fillers and the block's pictures (they are the block's `pics`), and with its page break lines as Word's page breaks. */
function inlineHtml(el) {
  const c = el.cloneNode(true);
  Array.from(c.querySelectorAll('img,figure[data-pic]')).forEach(x => x.remove());
  Array.from(c.querySelectorAll('[data-ai]')).forEach(x => x.removeAttribute('data-ai'));
  Array.from(c.querySelectorAll('[data-cid]')).forEach(x => { while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x); x.remove(); });
  Array.from(c.querySelectorAll('sup[data-fn],span[data-eq],span[data-shape]')).forEach(x => x.remove());
  return pbOut(c.innerHTML.replace(/\u200B/g, ''));
}

// ----- canonical runs, so browser markup and engine markup compare equal -----
const COLOR_NAMES = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080' };
export function colorHex(v) {
  if (!v) return null; v = String(v).trim().toLowerCase();
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/.exec(v);
  if (m) { if (m[4] !== undefined && +m[4] === 0) return null; return [m[1], m[2], m[3]].map(x => Math.min(255, +x).toString(16).padStart(2, '0')).join('').toUpperCase(); }
  if (v[0] === '#') { const h = v.slice(1); return (h.length === 3 ? h.split('').map(c => c + c).join('') : h.slice(0, 6)).toUpperCase(); }
  if (v === 'transparent' || v === 'inherit' || v === 'initial') return null;
  return COLOR_NAMES[v] || null;
}
/** Word's automatic text colour follows the colour behind it, whatever the theme: 'dark' ink on a light fill (RRGGBB), 'light'
 *  on a dark one. ponytail: the W3C brightness cut at half; Word does not document its own, tune here if a fill reads wrong. */
export const inkOn = fill => { const [r, g, b] = [0, 2, 4].map(i => parseInt(fill.substr(i, 2), 16)); return (r * 299 + g * 587 + b * 114) / 1000 < 128 ? 'light' : 'dark'; };
/** Marks every element under root that has a fill of its own (cell shading, a highlight, a pasted background) with the ink
 *  of its automatic text, which the Word editor's CSS draws; text with a colour of its own keeps it. Returns root. */
export function inkFills(root) {
  for (const el of root.querySelectorAll('[style*="background"],[data-ink]')) {
    const fill = colorHex(el.style.backgroundColor);
    const ink = fill && inkOn(fill); if (el.getAttribute('data-ink') !== ink) { if (ink) el.setAttribute('data-ink', ink); else el.removeAttribute('data-ink'); } // an unchanged mark restyles nothing
  }
  return root;
}
function ptOf(v) { const m = /^([\d.]+)\s*(pt|px)$/i.exec(String(v || '').trim()); if (!m) return null; const n = m[2].toLowerCase() === 'px' ? +m[1] * 0.75 : +m[1]; return String(Math.round(n * 2) / 2); }
export function runsOf(html) {
  const root = typeof html === 'string' ? parseHtml(html) : html, out = [];
  const add = (t, s) => { if (!t) return; const last = out[out.length - 1]; if (last && last.s === s) last.t += t; else out.push({ t, s }); };
  const key = f => JSON.stringify([f.b, f.i, f.u, f.s, f.c, f.a, f.color, f.bg, f.size, f.font, f.ins, f.del, f.rs, f.va, f.ls, f.sh, f.ol, f.caps].map(x => x || 0));
  const walk = (node, f) => {
    for (const c of Array.from(node.childNodes)) {
      if (c.nodeType === 3) { add(c.nodeValue.replace(/\u00a0/g, ' ').replace(/\u200B/g, ''), key(f)); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.tagName, st = c.style || {};
      if (tag === 'BR') { add(st.pageBreakBefore === 'always' || st.breakBefore === 'page' ? '\f' : '\n', key(f)); continue; } // a page break is not a line break
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'IMG') continue;
      if (BLOCK.test(tag) && out.length && !out[out.length - 1].t.endsWith('\n')) add('\n', key(f));
      const g = Object.assign({}, f);
      if (tag === 'B' || tag === 'STRONG' || st.fontWeight === 'bold' || st.fontWeight === 'bolder' || +st.fontWeight >= 600) g.b = 1;
      if (st.fontWeight === 'normal') g.b = 0;
      if (tag === 'I' || tag === 'EM' || st.fontStyle === 'italic') g.i = 1;
      if (tag === 'U' || (st.textDecoration || '').includes('underline')) g.u = /^(double|dotted|dashed|wavy)$/.test(st.textDecorationStyle) ? st.textDecorationStyle : 1; // a slide's double, dotted, dashed or wavy underline
      if (tag === 'INS') g.ins = 1;
      if (tag === 'DEL') g.del = 1;
      if (tag === 'S' || tag === 'STRIKE' || (st.textDecoration || '').includes('line-through')) g.s = 1;
      if (tag === 'CODE' || tag === 'TT' || tag === 'KBD') g.c = 1;
      if (tag === 'A' && c.getAttribute('href')) g.a = c.getAttribute('href');
      if (tag === 'MARK') g.bg = 'FFFF00';
      if (c.getAttribute && c.getAttribute('data-style')) g.rs = c.getAttribute('data-style'); // a character style
      if (tag === 'FONT') { if (c.getAttribute('color')) g.color = colorHex(c.getAttribute('color')); if (c.getAttribute('face')) g.font = c.getAttribute('face').split(',')[0].trim().replace(/["']/g, ''); }
      if (st.color) g.color = colorHex(st.color) || g.color;
      if (st.backgroundColor) g.bg = colorHex(st.backgroundColor) || g.bg;
      if (st.fontSize) g.size = ptOf(st.fontSize) || g.size;
      if (st.fontFamily) g.font = st.fontFamily.split(',')[0].trim().replace(/["']/g, '');
      if (tag === 'SUP' || st.verticalAlign === 'super') g.va = 'sup'; else if (tag === 'SUB' || st.verticalAlign === 'sub') g.va = 'sub'; // superscript, subscript
      if (st.letterSpacing && st.letterSpacing !== 'normal') g.ls = st.letterSpacing; // character spacing
      if (st.textTransform) g.caps = st.textTransform === 'uppercase' ? 'all' : 0; // a slide's all caps and small caps
      if (/small-caps/.test(st.fontVariant || st.fontVariantCaps || '')) g.caps = 'small';
      if (st.textShadow && st.textShadow !== 'none') g.sh = 1; // Word's text effects
      const stroke = st.webkitTextStroke || st.WebkitTextStroke || st.webkitTextStrokeWidth || st.WebkitTextStrokeWidth; if (stroke && !/^0(px)?(\s|$)/.test(stroke)) g.ol = 1;
      walk(c, g);
    }
  };
  walk(root, {});
  while (out.length && !out[out.length - 1].t.replace(/\n/g, '').length) out.pop();
  return JSON.stringify(out);
}
const sameRuns = (a, b) => a === b || runsOf(a || '') === runsOf(b || '');

// ----- docx: diff → commands -----
const seg = path => path.slice(path.lastIndexOf('/') + 1);
const kindOfSeg = s => s.replace(/\[.*$/, '');

/** Positions a survivor again after removals and additions: the block before it in the new order sits in the file already, so its
 *  ordinal is its place among the blocks of its kind placed so far (/body/paragraph[3] becomes /body/paragraph[2] when an earlier
 *  paragraph went away, [4] when one was added before it). Keyed segments (cell[B3]) and pictures named by id stay. */
function place(parentPath, o, counts) {
  if (o.kind === 'image' && o.path.startsWith('//')) { o.cur = o.path; return; } // named by id: the same wherever it stands
  const k = o.kind, s = seg(o.path), m = /\[([^\]]*)\]$/.exec(s), keyed = m && !/^-?\d+$/.test(m[1]);
  counts[k] = (counts[k] || 0) + 1;
  o.cur = keyed ? parentPath + '/' + s : parentPath + '/' + k + '[' + counts[k] + ']';
  if (o.rows) o.rows.forEach(r => { r.cur = o.cur + '/' + seg(r.path); (r.cells || []).forEach(c => c.cur = r.cur + '/' + seg(c.path)); });
  if (o.cells) o.cells.forEach(c => c.cur = o.cur + '/' + seg(c.path));
}

function propsArgs(props) { return Object.entries(props).flatMap(([k, v]) => v == null ? [] : ['--prop', k + '=' + v]); }

const blocksModel = rows => gridModel(rows.map(r => ({ id: r.path, ref: r, cells: r.cells.map(c => ({ id: c.path, ref: c, cs: c.props.colspan, rs: c.props.rowspan })) })));

/** A new table is added as a plain grid of texts (data), covered places left empty. */
function tableData(rows) {
  const m = blocksModel(rows), width = Math.max(1, ...coverage(m).map(r => r.length));
  return m.rows.map(row => { const v = Array(width).fill(''); for (const x of row.cells) v[x.c] = textOf(x.ref.props.html); return v; });
}

/** Commands that give a table just added from a plain grid its merges and formatting: colspans and cell props row by row
 * (a cell's place is its column less what merges on its left took), then rowspans top to bottom, left to right, when every
 * cell above is already merged and a cell's place is its place among the visible cells of its row. */
export function newTableCommands(tablePath, rows, file) {
  const m = blocksModel(rows), cmds = [], rowPath = r => `${tablePath}/row[${r + 1}]`;
  m.rows.forEach((row, r) => {
    if (Object.keys(rows[r].props || {}).length) cmds.push(['set', file, rowPath(r), ...propsArgs(rows[r].props)]);
    let taken = 0;
    for (const x of row.cells) {
      const at = `${rowPath(r)}/cell[${x.c + 1 - taken}]`;
      if (x.cs > 1) { cmds.push(['set', file, at, '--prop', 'colspan=' + x.cs]); taken += x.cs - 1; }
      const { colspan, rowspan, html, ...props } = x.ref.props;
      if (/<[a-z]/i.test(html || '') && !/^(<br>|&nbsp;)*$/.test(html || '')) props.html = html;
      if (Object.keys(props).length) cmds.push(['set', file, at, ...propsArgs(props)]);
    }
  });
  m.rows.forEach((row, r) => row.cells.forEach((x, i) => { if (x.rs > 1) cmds.push(['set', file, `${rowPath(r)}/cell[${i + 1}]`, '--prop', 'rowspan=' + x.rs]); }));
  return cmds;
}

function blockProps(b, forNew) {
  const p = {};
  if (b.kind === 'heading') { p.html = b.props.html; p.level = b.props.level; if (b.align || !forNew) p.align = b.align || 'left'; }
  else if (b.kind === 'paragraph') { p.html = b.props.html; p.list = b.props.list || (forNew ? null : 'none'); if (b.props.list) { p.level = b.props.level || '0'; if (b.props.restart) p.restart = 'true'; } if (b.props.style) p.style = b.props.style; if (b.align || !forNew) p.align = b.align || 'left'; }
  if (b.kind === 'heading' || b.kind === 'paragraph') { for (const k of PARA_OWN) if (b.props[k] && (b.props.sectionBreak || !SECTION_PROPS.includes(k))) p[k] = b.props[k]; }
  else if (b.kind === 'code') p.text = b.props.text;
  else if (b.kind === 'image') { p.src = b.props.src; if (b.width) p.width = Math.round(b.width) + 'px'; }
  else if (b.kind === 'table') { p.data = JSON.stringify(tableData(b.rows)); Object.assign(p, b.props); }
  else if (b.kind === 'row') { p.data = JSON.stringify(b.cells.map(c => textOf(c.props.html))); Object.assign(p, b.props); }
  else if (b.kind === 'cell') Object.assign(p, b.props);
  else if (b.kind === 'toc') { p.levels = b.props.levels || '3'; p.title = b.props.title || ''; if (b.props.style && b.props.style !== 'classic') p.style = b.props.style; }
  return p;
}
const textOf = html => { const d = parseHtml(html || ''); return d.innerText.replace(/ /g, ' ').trim(); };

const CLEARED = { fill: 'none', colspan: '1', rowspan: '1', header: 'false', style: '' };

function changedProps(orig, b) {
  const p = {};
  if (b.kind === 'heading' || b.kind === 'paragraph' || b.kind === 'cell') { if (!sameRuns(orig.props.html, b.props.html)) p.html = b.props.html; }
  if (b.kind === 'code' && (orig.props.text || '') !== (b.props.text || '')) p.text = b.props.text;
  if (b.kind === 'image') Object.assign(p, lookDiff(orig.props, b.props));
  if (b.kind === 'heading' && String(orig.props.level) !== String(b.props.level)) p.level = b.props.level;
  if (b.kind === 'heading' || b.kind === 'paragraph') for (const k of PARA_OWN) if ((orig.props[k] || '') !== (b.props[k] || '') && (b.props.sectionBreak || !SECTION_PROPS.includes(k))) p[k] = b.props[k] || PARA_OFF[k];
  if (b.kind === 'paragraph') {
    const ol = orig.props.list || 'none', nl = b.props.list || 'none';
    if (ol !== nl) p.list = nl;
    if (nl !== 'none' && String(orig.props.level || '0') !== String(b.props.level || '0')) p.level = b.props.level || '0';
    if (nl !== 'none' && nl !== 'bullet' && (orig.props.restart || '') !== (b.props.restart || '')) p.restart = b.props.restart || 'false';
    if ((orig.props.style || '') !== (b.props.style || '') && b.props.style) p.style = b.props.style;
  }
  if (b.kind === 'toc') {
    if (String(orig.props.levels || '3') !== String(b.props.levels || '3')) p.levels = b.props.levels || '3';
    if ((orig.props.title || '') !== (b.props.title || '')) p.title = b.props.title || '';
    if ((orig.props.style || 'classic') !== (b.props.style || 'classic')) p.style = b.props.style || 'classic';
  }
  if (b.kind === 'table' || b.kind === 'row' || b.kind === 'cell') {
    const keys = b.kind === 'table' ? ['style', 'header', 'borders', 'borderColor', 'width', 'widths', 'align'] : b.kind === 'row' ? ['header', 'height'] : ['fill', 'colspan', 'rowspan', 'borders', 'valign', 'width', 'align'];
    for (const k of keys) {
      const was = String(orig.props?.[k] ?? ''), now = String(b.props?.[k] ?? '');
      if (was === now) continue;
      if (now) p[k] = now;
      else if (k === 'borders') { if (b.kind === 'table') p[k] = 'style'; } // a cell has no value that brings the table's lines back
      else if (k in CLEARED) p[k] = CLEARED[k]; // width, align, height…: nothing to write that removes them, so they stay
    }
    if (orig.fresh && b.kind === 'cell' && !p.fill) p.fill = b.props?.fill || 'none'; // a new cell copied some row or cell in the engine: say which shading it has
  }
  if (b.kind !== 'code' && b.kind !== 'table' && b.kind !== 'image' && b.kind !== 'row' && b.kind !== 'cell' && b.kind !== 'toc' && b.kind !== 'pagebreak' && (orig.align || orig.props.align || 'left') !== (b.align || b.props.align || 'left')) p.align = b.align || b.props.align || 'left';
  return p;
}

/** Applies the difference between orig children and new children of one container as commands. Returns the number of commands run. */
async function planContainer(file, parentPath, origChildren, newChildren, log, exec = run) {
  let n = 0;
  const byPath = new Map(origChildren.map(o => [o.path, o]));
  const seen = new Set();
  for (const b of newChildren) {
    if (!b.path) continue;
    const o = byPath.get(b.path);
    if (!o || seen.has(b.path) || o.kind !== b.kind) { b.path = null; continue; } // duplicated by the editor, or changed kind: treat as new
    seen.add(b.path);
  }
  const removed = origChildren.filter(o => !seen.has(o.path));
  for (const o of removed.slice().reverse()) { await exec(['remove', file, o.path]); n++; log && log('remove', o.path); }
  // in document order: a new block goes in after the one just placed (kept or new), so a list item added continues the
  // list of the item before it, as it would when typed into Word; the first new block goes to the front
  const counts = {};
  let prev = null;
  for (const b of newChildren) {
    if (b.path) {
      const o = byPath.get(b.path);
      place(parentPath, o, counts);
      if (b.kind === 'table' || b.kind === 'row') {
        const diff = changedProps(o, b);
        if (Object.keys(diff).length) { await exec(['set', file, o.cur, ...propsArgs(diff)]); n++; log && log('set', o.cur, diff); }
        if (b.kind === 'table') {
          // the editor's row, column, merge and split commands first, replayed where the engine has the table now;
          // then the rows and cells, whose paths now follow the replayed table, differ only in content and formatting
          const plan = replayTable(o, b.ops, file);
          for (const argv of plan.cmds) { await exec(argv); n++; log && log(argv[0], argv[2]); }
          n += await planContainer(file, o.cur, plan.rows, b.rows, log, exec);
        } else n += await planContainer(file, o.cur, o.cells, b.cells, log, exec);
      }
      else {
        const diff = changedProps(o, b);
        if (Object.keys(diff).length) { await exec(['set', file, o.cur, ...propsArgs(diff)]); n++; log && log('set', o.cur, diff); }
      }
      prev = o.cur;
      continue;
    }
    const argv = ['add', file, parentPath, '--type', b.kind, ...propsArgs(blockProps(b, true)), ...(prev ? ['--after', prev] : ['--index', '1'])];
    const r = await exec(argv); n++; log && log('add', r.path, b.kind);
    b.path = r.path;
    counts[b.kind] = (counts[b.kind] || 0) + 1;
    if (b.kind === 'image' && r.props && r.props.id != null) b.id = String(r.props.id);
    if (b.kind === 'table') for (const argv of newTableCommands(r.path, b.rows, file)) { await exec(argv); n++; }
    if (b.kind === 'row') for (let ci = 0; ci < b.cells.length; ci++) {
      const { html, ...props } = b.cells[ci].props;
      if (/<[a-z]/i.test(html || '') && !/^(<br>|&nbsp;)*$/.test(html || '')) props.html = html;
      if (Object.keys(props).length) { await exec(['set', file, `${r.path}/cell[${ci + 1}]`, ...propsArgs(props)]); n++; }
    }
    prev = r.path;
  }
  return n;
}

/** The body's commands, as a save runs them: tables whose recorded commands do not fit are written anew. */
export function planDocxBlocks(file, before, after, exec, log) {
  for (const b of tablesToRewrite(before, after)) b.path = null;
  return planContainer(file, '/body', before, after, log, exec);
}

/** Engine paths are ordinals per kind under a parent, so after a save every block's path follows from the final order alone.
 * Writes the paths into the blocks and, when they came from a live editor, into the DOM. */
function renumber(parentPath, blocks) {
  const counts = {};
  for (const b of blocks) {
    if (b.kind === 'image' && /^\/\//.test(b.path || '')) continue; // named by id
    counts[b.kind] = (counts[b.kind] || 0) + 1;
    b.path = parentPath + '/' + b.kind + '[' + counts[b.kind] + ']';
    if (b.el) b.el.setAttribute('data-path', b.path);
    if (b.rows) renumber(b.path, b.rows);
    if (b.cells) renumber(b.path, b.cells);
  }
}
const strip = blocks => blocks.map(b => { const c = Object.assign({}, b); delete c.el; delete c.ops; if (c.rows) c.rows = strip(c.rows); if (c.cells) c.cells = strip(c.cells); if (c.pics) c.pics = strip(c.pics); return c; });

// ----- docx: pictures in paragraphs — moved, added and placed by id, around the block diff -----
/** Which pictures change between a block of their own and a paragraph in this save: `out` the paths of block pictures now in a
 *  paragraph, `in` the blocks that are pictures out of one. Both are the picture phases' business, not the block diff's. */
export function pictureMoves(origBlocks, blocks) {
  const inPara = new Set((origBlocks || []).flatMap(o => (o.pics || []).map(p => p.path))), nowIn = new Set(blocks.flatMap(b => (b.pics || []).map(p => p.path)));
  return { out: new Set((origBlocks || []).filter(o => o.kind === 'image' && nowIn.has(o.path)).map(o => o.path)), in: new Set(blocks.filter(b => b.kind === 'image' && inPara.has(b.path))) };
}
/** Every picture, before: path → { at: 'body' | its paragraph's path, pic }, and after: [{ pic, at, i }] with i the block's place. */
function picturesOf(origBlocks, blocks) {
  const before = new Map(), after = [];
  for (const o of origBlocks || []) { if (o.kind === 'image') before.set(o.path, { at: 'body', pic: o }); for (const p of o.pics || []) before.set(p.path, { at: o.path, pic: p }); }
  blocks.forEach((b, i) => { if (b.kind === 'image') after.push({ pic: b, at: 'body', i }); for (const p of b.pics || []) after.push({ pic: p, at: b.path, i }); });
  return { before, after };
}
/** Before the blocks are diffed: a paragraph's picture that is gone is removed, and every picture leaving its paragraph is parked in the
 *  body after it (the diff may remove or rewrite that paragraph). Block pictures moving into a paragraph wait in the body. */
export async function planPicturesBefore(file, origBlocks, blocks, exec, log) {
  const { before, after } = picturesOf(origBlocks, blocks);
  const now = new Map(after.filter(a => a.pic.path).map(a => [a.pic.path, a]));
  let n = 0;
  for (const [path, o] of before) {
    if (o.at === 'body') continue;
    const a = now.get(path);
    if (!a) { await exec(['remove', file, path]); n++; log && log('remove', path); }
    else if (a.at !== o.at) { await exec(['move', file, path, '--to', '/body', '--after', o.at]); n++; log && log('move', path); }
  }
  return n;
}
/** Where the engine puts a picture moved into a paragraph, and what a picture's place says by default. */
const FLOATED = { wrap: 'square', x: '0cm', y: '0cm', xFrom: 'column', yFrom: 'paragraph' }, INLINE = { wrap: 'inline' };
/** The set props that take a picture's place from a to b: inline says only that; an offset and an alignment on one axis replace each
 *  other, so whichever b has is sent when it differs. */
export function placeDiff(a, b) {
  a = Object.assign({}, INLINE, a); b = Object.assign({}, INLINE, b);
  if (b.wrap === 'inline') return a.wrap === 'inline' ? {} : { wrap: 'inline' };
  const p = {};
  if (a.wrap !== b.wrap) p.wrap = b.wrap;
  for (const k of ['xFrom', 'yFrom']) if (b[k] && (a[k] || '') !== b[k]) p[k] = b[k];
  for (const [off, al] of [['x', 'xAlign'], ['y', 'yAlign']]) {
    if (b[al]) { if (a[al] !== b[al]) p[al] = b[al]; }
    else if (b[off] != null && (a[off] == null || a[al] || Math.abs(cmOf(a[off]) - cmOf(b[off])) > 0.0005)) p[off] = b[off];
  }
  return p;
}
/** After the blocks stand in their final order with their final paths: pictures go to the paragraph they were dragged into (they float
 *  there), or back into the body before the block that follows them (inline, a block again); new pictures in a paragraph are added; then
 *  each picture's place and look that changed is set. `first`: an inline picture dropped before its paragraph's text goes in front of it
 *  (ponytail: anywhere else in the text lands at its end — the editor has no run offsets). */
export async function planPicturesAfter(file, origBlocks, blocks, exec, log, byId) {
  const { before, after } = picturesOf(origBlocks, blocks);
  let n = 0;
  const go = async argv => { const r = await exec(argv); n++; log && log(argv[0], argv[2]); return r; };
  const from = a => a.pic.path ? before.get(a.pic.path) : null, moved = a => !!from(a) && from(a).at !== a.at;
  for (const a of after) if (moved(a) && a.at !== 'body') await go(['move', file, a.pic.path, '--to', a.at, ...(a.pic.first ? ['--index', '1'] : [])]);
  for (const a of after.filter(a => moved(a) && a.at === 'body').reverse()) { const next = blocks[a.i + 1]; await go(['move', file, a.pic.path, '--to', '/body', ...(next && next.path ? ['--before', next.path] : [])]); }
  for (const a of after) {
    if (!a.pic.path) {
      if (a.at === 'body' || !a.at) continue; // a new block picture came with the blocks
      const r = await go(['add', file, a.at, '--type', 'image', ...propsArgs(Object.assign({ src: a.pic.props.src, width: a.pic.width ? Math.round(a.pic.width) + 'px' : null }, a.pic.place)), ...(a.pic.first ? ['--index', '1'] : [])]);
      a.pic.path = byId && r && r.props && r.props.id != null ? `//image[@id=${r.props.id}]` : r.path;
      if (a.pic.el) a.pic.el.setAttribute('data-path', a.pic.path);
      continue;
    }
    const o = from(a); if (!o) continue;
    const diff = Object.assign({}, a.at === 'body' && !moved(a) ? {} : lookDiff(o.pic.props, a.pic.props), a.at === 'body' ? {} : placeDiff(moved(a) ? FLOATED : o.pic.place, a.pic.place));
    if (Object.keys(diff).length) await go(['set', file, a.pic.path, ...propsArgs(diff)]);
  }
  return n;
}

/** A table whose recorded commands do not lead from the opened table to what the editor shows (commands lost, cells pasted
 * in) is written anew: every cell by path would otherwise land somewhere else. Returns the tables that are. */
export function tablesToRewrite(origBlocks, blocks) {
  const opened = new Map((origBlocks || []).filter(o => o.kind === 'table').map(o => [o.path, o])), out = [];
  for (const b of blocks) {
    if (b.kind !== 'table' || !b.path || !opened.has(b.path)) continue;
    let fits = false;
    try { fits = b.ops !== null && sameGrid(replayTable(opened.get(b.path), b.ops).m, blocksModel(b.rows)); } catch (e) { fits = false; }
    if (!fits) out.push(b);
  }
  return out;
}
const headingsOf = blocks => JSON.stringify(blocks.filter(b => b.kind === 'heading').map(b => [String(b.props.level), textOf(b.props.html || '')]));

/** Saves a Word document. `root` is the live editor element when the document is on screen, so new blocks learn their
 * paths in place and the user keeps typing; otherwise the model's html is parsed and written back. */
async function saveDocx(doc, root, log) {
  const orig = doc._orig || { blocks: [] };
  let n = 0;
  const pp = pageDiff(orig, doc);
  if (!!doc.track !== !!orig.track) pp.track = doc.track ? 'true' : 'false';
  if (Object.keys(pp).length) { await run(['set', doc.path, '/', ...propsArgs(pp)]); n++; log && log('set', '/', pp); }
  for (const s of doc.styleEdits || []) { await run(['set', doc.path, '/', '--prop', 'style=' + JSON.stringify(s)]); n++; log && log('set', '/', { style: s.id }); } // 修改样式 / 新建样式
  doc.styleEdits = [];
  const el = root || parseHtml(doc.html || '');
  const unsaved = p => p.path || /^data:/.test(p.props.src); // a picture of another file (a conflict copy) cannot be added
  const blocks = blocksFromHtml(el).filter(b => b.kind !== 'image' || unsaved(b));
  for (const b of blocks) if (b.pics) b.pics = b.pics.filter(unsaved);
  const opened = new Set((orig.blocks || []).map(o => o.path));
  const tocs = blocks.filter(b => b.kind === 'toc'), tocAdded = tocs.some(b => !b.path || !opened.has(b.path));
  const was = new Map(), remember = list => list.forEach(b => { was.set(b, b.path); if (b.rows) remember(b.rows); if (b.cells) remember(b.cells); });
  remember(blocks);
  const mv = pictureMoves(orig.blocks || [], blocks), byId = orig.ids !== false;
  n += await planPicturesBefore(doc.path, orig.blocks || [], blocks, run, log);
  n += await planDocxBlocks(doc.path, (orig.blocks || []).filter(o => !mv.out.has(o.path)), blocks.filter(b => !mv.in.has(b)), run, log);
  if (byId) for (const b of blocks) if (b.kind === 'image' && b.id && !/^\/\//.test(b.path || '')) { b.path = `//image[@id=${b.id}]`; if (b.el) b.el.setAttribute('data-path', b.path); }
  renumber('/body', blocks);
  n += await planPicturesAfter(doc.path, orig.blocks || [], blocks, run, log, byId);
  // a table of contents lists the headings: build it again in the file, and show it, once headings or contents changed
  if (tocs.length && (tocAdded || tocs.some(b => b.refresh) || headingsOf(orig.blocks || []) !== headingsOf(blocks))) {
    for (const b of tocs) { await run(['set', doc.path, b.path, '--prop', 'levels=' + (b.props.levels || '3')]); n++; log && log('set', b.path, { levels: b.props.levels }); }
    refreshTocs(el);
  }
  for (const b of tocs) { b.el?.removeAttribute('data-refresh'); b.refresh = false; }
  for (const b of blocks) if (b.kind === 'table') b.el?.removeAttribute('data-ops');
  const comments = await planComments(doc.path, orig.comments || [], commentsIn(doc, el, blocks, was), log);
  n += comments.count;
  const notes = await planNotes(doc.path, orig.notes || [], notesIn(doc, el, blocks), log);
  n += notes.count;
  const paras = objectsIn(blocks, was), origBy = new Map((orig.blocks || []).map(o => [o.path, o]));
  const eqs = await planEquations(doc.path, orig.eqs || [], paras, b => { const o = origBy.get(was.get(b)); return !o || !sameRuns(o.props.html, b.props.html); }, run, log);
  const shapes = await planShapes(doc.path, orig.shapes || [], paras, run, log);
  n += eqs.count + shapes.count;
  doc.notes = notes.list.map(x => ({ id: x.nid, kind: x.kind, text: x.text }));
  const pageProps = Object.fromEntries(['page', 'orientation', 'margin', 'columns'].filter(k => k in pp).map(k => [k, pp[k]]));
  const now = Object.assign({}, orig, { titlePg: String(!!orig.titlePg) }, pp); // headers and footers as the file has them now
  doc._orig = Object.assign({ blocks: strip(blocks), page: Object.assign({}, orig.page, pageProps) }, hfOf(now), { track: !!doc.track, lineNumbers: !!doc.lineNumbers, hyphenation: !!doc.hyphenation, noteFormat: doc.noteFormat || '', comments: comments.list, notes: notes.list, eqs: eqs.list, shapes: shapes.list });
  doc.html = el.innerHTML;
  return n;
}

// ----- xlsx -----
// Cell addresses (the sheet engine's helpers live in a module the bridge does not import).
const xCol = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const xRef = (r, c) => xCol(c) + (r + 1);
const xParse = a => { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(a || '').trim()); if (!m) return null; let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return { r: +m[2] - 1, c: c - 1 }; };
const CM_PX = 96 / 2.54, CHAR_PX = 7, CHAR_PAD = 5, PT_PX = 96 / 72;
/** Number format code → the editor's simple kind. `code` is kept so an untouched format round-trips byte for byte. */
export function fmtOf(code) {
  if (!code || code === 'General') return {};
  const cur = code.replace(/\[\$-[^\]]*\]/g, ''); // [$-409] names a locale, not a currency
  const bare = cur.replace(/\[[^\]]*\]|"[^"]*"|\\./g, ''), dec = (/\.(0+)/.exec(bare) || [, ''])[1].length;
  if (/[ymdhs]/i.test(bare) && !/[#0?]/.test(bare)) return { fmt: /[yd]/i.test(bare) || !/[hs]/i.test(bare) ? 'date' : 'time', code };
  if (/\*|_\(/.test(cur)) return { fmt: 'acct', dec, code };
  if (/%/.test(bare)) return { fmt: 'pct', dec, code };
  if (/\$/.test(cur)) return { fmt: 'usd', dec, code };
  if (/€/.test(cur)) return { fmt: 'eur', dec, code };
  if (/[¥£￥]/.test(cur)) return { fmt: 'money', dec, code };
  if (/[#0]/.test(bare)) return { fmt: /,/.test(bare) ? 'number' : 'plain', dec, code };
  if (/^@$/.test(bare)) return { fmt: 'text', code };
  return { fmt: 'custom', code };
}
/** The exact Excel code for a kind the user picked; a cell that still carries the file's code keeps it. */
export function codeOf(s) {
  if (!s || !s.fmt) return 'General';
  if (s.code) return s.code;
  const d = s.dec == null ? 2 : s.dec, zeros = d ? '.' + '0'.repeat(d) : '';
  switch (s.fmt) {
    case 'pct': return '0' + (s.dec ? '.' + '0'.repeat(s.dec) : '') + '%';
    case 'money': return '"¥"#,##0' + zeros;
    case 'usd': return '"$"#,##0' + zeros;
    case 'eur': return '"€"#,##0' + zeros;
    case 'acct': return `_("¥"* #,##0${zeros}_)`;
    case 'number': return '#,##0' + zeros;
    case 'plain': case 'custom': return '0' + zeros;
    case 'date': return 'yyyy/m/d';
    case 'time': return 'h:mm:ss';
    case 'text': return '@';
  }
  return 'General';
}
const BOOLS = { bold: 'b', italic: 'i', underline: 'u', strike: 'st', wrap: 'wrap' };
const SIDES = ['top', 'right', 'bottom', 'left'];
const sidesOf = bd => Object.fromEntries(SIDES.map(k => [k, !bd ? 'none' : typeof bd === 'object' ? bd[k] || 'none' : bd])); // s.bd (one style or per side) → all four sides
// The /json tree gives json-typed props (borders, merges, widths, series…) already parsed and sizes as "14pt"; `get` output and tests may carry strings.
const jsonOr = (s, d) => { if (s && typeof s === 'object') return s; try { return s ? JSON.parse(s) : d; } catch (e) { return d; } };
const ptOfSize = s => parseFloat(String(s == null ? '' : s)) || 0;
export function cellModel(p) {
  const v = p.formula != null ? '=' + p.formula : (p.value != null ? String(p.value) : '');
  const s = {};
  for (const k in BOOLS) if (p[k] === 'true') s[BOOLS[k]] = true;
  if (p.color) s.color = hex(p.color);
  if (p.fill && p.fill !== 'none') s.fill = hex(p.fill);
  if (ptOfSize(p.size)) s.fs = ptOfSize(p.size);
  if (p.font) s.font = p.font;
  if (p.align && p.align !== 'general') s.align = p.align;
  if (p.valign) s.va = p.valign;
  if (+p.indent) s.indent = +p.indent;
  if (+p.rotate) s.rotate = +p.rotate;
  const sides = jsonOr(p.borders, null);
  if (sides && typeof sides === 'object' && Object.keys(sides).length) s.bd = sides;
  else if (p.border && p.border !== 'none') s.bd = p.border;
  if (p.borderColor && p.borderColor !== 'none') s.bdc = hex(p.borderColor);
  if (p.link) s.link = p.link;
  if (p.note) s.note = p.note;
  Object.assign(s, fmtOf(p.format));
  if (p.type === 'date' && s.fmt !== 'date' && s.fmt !== 'time') s.fmt = 'date'; // the engine knows a date by its format; the formatter then reads the code
  if (p.type === 'string' && p.formula == null && v.trim() !== '' && !isNaN(Number(v.replace(/,/g, ''))) && !s.fmt) s.fmt = 'text'; // "007" stays text, not 7
  return Object.keys(s).length ? { v, s } : { v };
}
const pxOfCm = v => Math.round(cmOf(v) * CM_PX), cmOfPx = px => cmStr(px / CM_PX);
/** A chart on a sheet: `eid` is the engine's id, which every sheet numbers anew, so the editor's id adds the sheet. */
function chartModel(c, sheetPath) {
  const p = c.props || {}, ser = jsonOr(p.series, []);
  return { id: `${sheetPath}/chart[@id=${p.id}]`, eid: String(p.id), path: c.path, type: p.type || 'column', title: p.title || '', cat: p.categories || '', ser: Array.isArray(ser) ? ser : [], legend: p.legend || 'right', stacked: p.stacked === 'true' || p.stacked === true,
    x: p.x ? pxOfCm(p.x) : 0, y: p.y ? pxOfCm(p.y) : 0, w: p.w ? pxOfCm(p.w) : 460, h: p.h ? pxOfCm(p.h) : 300 };
}
/** Engine sheet node → editor sheet model. */
/** A picture on a sheet: addressed by its id (stable while sheets keep their places), placed in sheet px like a chart. Every sheet
 *  numbers its pictures anew, so the editor's id is the path, which names one picture in the workbook. */
function imageModel(n, sheetPath, file) {
  const p = n.props || {}, path = `${sheetPath}/image[@id=${p.id}]`;
  return { id: path, path, src: binaryUrl(file || '', path), x: p.x ? pxOfCm(p.x) : 0, y: p.y ? pxOfCm(p.y) : 0, w: p.w ? pxOfCm(p.w) : 96, h: p.h ? pxOfCm(p.h) : 96, alt: p.alt || '', look: lookFrom(p) };
}
export function sheetModel(s, file) {
  const p = s.props || {}, cells = {}, charts = [], images = [];
  const sheetPath = p.id != null ? `/sheet[@id=${p.id}]` : s.path; // by its stable id: saving the editor's sheet order moves sheets
  for (const r of s.children || []) {
    if (r.kind === 'chart') { charts.push(chartModel(r, sheetPath)); continue; }
    if (r.kind === 'image') { images.push(imageModel(r, sheetPath, file)); continue; }
    for (const c of r.children || []) { const ref = /\[([A-Z]+\d+)\]$/.exec(c.path); if (ref) cells[ref[1]] = cellModel(c.props || {}); }
  }
  const merges = jsonOr(p.merges, []).map(m => { const [a, b] = String(m).split(':'), p1 = xParse(a), p2 = xParse(b || a); return p1 && p2 ? { r: p1.r, c: p1.c, rs: p2.r - p1.r + 1, cs: p2.c - p1.c + 1 } : null; }).filter(Boolean);
  const colW = {}; Object.entries(jsonOr(p.widths, {})).forEach(([k, w]) => { if (+w) colW[k.toUpperCase()] = Math.round(+w * CHAR_PX + CHAR_PAD); });
  const rowH = {}; Object.entries(jsonOr(p.heights, {})).forEach(([k, h]) => { if (+h) rowH[k] = Math.round(+h * PT_PX); });
  const fz = p.freeze && p.freeze !== 'none' ? xParse(p.freeze) : null;
  const filter = p.filter && p.filter !== 'none' ? p.filter : null, filters = jsonOr(p.filters, {});
  // the sheet's rules keep the file's shape (ranges as text, see the engine's sheet props); colours get their # like a cell's
  const cf = jsonOr(p.cf, []).map(r => ruleColors(r, hex)), dv = jsonOr(p.validations, []);
  // rows the file hides inside a filtered range are the filter's doing (frows), which the editor recomputes; the rest were hidden by hand
  const hid = jsonOr(p.hidden, {}), fb = filter && Object.keys(filters).length ? filter.split(':').map(xParse) : null, hiddenRows = [], frows = [];
  (hid.rows || []).forEach(n => { const r = +n - 1; if (r < 0) return; (fb && fb[0] && r > fb[0].r && r <= (fb[1] || fb[0]).r ? frows : hiddenRows).push(r); });
  const hiddenCols = (hid.cols || []).map(k => xParse(String(k) + '1')).filter(Boolean).map(a => a.c);
  const m = { name: p.name, path: sheetPath, cells, colW, rowH, merges, frR: fz ? fz.r : 0, frC: fz ? fz.c : 0, filter, filters, frows, cf, dv, hiddenRows, hiddenCols, color: p.color && p.color !== 'none' ? hex(p.color) : null, charts, images };
  if (p.gridlines === 'false' || p.gridlines === false) m.noGrid = true; // the file hides them; otherwise the settings decide
  return m;
}
/** A conditional format's colours (fill, color, colors) through `f`: hex adds the #, unhex takes it away. */
const ruleColors = (r, f) => { const o = Object.assign({}, r); if (o.fill) o.fill = f(o.fill); if (o.color) o.color = f(o.color); if (Array.isArray(o.colors)) o.colors = o.colors.map(f); return o; };
/** The snapshot a later save is diffed against. */
const origOf = sheets => ({ sheets: JSON.parse(JSON.stringify(sheets.map(s => ({ path: s.path, name: s.name, cells: s.cells, colW: s.colW || {}, rowH: s.rowH || {}, merges: s.merges || [], frR: s.frR || 0, frC: s.frC || 0, filter: s.filter || null,
  filters: s.filters || {}, frows: s.frows || [], cf: s.cf || [], dv: s.dv || [], hiddenRows: s.hiddenRows || [], hiddenCols: s.hiddenCols || [], color: s.color || null, noGrid: !!s.noGrid, charts: s.charts || [], images: s.images || [] })))) });
async function openXlsx(doc) {
  const t = await tree(doc.path);
  const sheets = (t.children || []).filter(s => s.kind === 'sheet').map(s => sheetModel(s, doc.path)), tp = t.props || {};
  // the workbook's default font: cells without their own show it, as in Excel
  return { active: Math.min(doc.active || 0, sheets.length - 1), sheets, font: tp.font || null, fs: ptOfSize(tp.size) || null, _orig: origOf(sheets) };
}
const same = (a, b) => JSON.stringify(a || null) === JSON.stringify(b || null);
export function cellProps(o, c) {
  const p = {}, os = (o && o.s) || {}, ns = (c && c.s) || {};
  const ov = o ? o.v : '', nv = c ? (c.v == null ? '' : String(c.v)) : '';
  if (ov !== nv) { if (nv[0] === '=') p.formula = nv.slice(1); else { p.value = nv; if (ns.fmt === 'text' && nv !== '') p.type = 'string'; } }
  for (const k in BOOLS) if (!!os[BOOLS[k]] !== !!ns[BOOLS[k]]) p[k] = ns[BOOLS[k]] ? 'true' : 'false';
  if ((os.color || null) !== (ns.color || null)) p.color = unhex(ns.color);
  if ((os.fill || null) !== (ns.fill || null)) p.fill = unhex(ns.fill);
  if ((os.fs || null) !== (ns.fs || null)) p.size = String(ns.fs || 11);
  if ((os.font || '') !== (ns.font || '')) p.font = ns.font || 'Calibri';
  if ((os.align || '') !== (ns.align || '')) p.align = ns.align || 'general';
  if ((os.va || '') !== (ns.va || '')) p.valign = ns.va || 'bottom';
  if ((os.indent || 0) !== (ns.indent || 0)) p.indent = String(ns.indent || 0);
  if ((os.rotate || 0) !== (ns.rotate || 0)) p.rotate = String(ns.rotate || 0);
  if (!same(os.bd, ns.bd)) {
    // `borders` writes only the sides it names, so a side that was taken away is sent as none
    if (!ns.bd || typeof ns.bd !== 'object') p.border = ns.bd || 'none';
    else { const a = sidesOf(os.bd), b = sidesOf(ns.bd), d = SIDES.filter(k => a[k] !== b[k]); if (d.length) p.borders = JSON.stringify(Object.fromEntries(d.map(k => [k, b[k]]))); }
  }
  if ((os.bdc || null) !== (ns.bdc || null)) p.borderColor = unhex(ns.bdc);
  if ((os.link || '') !== (ns.link || '')) p.link = ns.link || '';
  if ((os.note || '') !== (ns.note || '')) p.note = ns.note || '';
  if ((os.fmt || '') !== (ns.fmt || '') || (os.dec ?? null) !== (ns.dec ?? null) || (os.code || '') !== (ns.code || '')) p.format = codeOf(ns);
  return p;
}
const CELL_ONLY = new Set(['value', 'formula', 'type', 'link', 'note']);
const RANGE_MIN = 8;
/** Sheet-level props that differ between the saved snapshot and the model. */
export function sheetProps(o, s) {
  const p = {};
  const mg = x => JSON.stringify((x.merges || []).map(m => xRef(m.r, m.c) + ':' + xRef(m.r + m.rs - 1, m.c + m.cs - 1)));
  const wd = x => Object.fromEntries(Object.entries(x.colW || {}).filter(([, px]) => px > 0).map(([k, px]) => [k, +((px - CHAR_PAD) / CHAR_PX).toFixed(2)]));
  const ht = x => Object.fromEntries(Object.entries(x.rowH || {}).filter(([, px]) => px > 0).map(([k, px]) => [k, +(px / PT_PX).toFixed(2)]));
  // widths/heights: the engine keeps keys it is not given, so send only changed keys, and null to reset a removed one
  const sizes = (a, b) => { const d = {}; for (const k in b) if (a[k] !== b[k]) d[k] = b[k]; for (const k in a) if (!(k in b)) d[k] = null; return Object.keys(d).length ? JSON.stringify(d) : null; };
  const fz = x => (x.frR || x.frC) ? xRef(x.frR || 0, x.frC || 0) : 'none';
  if (mg(o) !== mg(s)) p.merges = mg(s);
  const w = sizes(wd(o), wd(s)), h = sizes(ht(o), ht(s));
  if (w) p.widths = w;
  if (h) p.heights = h;
  if (fz(o) !== fz(s)) p.freeze = fz(s);
  if ((o.filter || 'none') !== (s.filter || 'none')) p.filter = s.filter || 'none';
  // the rules: whole sets, sent when they differ (filters only while there is a filter range to hold them; the range goes first in the same set)
  const js = x => JSON.stringify(x || null);
  const fl = x => js(Object.keys(x.filters || {}).length ? x.filters : {});
  if (fl(o) !== fl(s) && (s.filter || fl(s) === '{}')) p.filters = fl(s);
  const cf = x => js((x.cf || []).map(r => ruleColors(r, unhex)));
  if (cf(o) !== cf(s)) p.cf = cf(s);
  if (js(o.dv || []) !== js(s.dv || [])) p.validations = js(s.dv || []);
  const hid = x => js({ rows: [...new Set([...(x.hiddenRows || []), ...(x.frows || [])])].sort((a, b) => a - b).map(r => r + 1), cols: (x.hiddenCols || []).slice().sort((a, b) => a - b).map(xCol) });
  if (hid(o) !== hid(s)) p.hidden = hid(s);
  if ((o.color || null) !== (s.color || null)) p.color = unhex(s.color);
  if (!!o.noGrid !== !!s.noGrid) p.gridlines = s.noGrid ? 'false' : 'true';
  return p;
}
export function chartProps(ch) {
  return { type: ch.type || 'column', title: ch.title || '', categories: ch.cat || '', series: JSON.stringify(ch.ser || []), legend: ch.legend || 'right', stacked: ch.stacked ? 'true' : 'false', x: cmOfPx(ch.x || 0), y: cmOfPx(ch.y || 0), w: cmOfPx(ch.w || 460), h: cmOfPx(ch.h || 300) };
}
const chartKey = ch => ch.eid || ch.id;
/** Turns cell diffs into commands; identical style changes over a full rectangle of ≥ RANGE_MIN cells become one range set.
 * ponytail: only exact rectangles are batched, a ragged block falls back to per-cell sets. */
function planCells(file, sheetPath, o, s, exec, log) {
  const cmds = [], sets = [];
  for (const ref of new Set([...Object.keys(o.cells), ...Object.keys(s.cells)])) {
    const oc = o.cells[ref], nc = s.cells[ref];
    const styled = nc && nc.s && Object.keys(nc.s).some(k => k !== 'dv');
    if (!nc || ((nc.v == null || nc.v === '') && !styled)) { if (oc) cmds.push({ argv: ['remove', file, `${sheetPath}/cell[${ref}]`], what: ['remove', `${sheetPath}/cell[${ref}]`] }); continue; }
    if (same(oc, nc)) continue;
    const p = cellProps(oc, nc), a = xParse(ref);
    if (Object.keys(p).length) sets.push({ ref, r: a.r, c: a.c, p });
  }
  const groups = new Map();
  for (const x of sets) { const st = Object.keys(x.p).filter(k => !CELL_ONLY.has(k)).sort().map(k => [k, x.p[k]]); if (st.length) { const key = JSON.stringify(st); (groups.get(key) || groups.set(key, []).get(key)).push(x); } }
  for (const [key, g] of groups) {
    if (g.length < RANGE_MIN) continue;
    const r1 = Math.min(...g.map(x => x.r)), r2 = Math.max(...g.map(x => x.r)), c1 = Math.min(...g.map(x => x.c)), c2 = Math.max(...g.map(x => x.c));
    if ((r2 - r1 + 1) * (c2 - c1 + 1) !== g.length) continue;
    const props = Object.fromEntries(JSON.parse(key)), path = `${sheetPath}/range[${xRef(r1, c1)}:${xRef(r2, c2)}]`;
    cmds.push({ argv: ['set', file, path, ...propsArgs(props)], what: ['set', path, props] });
    g.forEach(x => { for (const k in props) delete x.p[k]; });
  }
  for (const x of sets) if (Object.keys(x.p).length) cmds.push({ argv: ['set', file, `${sheetPath}/cell[${x.ref}]`, ...propsArgs(x.p)], what: ['set', `${sheetPath}/cell[${x.ref}]`, x.p] });
  return cmds;
}
/** A picture's bytes as a data: URL, which `add --type image --prop src=` takes. */
async function dataUrlOf(src) {
  if (src.startsWith('data:')) return src;
  const res = await http(src), bytes = new Uint8Array(await res.arrayBuffer());
  let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return 'data:' + (res.headers.get('Content-Type') || 'image/png') + ';base64,' + btoa(s);
}
/** Plans and runs the commands that turn the saved snapshot into the model. `exec(argv)` runs one command and resolves with its result
 * (`add` results carry `path` and `props.id`); the model is updated in place with the paths and ids the engine assigned. Returns the count. */
export async function planXlsx(file, origSheets, sheets, exec, log) {
  let n = 0;
  const go = async (argv, what) => { const r = await exec(argv); n++; log && log(...what); return r; };
  const used = new Set();
  // the file's sheet order as it changes: each sheet goes right after the one before it in the editor (sheets are addressed by id,
  // so a move leaves every path valid); sheets that go away are removed last, since a workbook keeps at least one
  const order = origSheets.map(o => o.path), stays = new Set(sheets.map(s => s.path).filter(Boolean));
  const place = (path, prev) => { if (order.includes(path)) order.splice(order.indexOf(path), 1); order.splice(prev ? order.indexOf(prev) + 1 : 0, 0, path); };
  let prev = null;
  for (const s of sheets) {
    let o = s.path ? origSheets.find(x => x.path === s.path && !used.has(x.path)) : null;
    const where = prev ? ['--after', prev] : ['--index', '1'];
    if (!o) {
      const r = await go(['add', file, '/', '--type', 'sheet', '--prop', 'name=' + s.name, ...where], ['add', '(sheet)', 'sheet']);
      s.path = r.props && r.props.id != null ? `/sheet[@id=${r.props.id}]` : r.path; o = { path: s.path, name: s.name, cells: {}, charts: [] }; stays.add(s.path); place(s.path, prev);
    } else {
      if (o.name !== s.name) await go(['set', file, o.path, '--prop', 'name=' + s.name], ['set', o.path, { name: s.name }]);
      const kept = order.filter(p => stays.has(p));
      if (kept[prev ? kept.indexOf(prev) + 1 : 0] !== o.path) { await go(['move', file, o.path, '--to', '/', ...where], ['move', o.path, prev || '/']); place(o.path, prev); }
    }
    used.add(o.path); prev = o.path;
    for (const c of planCells(file, o.path, o, s, exec, log)) await go(c.argv, c.what);
    const sp = sheetProps(o, s);
    if (Object.keys(sp).length) await go(['set', file, o.path, ...propsArgs(sp)], ['set', o.path, sp]);
    // chart ids are unique per sheet only (every sheet's first chart is id 2), so a chart is addressed under its sheet
    const oc = new Map((o.charts || []).map(ch => [chartKey(ch), ch])), at = id => `${o.path}/chart[@id=${id}]`;
    for (const ch of s.charts || []) {
      const old = ch.eid != null ? oc.get(ch.eid) : null, np = chartProps(ch);
      if (!old) { const r = await go(['add', file, o.path, '--type', 'chart', ...propsArgs(np)], ['add', o.path, 'chart']); ch.path = r.path; ch.eid = String(r.props && r.props.id != null ? r.props.id : ch.id); continue; }
      const op = chartProps(old), d = {}; for (const k in np) if (np[k] !== op[k]) d[k] = np[k];
      if (Object.keys(d).length) await go(['set', file, at(ch.eid), ...propsArgs(d)], ['set', ch.path || at(ch.eid), d]);
    }
    const keep = new Set((s.charts || []).map(ch => ch.eid).filter(x => x != null));
    for (const old of (o.charts || []).slice().reverse()) if (!keep.has(chartKey(old))) await go(['remove', file, at(chartKey(old))], ['remove', old.path || at(chartKey(old))]);
    // pictures: the 图片 tab writes each change at once, so a difference here is an undo, sent back as the look it returns to. One
    // with no path yet (a copied sheet's) is new: added from the bytes it shows, then given its look with the frame after it (a crop shrinks the frame)
    for (const im of s.images || []) {
      if (!im.path) {
        const box = { x: cmOfPx(im.x), y: cmOfPx(im.y), w: cmOfPx(im.w), h: cmOfPx(im.h) };
        const r = await go(['add', file, o.path, '--type', 'image', '--prop', 'src=' + await dataUrlOf(im.src), ...propsArgs(Object.assign({}, box, im.alt ? { alt: im.alt } : {}))], ['add', o.path, 'image']);
        im.path = `${o.path}/image[@id=${r.props.id}]`; im.src = binaryUrl(file, im.path);
        if (Object.keys(im.look || {}).length) await go(['set', file, im.path, ...propsArgs(Object.assign({}, im.look, box))], ['set', im.path, im.look]);
        continue;
      }
      const old = (o.images || []).find(x => x.id === im.id), d = old ? lookDiff(old.look, im.look) : {};
      if (Object.keys(d).length) await go(['set', file, im.path, ...propsArgs(d)], ['set', im.path, d]);
    }
  }
  for (const o of origSheets.slice().reverse()) if (!used.has(o.path)) await go(['remove', file, o.path], ['remove', o.path]);
  return n;
}
async function saveXlsx(doc, log) { return planXlsx(doc.path, doc._orig ? doc._orig.sheets : [], doc.sheets, run, log); }

// ----- pptx -----
/** The editor's shape key for a DrawingML preset (POLY's are their own), and back. Presets the editor cannot draw show as rectangles. */
const geomOf = prst => prst === 'roundRect' ? 'round' : prst === 'ellipse' ? 'ellipse' : POLY[prst] ? prst : 'rect';
const GEOM_BACK = { round: 'roundRect', pill: 'roundRect', arrow: 'rightArrow' };
const geomBack = key => GEOM_BACK[key] || key || 'rect';
/** Outline width: the file's points to slide units and back. */
const swOf = (p, ptPx) => p.line && p.line !== 'none' ? Math.max(1, Math.round((parseFloat(p.lineWidth) || 0.75) * ptPx)) : 0; // lineWidth prints as "3pt"
const linePt = (sw, ptPx) => (Math.round(sw / ptPx * 4) / 4) + 'pt';
function paraHtml(paragraphs) {
  let out = '', inList = null; // the open list's tag
  for (const p of paragraphs) {
    const pr = p.props || {}, h = pr.html != null ? pr.html : esc(pr.text || ''); // a stub (get --depth 2) has no props
    const lvl = +pr.level > 0 ? ` data-lvl="${Math.min(4, +pr.level)}"` : '';
    if (pr.list && pr.list !== 'none') { const tag = pr.list === 'number' ? 'ol' : 'ul'; if (inList !== tag) { if (inList) out += `</${inList}>`; out += `<${tag}>`; inList = tag; } out += `<li${lvl}>` + (h || '<br>') + '</li>'; }
    else { if (inList) { out += `</${inList}>`; inList = null; } out += `<p${lvl}>` + (h || '<br>') + '</p>'; }
  }
  if (inList) out += `</${inList}>`;
  return out;
}
/** The box-wide text settings of a shape as the editor keeps them (spacing in slide units, the rest as the file names it). */
function textBox(p, ptPx) {
  const t = {}, pt = v => Math.round(parseFloat(v) * ptPx);
  if (p.lineSpacing) t.lh = +p.lineSpacing; if (p.spaceBefore) t.sb = pt(p.spaceBefore); if (p.spaceAfter) t.sa = pt(p.spaceAfter); if (p.charSpacing) t.cs = Math.round(parseFloat(p.charSpacing) * ptPx * 10) / 10;
  if (+p.columns > 1) t.cols = +p.columns; if (p.direction && p.direction !== 'horz') t.vert = p.direction;
  if (p.autofit) { const m = /^shrink(?::(\d+))?$/.exec(p.autofit); if (m) { t.autofit = 'shrink'; if (m[1]) t.fit = +m[1] / 100; } else if (p.autofit === 'resize') t.autofit = 'resize'; }
  if (p.textOutline && p.textOutline !== 'none') t.tOutline = hex(p.textOutline); if (p.textShadow === 'true') t.tShadow = true;
  if (p.textGradient) t.tGrad = p.textGradient.split(',').map((v, i) => i < 2 ? '#' + v : v).join(',');
  return t;
}
/** The engine's placeholder roles as the editor names them; footer ones (dt, ftr, sldNum) are ordinary objects to it. */
const PH = { title: 'title', subtitle: 'sub', body: 'body', obj: 'body', pic: 'pic' };
/** A slide's child in the engine's tree (a shape, picture, table, or decor inherited from the layout and master) as an editor
 *  object that remembers its path. sp is the slide's path, geo the deck's scale (openPptx). */
function pptxObject(file, sp, geo, n) {
  const { kx, ky, ptPx } = geo;
  const p = Object.assign({}, n.computed, n.props), decor = n.kind === 'decor'; // computed: what the shape inherits from its theme, layout and master
  const kind = decor ? p.type : n.kind;
  const box = { x: Math.round(cmOf(p.x) * kx), y: Math.round(cmOf(p.y) * ky), w: Math.round(cmOf(p.w) * kx), h: Math.round(cmOf(p.h) * ky) };
  const path = decor || !p.id ? n.path : sp + '/' + n.kind + '[@id=' + p.id + ']';
  // an id names one object in the whole deck, and a cNvPr id does not: every slide numbers its shapes anew, and master and layout shapes may share one
  const id = (decor ? 'd' : 'e') + path;
  const look = { rot: Number(p.rotation) || 0, shadow: p.shadow === 'true' };
  if (kind === 'image') { const lk = lookFrom(p); delete lk.rotation; return txt(Object.assign({ id, path, kind, t: 'image', src: binaryUrl(file, path), html: '', look: lk, rot: look.rot }, box)); }
  if (n.kind === 'table') {
    const merges = [], cells = {};
    (n.children || []).forEach((r, ri) => (r.children || []).forEach((c, ci) => { const cp = c.props || {}; if (+cp.colspan > 1 || +cp.rowspan > 1) merges.push({ r: ri, c: ci, rs: +cp.rowspan || 1, cs: +cp.colspan || 1 }); const own = {}; if (cp.fill) own.fill = hex(cp.fill); if (cp.line) own.line = cp.line === 'none' ? 'none' : hex(cp.line); if (cp.align && cp.align !== 'left') own.align = cp.align; if (Object.keys(own).length) cells[ri + ':' + ci] = own; }));
    let colW; try { const ws = typeof p.widths === 'string' ? JSON.parse(p.widths) : p.widths; colW = Array.isArray(ws) ? ws.map(v => Math.round(cmOf(v) * kx)) : undefined; } catch (e) { colW = undefined; } // the tree gives JSON props as values, a get as text
    return txt(Object.assign({ id, path, kind: 'table', t: 'table', html: '', fs: 24, rows: (n.children || []).map(r => (r.children || []).map(c => (c.props || {}).text || '')), merges, cells, colW, tstyle: p.style || 'MediumStyle2Accent1', header: p.header === 'true', banded: p.banded === 'true', firstCol: p.firstCol === 'true' }, box));
  }
  if (n.kind === 'group') return Object.assign({ id, path, kind: 'group', t: 'group', op: 1, kids: (n.children || []).map(c => pptxObject(file, path, geo, c)) }, box, look);
  const cxn = ref => { const m = /^(\d+),(\d+)$/.exec(ref || ''); return m ? { cnv: +m[1], idx: +m[2] } : null; }; // resolved to an editor id once the slide's objects are known (openPptx)
  if (n.kind === 'connector') return mkLine(Object.assign({ id, path, kind: 'connector', stroke: hex(p.line) || '', sw: swOf(p, ptPx), dash: p.dash || 'solid', head: p.head || 'none', tail: p.tail || 'none', flipH: p.flipH === 'true', flipV: p.flipV === 'true', bent: /^bent|^curved/.test(p.geometry || ''), start: cxn(p.start), end: cxn(p.end) }, box, look));
  const ph = decor ? null : PH[p.placeholder] || null, isTitle = ph === 'title', filled = (p.fill && p.fill !== 'none') || p.gradient;
  // a connector has no height: draw it as a thin bar in its line colour
  const from = decor ? { source: p.source } : {}; // master or layout: 版式 keeps the master's decor and swaps the layout's
  if (decor && p.geometry === 'line') return mkShape(Object.assign({ id, path, kind: 'shape', html: '', fill: hex(p.line), stroke: '', sw: 0, shape: 'rect' }, box, from, { h: Math.max(box.h, 2) }));
  // an empty placeholder is '' rather than an empty paragraph: the editor shows its hint, and nothing is written until the user types
  const html = decor ? (p.html || (p.text ? '<p>' + esc(p.text) + '</p>' : '')) : ph && !String(p.text || '').trim() ? '' : paraHtml((n.children || []).filter(c => c.kind === 'paragraph'));
  const base = Object.assign({ id, path, kind: 'shape', html, fs: p.size ? Math.round(cmOf(p.size) / UNIT.pt * ptPx) : (isTitle ? Math.round(40 * ptPx) : Math.round(20 * ptPx)), color: hex(p.color), font: p.font || null, ph, bold: p.bold != null ? p.bold === 'true' : isTitle, lockAspect: p.lockAspect === 'true' }, box, from, look, decor ? {} : textBox(p, ptPx));
  const outline = { stroke: p.line && p.line !== 'none' ? hex(p.line) : '', sw: swOf(p, ptPx), dash: p.dash || 'solid' };
  if (filled || (p.geometry && p.geometry !== 'rect' && p.geometry !== 'textbox' && p.geometry !== 'custom')) return mkShape(Object.assign(base, outline, { fill: p.gradient ? 'grad:' + p.gradient.split(',').map((v, i) => i < 2 ? '#' + v : v).join(',') : filled ? hex(p.fill) : null, shape: geomOf(p.geometry), align: 'center', va: 'middle' }));
  return txt(Object.assign(base, outline));
}
/** A connector's ends name shapes by their drawing id; the editor's lines hold the object's id. */
function linkLines(objs) {
  const byCnv = new Map(flatObjs(objs).filter(o => o.path).map(o => [+(/\[@id=(\d+)\]$/.exec(o.path) || [])[1], o.id]));
  for (const o of objs) if (o.t === 'line') for (const k of ['start', 'end']) { const r = o[k]; if (r && r.cnv != null) o[k] = byCnv.has(r.cnv) ? { id: byCnv.get(r.cnv), idx: r.idx } : null; }
}
async function openPptx(doc) {
  const t = await tree(doc.path);
  const wcm = cmOf(t.props.width) || 33.867, hcm = cmOf(t.props.height) || 19.05;
  const ratio = hcm / wcm > 0.7 ? '4:3' : '16:9', H = slideH(ratio), kx = SW / wcm, ky = H / hcm, ptPx = SW / (wcm / 2.54 * 72);
  const geo = { wcm, hcm, kx, ky, ptPx };
  const slides = (t.children || []).filter(s => s.kind === 'slide').map(s => {
    const sp = '/slide[@id=' + s.props.id + ']', children = s.children || [];
    const decor = children.filter(n => n.kind === 'decor').map(n => pptxObject(doc.path, sp, geo, n));
    const objs = children.filter(n => n.kind !== 'decor').map(n => pptxObject(doc.path, sp, geo, n));
    linkLines(objs);
    return { id: 's' + s.props.id, path: sp, layout: s.props.layout, decor, objs, anims: animsFrom(s.props.animations, objs), sec: s.props.section || '', notes: s.props.notes || '', trans: s.props.transition || 'none', duration: s.props.duration == null ? null : Number(s.props.duration), hidden: s.props.hidden === true || s.props.hidden === 'true', bg: p2bg(s.props.background || (s.computed || {}).background) };
  });
  // the palette the deck wears (its theme, written by 设计 or the assistant) is the editor's theme; a deck without one keeps the editor's
  const palette = THEMES[t.props.palette] ? t.props.palette : null;
  const orig = { geo, palette, slides: pptxSnapshot(slides) };
  return { theme: palette || doc.theme || 'paper', ratio, slides, _orig: orig };
}
const p2bg = c => c ? '#' + c : '#FFFFFF';
const pptxSnapshot = slides => JSON.parse(JSON.stringify(slides.map(s => ({ id: s.id, path: s.path, bg: s.bg, layout: s.layout, sec: s.sec || '', notes: s.notes || '', trans: s.trans || 'none', duration: s.duration ?? null, hidden: !!s.hidden, objs: s.objs.map(objKey), anims: animJson(s) }))));
const cnvId = o => { const m = o && o.path && /\[@id=(\d+)\]$/.exec(o.path); return m ? m[1] : null; };
/** A slide's animations (the engine's animations prop) as the editor keeps them: each effect names its object by editor id, or by
 *  drawing id (sp) when it animates something the editor does not show (a shape of the layout); an effect the engine does not model
 *  keeps its class and markup. */
function animsFrom(json, objs) {
  let list; try { list = typeof json === 'string' ? JSON.parse(json) : json; } catch (e) { list = null; } // the tree gives JSON props as values, a get as text
  if (!Array.isArray(list)) return [];
  const byCnv = new Map(flatObjs(objs).map(o => [cnvId(o), o.id]));
  return list.map(e => Object.assign({ fx: e.effect, start: e.start || 'click', dur: +e.duration || 500, delay: +e.delay || 0 },
    byCnv.has(String(e.shape)) ? { id: byCnv.get(String(e.shape)) } : { sp: e.shape == null ? null : String(e.shape) }, e.effect === 'other' ? { cls: e.class, xml: e.xml } : {}));
}
/** And back: the prop as the slide is now, each object by its drawing id; the effects of objects that are gone are left out. */
function animJson(s) {
  const out = [], all = flatObjs(s.objs || []);
  for (const a of s.anims || []) {
    const shape = a.id ? cnvId(all.find(o => o.id === a.id)) : a.sp;
    if (shape == null && (a.id || a.fx !== 'other')) continue;
    const e = { effect: a.fx, start: a.start || 'click', duration: a.dur, delay: a.delay || 0 };
    if (shape != null) e.shape = shape;
    if (a.fx === 'other') Object.assign(e, { class: a.cls, xml: a.xml });
    out.push(e);
  }
  return out;
}
export function slideProps(o, s) {
  const p = {};
  if ((o.notes || '') !== (s.notes || '')) p.notes = s.notes || '';
  if ((o.sec || '') !== (s.sec || '')) p.section = s.sec || '';
  if (!!o.hidden !== !!s.hidden) p.hidden = s.hidden ? 'true' : 'false';
  if ((o.trans || 'none') !== (s.trans || 'none')) p.transition = s.trans || 'none';
  if (s.duration != null && s.trans !== 'none' && s.trans !== 'other' && o.duration !== s.duration) p.duration = String(s.duration);
  return p;
}
function objKey(o) {
  return Object.assign({ id: o.id, path: o.path, kind: o.kind, t: o.t, x: o.x, y: o.y, w: o.w, h: o.h, html: o.html, fill: o.fill, color: o.color, font: o.font, fs: o.fs, shape: o.shape, rows: o.rows, src: o.src && o.src.startsWith('data:') ? 'data' : o.src, rot: o.rot || 0,
    stroke: o.stroke, sw: o.sw, dash: o.dash, shadow: !!o.shadow, lockAspect: !!o.lockAspect,
    lh: o.lh, sb: o.sb || 0, sa: o.sa || 0, cs: o.cs || 0, cols: o.cols || 1, vert: o.vert || 'horz', autofit: o.autofit || 'none', fit: o.fit || 1, tOutline: o.tOutline || '', tShadow: !!o.tShadow, tGrad: o.tGrad || '' },
    o.t === 'image' ? { look: o.look || {} } : {}, o.t === 'line' ? { head: o.head, tail: o.tail, flipH: !!o.flipH, flipV: !!o.flipV, bent: !!o.bent, start: o.start || null, end: o.end || null } : {},
    o.t === 'table' ? { colW: o.colW || null, merges: o.merges || [], cells: o.cells || {}, tstyle: o.tstyle || 'MediumStyle2Accent1', header: o.header !== false, banded: o.banded !== false, firstCol: !!o.firstCol } : {},
    o.t === 'group' ? { kids: (o.kids || []).map(objKey) } : {});
}
function boxProps(o, g) { return { x: cmStr(o.x / g.kx), y: cmStr(o.y / g.ky), w: cmStr(o.w / g.kx), h: cmStr(o.h / g.ky) }; }
/** The drawing id a line's end names in the file: the id in the target's path, once it has one. */
const cnvOf = (ref, slide) => { const t = ref && flatObjs(slide.objs).find(x => x.id === ref.id), m = t && t.path && /\[@id=(\d+)\]$/.exec(t.path); return m ? m[1] + ',' + ref.idx : ''; };
function objProps(o, g, orig, slide) {
  const p = {};
  const nb = boxProps(o, g), ob = orig ? boxProps(orig, g) : {};
  for (const k of ['x', 'y', 'w', 'h']) if (nb[k] !== ob[k]) p[k] = nb[k];
  const hexOf = c => unhex(g.th && c[0] !== '#' ? resolveColor(c, g.th, c) : c); // acc, card, sub, fg: the editor's theme colours, as hex for the file
  const was = k => orig ? orig[k] : undefined;
  if (o.t === 'text' || o.t === 'shape') {
    if (!orig || !sameRuns(orig.html, o.html)) p.html = o.html || '';
    const grad = f => f && String(f).startsWith('grad:');
    // a shape's null fill is the theme's accent (objView draws it so): the file gets the colour, not noFill
    if (was('fill') !== o.fill && (o.t === 'shape' || o.fill)) { if (grad(o.fill)) p.gradient = o.fill.slice(5).split(',').map((v, i) => i < 2 ? hexOf(v) : v).join(','); else { if (grad(was('fill'))) p.gradient = ''; p.fill = o.fill ? hexOf(o.fill) : o.fill == null && o.t === 'shape' ? hexOf('acc') : 'none'; } }
    if (was('color') !== o.color && o.color) p.color = hexOf(o.color);
    if (was('font') !== o.font && o.font) p.font = o.font;
    if (was('fs') !== o.fs && o.fs) p.size = (Math.round(o.fs / g.ptPx * 2) / 2) + 'pt';
    if (!orig && o.t === 'shape') p.geometry = geomBack(o.shape);
    if (!!was('lockAspect') !== !!o.lockAspect) p.lockAspect = o.lockAspect ? 'true' : 'false';
    // the box-wide text settings: a new object writes only what differs from the editor's defaults
    const pt = v => (Math.round(v / g.ptPx * 2) / 2) + 'pt';
    if (orig ? was('lh') !== o.lh : o.lh && o.lh !== 1.35 && o.lh !== 1.5) p.lineSpacing = String(o.lh || 1);
    if ((was('sb') || 0) !== (o.sb || 0)) p.spaceBefore = pt(o.sb || 0);
    if ((was('sa') || 0) !== (o.sa || 0)) p.spaceAfter = pt(o.sa || 0);
    if ((was('cs') || 0) !== (o.cs || 0)) p.charSpacing = String(Math.round((o.cs || 0) / g.ptPx * 10) / 10);
    if ((was('cols') || 1) !== (o.cols || 1)) p.columns = String(o.cols || 1);
    if ((was('vert') || 'horz') !== (o.vert || 'horz')) p.direction = o.vert || 'horz';
    const fitOf = x => (x.autofit || 'none') === 'shrink' && x.fit && x.fit < 1 ? 'shrink:' + Math.round(x.fit * 100) : x.autofit || 'none';
    if ((orig ? fitOf(orig) : 'none') !== fitOf(o)) p.autofit = fitOf(o);
    if ((was('tOutline') || '') !== (o.tOutline || '')) p.textOutline = o.tOutline ? hexOf(o.tOutline) : 'none';
    if (!!was('tShadow') !== !!o.tShadow) p.textShadow = o.tShadow ? 'true' : 'false';
    if ((was('tGrad') || '') !== (o.tGrad || '')) p.textGradient = o.tGrad ? o.tGrad.split(',').map((v, i) => i < 2 ? hexOf(v) : v).join(',') : '';
  }
  if (o.t === 'text' || o.t === 'shape' || o.t === 'line') {
    const on = o.sw > 0 && o.stroke, wasOn = orig && orig.sw > 0 && orig.stroke;
    if (on ? (!wasOn || orig.stroke !== o.stroke) : wasOn) p.line = on ? hexOf(o.stroke) : 'none';
    if (on && (!wasOn || orig.sw !== o.sw)) p.lineWidth = linePt(o.sw, g.ptPx);
    if ((was('dash') || 'solid') !== (o.dash || 'solid')) p.dash = o.dash || 'solid';
    if (!!was('shadow') !== !!o.shadow) p.shadow = o.shadow ? 'true' : 'false';
  }
  if (o.t === 'line') {
    if (!orig) p.geometry = o.bent ? 'bentConnector3' : 'straightConnector1';
    for (const k of ['head', 'tail']) if ((was(k) || 'none') !== (o[k] || 'none')) p[k] = o[k] || 'none';
    for (const k of ['flipH', 'flipV']) if (!!was(k) !== !!o[k]) p[k] = o[k] ? 'true' : 'false';
    if (slide) for (const k of ['start', 'end']) { const now = cnvOf(o[k], slide), before = orig ? cnvOf(orig[k], slide) : ''; if (now !== before || (!orig && now)) p[k] = now; }
  }
  if (o.t === 'table') {
    if (!orig || !same(orig.rows, o.rows)) p.data = JSON.stringify(o.rows || []);
    const widths = t => t.colW ? JSON.stringify(t.colW.map(w => cmStr(w / g.kx))) : null, nw = widths(o);
    if (nw && nw !== (orig ? widths(orig) : null) && (!p.w || orig)) p.widths = nw; // a new table's w already says the total; its own widths follow
    if ((was('tstyle') || 'MediumStyle2Accent1') !== (o.tstyle || 'MediumStyle2Accent1')) p.style = o.tstyle;
    for (const [k, f] of [['header', 'header'], ['banded', 'banded'], ['firstCol', 'firstCol']]) { const now = k === 'firstCol' ? !!o[k] : o[k] !== false, before = orig ? (k === 'firstCol' ? !!orig[k] : orig[k] !== false) : (k !== 'firstCol'); if (now !== before) p[f] = now ? 'true' : 'false'; }
  }
  if (o.t === 'image' && !orig) p.src = o.src;
  if (o.t === 'image') Object.assign(p, lookDiff(orig && orig.look, o.look));
  if (Math.round((orig && orig.rot) || 0) !== Math.round(o.rot || 0)) p.rotation = String(Math.round(o.rot || 0));
  return p;
}
/** The list marker and level of each paragraph of a text box, from its html: li in a ul is a bullet, in an ol a number; data-lvl
 *  (Tab / ⇧Tab) is the level. Nothing is written for a box without lists or levels. */
async function setListProps(file, shapePath, html, force) {
  const root = parseHtml(html || ''); const items = [];
  const walk = (n, inOl) => { for (const c of Array.from(n.children)) { if (c.tagName === 'UL' || c.tagName === 'OL') walk(c, c.tagName === 'OL'); else if (BLOCK.test(c.tagName)) items.push({ list: c.tagName === 'LI' ? (inOl ? 'number' : 'bullet') : 'none', level: Math.min(8, Math.max(0, +c.getAttribute('data-lvl') || 0)) }); } };
  walk(root, false);
  if (!force && !items.some(x => x.list !== 'none' || x.level)) return 0;
  let n = 0;
  for (let i = 0; i < items.length; i++) { await run(['set', file, `${shapePath}/paragraph[${i + 1}]`, '--prop', 'list=' + items[i].list, '--prop', 'level=' + items[i].level]); n++; }
  return n;
}
/** A picture's bytes as a data: URL, read in turn with the file's saves from its place in the file (or from the picture it copies);
 *  null when it has none there. */
export function pictureData(file, o) {
  const at = !String(o.src || '').startsWith('data:') && (o.path || o.from);
  return at ? inLane(file, () => dataUrl(binaryUrl(file, at))).catch(() => null) : Promise.resolve(null);
}
async function dataUrl(url) {
  const r = await http(url), b = new Uint8Array(await r.arrayBuffer());
  let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode(...b.subarray(i, i + 32768));
  return 'data:' + (r.headers.get('content-type') || 'image/png') + ';base64,' + btoa(s);
}
const slideOf = path => path.slice(0, path.indexOf(']') + 1);
const idPath = (slide, node) => slide + '/' + node.kind + '[@id=' + node.props.id + ']';
/** A copied object takes the path its copy has in the file (moved: the original's path → the copy's), its group members with it. */
function adoptCopy(doc, x, path, moved) {
  x.path = path; if (x.t === 'image') x.src = binaryUrl(doc.path, x.path);
  for (const k of x.kids || []) { const src = k.from && !k.path ? k.from : k.path; if (src) adoptCopy(doc, k, moved(src), moved); }
}
/** A snapshot entry with its paths moved, members included. */
const movedKey = (y, moved) => Object.assign({}, y, { path: moved(y.path) }, y.kids ? { kids: y.kids.map(k => movedKey(k, moved)) } : {});
/** After a slide took a layout (a new slide, or set layout=), the engine's placeholders — empty, placed by the layout — are
 *  bound to the editor's placeholder objects the way the engine binds them: by family, in drawing order. A bound one is saved
 *  through its path, and nothing is written while both stay empty; an engine placeholder nobody claims goes; an editor object
 *  nobody has is added as an ordinary shape by the main pass. Snapshot o learns the slide's shapes as the engine has them. */
async function bindPlaceholders(doc, g, s, o) {
  let n = 0;
  const made = await run(['get', doc.path, s.path, '--depth', '2']); // depth 2: the shapes with their props (placeholder, id, box); their paragraphs as stubs
  const live = (made.children || []).filter(c => c.kind !== 'decor').map(c => pptxObject(doc.path, s.path, g, c)), paths = new Set(live.map(y => y.path));
  for (const x of s.objs) if (x.path && !paths.has(x.path)) delete x.path; // dropped with the old layout (an empty leftover); made anew if the editor still shows it
  o.objs = o.objs.filter(y => paths.has(y.path));
  for (const y of live) {
    const kept = o.objs.find(z => z.path === y.path);
    const x = kept ? s.objs.find(z => z.path === y.path) : y.ph && s.objs.find(z => !z.path && z.ph && phFamily(z.ph) === phFamily(y.ph));
    if (!x) { if (!kept) { await run(['remove', doc.path, y.path]); n++; } continue; } // what the editor deleted, the main pass removes
    if (!kept) { x.path = y.path; x.kind = 'shape'; o.objs.push(objKey(Object.assign({}, y, { id: x.id, fs: x.fs }))); }
    // ponytail: a placeholder's snapshot takes the editor's box, so the layout's place is never pinned into the file; a deck whose own layout
    // differs from LAYOUT_SPECS shows the editor's boxes until it is reopened (the file has the layout's)
    if (x.ph) Object.assign(o.objs.find(z => z.path === y.path), { x: x.x, y: x.y, w: x.w, h: x.h, lh: x.lh }); // and the line height the editor draws a body with
  }
  return n;
}
async function savePptx(doc, log) {
  let n = 0;
  const orig = doc._orig || { geo: null, slides: [] }, g = Object.assign({ kx: SW / 33.867, ky: 900 / 19.05, ptPx: SW / 960 }, orig.geo, { th: THEMES[doc.theme] || THEMES.paper });
  const origSlides = orig.slides, at = path => path && origSlides.find(x => x.path === path);
  // 设计's palette goes into the deck's theme (every slide's inherited colours and fonts follow); a deck that never wore one keeps its own look under 素白
  if (THEMES[doc.theme] && doc.theme !== (orig.palette || 'paper')) { await run(['set', doc.path, '/', '--prop', 'palette=' + doc.theme]); n++; orig.palette = doc.theme; log && log('set', '/', { palette: doc.theme }); }
  // ponytail: the markup of every object a save removes is kept for the session (a few KB each), so an undo can put it back exactly
  const removed = orig.removed = orig.removed || {};
  // What the file lacks goes in first, while the file still is what the snapshot says. A copy of a slide or an object the file has is
  // copied there by the engine: the same markup, so what the editor does not model comes too (a title's inherited bold, table
  // styles, effects, groups, charts). An object a save removed and an undo brought back returns from the markup that save kept.
  // Only what the file never had is made from the model.
  for (const s of doc.slides) {
    if (at(s.path)) continue;
    const src = at(s.from);
    if (src) {
      const r = await run(['copy', doc.path, s.from, '--to', '/', '--after', s.from]); n++;
      s.path = '/slide[@id=' + r.props.id + ']'; log && log('copy', s.path, s.from);
      const moved = p => s.path + p.slice(src.path.length), taken = new Set();
      origSlides.splice(origSlides.indexOf(src) + 1, 0, Object.assign(JSON.parse(JSON.stringify(src)), { id: s.id, path: s.path, sec: '', objs: src.objs.map(y => movedKey(y, moved)) })); // a copy starts no section
      for (const x of s.objs) if (!x.path && x.from && slideOf(x.from) === src.path && !taken.has(x.from)) { taken.add(x.from); adoptCopy(doc, x, moved(x.from), moved); }
    } else {
      // a slide pasted from another deck may name a layout this one lacks: it goes on a blank one
      const r = await run(['add', doc.path, '/', '--type', 'slide', '--prop', 'layout=' + (s.layout || 'Blank')]).catch(() => { s.layout = 'Blank'; return run(['add', doc.path, '/', '--type', 'slide', '--prop', 'layout=Blank']); }); n++;
      s.path = '/slide[@id=' + r.props.id + ']'; const made = { path: s.path, layout: s.layout, bg: '#FFFFFF', sec: '', notes: '', trans: 'none', duration: null, hidden: false, objs: [] };
      origSlides.push(made); log && log('add', s.path, 'slide');
      n += await bindPlaceholders(doc, g, s, made);
    }
  }
  for (const s of doc.slides) {
    const o = at(s.path);
    for (const x of s.objs) {
      if (x.path ? o.objs.some(y => y.path === x.path) : !x.from) continue; // in the file, or never was
      const was = !x.path && origSlides.map(os => os.objs.find(y => y.path === x.from)).find(Boolean), gone = removed[x.path || x.from];
      let r = null;
      if (was) { r = await run(['copy', doc.path, x.from, '--to', s.path]); n++; log && log('copy', x.from, s.path); }
      else if (gone && at(slideOf(gone.snap.path))) {
        const home = slideOf(gone.snap.path);
        r = await run(['add', doc.path, home, '--raw', gone.xml]); n++; log && log('add', home, 'raw');
        if (home !== s.path) { const put = idPath(home, r); r = await run(['copy', doc.path, put, '--to', s.path]); await run(['remove', doc.path, put]); n += 2; }
      }
      if (!r) continue; // made from the model below
      const srcKey = was || gone.snap, moved = p => idPath(s.path, r) + p.slice(srcKey.path.length);
      adoptCopy(doc, x, idPath(s.path, r), moved); x.kind = r.kind;
      o.objs.push(Object.assign(movedKey(srcKey, moved), { id: x.id }));
    }
  }
  const keep = new Set(doc.slides.map(s => s.path));
  const order = [];
  for (const o of origSlides) { if (!keep.has(o.path)) { await run(['remove', doc.path, o.path]); n++; log && log('remove', o.path); } else order.push(o.path); }
  // 取消组合: a group the editor dissolved (its members now on the slide, still under their group paths) is dissolved in the file too,
  // which keeps the members' markup; their paths become the slide's
  for (const s of doc.slides) {
    const o = at(s.path);
    for (const og of o.objs.filter(y => y.t === 'group' && !s.objs.some(x => x.path === y.path))) {
      const kids = s.objs.filter(x => x.path && x.path.startsWith(og.path + '/'));
      if (!kids.length) continue;
      await run(['set', doc.path, og.path, '--prop', 'ungroup=true']); n++; log && log('set', og.path, { ungroup: 'true' });
      const tail = p => s.path + p.slice(p.lastIndexOf('/'));
      for (const x of kids) { x.path = tail(x.path); if (x.t === 'image') x.src = binaryUrl(doc.path, x.path); }
      o.objs.splice(o.objs.indexOf(og), 1, ...(og.kids || []).map(k => Object.assign({}, k, { path: tail(k.path) })));
    }
  }
  for (let i = 0; i < doc.slides.length; i++) {
    const s = doc.slides[i], o = at(s.path);
    if (order.indexOf(s.path) !== i) { await run(['move', doc.path, s.path, '--to', '/', '--index', String(i + 1)]); n++; order.splice(order.indexOf(s.path), 1); order.splice(i, 0, s.path); }
    if ((o.bg || '#FFFFFF') !== (s.bg || '#FFFFFF')) { await run(['set', doc.path, s.path, '--prop', 'background=' + unhex(s.bg || '#FFFFFF')]); n++; }
    const sp = slideProps(o, s);
    if (Object.keys(sp).length) { await run(['set', doc.path, s.path, ...propsArgs(sp)]); n++; log && log('set', s.path, sp); }
    if (s.layout && o.layout && s.layout !== o.layout) {
      // 版式: the engine keeps a placeholder with text and drops an empty one, so it must see the editor's text first
      for (const x of s.objs) { const oo = x.path && x.ph && o.objs.find(y => y.path === x.path); if (oo && !sameRuns(oo.html, x.html)) { await run(['set', doc.path, x.path, '--prop', 'html=' + (x.html || '')]); n++; oo.html = x.html; } }
      await run(['set', doc.path, s.path, '--prop', 'layout=' + s.layout]); n++; o.layout = s.layout; log && log('set', s.path, { layout: s.layout });
      n += await bindPlaceholders(doc, g, s, o);
    }
    const keepObjs = new Set(s.objs.filter(x => x.path).map(x => x.path));
    for (const oo of o.objs.slice().reverse()) if (!keepObjs.has(oo.path)) {
      removed[oo.path] = { xml: String(await run(['get', doc.path, oo.path, '--raw'])).trim(), snap: oo };
      await run(['remove', doc.path, oo.path]); n++; log && log('remove', oo.path);
    }
    for (const x of s.objs) {
      if (x.t === 'group') { n += await saveGroup(doc, g, s, o, x, log); continue; }
      n += await saveObj(doc, g, s, o, x, x.path ? o.objs.find(y => y.path === x.path) : null, log);
    }
    // the animations last: they name the objects by drawing id, which a new object has only now
    const anims = animJson(s);
    if (!same(anims, o.anims || [])) { await run(['set', doc.path, s.path, '--prop', 'animations=' + JSON.stringify(anims)]); n++; o.anims = anims; log && log('set', s.path, { animations: anims.length }); }
  }
  return n;
}
/** One object into the file: set what changed against its snapshot oo, or add it (under `into`, the slide or a group's path). */
async function saveObj(doc, g, s, o, x, oo, log, into) {
  let n = 0;
  const p = objProps(x, g, oo, s);
  if (oo) {
    if (Object.keys(p).length) { await run(['set', doc.path, x.path, ...propsArgs(p)]); n++; log && log('set', x.path, p); }
    // the markers and levels: after a text rewrite when either side had any, or when only the markup changed (the runs are the same)
    const lists = h => /<li|data-lvl/i.test(h || '');
    if (p.html != null ? lists(x.html) || lists(oo.html) : (x.t === 'text' || x.t === 'shape') && (oo.html || '') !== (x.html || '')) n += await setListProps(doc.path, x.path, x.html, true);
    if (x.t === 'table') n += await setCellProps(doc.path, x, oo, p.data != null);
    return n;
  }
  if (x.t === 'image' && !(x.src || '').startsWith('data:')) return n;
  const kind = x.t === 'image' ? 'image' : x.t === 'table' ? 'table' : x.t === 'line' ? 'connector' : 'shape';
  const r = await run(['add', doc.path, into || s.path, '--type', kind, ...propsArgs(p)]); n++;
  x.path = (into || s.path) + '/' + kind + '[@id=' + r.props.id + ']'; x.kind = kind; log && log('add', x.path, kind);
  if (kind === 'shape') n += await setListProps(doc.path, x.path, x.html);
  if (kind === 'table') { if (x.colW) { await run(['set', doc.path, x.path, '--prop', 'widths=' + JSON.stringify(x.colW.map(w => cmStr(w / g.kx)))]); n++; } n += await setCellProps(doc.path, x, null, true); }
  if (kind === 'image') {
    x.src = binaryUrl(doc.path, x.path); // from now on it loads from its own place in the file
    if (p.crop) { await run(['set', doc.path, x.path, ...propsArgs(boxProps(x, g))]); n++; } // a crop keeps the picture's scale, so it moved the frame
  }
  return n;
}
/** A table's cells: merges as the anchor's colspan / rowspan (a merge that went sets them back to 1), and each cell's own fill,
 *  line and alignment. Only what changed against the snapshot oo; everything after a data rewrite that changed the grid's size. */
async function setCellProps(file, x, oo, rewritten) {
  const key = m => m.r + ':' + m.c, cellPath = k => { const [r, c] = k.split(':'); return `${x.path}/row[${+r + 1}]/cell[${+c + 1}]`; };
  const before = oo || { merges: [], cells: {}, rows: [] }, resized = !oo || rewritten && (before.rows.length !== x.rows.length || (before.rows[0] || []).length !== (x.rows[0] || []).length);
  const want = {};
  const put = (k, props) => { want[k] = Object.assign(want[k] || {}, props); };
  const newM = Object.fromEntries((x.merges || []).map(m => [key(m), m])), oldM = Object.fromEntries((before.merges || []).map(m => [key(m), m]));
  for (const k of new Set([...Object.keys(newM), ...Object.keys(oldM)])) { const a = newM[k], b = oldM[k]; if (resized || !b || !a || a.rs !== b.rs || a.cs !== b.cs) put(k, { colspan: String(a ? a.cs : 1), rowspan: String(a ? a.rs : 1) }); }
  const cellsNow = x.cells || {}, cellsThen = before.cells || {};
  for (const k of new Set([...Object.keys(cellsNow), ...Object.keys(cellsThen)])) {
    const a = cellsNow[k] || {}, b = cellsThen[k] || {}; if (!resized && same(a, b)) continue;
    const props = {}; if (resized ? a.fill : a.fill !== b.fill) props.fill = a.fill ? unhex(a.fill) : 'none'; if (resized ? a.line : a.line !== b.line) props.line = a.line && a.line !== 'none' ? unhex(a.line) : 'none'; if (resized ? a.align : (a.align || 'left') !== (b.align || 'left')) props.align = a.align || 'left';
    if (Object.keys(props).length) put(k, props);
  }
  let n = 0;
  const inGrid = k => { const [r, c] = k.split(':').map(Number); return r < x.rows.length && c < (x.rows[r] || []).length; };
  // anchors whose span shrinks go first, so a cell another merge grows over is free by then
  const release = k => want[k].colspan === '1' && want[k].rowspan === '1' ? 0 : 1;
  for (const k of Object.keys(want).sort((a, b) => release(a) - release(b))) if (inGrid(k)) { await run(['set', file, cellPath(k), ...propsArgs(want[k])]); n++; }
  return n;
}
/** A group: new ones are made of members put on the slide first; a saved one moves and resizes as one (the engine carries its
 *  members, whose snapshot follows), then each member's own changes go in. */
async function saveGroup(doc, g, s, o, x, log) {
  let n = 0;
  const oo = x.path && o.objs.find(y => y.path === x.path);
  if (!oo) {
    for (const k of x.kids) n += await saveObj(doc, g, s, o, k, null, log);
    const members = x.kids.filter(k => k.path).map(k => k.path.slice(k.path.lastIndexOf('/') + 1));
    if (!members.length) return n;
    const r = await run(['add', doc.path, s.path, '--type', 'group', '--prop', 'members=' + members.join(',')]); n++;
    x.path = s.path + '/group[@id=' + r.props.id + ']'; x.kind = 'group'; log && log('add', x.path, 'group');
    for (const k of x.kids) if (k.path) { k.path = x.path + k.path.slice(k.path.lastIndexOf('/')); if (k.t === 'image') k.src = binaryUrl(doc.path, k.path); }
    if (x.rot) { await run(['set', doc.path, x.path, '--prop', 'rotation=' + Math.round(x.rot)]); n++; }
    o.objs.push(objKey(x));
    return n;
  }
  const box = union(x.kids), bp = objProps(Object.assign({}, x, box, { t: 'group' }), g, oo, s);
  if (Object.keys(bp).length) {
    await run(['set', doc.path, x.path, ...propsArgs(bp)]); n++; log && log('set', x.path, bp);
    const sx = oo.w ? box.w / oo.w : 1, sy = oo.h ? box.h / oo.h : 1; // the engine moved the members with the group: so does their snapshot
    for (const k of oo.kids || []) Object.assign(k, { x: Math.round(box.x + (k.x - oo.x) * sx), y: Math.round(box.y + (k.y - oo.y) * sy), w: Math.round(k.w * sx), h: Math.round(k.h * sy) });
    Object.assign(oo, box, { rot: x.rot || 0 });
  }
  for (const k of x.kids) n += await saveObj(doc, g, s, o, k, k.path ? (oo.kids || []).find(y => y.path === k.path) : null, log, x.path);
  oo.kids = x.kids.map(objKey);
  return n;
}

// ----- mm (FreeMind mind map): the topic tree is the model; ids are the engine's, new nodes carry tmp_ ids until saved -----
async function openMm(doc) {
  let t = await tree(doc.path);
  // topics without an id (FreeMind files, or a map nobody has edited yet) cannot be addressed: one no-op write makes the engine assign ids
  const missing = n => n.kind === 'topic' && !(n.props && n.props.id) || (n.children || []).some(missing);
  const root = (t.children || []).find(c => c.kind === 'topic');
  if (root && missing(root)) { await run(['set', doc.path, '/topic[1]', '--prop', 'text=' + ((root.props && root.props.text) || '')]); t = await tree(doc.path); }
  const conv = n => {
    const p = n.props || {}, m = { id: p.id, text: p.text || '', children: (n.children || []).filter(c => c.kind === 'topic').map(conv) };
    for (const k of ['collapsed', 'bold', 'italic', 'strike', 'mono']) if (p[k] === 'true' || p[k] === true) m[k] = true;
    for (const k of ['side', 'note', 'link', 'color', 'fill', 'icon', 'font', 'free', 'labels', 'image', 'imageSize', 'cloud', 'summary', 'structure', 'theme', 'lines']) if (p[k]) m[k] = p[k];
    if (+p.size > 0) m.size = +p.size;
    if (p.rels) { try { const r = MM.relList({ rels: JSON.parse(p.rels) }); if (r.length) m.rels = r; } catch (e) { } }
    return m;
  };
  const top = (t.children || []).find(c => c.kind === 'topic');
  const map = top ? conv(top) : { id: 'root', text: doc.title, children: [] };
  return { map, _orig: MM.flatten(map) };
}
/** An .xmind file: the engine writes it as a FreeMind map next to it (plan.xmind → plan.mm, a free name), and that map is
 *  what opens — under the .xmind's own doc id, so the tab that asked for it shows it. The .xmind itself stays untouched. */
async function openXmind(doc) {
  const path = await freeName(titleOf(doc.path), 'mm', dirOf(doc.path));
  await run(['export', doc.path, '--to', path]);
  const mm = lazy(path, 'mm'), model = await openMm(mm), st = await stat(path).catch(() => null);
  return Object.assign({}, doc, mm, model, { id: doc.id, loaded: true, dirty: false, _mtime: st ? st.mtime : null });
}
/** Runs the planner's steps in order; added nodes get their engine id written into the model and into doc._tmpToId (old id → real id). */
async function saveMm(doc, log) {
  const { steps, ids } = MM.plan(doc._orig, doc.map, doc.path);
  let n = 0;
  for (const s of steps) { const argv = s.argv, r = await run(argv); s.onResult && s.onResult(r); n++; log && log(argv[0], argv[2]); doc._tmpToId = Object.assign({}, doc._tmpToId, ids); }
  if (n) { MM.walk(doc.map, x => MM.retarget(x, doc._tmpToId || {})); doc._orig = MM.flatten(doc.map); }
  return n;
}
function adoptMm(cur, saved) {
  cur._tmpToId = Object.assign({}, cur._tmpToId, saved._tmpToId);
  const m = cur._tmpToId, real = id => { for (let i = 0; m[id] && i < 20; i++) id = m[id]; return id; };
  if (cur.map) MM.walk(cur.map, x => { x.id = real(x.id); MM.retarget(x, m); });
}

// ---------- save ----------
// A file's saves and picture commands take turns: each engine command writes the whole file, so two at once would lose one.
const lanes = new Map();
function inLane(file, fn) {
  const turn = (lanes.get(file) || Promise.resolve()).then(fn);
  lanes.set(file, turn.then(() => { }, () => { }));
  return turn;
}

/** Writes the model's changes to the file. Resolves with the number of commands run; 0 means nothing differed. */
export async function save(doc, opts) {
  if (!doc || !doc.loaded) return 0;
  return inLane(doc.path, () => saveNow(doc, opts || {}));
}
async function saveNow(doc, opts) {
  const log = opts.log, root = opts.root;
  if (doc.type === 'md') {
    if ((doc.text || '') === (doc._orig || '')) return 0;
    // ifMtime is the real guard: if another program changed the file since it was read, the engine refuses (409) instead
    // of landing on top of it, and throws — the caller reconciles rather than this function silently overwriting it
    const r = await http('/file?file=' + enc(doc.path) + (doc._mtime != null ? '&ifMtime=' + doc._mtime : ''), { method: 'PUT', body: doc.text || '' });
    doc._orig = doc.text || ''; doc._mtime = (await r.json()).mtime;
    return 1;
  }
  if (doc.type === 'pdf') return savePdf(doc);
  let n = 0;
  if (doc.type === 'mm') n = await saveMm(doc, log);
  else if (doc.type === 'docx') n = await saveDocx(doc, root, log);
  else if (doc.type === 'xlsx') { n = await saveXlsx(doc, log); if (n) doc._orig = origOf(doc.sheets); }
  else if (doc.type === 'pptx') { n = await savePptx(doc, log); if (n) doc._orig = Object.assign({}, doc._orig, { slides: pptxSnapshot(doc.slides) }); }
  // these save through many small engine commands, not one PUT, so there is no ifMtime guard for them (see checkExternal
  // in the shell, which catches a conflict before scheduling a save rather than mid-flight): just remember the mtime our
  // own write left, so the next poll does not mistake it for an external change
  if (n) { const st = await stat(doc.path).catch(() => null); if (st) doc._mtime = st.mtime; }
  return n;
}

/** Carries the paths and the saved snapshot from a doc that was just saved onto the doc the editor holds now
 * (they differ when the user kept editing during the save). */
export function adopt(cur, saved, withHtml) {
  if (cur === saved) return cur;
  cur._orig = saved._orig;
  cur._mtime = saved._mtime;
  if (cur.type === 'docx' && withHtml) cur.html = saved.html;
  if (cur.type === 'pptx') {
    const bySlide = new Map(saved.slides.map(s => [s.id, s]));
    // a save gives paths to what it put in the file, and new ones to what it put back or copied under ids another shape had taken
    const take = (xs, ys) => { const byObj = new Map(ys.map(x => [x.id, x])); xs.forEach(x => { const y = byObj.get(x.id); if (y) { if (y.path && x.path !== y.path) { x.path = y.path; if (x.t === 'image') x.src = y.src; } if (!x.kind) x.kind = y.kind; if (x.kids && y.kids) take(x.kids, y.kids); } }); };
    cur.slides.forEach(s => { const o = bySlide.get(s.id); if (!o) return; if (o.path && s.path !== o.path) s.path = o.path; take(s.objs, o.objs); });
  }
  if (cur.type === 'mm') adoptMm(cur, saved);
  if (cur.type === 'pdf' && saved._saved) { // the saved pages now come from the new generation; what was written leaves the pending lists
    const s = saved._saved, by = new Map(saved.pages.map(p => [p.id, p])), was = new Map(s.before.map(p => [p.id, p]));
    cur.pages = cur.pages.map(p => { const n = by.get(p.id), w = was.get(p.id); return n && w && w.from === p.from && w.src === p.src ? Object.assign({}, n, { rot: (((p.rot || 0) - (w.rot || 0)) % 360 + 360) % 360 }) : p; });
    const flushed = new Set(s.annots), rm = new Set(s.removed);
    cur.annots = (cur.annots || []).filter(a => !flushed.has(a.id)); cur.removed = (cur.removed || []).filter(x => !rm.has(x));
    cur._gen = saved._gen; cur._n = saved._n; cur._formSaved = saved._formSaved;
  }
  if (cur.type === 'xlsx') cur.sheets.forEach((s, i) => { const o = saved.sheets.find(x => x.name === s.name) || saved.sheets[i]; if (!o) return; if (!s.path) s.path = o.path; (s.charts || []).forEach(ch => { const oc = (o.charts || []).find(x => x.id === ch.id); if (oc) { if (!ch.path) ch.path = oc.path; if (ch.eid == null) ch.eid = oc.eid; } });
    (s.images || []).forEach(im => { const oi = !im.path && (o.images || []).find(x => x.id === im.id); if (oi && oi.path) Object.assign(im, { path: oi.path, src: oi.src }); }); });
  return cur;
}

// ---------- reload when the file changes outside Writer ----------
/** What a changed on-disk mtime means for an open document: 'none' — tracked is unset or unchanged, nothing to do;
 *  'reload' — no local edits are pending, safe to load the new version in place; 'conflict' — local edits (a scheduled or
 *  just-failed save) exist too, so loading the new version would lose them: keep both instead. Pure, so it is the same
 *  decision whether a poll, a focus/visibilitychange, or a 409 from a save is what noticed the change. */
export function reloadDecision(trackedMtime, diskMtime, pending) {
  if (trackedMtime == null || diskMtime === trackedMtime) return 'none';
  return pending ? 'conflict' : 'reload';
}

/** The path for a conflict copy beside `path`: name（冲突副本）.ext, or name（冲突副本 2）.ext if that is taken. */
export function conflictPath(path, dir) { return freeName(titleOf(path) + _t('（冲突副本）'), path.split('.').pop(), dir); }

/** A docx html string with every pre-existing image dropped: its src is /binary?file=doc.path&path=..., which only doc.path
 *  (not the blank document a conflict copy replays onto) can resolve. A freshly inserted, still-unsaved image (a data: src)
 *  is kept. ponytail: known ceiling — a conflict copy of a docx loses images that were already in the file; the reloaded
 *  original (still doc.path) keeps them safely. Upgrade: fetch each as a data: URL first if this needs to be lossless. */
function dropOldImages(html) {
  const root = parseHtml(html || '');
  root.querySelectorAll('img').forEach(img => { if (!/^data:/.test(img.getAttribute('src') || '')) (img.closest('[data-path]') || img).remove(); });
  return root.innerHTML;
}

/** Materializes doc's current in-memory edits (not yet on disk at doc.path, which now holds someone else's change) into a
 *  new file beside it, so neither side is lost. For md it is one PUT of the live text. Every other format saves through
 *  many small engine commands that assume the file they are opened against, so this replays them onto a freshly created
 *  blank document instead of doc.path — the same replay a brand new document's first save already does — never against
 *  the live editor (opts carries no root, so this parses a copy of doc.html rather than touching the on-screen DOM; the
 *  caller must still replace doc with a fresh EN.open right after, since the paths this leaves on doc belong to the copy,
 *  not doc.path). Falls back to the drafts folder when the sibling name is refused (a document opened from outside the
 *  workspace). A pptx/xlsx/mm conflict that references binary content the blank document cannot resolve throws rather
 *  than writing a broken copy — the caller leaves the local edits in the editor rather than reload over them. */
export async function saveConflictCopy(doc) {
  const make = async dir => {
    const path = await conflictPath(doc.path, dir);
    if (doc.type === 'md') await http('/file?file=' + enc(path), { method: 'PUT', body: doc.text || '' });
    else if (doc.type === 'pdf') await http('/file?file=' + enc(path), { method: 'PUT', body: await PK.saveBytes(doc.id, doc) });
    else {
      await run(['create', path]);
      const blank = await open(lazy(path, doc.type));
      const tmp = Object.assign({}, doc, { path, _orig: blank._orig });
      if (tmp.type === 'docx') tmp.html = dropOldImages(tmp.html);
      await saveNow(tmp, {});
    }
    return path;
  };
  try { return await make(dirOf(doc.path)); }
  catch (e) { return await make(state.drafts || ''); }
}

// ---------- change detection between two loaded models (for AI change cards) ----------
/** Marks what differs between before and after in place on `after` and returns [[kind, label], ...]. */
export function diffMark(before, after) {
  const items = [];
  if (!before || !after || before.type !== after.type) return items;
  if (after.type === 'docx') {
    const root = parseHtml(after.html || '');
    items.push(...diffBlocks(blocksFromHtml(before.html || ''), blocksFromHtml(root)));
    after.html = root.innerHTML;
    const bp = before.page || {}, ap = after.page || {};
    if (['size', 'orient', 'margin', 'cols'].some(k => String(bp[k] == null ? '' : bp[k]) !== String(ap[k] == null ? '' : ap[k]))) items.push([_t('修改'), _t('页面设置')]);
    if ((before.header || '') !== (after.header || '')) items.push([_t('修改'), _t('页眉')]);
    if ((before.footer || '') !== (after.footer || '')) items.push([_t('修改'), _t('页脚')]);
  } else if (after.type === 'xlsx') {
    const to = pairUp(before.sheets.map(x => x.name), after.sheets.map(x => x.name)); // by name: a sheet added in front moves the others' paths (/sheet[2])
    after.sheets.forEach((s, i) => {
      const o = before.sheets[to[i]];
      if (!o) { items.push([_t('新增'), _t('工作表 {name}', { name: s.name })]); return; }
      let n = 0;
      for (const ref of new Set([...Object.keys(o.cells), ...Object.keys(s.cells)])) if (!same(o.cells[ref], s.cells[ref])) { if (s.cells[ref]) s.cells[ref].ai = true; n++; }
      if (n) items.push([_t('修改'), _t('{name} · {n} 个单元格', { name: s.name, n })]);
      const oc = new Map((o.charts || []).map(ch => [chartKey(ch), ch]));
      (s.charts || []).forEach(ch => { const x = oc.get(chartKey(ch)); if (!x || !same(chartProps(x), chartProps(ch))) { ch.ai = true; items.push([_t(x ? '修改' : '新增'), _t('{name} · 图表 {title}', { name: s.name, title: ch.title || ch.type })]); } });
      const gone = (o.charts || []).filter(ch => !(s.charts || []).some(x => chartKey(x) === chartKey(ch))).length;
      if (gone) items.push([_t('删除'), _t('{name} · {n} 个图表', { name: s.name, n: gone })]);
    });
    const gone = before.sheets.length - to.filter(i => i >= 0).length;
    if (gone) items.push([_t('删除'), _t('{n} 个工作表', { n: gone })]);
  } else if (after.type === 'pptx') {
    after.slides.forEach((s, i) => {
      const o = before.slides.find(x => x.path === s.path);
      if (!o) { s.ai = true; items.push([_t('新增'), _t('第 {n} 页', { n: i + 1 })]); return; }
      if (!same(o.objs.map(objKey), s.objs.map(objKey)) || o.bg !== s.bg || o.notes !== s.notes || o.trans !== s.trans || o.duration !== s.duration || o.hidden !== s.hidden || !same(o.anims || [], s.anims || []) || (o.sec || '') !== (s.sec || '')) { s.ai = true; items.push([_t('修改'), _t('第 {n} 页', { n: i + 1 })]); }
    });
    const kept = new Set(after.slides.map(s => s.path)), gone = before.slides.filter(o => !kept.has(o.path)).length; // by the engine's slide id, which the path carries
    if (gone) items.push([_t('删除'), _t('{n} 页', { n: gone })]);
  } else if (after.type === 'mm') { items.push(...MM.markAi(before.map, after.map));
  } else if (after.type === 'md' && (before.text || '') !== (after.text || '')) items.push([_t('修改'), _t('文本')]);
  return items;
}
/** The Word part of diffMark. Blocks are paired like the lines of a diff, not by path: the engine numbers them by position, so
 *  one added in the middle moves the paths of all after it. Marks each new or changed block's element; returns its entries. */
export function diffBlocks(before, after) {
  const ta = before.map(blockText), tb = after.map(blockText), ka = ta.map((t, i) => before[i].kind + ':' + t), kb = tb.map((t, j) => after[j].kind + ':' + t);
  // left between two unchanged blocks: the same text pairs first (a paragraph that became a heading), then the same kind (an edit)
  const items = [], to = pairUp(ka, kb, [(i, j) => ta[i] !== '' && ta[i] === tb[j], (i, j) => before[i].kind === after[j].kind]);
  after.forEach((b, j) => {
    const o = before[to[j]];
    const differs = !o || ka[to[j]] !== kb[j] || lookOf(o) !== lookOf(b) || (b.kind === 'table' ? !same(o.rows.map(r => r.cells.map(c => runsOf(c.props.html))), b.rows.map(r => r.cells.map(c => runsOf(c.props.html)))) : b.kind === 'code' ? o.props.text !== b.props.text : b.kind === 'image' || b.kind === 'pagebreak' ? false : !sameRuns(o.props.html, b.props.html));
    if (differs) { b.el?.setAttribute('data-ai', '1'); items.push([_t(o ? '修改' : '新增'), labelOf(b)]); }
  });
  const gone = before.length - to.filter(i => i >= 0).length;
  if (gone) items.push([_t('删除'), _t('{n} 处内容', { n: gone })]);
  return items;
}
/** What a block says, without markup or spaces (the editor's markup and the engine's differ there); with its kind, what pairs it. */
const blockText = b => (b.rows ? b.rows.map(r => r.cells.map(c => c.props.html).join('|')).join('|') : b.props.html || b.props.text || '').replace(/<[^>]*>|&nbsp;|[\s\u200B]+/g, '');
/** How a block shows besides its kind and text: heading or list level, list, paragraph style (Normal named or not), alignment.
 *  Compared once blocks are paired, never part of the pairing: a restyled paragraph still pairs with itself. */
const lookOf = b => [b.props.level, b.props.list, (b.props.style || '').replace(/^Normal$/i, ''), b.align || b.props.align || 'left', ...PARA_OWN.map(k => b.props[k] || '')].join('|');
/** Pairs two lists of keys like a line diff: equal keys along a longest common subsequence, then between two of those what is
 *  left on both sides, in order, by each test of akin in turn (an edited item). Returns for each item of b its partner's index in a, or -1. */
function pairUp(a, b, akin = [() => true]) {
  const to = new Array(b.length).fill(-1);
  let s = 0, n = a.length, m = b.length;
  while (s < n && s < m && a[s] === b[s]) { to[s] = s; s++; }
  while (n > s && m > s && a[n - 1] === b[m - 1]) { n--; m--; to[m] = n; }
  // ponytail: an N×M table over what lies between the common start and end; past 16M cells (4000+ blocks, changed at both ends)
  // the whole middle is one gap for akin, which is quadratic (0.7 s for 10 000 blocks all rewritten). Myers' O(ND) diff if that matters.
  const N = n - s, M = m - s, W = M + 1, L = N * M <= 16e6 ? new Uint16Array((N + 1) * W) : null;
  if (L) for (let i = N - 1; i >= 0; i--) for (let j = M - 1; j >= 0; j--) L[i * W + j] = a[s + i] === b[s + j] ? L[(i + 1) * W + j + 1] + 1 : Math.max(L[(i + 1) * W + j], L[i * W + j + 1]);
  let gi = [], gj = [];
  const flush = () => { for (const ok of akin) for (const j of gj) if (to[s + j] < 0) { const k = gi.findIndex(i => ok(s + i, s + j)); if (k >= 0) to[s + j] = s + gi.splice(k, 1)[0]; } gi = []; gj = []; };
  for (let i = 0, j = 0; i < N || j < M;) {
    if (L && i < N && j < M && a[s + i] === b[s + j]) { flush(); to[s + j] = s + i; i++; j++; }
    else if (j === M || i < N && (!L || L[(i + 1) * W + j] >= L[i * W + j + 1])) gi.push(i++);
    else gj.push(j++);
  }
  flush();
  return to;
}
/** A block's entry in the change list: its kind, then what tells it apart (its text, a table's rows), each said once. */
function labelOf(b) {
  const p = b.props, fixed = { image: '图片', code: '代码块', pagebreak: '分页符' }[b.kind];
  if (fixed) return _t(fixed);
  if (b.kind === 'table') return _t('表格 · {n} 行', { n: b.rows.length });
  if (b.kind === 'toc') return p.title && p.title !== '目录' ? _t('目录') + ' · ' + p.title : _t('目录');
  return _t(b.kind === 'heading' ? '标题' : p.list ? '列表项' : '段落') + ' · ' + (plainOf(p.html || '').slice(0, 24) || _t('（空）'));
}

// ---------- pictures: the 图片 tab of the Word, slide and sheet editors ----------
// Each tool is one `set <picture> --prop …`, in turn with the file's saves. The answer goes into the model and into its saved
// snapshot (pictureSaved), so the next save sends nothing for it, and an undo — the model going back — sends the way back.

/** The picture props the editors keep, as `get` prints them (picture.js draws them). */
export const LOOK = ['crop', 'rotation', 'flipH', 'flipV', 'brightness', 'contrast', 'grayscale', 'transparency', 'line', 'lineWidth', 'shadow', 'geometry'];
const LOOK_OFF = { crop: '0,0,0,0', rotation: '0', flipH: 'false', flipV: 'false', brightness: '0', contrast: '0', grayscale: 'false', transparency: '0', line: 'none', shadow: 'false', geometry: 'rect' };

/** The props that take a picture from look a to look b; a prop b leaves out goes back to its neutral value. */
export function lookDiff(a, b) {
  const x = lookFrom(a), y = lookFrom(b), p = {};
  for (const k of LOOK) {
    if ((x[k] || '') === (y[k] || '')) continue;
    const v = y[k] != null ? y[k] : LOOK_OFF[k];
    if (v != null) p[k] = v;
  }
  return p;
}

/** Runs one picture command in turn with the file's saves. `path` may be a function, read when the command's turn comes (a
 *  save before it renumbers Word pictures). Resolves with { path, props }: the picture's props afterwards, as `get` prints them;
 *  a compress also brings `before`, the stored size it started from, so the editor can say what it saved. */
export function setPicture(file, path, props) {
  return inLane(file, async () => {
    const at = typeof path === 'function' ? path() : path;
    if (!at) throw new EngineError(_t('这张图片还没有存进文件，稍等片刻再试'), 'NOT_SAVED', _t('新插入的图片会在自动保存后可用'));
    const before = props.compress ? Number(((await run(['get', file, at])).props || {}).bytes) || 0 : undefined;
    const r = await run(['set', file, at, ...propsArgs(props)]);
    return Object.assign({ path: at, props: (r && r.props) || {} }, before != null ? { before } : {});
  });
}

/** Takes a picture command's answer (the picture's props) into the doc's saved snapshot, since the file has it now. Returns what the
 *  command changed, for the editor's model — which may hold edits not saved yet, so it takes the change, not the result:
 *  look: the look props that changed (null: gone, see mergeLook); docx fw, fh: the frame's size now, px;
 *  pptx and xlsx dx, dy, dw, dh: how the frame moved and grew (slide units, sheet px); pptx drot: how far it turned. */
export function pictureSaved(doc, path, props) {
  const after = lookFrom(props), o = doc._orig || {};
  const changed = before => { const c = {}; for (const k of LOOK) if (((before || {})[k] || '') !== (after[k] || '')) c[k] = after[k] != null ? after[k] : null; return c; };
  if (doc.type === 'docx') {
    const b = (o.blocks || []).flatMap(x => [x, ...(x.pics || [])]).find(x => x.path === path), look = changed(b && lookFrom(b.props));
    if (b) b.props = Object.assign({ src: b.props.src }, after);
    return { look, fw: props.width ? cmOf(props.width) / 2.54 * 96 : 0, fh: props.height ? cmOf(props.height) / 2.54 * 96 : 0 };
  }
  const moved = (snap, box) => ({ dx: snap ? box.x - snap.x : 0, dy: snap ? box.y - snap.y : 0, dw: snap ? box.w - snap.w : 0, dh: snap ? box.h - snap.h : 0 });
  if (doc.type === 'pptx') {
    const g = o.geo || { kx: SW / 33.867, ky: 900 / 19.05 }, rot = Number(after.rotation) || 0, own = Object.assign({}, after);
    delete own.rotation;
    const box = { x: Math.round(cmOf(props.x) * g.kx), y: Math.round(cmOf(props.y) * g.ky), w: Math.round(cmOf(props.w) * g.kx), h: Math.round(cmOf(props.h) * g.ky) };
    const snap = (o.slides || []).flatMap(s => s.objs).find(x => x.path === path);
    const change = Object.assign({ look: changed(snap && Object.assign({}, snap.look, snap.rot ? { rotation: String(snap.rot) } : {})), drot: rot - ((snap && snap.rot) || 0) }, moved(snap, box));
    delete change.look.rotation;
    if (snap) Object.assign(snap, box, { look: own, rot });
    return change;
  }
  const box = { x: pxOfCm(props.x), y: pxOfCm(props.y), w: pxOfCm(props.w), h: pxOfCm(props.h) };
  const snap = (o.sheets || []).flatMap(s => s.images || []).find(im => im.path === path);
  const change = Object.assign({ look: changed(snap && snap.look) }, moved(snap, box));
  if (snap) Object.assign(snap, box, { look: after });
  return change;
}

/** A look with changes from pictureSaved put in (null removes a prop). */
export function mergeLook(look, changes) {
  const l = Object.assign({}, look);
  for (const [k, v] of Object.entries(changes || {})) { if (v == null) delete l[k]; else l[k] = v; }
  return l;
}

const WORD_PT = 96 / 72;
/** A Word picture in the editor: a bare img, or — once it has a look — an img in a figure that frames and clips it. The look rides
 *  on the img as data-w-* attributes; data-fw and data-fh keep the frame's size in px. */
export function picHtml(b, inPara) {
  const look = lookFrom(b.props), keys = Object.keys(look), w = Math.round(b.width || 0), h = Math.round(b.height || b.width || 0), place = b.place || {};
  const attrs = ` data-path="${esc(b.path)}"` + (w ? ` data-fw="${w}" data-fh="${h}"` : '') + keys.map(k => ` data-w-${k.toLowerCase()}="${esc(look[k])}"`).join('')
    + PLACE.map(k => place[k] != null ? ` data-w-${k.toLowerCase()}="${esc(place[k])}"` : '').join('');
  const src = esc(picSrc(b.props.src)), placed = inPara ? placeCss(placeStyle(place, w || 300, h || w || 300, null)) : ''; // the editor lays it out again with the page's geometry
  if (!keys.length) return `<img${attrs}${placed ? ' data-placed="1"' : ''} src="${src}" style="${inPara ? 'max-width:100%;vertical-align:middle' : 'max-width:100%;display:block;margin:8px auto'}${w ? ';width:' + w + 'px' : ''}${placed}">`;
  const fw = w || 300, fh = h || fw, v = pictureView(look, fw, fh, WORD_PT);
  return `<figure data-pic="1"${placed ? ' data-placed="1"' : ''} style="${figureStyle(look, fw, fh, v, inPara)}${placed}"><img${attrs} src="${src}" style="${v.image}"></figure>`;
}
const placeCss = s => Object.entries(s).filter(([, v]) => v !== '').map(([k, v]) => ';' + k.replace(/[A-Z]/g, c => '-' + c.toLowerCase()) + ':' + v).join('');
/** A figure's own style: a block of its own, centred, or in a paragraph's line of text (a floating one is placed on top of this). */
const figureStyle = (look, w, h, v, inPara) => `position:relative;${inPara ? 'display:inline-block;vertical-align:middle;margin:0 2px' : 'display:block;margin:8px auto'};width:${Math.round(w)}px;max-width:100%;aspect-ratio:${Math.round(w)} / ${Math.round(h)};transform:rotate(${Number(look.rotation) || 0}deg);${v.frame}`;
const IN_PARA = /^(P|H[1-6]|LI|BLOCKQUOTE)$/;

/** Draws a look on a live Word picture, wrapping it in its figure the first time it needs one (a paragraph that held only the
 *  picture becomes the figure, unless the picture floats in it). fw, fh: a new frame size in px, when it changed. A preview draws
 *  without recording the look in the data-w-* attributes, which a save compares with the file. Returns the element that frames the picture. */
export function paintPic(img, look, fw, fh, preview) {
  if (!preview) for (const k of LOOK) { const a = 'data-w-' + k.toLowerCase(); if (look[k] != null) img.setAttribute(a, look[k]); else img.removeAttribute(a); }
  if (fw) { img.setAttribute('data-fw', String(Math.round(fw))); img.setAttribute('data-fh', String(Math.round(fh || fw))); }
  const w = +img.getAttribute('data-fw') || img.offsetWidth || 300, h = +img.getAttribute('data-fh') || img.offsetHeight || w;
  let fig = img.parentElement && img.parentElement.matches('figure[data-pic]') ? img.parentElement : null;
  if (!fig) {
    if (!Object.keys(look).length) { if (fw) img.style.width = Math.round(w) + 'px'; return img; }
    fig = img.ownerDocument.createElement('figure'); fig.setAttribute('data-pic', '1');
    const p = img.parentElement, alone = p && p.tagName === 'P' && !p.textContent.trim() && p.querySelectorAll('img').length === 1 && (img.getAttribute('data-w-wrap') || 'inline') === 'inline';
    (alone ? p : img).replaceWith(fig); fig.appendChild(img);
  }
  const v = pictureView(look, w, h, WORD_PT);
  fig.style.cssText = figureStyle(look, w, h, v, !!fig.parentElement && IN_PARA.test(fig.parentElement.tagName));
  img.style.cssText = v.image;
  return fig;
}
