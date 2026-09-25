// engine.js — the bridge between the editors and the writer engine serving this page.
// Open: the engine's JSON tree becomes an editor model, every block remembering its path.
// Save: the model is diffed against what was opened and the difference becomes writer commands.
// The file on disk is the only source of truth; the engine keeps everything the editor does not model.
import { txt, shape as mkShape, SW, slideH } from './office-io.js';
import { lookFrom, pictureView, picSrc } from './picture.js';
import * as PK from './pdf-kit.js';
import * as MM from './mindmap.js';

export const state = { workspace: '', chat: false, model: '', version: '', drafts: '' };
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
const FORMATS = { docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', md: 'md', pdf: 'pdf' };
FORMATS.mm = 'mm';

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
  return r.files.filter(f => FORMATS[f.format]).map(f => lazy(f.path, f.format));
}

export function lazy(path, format) {
  return { id: idOf(path), type: format || FORMATS[path.split('.').pop().toLowerCase()] || 'md', title: titleOf(path), path, loaded: false };
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
  return lazy(path, type);
}

export async function upload(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  const path = await freeName(file.name.replace(/\.[^.]+$/, ''), ext);
  await http('/file?file=' + enc(path), { method: 'PUT', body: file });
  return lazy(path, FORMATS[ext]);
}

export async function copy(doc) {
  const ext = doc.path.split('.').pop();
  const path = await freeName(doc.title + _t(' 副本'), ext, newDir());
  const bytes = await (await http(fileUrl(doc.path))).arrayBuffer();
  await http('/file?file=' + enc(path), { method: 'PUT', body: bytes });
  return lazy(path, doc.type);
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
  const t = doc.type;
  let model;
  if (t === 'docx') model = await openDocx(doc);
  else if (t === 'xlsx') model = await openXlsx(doc);
  else if (t === 'pptx') model = await openPptx(doc);
  else if (t === 'pdf') model = await openPdf(doc);
  else if (t === 'mm') model = await openMm(doc);
  else model = await openMd(doc);
  const st = await stat(doc.path).catch(() => null); // best-effort: a missing stat just turns off the external-change check for this doc
  return Object.assign({}, doc, model, { loaded: true, dirty: false, _mtime: st ? st.mtime : null });
}

async function openMd(doc) {
  const text = await (await http(fileUrl(doc.path))).text();
  return { text, _orig: text };
}

async function openPdf(doc) {
  const bytes = new Uint8Array(await (await http(fileUrl(doc.path))).arrayBuffer());
  PK.store[doc.id] = bytes;
  const pdf = await PK.openPdf(doc.id);
  return { pages: Array.from({ length: pdf.numPages }, (_, i) => ({ id: 'p' + i + '_' + doc.id, src: i, rot: 0 })), annots: [] };
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
  if (/^ThreeLineTable$/i.test(style)) return 'three';
  if (/^PlainTable1$/i.test(style)) return 'horizontal';
  return 'grid';
}
const LINE = '1px solid #C7C7CC', GUIDE = '1px dashed #E5E5EA', RULE = '1.5px solid #1D1D1F', HAIR = '1px solid #1D1D1F';
/** Border CSS of a cell at grid row r, place x ({ c, cs, rs }) in an R×C table. Lines Word would not draw show as faint guides. */
export function cellLines(look, r, x, R, C) {
  const top = r === 0, bottom = r + x.rs >= R, left = x.c === 0, right = x.c + x.cs >= C;
  const edge = (outer, isOuter) => look === 'grid' ? LINE : look === 'outside' ? (isOuter ? LINE : GUIDE) : look === 'inside' ? (isOuter ? GUIDE : LINE) : outer;
  const t = look === 'three' ? (top ? RULE : GUIDE) : look === 'horizontal' ? LINE : edge(GUIDE, top);
  const b = look === 'three' ? (bottom ? RULE : r === 0 ? HAIR : GUIDE) : look === 'horizontal' ? LINE : edge(GUIDE, bottom);
  const l = look === 'three' || look === 'horizontal' ? GUIDE : edge(GUIDE, left), rt = look === 'three' || look === 'horizontal' ? GUIDE : edge(GUIDE, right);
  return `border-top:${t};border-right:${rt};border-bottom:${b};border-left:${l}`;
}
const cellStyle = (p, look, r, x, R, C) => `padding:6px 8px;${cellLines(look, r, x, R, C)}${p.fill && p.fill !== 'none' ? ';background:#' + esc(p.fill) : ''}${p.valign && p.valign !== 'top' ? ';vertical-align:' + esc(p.valign) : ''}${p.align && p.align !== 'left' ? ';text-align:' + esc(p.align) : ''}`;

const TOC_STYLE = "border:1px solid #E5E5EA;border-radius:6px;padding:14px 18px;margin:12px 0;font-family:'Noto Sans SC',sans-serif;font-size:14px;line-height:1.9;background:#FFFFFF;color:#1D1D1F";
/** A table of contents in the editor: read-only, its entries indented by level, page numbers when the file has them. */
export function tocHtml({ path, levels, title, entries }) {
  const head = title ? `<div style="font-weight:600;margin-bottom:4px">${esc(title)}</div>` : '';
  const body = entries.length ? entries.map(e => `<div style="display:flex;gap:12px;padding-left:${(Math.max(1, e.level) - 1) * 18}px"><span style="flex:1">${e.href ? `<a href="#${esc(e.href)}">${esc(e.text)}</a>` : esc(e.text)}</span>${e.page ? `<span style="color:#8E8E93">${esc(e.page)}</span>` : ''}</div>`).join('')
    : `<div style="color:#8E8E93">${_t('添加标题后，目录会在保存时生成')}</div>`;
  return `<nav data-toc="1"${path ? ` data-path="${esc(path)}"` : ''} data-levels="${esc(levels || '3')}" data-title="${esc(title || '')}" contenteditable="false" style="${TOC_STYLE}">${head}${body}</nav>`;
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
    const tmp = parseHtml(tocHtml({ path: nav.getAttribute('data-path'), levels, title, entries: editorHeadings(root, levels) })).firstChild;
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
  drawTableLines(table);
  return r.m;
}
/** Draws the live table's lines from its style and borders attributes, after a command or a new look. */
export function drawTableLines(table) {
  const m = tableModelOf(table), look = tableLook({ style: table.getAttribute('data-w-style'), borders: table.getAttribute('data-w-borders') });
  const R = m.rows.length, C = Math.max(1, ...coverage(m).map(r => r.length));
  m.rows.forEach((row, r) => row.cells.forEach(x => { x.ref.style.border = ''; x.ref.style.cssText += ';' + cellLines(look, r, x, R, C); }));
}

/** Cell contents after a merge: both, one under the other, as the engine joins their paragraphs (empty ones dropped). */
const blankHtml = h => !/<img/i.test(h || '') && !String(h || '').replace(/<[^>]*>/g, '').replace(/&nbsp;| |​/g, '').trim();
export const joinCells = (a, b) => [a, b].filter(h => !blankHtml(h)).join('<br>') || '<br>';

/** A page break inside a paragraph: the engine writes it as Word's HTML does; the editor draws it as a line across the page that
 *  cannot be typed into (a span, which a paragraph can hold), and sends it back as the engine's. */
const PB_WORD = '<br style="page-break-before:always">', PB_LINE = '<span data-pb="1" contenteditable="false"></span>';
export const pbIn = html => String(html).split(PB_WORD).join(PB_LINE);
export const pbOut = html => String(html).replace(/<span data-pb="1"[^>]*><\/span>/g, PB_WORD);
/** A heading or paragraph's own props (kept for the save), and how it shows, its own or its style's: a page break above it, a fill. */
const paraOf = (b, p, n) => { const c = n.computed || {}; for (const k of PARA_OWN) if (p[k]) b.props[k] = p[k]; b.pbb = (p.pageBreakBefore || c.pageBreakBefore) === 'true'; b.shade = p.fill || c.fill; };

export function blocksOf(nodes, file) {
  return (nodes || []).map(n => {
    const p = n.props || {}, b = { kind: n.kind, path: n.path, props: {} };
    if (n.kind === 'heading') { b.props.html = p.html || esc(p.text); b.props.level = p.level || '1'; if (p.align) b.props.align = p.align; paraOf(b, p, n); }
    else if (n.kind === 'paragraph') { b.props.html = p.html || esc(p.text); if (p.list && p.list !== 'none') { b.props.list = p.list; b.props.level = p.level || '0'; } if (p.align) b.props.align = p.align; if (p.style) b.props.style = p.style; paraOf(b, p, n); }
    else if (n.kind === 'code') b.props.text = p.text || '';
    else if (n.kind === 'table') {
      // the tree gives json props (widths) parsed; the editor keeps them as the JSON text the engine takes back
      b.props = Object.fromEntries(['style', 'borders', 'borderColor', 'width', 'widths', 'align'].filter(k => p[k] != null).map(k => [k, typeof p[k] === 'object' ? JSON.stringify(p[k]) : p[k]]));
      b.rows = (n.children || []).filter(r => r.kind === 'row').map(r => ({ kind: 'row', path: r.path,
        props: Object.fromEntries(['header', 'height'].filter(k => r.props?.[k] != null).map(k => [k, r.props[k]])),
        cells: (r.children || []).filter(c => c.kind === 'cell').map(c => ({ kind: 'cell', path: c.path,
          props: Object.assign({ html: c.props.html != null ? c.props.html : esc(c.props.text) },
            Object.fromEntries(['fill', 'colspan', 'rowspan', 'borders', 'valign', 'width', 'align'].filter(k => c.props?.[k] != null).map(k => [k, c.props[k]]))) })) }));
    }
    else if (n.kind === 'toc') b.props = { levels: p.levels || '3', title: p.title || '', text: p.text || '' };
    else if (n.kind === 'image') { b.props = Object.assign({ src: binaryUrl(file, n.path) }, lookFrom(p)); b.width = p.width ? cmOf(p.width) / 2.54 * 96 : 0; b.height = p.height ? cmOf(p.height) / 2.54 * 96 : 0; }
    else if (n.kind === 'pagebreak') { }
    else b.props.html = esc(p.text || '');
    return b;
  });
}

export function blocksToHtml(blocks) {
  let out = ''; const stack = []; // open lists: {level, type}
  const closeLists = n => { while (stack.length > n) { out += '</' + stack.pop().type + '>'; } };
  const pa = (b, tag, extra) => {
    const css = [b.props.align && b.props.align !== 'left' ? 'text-align:' + b.props.align : '', b.shade ? 'background:#' + esc(b.shade) : ''].filter(Boolean).join(';');
    return `<${tag} data-path="${esc(b.path)}"${extra || ''}${attrs(b.props, PARA_OWN)}${b.pbb ? ' data-pb="before"' : ''}${css ? ` style="${css}"` : ''}>${pbIn(b.props.html || '<br>')}</${tag}>`;
  };
  const attrs = (p, keys) => keys.map(k => p?.[k] != null ? ` data-w-${k.toLowerCase()}="${esc(p[k])}"` : '').join('');
  for (const b of blocks) {
    if (b.kind === 'paragraph' && b.props.list) {
      const level = +b.props.level || 0, type = b.props.list === 'number' ? 'ol' : 'ul';
      while (stack.length > level + 1) out += '</' + stack.pop().type + '>';
      if (stack.length === level + 1 && stack[level].type !== type) { out += '</' + stack.pop().type + '>'; }
      while (stack.length < level + 1) { const t = stack.length === level ? type : 'ul'; out += '<' + t + '>'; stack.push({ type: t }); }
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
      out += `<table data-path="${esc(b.path)}"${attrs(b.props, ['style', 'borders', 'borderColor', 'width', 'widths', 'align'])} style="border-collapse:collapse;width:100%;margin:8px 0"><tbody>` + m.rows.map((row, r) => `<tr data-path="${esc(row.ref.path)}"${attrs(row.ref.props, ['header', 'height'])}>` + row.cells.map(x => { const c = x.ref;
        return `<td data-path="${esc(c.path)}"${attrs(c.props, ['fill', 'borders', 'valign', 'width', 'align'])}${x.cs > 1 ? ` colspan="${x.cs}"` : ''}${x.rs > 1 ? ` rowspan="${x.rs}"` : ''} style="${cellStyle(c.props, look, r, x, R, C)}">${pbIn(c.props.html || '<br>')}</td>`; }).join('') + '</tr>').join('') + '</tbody></table>';
    }
    else if (b.kind === 'toc') {
      const oneLine = h => plainOf(h).replace(/\s+/g, ' ').trim(), levels = new Map(blocks.filter(h => h.kind === 'heading').map(h => [oneLine(h.props.html), +h.props.level || 1]));
      const entries = (b.props.text || '').split('\n').filter(Boolean).map(line => { const tab = line.lastIndexOf('\t'); const text = tab < 0 ? line : line.slice(0, tab); return { text, page: tab < 0 ? '' : line.slice(tab + 1), level: levels.get(text.replace(/\s+/g, ' ').trim()) || 1 }; });
      out += tocHtml({ path: b.path, levels: b.props.levels, title: b.props.title, entries });
    }
    else if (b.kind === 'image') out += picHtml(b);
    else if (b.kind === 'pagebreak') out += `<hr data-pb="1" data-path="${esc(b.path)}">`;
    else out += pa(b, 'p');
  }
  closeLists(0);
  return out || '<p><br></p>';
}

// ----- docx: page setup, header and footer (document props) -----
const PAPERS = ['A4', 'Letter', 'A5'], MARGINS = ['narrow', 'normal', 'wide'];
/** The editor's page model from the document props; sizes and margins the editor cannot show fall back but stay in `raw`. */
export function pageOf(p) {
  p = p || {};
  return { size: PAPERS.includes(p.page) ? p.page : 'A4', orient: p.orientation === 'landscape' ? 'landscape' : 'portrait', margin: p.margin === 'moderate' ? 'normal' : MARGINS.includes(p.margin) ? p.margin : 'normal', cols: Math.max(1, +p.columns || 1), raw: { size: p.page || '', margin: p.margin || '' } };
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
  const html = inkFills(parseHtml(anchorComments(trackHtml(blocksToHtml(blocks)), comments))).innerHTML;
  return { html, rev: (doc.rev || 0) + 1, track, comments: comments.map(c => ({ id: c.cid, author: c.author, initials: c.initials, mine: c.mine, time: c.time, text: c.text, quote: c.quote, path: c.path, resolved: c.resolved })),
    page, ...hfOf(p), _orig: { blocks, page: { page: p.page, orientation: p.orientation, margin: p.margin, columns: p.columns }, ...hfOf(p), track, comments } };
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
  const walk = list => { for (const n of list || []) { if (n.kind === 'comment') { const q = n.props || {}; out.push({ cid: String(q.id), id: String(q.id), path: n.path.replace(/\/comment\[[^\]]*\]$/, ''), author: q.author || '', initials: q.initials || '', mine: !!me && q.author === me, date: q.date || '', time: [q.author, fmtTime(q.date)].filter(Boolean).join(' · '), text: q.text || '', quote: q.quote || '', resolved: q.resolved === 'true' }); } walk(n.children); } };
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
    const span = el.querySelector(`[data-cid="${c.id}"]`);
    const b = span && (blocks.find(x => (x.kind === 'paragraph' || x.kind === 'heading') && x.el && x.el.contains(span)) || cells.find(x => x.el && x.el.contains(span)));
    const at = path => path ? (b.kind === 'cell' ? path + '/paragraph[1]' : path) : null;
    const q = span && span.cloneNode(true); if (q) q.querySelectorAll('del').forEach(x => x.remove());
    return { cid: String(c.id), text: c.text || '', resolved: !!c.resolved, parent: b ? at(b.path) : null, origin: b ? at(was.get(b)) : null,
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
    const props = Object.assign({ text: c.text }, keep || {}, c.quote ? { quote: c.quote } : {}, c.resolved ? { resolved: 'true' } : {});
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
      if (tag === 'UL' || tag === 'OL') { walk(c, tag === 'OL' ? 'number' : 'bullet', level + 1); continue; }
      if (tag === 'LI') {
        const inner = c.cloneNode(true); Array.from(inner.querySelectorAll('ul,ol')).forEach(x => x.remove());
        out.push({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(inner), list: listType || 'bullet', level: String(Math.max(0, level)) }, paraAttrs(c)), el: c, align: alignOf(c) });
        Array.from(c.children).filter(x => /^(UL|OL)$/.test(x.tagName)).forEach(x => walk(x, x.tagName === 'OL' ? 'number' : 'bullet', level + 1));
        continue;
      }
      if (/^H[1-6]$/.test(tag)) {
        const st = c.getAttribute('data-style');
        if (st) out.push({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), style: st }, paraAttrs(c)), el: c, align: alignOf(c) });
        else out.push({ kind: 'heading', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), level: c.getAttribute('data-level') || tag[1] }, paraAttrs(c)), el: c, align: alignOf(c) });
        continue;
      }
      if (tag === 'PRE') { out.push({ kind: 'code', path: pathOf(c), props: { text: c.innerText.replace(/\n$/, '') }, el: c }); continue; }
      if (tag === 'BLOCKQUOTE') { out.push({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c), style: c.getAttribute('data-style') || 'Quote' }, paraAttrs(c)), el: c, align: alignOf(c) }); continue; }
      if (c.hasAttribute('data-toc')) { out.push({ kind: 'toc', path: pathOf(c), props: { levels: c.getAttribute('data-levels') ?? '3', title: c.getAttribute('data-title') ?? '目录' }, refresh: c.hasAttribute('data-refresh'), el: c }); continue; }
      if (tag === 'TABLE') {
        let ops = []; try { ops = JSON.parse(c.getAttribute('data-ops') || '[]'); } catch (e) { ops = null; } // unreadable: the save writes the table anew
        out.push({ kind: 'table', path: pathOf(c), el: c, ops, props: docxAttrs(c, ['style', 'borders', 'borderColor', 'width', 'widths', 'align']),
          rows: Array.from(c.querySelectorAll('tr')).filter(tr => tr.closest('table') === c).map(tr => ({ kind: 'row', path: pathOf(tr), el: tr,
            props: docxAttrs(tr, ['header', 'height']), cells: Array.from(tr.children).filter(td => /^(TD|TH)$/.test(td.tagName)).map(td => ({ kind: 'cell', path: pathOf(td), el: td,
              props: Object.assign({ html: inlineHtml(td) }, docxAttrs(td, ['fill', 'borders', 'valign', 'width', 'align']), td.colSpan > 1 ? { colspan: String(td.colSpan) } : {}, td.rowSpan > 1 ? { rowspan: String(td.rowSpan) } : {}) })) })) });
        continue;
      }
      if (tag === 'IMG') { out.push(imgBlock(c)); continue; }
      if (tag === 'HR') { if (c.getAttribute('data-pb')) out.push({ kind: 'pagebreak', path: pathOf(c), props: {}, el: c }); continue; }
      if (tag === 'BR' && node === el) continue;
      if (tag === 'P' || tag === 'DIV' || tag === 'FIGURE' || tag === 'SECTION' || tag === 'ARTICLE') {
        if (Array.from(c.children).some(x => BLOCK.test(x.tagName) && x.tagName !== 'IMG' && x.tagName !== 'BR')) { walk(c, listType, level); continue; }
        const img = c.querySelector('img');
        if (img && !c.textContent.trim()) { out.push(imgBlock(img)); continue; }
        const st = c.getAttribute('data-style');
        out.push({ kind: 'paragraph', path: pathOf(c), props: Object.assign({ html: inlineHtml(c) }, st ? { style: st } : {}, paraAttrs(c)), el: c, align: alignOf(c) });
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
/** A heading or paragraph's own props that its element keeps as data-w-* (the file's; not what its style gives), and the value
 *  that turns each off. */
const PARA_OWN = ['pageBreakBefore', 'fill'], PARA_OFF = { pageBreakBefore: 'false', fill: 'none' };
const paraAttrs = el => docxAttrs(el, PARA_OWN);
const imgBlock = img => ({ kind: 'image', path: pathOf(img), props: Object.assign({ src: img.getAttribute('src') || '' }, docxAttrs(img, LOOK)), el: img, width: +img.getAttribute('data-fw') || (img.style.width ? parseFloat(img.style.width) : 0) });
const alignOf = el => { const a = el.style && el.style.textAlign; return a && a !== 'start' && a !== 'left' ? (a === 'end' ? 'right' : a) : null; };
/** A block's inline html for the engine: without the AI change marks, comment anchor spans (the file keeps anchors itself) and the
 *  editor's zero-width fillers, and with its page break lines as Word's page breaks. */
function inlineHtml(el) {
  const c = el.cloneNode(true);
  Array.from(c.querySelectorAll('[data-ai]')).forEach(x => x.removeAttribute('data-ai'));
  Array.from(c.querySelectorAll('[data-cid]')).forEach(x => { while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x); x.remove(); });
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
    if (fill) el.setAttribute('data-ink', inkOn(fill)); else el.removeAttribute('data-ink');
  }
  return root;
}
function ptOf(v) { const m = /^([\d.]+)\s*(pt|px)$/i.exec(String(v || '').trim()); if (!m) return null; const n = m[2].toLowerCase() === 'px' ? +m[1] * 0.75 : +m[1]; return String(Math.round(n * 2) / 2); }
export function runsOf(html) {
  const root = typeof html === 'string' ? parseHtml(html) : html, out = [];
  const add = (t, s) => { if (!t) return; const last = out[out.length - 1]; if (last && last.s === s) last.t += t; else out.push({ t, s }); };
  const key = f => JSON.stringify([f.b, f.i, f.u, f.s, f.c, f.a, f.color, f.bg, f.size, f.font, f.ins, f.del].map(x => x || 0));
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
      if (tag === 'U' || (st.textDecoration || '').includes('underline')) g.u = 1;
      if (tag === 'INS') g.ins = 1;
      if (tag === 'DEL') g.del = 1;
      if (tag === 'S' || tag === 'STRIKE' || (st.textDecoration || '').includes('line-through')) g.s = 1;
      if (tag === 'CODE' || tag === 'TT' || tag === 'KBD') g.c = 1;
      if (tag === 'A' && c.getAttribute('href')) g.a = c.getAttribute('href');
      if (tag === 'MARK') g.bg = 'FFFF00';
      if (tag === 'FONT') { if (c.getAttribute('color')) g.color = colorHex(c.getAttribute('color')); if (c.getAttribute('face')) g.font = c.getAttribute('face').split(',')[0].trim().replace(/["']/g, ''); }
      if (st.color) g.color = colorHex(st.color) || g.color;
      if (st.backgroundColor) g.bg = colorHex(st.backgroundColor) || g.bg;
      if (st.fontSize) g.size = ptOf(st.fontSize) || g.size;
      if (st.fontFamily) g.font = st.fontFamily.split(',')[0].trim().replace(/["']/g, '');
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

/** Positions survivors again after removals: /body/paragraph[3] becomes /body/paragraph[2] when an earlier paragraph went away. */
function reindex(parentPath, survivors) {
  const counts = {};
  for (const o of survivors) {
    const k = o.kind, s = seg(o.path), m = /\[([^\]]*)\]$/.exec(s), keyed = m && !/^-?\d+$/.test(m[1]);
    counts[k] = (counts[k] || 0) + 1;
    o.cur = keyed ? parentPath + '/' + s : parentPath + '/' + k + '[' + counts[k] + ']';
    if (o.rows) o.rows.forEach(r => { r.cur = o.cur + '/' + seg(r.path); (r.cells || []).forEach(c => c.cur = r.cur + '/' + seg(c.path)); });
    if (o.cells) o.cells.forEach(c => c.cur = o.cur + '/' + seg(c.path));
  }
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
  else if (b.kind === 'paragraph') { p.html = b.props.html; p.list = b.props.list || (forNew ? null : 'none'); if (b.props.list) p.level = b.props.level || '0'; if (b.props.style) p.style = b.props.style; if (b.align || !forNew) p.align = b.align || 'left'; }
  if (b.kind === 'heading' || b.kind === 'paragraph') for (const k of PARA_OWN) if (b.props[k]) p[k] = b.props[k];
  else if (b.kind === 'code') p.text = b.props.text;
  else if (b.kind === 'image') { p.src = b.props.src; if (b.width) p.width = Math.round(b.width) + 'px'; }
  else if (b.kind === 'table') { p.data = JSON.stringify(tableData(b.rows)); Object.assign(p, b.props); }
  else if (b.kind === 'row') { p.data = JSON.stringify(b.cells.map(c => textOf(c.props.html))); Object.assign(p, b.props); }
  else if (b.kind === 'cell') Object.assign(p, b.props);
  else if (b.kind === 'toc') { p.levels = b.props.levels || '3'; p.title = b.props.title || ''; }
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
  if (b.kind === 'heading' || b.kind === 'paragraph') for (const k of PARA_OWN) if ((orig.props[k] || '') !== (b.props[k] || '')) p[k] = b.props[k] || PARA_OFF[k];
  if (b.kind === 'paragraph') {
    const ol = orig.props.list || 'none', nl = b.props.list || 'none';
    if (ol !== nl) p.list = nl;
    if (nl !== 'none' && String(orig.props.level || '0') !== String(b.props.level || '0')) p.level = b.props.level || '0';
    if ((orig.props.style || '') !== (b.props.style || '') && b.props.style) p.style = b.props.style;
  }
  if (b.kind === 'toc') {
    if (String(orig.props.levels || '3') !== String(b.props.levels || '3')) p.levels = b.props.levels || '3';
    if ((orig.props.title || '') !== (b.props.title || '')) p.title = b.props.title || '';
  }
  if (b.kind === 'table' || b.kind === 'row' || b.kind === 'cell') {
    const keys = b.kind === 'table' ? ['style', 'borders', 'borderColor', 'width', 'widths', 'align'] : b.kind === 'row' ? ['header', 'height'] : ['fill', 'colspan', 'rowspan', 'borders', 'valign', 'width', 'align'];
    for (const k of keys) {
      const was = String(orig.props?.[k] ?? ''), now = String(b.props?.[k] ?? '');
      if (was === now) continue;
      if (now) p[k] = now;
      else if (k === 'borders') { if (b.kind === 'table') p[k] = 'style'; } // a cell has no value that brings the table's lines back
      else if (k in CLEARED) p[k] = CLEARED[k]; // width, align, height…: nothing to write that removes them, so they stay
    }
    if (orig.fresh && b.kind === 'cell' && !p.fill) p.fill = b.props?.fill || 'none'; // a new cell copied some row or cell in the engine: say which shading it has
  }
  if (b.kind !== 'code' && b.kind !== 'table' && b.kind !== 'image' && b.kind !== 'row' && b.kind !== 'cell' && b.kind !== 'toc' && b.kind !== 'pagebreak' && (orig.align || orig.props.align || 'left') !== (b.align || 'left')) p.align = b.align || 'left';
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
  const survivors = origChildren.filter(o => seen.has(o.path));
  reindex(parentPath, survivors);
  let anchor = null;
  for (let i = newChildren.length - 1; i >= 0; i--) {
    const b = newChildren[i];
    if (b.path) {
      const o = byPath.get(b.path);
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
      anchor = o.cur;
      continue;
    }
    const argv = ['add', file, parentPath, '--type', b.kind, ...propsArgs(blockProps(b, true))];
    if (anchor) argv.push('--before', anchor);
    const r = await exec(argv); n++; log && log('add', r.path, b.kind);
    b.path = r.path;
    if (b.kind === 'table') for (const argv of newTableCommands(r.path, b.rows, file)) { await exec(argv); n++; }
    if (b.kind === 'row') for (let ci = 0; ci < b.cells.length; ci++) {
      const { html, ...props } = b.cells[ci].props;
      if (/<[a-z]/i.test(html || '') && !/^(<br>|&nbsp;)*$/.test(html || '')) props.html = html;
      if (Object.keys(props).length) { await exec(['set', file, `${r.path}/cell[${ci + 1}]`, ...propsArgs(props)]); n++; }
    }
    anchor = r.path;
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
    counts[b.kind] = (counts[b.kind] || 0) + 1;
    b.path = parentPath + '/' + b.kind + '[' + counts[b.kind] + ']';
    if (b.el) b.el.setAttribute('data-path', b.path);
    if (b.rows) renumber(b.path, b.rows);
    if (b.cells) renumber(b.path, b.cells);
  }
}
const strip = blocks => blocks.map(b => { const c = Object.assign({}, b); delete c.el; delete c.ops; if (c.rows) c.rows = strip(c.rows); if (c.cells) c.cells = strip(c.cells); return c; });

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
  const el = root || parseHtml(doc.html || '');
  const blocks = blocksFromHtml(el).filter(b => !(b.kind === 'image' && !b.path && !/^data:/.test(b.props.src)));
  const opened = new Set((orig.blocks || []).map(o => o.path));
  const tocs = blocks.filter(b => b.kind === 'toc'), tocAdded = tocs.some(b => !b.path || !opened.has(b.path));
  const was = new Map(), remember = list => list.forEach(b => { was.set(b, b.path); if (b.rows) remember(b.rows); if (b.cells) remember(b.cells); });
  remember(blocks);
  n += await planDocxBlocks(doc.path, orig.blocks || [], blocks, run, log);
  renumber('/body', blocks);
  // a table of contents lists the headings: build it again in the file, and show it, once headings or contents changed
  if (tocs.length && (tocAdded || tocs.some(b => b.refresh) || headingsOf(orig.blocks || []) !== headingsOf(blocks))) {
    for (const b of tocs) { await run(['set', doc.path, b.path, '--prop', 'levels=' + (b.props.levels || '3')]); n++; log && log('set', b.path, { levels: b.props.levels }); }
    refreshTocs(el);
  }
  for (const b of tocs) { b.el?.removeAttribute('data-refresh'); b.refresh = false; }
  for (const b of blocks) if (b.kind === 'table') b.el?.removeAttribute('data-ops');
  const comments = await planComments(doc.path, orig.comments || [], commentsIn(doc, el, blocks, was), log);
  n += comments.count;
  const pageProps = Object.fromEntries(['page', 'orientation', 'margin', 'columns'].filter(k => k in pp).map(k => [k, pp[k]]));
  const now = Object.assign({}, orig, { titlePg: String(!!orig.titlePg) }, pp); // headers and footers as the file has them now
  doc._orig = Object.assign({ blocks: strip(blocks), page: Object.assign({}, orig.page, pageProps) }, hfOf(now), { track: !!doc.track, comments: comments.list });
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
  const bare = code.replace(/\[[^\]]*\]|"[^"]*"|\\./g, ''), dec = (/\.(0+)/.exec(bare) || [, ''])[1].length;
  if (/\*|_\(/.test(code)) return { fmt: 'acct', dec, code };
  if (/%/.test(bare)) return { fmt: 'pct', dec, code };
  if (/\$/.test(code)) return { fmt: 'usd', dec, code };
  if (/[¥€£￥]/.test(code)) return { fmt: 'money', dec, code };
  if (/[#0]/.test(bare)) return { fmt: /,/.test(bare) ? 'number' : 'plain', dec, code };
  if (/[hs]/i.test(bare) && !/[yd]/i.test(bare)) return { fmt: 'time', code };
  if (/[ymd]/i.test(bare)) return { fmt: 'date', code };
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
  const sides = jsonOr(p.borders, null);
  if (sides && typeof sides === 'object' && Object.keys(sides).length) s.bd = sides;
  else if (p.border && p.border !== 'none') s.bd = p.border;
  if (p.borderColor && p.borderColor !== 'none') s.bdc = hex(p.borderColor);
  if (p.link) s.link = p.link;
  if (p.note) s.note = p.note;
  Object.assign(s, fmtOf(p.format));
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
  return { name: p.name, path: sheetPath, cells, colW, rowH, merges, frR: fz ? fz.r : 0, frC: fz ? fz.c : 0, filter: p.filter && p.filter !== 'none' ? p.filter : null, charts, images };
}
/** The snapshot a later save is diffed against. */
const origOf = sheets => ({ sheets: JSON.parse(JSON.stringify(sheets.map(s => ({ path: s.path, name: s.name, cells: s.cells, colW: s.colW || {}, rowH: s.rowH || {}, merges: s.merges || [], frR: s.frR || 0, frC: s.frC || 0, filter: s.filter || null, charts: s.charts || [], images: s.images || [] })))) });
async function openXlsx(doc) {
  const t = await tree(doc.path);
  const sheets = (t.children || []).filter(s => s.kind === 'sheet').map(s => sheetModel(s, doc.path));
  return { active: Math.min(doc.active || 0, sheets.length - 1), sheets, _orig: origOf(sheets) };
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
const GEOM = { roundRect: 'round', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', rightArrow: 'arrow' };
const GEOM_BACK = { round: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', arrow: 'rightArrow', rect: 'rect' };
function paraHtml(paragraphs) {
  let out = '', inList = false;
  for (const p of paragraphs) {
    const h = p.props.html != null ? p.props.html : esc(p.props.text || '');
    if (p.props.list && p.props.list !== 'none') { if (!inList) { out += '<ul>'; inList = true; } out += '<li>' + (h || '<br>') + '</li>'; }
    else { if (inList) { out += '</ul>'; inList = false; } out += '<p>' + (h || '<br>') + '</p>'; }
  }
  if (inList) out += '</ul>';
  return out;
}
async function openPptx(doc) {
  const t = await tree(doc.path);
  const wcm = cmOf(t.props.width) || 33.867, hcm = cmOf(t.props.height) || 19.05;
  const ratio = hcm / wcm > 0.7 ? '4:3' : '16:9', H = slideH(ratio), kx = SW / wcm, ky = H / hcm, ptPx = SW / (wcm / 2.54 * 72);
  const geo = { wcm, hcm, kx, ky, ptPx };
  const slides = (t.children || []).filter(s => s.kind === 'slide').map(s => {
    const sp = '/slide[@id=' + s.props.id + ']';
    const asObject = n => {
      const p = Object.assign({}, n.computed, n.props), decor = n.kind === 'decor'; // computed: what the shape inherits from its theme, layout and master
      const kind = decor ? p.type : n.kind;
      const box = { x: Math.round(cmOf(p.x) * kx), y: Math.round(cmOf(p.y) * ky), w: Math.round(cmOf(p.w) * kx), h: Math.round(cmOf(p.h) * ky) };
      const path = decor ? n.path : sp + '/' + n.kind + '[@id=' + p.id + ']';
      // an id names one object in the whole deck, and a cNvPr id does not: every slide numbers its shapes anew, and master and layout shapes may share one
      const id = (decor ? 'd' : 'e') + path;
      if (kind === 'image') { const look = lookFrom(p); delete look.rotation; return txt(Object.assign({ id, path, kind, t: 'image', src: binaryUrl(doc.path, path), html: '', look, rot: Number(p.rotation) || 0 }, box)); }
      if (n.kind === 'table') return txt(Object.assign({ id, path, kind: 'table', t: 'table', html: '', fs: 24, rows: (n.children || []).map(r => (r.children || []).map(c => c.props.text || '')) }, box));
      const isTitle = p.placeholder === 'title', filled = p.fill && p.fill !== 'none';
      // a connector has no height: draw it as a thin bar in its line colour
      if (decor && p.geometry === 'line') return mkShape(Object.assign({ id, path, kind: 'shape', html: '', fill: hex(p.line), stroke: '', sw: 0, shape: 'rect' }, box, { h: Math.max(box.h, 2) }));
      const base = Object.assign({ id, path, kind: 'shape', html: decor ? (p.html || (p.text ? '<p>' + esc(p.text) + '</p>' : '')) : paraHtml((n.children || []).filter(c => c.kind === 'paragraph')), fs: p.size ? Math.round(cmOf(p.size) / UNIT.pt * ptPx) : (isTitle ? Math.round(40 * ptPx) : Math.round(20 * ptPx)), color: hex(p.color), font: p.font || null, ph: isTitle ? 'title' : p.placeholder === 'subtitle' ? 'sub' : null, bold: p.bold != null ? p.bold === 'true' : isTitle }, box);
      if (filled || (p.geometry && p.geometry !== 'rect' && p.geometry !== 'textbox' && p.geometry !== 'custom')) return mkShape(Object.assign(base, { fill: filled ? hex(p.fill) : null, stroke: p.line && p.line !== 'none' ? hex(p.line) : '', sw: p.line && p.line !== 'none' ? 1 : 0, shape: GEOM[p.geometry] || 'rect', align: 'center', va: 'middle' }));
      return txt(base);
    };
    const children = s.children || [];
    const decor = children.filter(n => n.kind === 'decor').map(asObject);
    const objs = children.filter(n => n.kind !== 'decor').map(asObject);
    return { id: 's' + s.props.id, path: sp, layout: s.props.layout, decor, objs, notes: s.props.notes || '', trans: s.props.transition || 'none', duration: s.props.duration == null ? null : Number(s.props.duration), hidden: s.props.hidden === true || s.props.hidden === 'true', bg: p2bg(s.props.background || (s.computed || {}).background) };
  });
  const orig = { geo, slides: pptxSnapshot(slides) };
  return { theme: doc.theme || 'paper', ratio, slides, _orig: orig };
}
const p2bg = c => c ? '#' + c : '#FFFFFF';
const pptxSnapshot = slides => JSON.parse(JSON.stringify(slides.map(s => ({ id: s.id, path: s.path, bg: s.bg, layout: s.layout, notes: s.notes || '', trans: s.trans || 'none', duration: s.duration ?? null, hidden: !!s.hidden, objs: s.objs.map(objKey) }))));
export function slideProps(o, s) {
  const p = {};
  if ((o.notes || '') !== (s.notes || '')) p.notes = s.notes || '';
  if (!!o.hidden !== !!s.hidden) p.hidden = s.hidden ? 'true' : 'false';
  if ((o.trans || 'none') !== (s.trans || 'none')) p.transition = s.trans || 'none';
  if (s.duration != null && s.trans !== 'none' && s.trans !== 'other' && o.duration !== s.duration) p.duration = String(s.duration);
  return p;
}
function objKey(o) { return Object.assign({ id: o.id, path: o.path, kind: o.kind, t: o.t, x: o.x, y: o.y, w: o.w, h: o.h, html: o.html, fill: o.fill, color: o.color, font: o.font, fs: o.fs, shape: o.shape, rows: o.rows, src: o.src && o.src.startsWith('data:') ? 'data' : o.src }, o.t === 'image' ? { look: o.look || {}, rot: o.rot || 0 } : {}); }
function boxProps(o, g) { return { x: cmStr(o.x / g.kx), y: cmStr(o.y / g.ky), w: cmStr(o.w / g.kx), h: cmStr(o.h / g.ky) }; }
function objProps(o, g, orig) {
  const p = {};
  const nb = boxProps(o, g), ob = orig ? boxProps(orig, g) : {};
  for (const k of ['x', 'y', 'w', 'h']) if (nb[k] !== ob[k]) p[k] = nb[k];
  if (o.t === 'text' || o.t === 'shape') {
    if (!orig || !sameRuns(orig.html, o.html)) p.html = o.html || '';
    if ((orig ? orig.fill : undefined) !== o.fill && (o.t === 'shape' || o.fill)) p.fill = o.fill ? unhex(o.fill) : 'none';
    if ((orig ? orig.color : undefined) !== o.color && o.color) p.color = unhex(o.color);
    if ((orig ? orig.font : undefined) !== o.font && o.font) p.font = o.font;
    if ((orig ? orig.fs : undefined) !== o.fs && o.fs) p.size = (Math.round(o.fs / g.ptPx * 2) / 2) + 'pt';
    if (!orig && o.t === 'shape') p.geometry = GEOM_BACK[o.shape] || 'rect';
  }
  if (o.t === 'table' && (!orig || !same(orig.rows, o.rows))) p.data = JSON.stringify(o.rows || []);
  if (o.t === 'image' && !orig) p.src = o.src;
  if (o.t === 'image') { Object.assign(p, lookDiff(orig && orig.look, o.look)); if (Math.round((orig && orig.rot) || 0) !== Math.round(o.rot || 0)) p.rotation = String(Math.round(o.rot || 0)); }
  return p;
}
async function setListProps(file, shapePath, html) {
  const root = parseHtml(html || ''); const items = [];
  const walk = n => { for (const c of Array.from(n.children)) { if (c.tagName === 'UL' || c.tagName === 'OL') walk(c); else if (BLOCK.test(c.tagName)) items.push(c.tagName === 'LI'); } };
  walk(root);
  if (!items.some(Boolean)) return 0;
  let n = 0;
  for (let i = 0; i < items.length; i++) { await run(['set', file, `${shapePath}/paragraph[${i + 1}]`, '--prop', 'list=' + (items[i] ? 'bullet' : 'none')]); n++; }
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
async function savePptx(doc, log) {
  let n = 0;
  const orig = doc._orig || { geo: null, slides: [] }, g = orig.geo || { kx: SW / 33.867, ky: 900 / 19.05, ptPx: SW / 960 };
  const origSlides = orig.slides, at = path => path && origSlides.find(x => x.path === path);
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
      origSlides.splice(origSlides.indexOf(src) + 1, 0, Object.assign(JSON.parse(JSON.stringify(src)), { id: s.id, path: s.path, objs: src.objs.map(y => Object.assign({}, y, { path: moved(y.path) })) }));
      for (const x of s.objs) if (!x.path && x.from && slideOf(x.from) === src.path && !taken.has(x.from)) { taken.add(x.from); x.path = moved(x.from); if (x.t === 'image') x.src = binaryUrl(doc.path, x.path); }
    } else {
      const r = await run(['add', doc.path, '/', '--type', 'slide', '--prop', 'layout=' + (s.layout || 'Blank')]); n++;
      s.path = '/slide[@id=' + r.props.id + ']'; origSlides.push({ path: s.path, bg: '#FFFFFF', notes: '', trans: 'none', duration: null, hidden: false, objs: [] }); log && log('add', s.path, 'slide');
      // a blank layout may still carry placeholders: clear them so the new slide holds only what the editor shows
      const made = await run(['get', doc.path, s.path, '--depth', '1']);
      for (const c of (made.children || []).slice().reverse()) { if (c.kind === 'decor') continue; await run(['remove', doc.path, c.path]); n++; }
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
      x.path = idPath(s.path, r); x.kind = r.kind;
      if (x.t === 'image') x.src = binaryUrl(doc.path, x.path);
      o.objs.push(Object.assign({}, was || gone.snap, { id: x.id, path: x.path }));
    }
  }
  const keep = new Set(doc.slides.map(s => s.path));
  const order = [];
  for (const o of origSlides) { if (!keep.has(o.path)) { await run(['remove', doc.path, o.path]); n++; log && log('remove', o.path); } else order.push(o.path); }
  for (let i = 0; i < doc.slides.length; i++) {
    const s = doc.slides[i], o = at(s.path);
    if (order.indexOf(s.path) !== i) { await run(['move', doc.path, s.path, '--to', '/', '--index', String(i + 1)]); n++; order.splice(order.indexOf(s.path), 1); order.splice(i, 0, s.path); }
    if ((o.bg || '#FFFFFF') !== (s.bg || '#FFFFFF')) { await run(['set', doc.path, s.path, '--prop', 'background=' + unhex(s.bg || '#FFFFFF')]); n++; }
    const sp = slideProps(o, s);
    if (Object.keys(sp).length) { await run(['set', doc.path, s.path, ...propsArgs(sp)]); n++; log && log('set', s.path, sp); }
    const keepObjs = new Set(s.objs.filter(x => x.path).map(x => x.path));
    for (const oo of o.objs.slice().reverse()) if (!keepObjs.has(oo.path)) {
      removed[oo.path] = { xml: String(await run(['get', doc.path, oo.path, '--raw'])).trim(), snap: oo };
      await run(['remove', doc.path, oo.path]); n++; log && log('remove', oo.path);
    }
    for (const x of s.objs) {
      const oo = x.path ? o.objs.find(y => y.path === x.path) : null;
      const p = objProps(x, g, oo);
      if (oo) {
        if (Object.keys(p).length) { await run(['set', doc.path, x.path, ...propsArgs(p)]); n++; log && log('set', x.path, p); if (p.html != null) n += await setListProps(doc.path, x.path, x.html); }
        continue;
      }
      if (x.t === 'image' && !(x.src || '').startsWith('data:')) continue;
      const kind = x.t === 'image' ? 'image' : x.t === 'table' ? 'table' : 'shape';
      const r = await run(['add', doc.path, s.path, '--type', kind, ...propsArgs(p)]); n++;
      x.path = s.path + '/' + kind + '[@id=' + r.props.id + ']'; x.kind = kind; log && log('add', x.path, kind);
      if (kind === 'shape') n += await setListProps(doc.path, x.path, x.html);
      if (kind === 'image') {
        x.src = binaryUrl(doc.path, x.path); // from now on it loads from its own place in the file
        if (p.crop) { await run(['set', doc.path, x.path, ...propsArgs(boxProps(x, g))]); n++; } // a crop keeps the picture's scale, so it moved the frame
      }
    }
  }
  return n;
}

// ----- mm (FreeMind mind map): the topic tree is the model; ids are the engine's, new nodes carry tmp_ ids until saved -----
async function openMm(doc) {
  let t = await tree(doc.path);
  // topics without an id (FreeMind files, or a map nobody has edited yet) cannot be addressed: one no-op write makes the engine assign ids
  const missing = n => n.kind === 'topic' && !(n.props && n.props.id) || (n.children || []).some(missing);
  const root = (t.children || []).find(c => c.kind === 'topic');
  if (root && missing(root)) { await run(['set', doc.path, '/topic[1]', '--prop', 'text=' + ((root.props && root.props.text) || '')]); t = await tree(doc.path); }
  const conv = n => { const p = n.props || {}, m = { id: p.id, text: p.text || '', children: (n.children || []).filter(c => c.kind === 'topic').map(conv) }; if (p.collapsed === 'true' || p.collapsed === true) m.collapsed = true; for (const k of ['side', 'note', 'link', 'color', 'fill', 'icon']) if (p[k]) m[k] = p[k]; return m; };
  const top = (t.children || []).find(c => c.kind === 'topic');
  const map = top ? conv(top) : { id: 'root', text: doc.title, children: [] };
  return { map, _orig: MM.flatten(map) };
}
/** Runs the planner's steps in order; added nodes get their engine id written into the model and into doc._tmpToId (old id → real id). */
async function saveMm(doc, log) {
  const { steps, ids } = MM.plan(doc._orig, doc.map, doc.path);
  let n = 0;
  for (const s of steps) { const argv = s.argv, r = await run(argv); s.onResult && s.onResult(r); n++; log && log(argv[0], argv[2]); doc._tmpToId = Object.assign({}, doc._tmpToId, ids); }
  if (n) doc._orig = MM.flatten(doc.map);
  return n;
}
function adoptMm(cur, saved) {
  cur._tmpToId = Object.assign({}, cur._tmpToId, saved._tmpToId);
  const m = cur._tmpToId, real = id => { for (let i = 0; m[id] && i < 20; i++) id = m[id]; return id; };
  if (cur.map) MM.walk(cur.map, x => { x.id = real(x.id); });
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
  if (doc.type === 'pdf') return 0;
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
    cur.slides.forEach(s => { const o = bySlide.get(s.id); if (!o) return; if (o.path && s.path !== o.path) s.path = o.path; const byObj = new Map(o.objs.map(x => [x.id, x])); s.objs.forEach(x => { const y = byObj.get(x.id); if (y) { if (y.path && x.path !== y.path) { x.path = y.path; if (x.t === 'image') x.src = y.src; } if (!x.kind) x.kind = y.kind; } }); });
  }
  if (cur.type === 'mm') adoptMm(cur, saved);
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
      if (!same(o.objs.map(objKey), s.objs.map(objKey)) || o.bg !== s.bg || o.notes !== s.notes || o.trans !== s.trans || o.duration !== s.duration || o.hidden !== s.hidden) { s.ai = true; items.push([_t('修改'), _t('第 {n} 页', { n: i + 1 })]); }
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
    const b = (o.blocks || []).find(x => x.path === path), look = changed(b && lookFrom(b.props));
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
export function picHtml(b) {
  const look = lookFrom(b.props), keys = Object.keys(look), w = Math.round(b.width || 0), h = Math.round(b.height || b.width || 0);
  const attrs = ` data-path="${esc(b.path)}"` + (w ? ` data-fw="${w}" data-fh="${h}"` : '') + keys.map(k => ` data-w-${k.toLowerCase()}="${esc(look[k])}"`).join('');
  const src = esc(picSrc(b.props.src));
  if (!keys.length) return `<img${attrs} src="${src}" style="max-width:100%;display:block;margin:8px auto${w ? ';width:' + w + 'px' : ''}">`;
  const fw = w || 300, fh = h || fw, v = pictureView(look, fw, fh, WORD_PT);
  return `<figure data-pic="1" style="${figureStyle(look, fw, fh, v)}"><img${attrs} src="${src}" style="${v.image}"></figure>`;
}
const figureStyle = (look, w, h, v) => `position:relative;display:block;margin:8px auto;width:${Math.round(w)}px;max-width:100%;aspect-ratio:${Math.round(w)} / ${Math.round(h)};transform:rotate(${Number(look.rotation) || 0}deg);${v.frame}`;

/** Draws a look on a live Word picture, wrapping it in its figure the first time it needs one (a paragraph that held only the
 *  picture becomes the figure). fw, fh: a new frame size in px, when it changed. A preview draws without recording the look in
 *  the data-w-* attributes, which a save compares with the file. Returns the element that frames the picture. */
export function paintPic(img, look, fw, fh, preview) {
  if (!preview) for (const k of LOOK) { const a = 'data-w-' + k.toLowerCase(); if (look[k] != null) img.setAttribute(a, look[k]); else img.removeAttribute(a); }
  if (fw) { img.setAttribute('data-fw', String(Math.round(fw))); img.setAttribute('data-fh', String(Math.round(fh || fw))); }
  const w = +img.getAttribute('data-fw') || img.offsetWidth || 300, h = +img.getAttribute('data-fh') || img.offsetHeight || w;
  let fig = img.parentElement && img.parentElement.matches('figure[data-pic]') ? img.parentElement : null;
  if (!fig) {
    if (!Object.keys(look).length) { if (fw) img.style.width = Math.round(w) + 'px'; return img; }
    fig = img.ownerDocument.createElement('figure'); fig.setAttribute('data-pic', '1');
    const p = img.parentElement, alone = p && p.tagName === 'P' && !p.textContent.trim() && p.querySelectorAll('img').length === 1;
    (alone ? p : img).replaceWith(fig); fig.appendChild(img);
  }
  const v = pictureView(look, w, h, WORD_PT);
  fig.style.cssText = figureStyle(look, w, h, v);
  img.style.cssText = v.image;
  return fig;
}
