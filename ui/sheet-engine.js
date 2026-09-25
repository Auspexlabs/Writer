// Spreadsheet formula engine: tokenizer, parser, evaluator, formatting, ref shifting.
// Values: number | string | boolean | '' (empty) | {err} (error) | {range: rows[][]} (range or array result).
export const NC = 20, NR = 80;
export const colName = c => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
export const colIdx = s => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
export const A = (r, c) => colName(c) + (r + 1);
export const parseA = a => { const m = /^\$?([A-Z]{1,2})\$?(\d+)$/i.exec(a); return m ? { r: +m[2] - 1, c: colIdx(m[1]) } : null; };
export const isErr = v => !!v && typeof v === 'object' && 'err' in v;
class FErr { constructor(e) { this.e = e; } }
const fail = e => { throw new FErr(e); };
const E = { DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!', NAME: '#NAME?', NA: '#N/A', NUM: '#NUM!', NULL: '#NULL!', SPILL: '#SPILL!', CALC: '#CALC!', CIRC: '#CIRC!' };
const ERRS = Object.values(E);
const MAXCELLS = 1e6; // ponytail: hard cap on range size, sparse ranges if someone needs bigger grids

// ---- tokenizer / parser ----
// A reference: optional sheet prefix, then A1 / A1:B2 / A:C (whole columns) / 3:5 (whole rows), $ allowed.
const SHEET = "(?:(?:'[^']+'|[A-Za-z_\\u4e00-\\u9fa5][\\w\\u4e00-\\u9fa5]*)!)?";
const REFSRC = SHEET + '(?:\\$?[A-Za-z]{1,2}\\$?\\d+(?::\\$?[A-Za-z]{1,2}\\$?\\d+)?|\\$?[A-Za-z]{1,2}:\\$?[A-Za-z]{1,2}|\\$?\\d+:\\$?\\d+)';
const REFSTR = new RegExp('^' + REFSRC + '$', 'i');
const TK = new RegExp('\\s*(?:(' + REFSRC + ')(?![\\w(])|(\\d+(?:\\.\\d+)?(?:[eE][+-]?\\d+)?|\\.\\d+(?:[eE][+-]?\\d+)?)|("(?:[^"]|"")*")|(#N\\/A|#DIV\\/0!|#VALUE!|#REF!|#NAME\\?|#NUM!|#NULL!|#SPILL!|#CALC!)|([A-Za-z_][\\w.]*)(?=\\s*\\()|(TRUE|FALSE)\\b|([A-Za-z_][\\w.]*)|(<=|>=|<>|[-+*/^&=<>%(),{};]))', 'iy');
function tokenize(s) {
  if (/[:!]\s*#REF!|#REF!\s*:/.test(s)) fail(E.REF);
  const out = []; let i = 0;
  while (i < s.length) {
    if (/^\s*$/.test(s.slice(i))) break;
    TK.lastIndex = i; const m = TK.exec(s); if (!m) fail(E.NAME); i = TK.lastIndex;
    if (m[1]) out.push({ t: 'ref', v: m[1] });
    else if (m[2]) out.push({ t: 'n', v: +m[2] });
    else if (m[3]) out.push({ t: 's', v: m[3].slice(1, -1).replace(/""/g, '"') });
    else if (m[4]) out.push({ t: 'e', v: m[4].toUpperCase() });
    else if (m[5]) out.push({ t: 'fn', v: m[5].toUpperCase() });
    else if (m[6]) out.push({ t: 'b', v: m[6].toUpperCase() === 'TRUE' });
    else if (m[7]) out.push({ t: 'name', v: m[7].toUpperCase() });
    else out.push({ t: 'op', v: m[8] });
  }
  return out;
}
// Whole-column ranges carry r2 = null, whole-row ranges c2 = null; Calc resolves them against the used range.
function refNode(s) {
  let sh = null; const i = s.lastIndexOf('!');
  if (i >= 0) { sh = s.slice(0, i).replace(/^'|'$/g, ''); s = s.slice(i + 1); }
  const [a, b] = s.split(':');
  if (b === undefined) { const p = parseA(a); if (!p) fail(E.REF); return { t: 'ref', sh, r: p.r, c: p.c }; }
  let m;
  if ((m = /^\$?([A-Za-z]{1,2})$/.exec(a))) { const c1 = colIdx(m[1]), c2 = colIdx(b.replace('$', '')); return { t: 'rng', sh, r1: 0, c1: Math.min(c1, c2), r2: null, c2: Math.max(c1, c2) }; }
  if ((m = /^\$?(\d+)$/.exec(a))) { const r1 = +m[1] - 1, r2 = +b.replace('$', '') - 1; return { t: 'rng', sh, r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: null }; }
  const p1 = parseA(a), p2 = parseA(b); if (!p1 || !p2) fail(E.REF);
  return { t: 'rng', sh, r1: Math.min(p1.r, p2.r), c1: Math.min(p1.c, p2.c), r2: Math.max(p1.r, p2.r), c2: Math.max(p1.c, p2.c) };
}
function parse(src) {
  const T = tokenize(src); let p = 0;
  const isOp = v => T[p] && T[p].t === 'op' && T[p].v === v;
  const next = () => T[p++];
  const cmp = () => { let a = cat(); while (T[p] && T[p].t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(T[p].v)) { const op = next().v; a = { t: 'op', op, a, b: cat() }; } return a; };
  const cat = () => { let a = add(); while (isOp('&')) { next(); a = { t: 'op', op: '&', a, b: add() }; } return a; };
  const add = () => { let a = mul(); while (isOp('+') || isOp('-')) { const op = next().v; a = { t: 'op', op, a, b: mul() }; } return a; };
  const mul = () => { let a = pw(); while (isOp('*') || isOp('/')) { const op = next().v; a = { t: 'op', op, a, b: pw() }; } return a; };
  const pw = () => { let a = un(); while (isOp('^')) { next(); a = { t: 'op', op: '^', a, b: un() }; } return a; };
  const un = () => { if (isOp('-')) { next(); return { t: 'neg', a: un() }; } if (isOp('+')) { next(); return un(); } let a = prim(); while (isOp('%')) { next(); a = { t: 'pct', a }; } return a; };
  const prim = () => {
    const k = next(); if (!k) fail(E.VALUE);
    if (k.t === 'n' || k.t === 's' || k.t === 'b') return { t: k.t, v: k.v };
    if (k.t === 'e') return { t: 'err', e: k.v };
    if (k.t === 'name') return { t: 'name', v: k.v };
    if (k.t === 'ref') return refNode(k.v);
    if (k.t === 'fn') {
      if (!isOp('(')) fail(E.NAME); next(); const args = [];
      if (!isOp(')')) { const arg = () => args.push(isOp(',') || isOp(')') ? { t: 'empty' } : cmp()); arg(); while (isOp(',')) { next(); arg(); } }
      if (!isOp(')')) fail(E.VALUE); next(); return { t: 'fn', name: k.v, args };
    }
    if (k.t === 'op' && k.v === '(') { const e = cmp(); if (!isOp(')')) fail(E.VALUE); next(); return e; }
    if (k.t === 'op' && k.v === '{') { // array constant {1,2;3,4}: "," separates columns, ";" rows
      const rows = [[]];
      for (;;) { rows[rows.length - 1].push(cmp()); if (isOp(',')) { next(); continue; } if (isOp(';')) { next(); rows.push([]); continue; } break; }
      if (!isOp('}') || rows.some(r => r.length !== rows[0].length)) fail(E.VALUE); next(); return { t: 'arr', v: rows };
    }
    fail(E.VALUE);
  };
  const e = cmp(); if (p < T.length) fail(E.VALUE); return e;
}
const AST = new Map();
function ast(f) { if (!AST.has(f)) { try { AST.set(f, parse(f)); } catch (e) { AST.set(f, { t: 'err', e: e instanceof FErr ? e.e : E.NAME }); } } return AST.get(f); }

// ---- value coercion ----
const isRng = v => !!v && typeof v === 'object' && 'range' in v;
const sc = v => { if (isRng(v)) v = v.range[0][0]; if (isErr(v)) fail(v.err); return v; };
const P15 = x => typeof x === 'number' && isFinite(x) ? +x.toPrecision(15) : x;
const num = v => {
  v = sc(v); if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; if (v === '' || v == null) return 0;
  const s = String(v).trim(); let n = Number(s.replace(/,/g, '').replace(/^[¥￥$]/, '')); if (!isNaN(n) && s !== '') return n;
  if (/%$/.test(s)) { n = Number(s.slice(0, -1).replace(/,/g, '')); if (!isNaN(n)) return n / 100; }
  const d = parseDate(s); if (d != null) return d;
  fail(E.VALUE);
};
const str = v => { v = sc(v); if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return String(P15(v)); return v == null ? '' : String(v); };
const truthy = v => { v = sc(v); if (typeof v === 'string') { if (/^true$/i.test(v)) return true; if (/^false$/i.test(v)) return false; if (v === '') return false; fail(E.VALUE); } return !!num(v); };
const rank = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
function cmpv(op, a, b) {
  let x = a, y = b;
  if (x === '' || x == null) x = typeof y === 'number' ? 0 : typeof y === 'boolean' ? false : '';
  if (y === '' || y == null) y = typeof x === 'number' ? 0 : typeof x === 'boolean' ? false : '';
  const rx = rank(x), ry = rank(y);
  if (rx !== ry) { switch (op) { case '=': return false; case '<>': return true; case '<': case '<=': return rx < ry; default: return rx > ry; } }
  if (rx === 0) { x = P15(x); y = P15(y); } else if (rx === 1) { x = x.toLowerCase(); y = y.toLowerCase(); } else { x = +x; y = +y; }
  switch (op) { case '=': return x === y; case '<>': return x !== y; case '<': return x < y; case '>': return x > y; case '<=': return x <= y; default: return x >= y; }
}
// Excel sort order for SORT/UNIQUE: numbers < text (case-insensitive) < logicals < errors; blanks last.
function order(a, b) {
  const blank = v => v === '' || v == null; if (blank(a) && blank(b)) return 0; if (blank(a)) return 1; if (blank(b)) return -1;
  const ra = isErr(a) ? 3 : rank(a), rb = isErr(b) ? 3 : rank(b); if (ra !== rb) return ra - rb;
  if (ra === 0) return a - b; if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; } if (ra === 2) return +a - +b; return 0;
}
const has = (a, i) => !!a[i] && a[i].t !== 'empty';
const evo = (a, i, ev) => has(a, i) ? ev(a[i]) : undefined;
function flat(args, ev) { const out = []; args.forEach(a => { if (a.t === 'empty') return; const v = ev(a); if (isRng(v)) v.range.forEach(r => r.forEach(x => out.push({ v: x, ref: true }))); else out.push({ v, ref: false }); }); return out; }
function nums(args, ev) { const o = []; flat(args, ev).forEach(({ v, ref }) => { if (isErr(v)) fail(v.err); if (ref) { if (typeof v === 'number') o.push(v); } else o.push(num(v)); }); return o; }
const rng = v => isRng(v) ? v.range : [[sc(v)]];
const rngv = (a, i, ev) => rng(ev(a[i]));
const grid = (R, C, f) => { if (R * C > MAXCELLS) fail(E.NUM); const rows = []; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) row.push(f(i, j)); rows.push(row); } return rows; };
const dims = m => [m.length, m[0] ? m[0].length : 0];
// Broadcast f over array arguments (Excel array lifting). Errors become per-element error values.
function lift2(vals, f) {
  const arrs = vals.map(v => isRng(v) ? v.range : null);
  if (!arrs.some(Boolean)) return f(...vals.map(v => v === undefined ? undefined : sc(v)));
  const R = Math.max(...arrs.map(x => x ? x.length : 1)), C = Math.max(...arrs.map(x => x ? x[0].length : 1));
  return { range: grid(R, C, (i, j) => { try { return f(...vals.map((v, k) => { const x = arrs[k]; if (!x) return v === undefined ? undefined : sc(v); const ri = x.length === 1 ? 0 : i, ci = x[0].length === 1 ? 0 : j; if (ri >= x.length || ci >= x[0].length) fail(E.NA); const e = x[ri][ci]; if (isErr(e)) fail(e.err); return e; })); } catch (e) { return { err: e instanceof FErr ? e.e : E.VALUE }; } }) };
}
// S(f, min): scalar function; args are evaluated, arrays broadcast element-wise, empty/missing args arrive as undefined.
const S = (f, min = 0) => (a, ev) => { if (a.length < min) fail(E.VALUE); return lift2(a.map(n => n.t === 'empty' ? undefined : ev(n)), f); };
const oNum = (v, d) => v === undefined ? d : num(v);
const oStr = (v, d) => v === undefined ? d : str(v);
const oBool = (v, d) => v === undefined ? d : truthy(v);
// Criteria for *IF/*IFS: ">10", "<>x", wildcards * ? with ~ escape, numbers/dates/booleans; text compares case-insensitively.
const wild = s => new RegExp('^' + s.replace(/~([*?~])|([.+^${}()|[\]\\])|(\*)|(\?)/g, (m, e, x, st, q) => e ? e.replace(/[.*?+^${}()|[\]\\]/g, '\\$&') : x ? '\\' + x : st ? '.*' : '.') + '$', 'i');
function crit(c) {
  c = sc(c);
  if (typeof c === 'number') return v => typeof v === 'number' ? P15(v) === P15(c) : str(v) === String(c);
  if (typeof c === 'boolean') return v => v === c;
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(String(c)); const op = m[1] || '=', rhs = m[2];
  if (rhs === '') return op === '<>' ? (v => v !== '' && v != null) : op === '=' ? (v => v === '' || v == null) : (() => false);
  let rn = null; const plain = rhs.replace(/,/g, '');
  if (plain.trim() !== '' && !isNaN(Number(plain))) rn = Number(plain);
  else if (/^true$/i.test(rhs)) return v => op === '<>' ? v !== true : v === true;
  else if (/^false$/i.test(rhs)) return v => op === '<>' ? v !== false : v === false;
  else { const d = parseDate(rhs); if (d != null) rn = d; else if (/%$/.test(plain) && !isNaN(Number(plain.slice(0, -1)))) rn = Number(plain.slice(0, -1)) / 100; }
  if (rn !== null) return v => { if (typeof v !== 'number') return op === '<>' ? true : op === '=' ? typeof v === 'string' && v.trim().toLowerCase() === rhs.trim().toLowerCase() : false; return cmpv(op, v, rn); };
  if ((op === '=' || op === '<>') && /[*?]/.test(rhs.replace(/~./g, ''))) { const re = wild(rhs); return v => { const ok = typeof v === 'string' && re.test(v); return op === '=' ? ok : !ok; }; }
  const lit = rhs.replace(/~([*?~])/g, '$1');
  if (op === '=') return v => typeof v === 'string' && v.toLowerCase() === lit.toLowerCase();
  if (op === '<>') return v => !(typeof v === 'string' && v.toLowerCase() === lit.toLowerCase());
  return v => typeof v === 'string' && cmpv(op, v, lit);
}
// Shared *IFS walker: (range, criteria) pairs from index `from` → [i,j] hits; every range must match the first's dims.
function ifsHits(a, ev, from) {
  const tests = []; for (let k = from; k < a.length; k += 2) { const r = rngv(a, k, ev); if (!has(a, k + 1)) fail(E.VALUE); tests.push([r, crit(ev(a[k + 1]))]); }
  if (!tests.length) fail(E.VALUE); const [R, C] = dims(tests[0][0]); const hits = [];
  tests.forEach(([r]) => { const [r2, c2] = dims(r); if (r2 !== R || c2 !== C) fail(E.VALUE); });
  for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) if (tests.every(([r, f]) => f(r[i][j]))) hits.push([i, j]);
  return hits;
}
function ifsValues(a, ev) {
  const values = rngv(a, 0, ev), criteria = rngv(a, 1, ev);
  if (dims(values).some((n, i) => n !== dims(criteria)[i])) fail(E.VALUE);
  return ifsHits(a, ev, 1).map(([i, j]) => values[i][j]).filter(x => typeof x === 'number');
}
// ---- dates: Excel 1900 serial system (serial 60 = the fictitious 1900-02-29, serials 1..59 sit one day early) ----
const EPOCH = Date.UTC(1899, 11, 30);
const leap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate(); // m is 1-based
function ymd(v) {
  const s = Math.floor(v); if (s < 0 || s > 2958465) fail(E.NUM);
  if (s === 0) return { y: 1900, m: 1, d: 0 }; if (s === 60) return { y: 1900, m: 2, d: 29 };
  const dt = new Date(EPOCH + (s < 60 ? s + 1 : s) * 864e5); return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}
function serial(y, m, d) { // DATE() semantics: month/day overflow rolls over, years 0..1899 are offset by 1900
  y = Math.trunc(y); m = Math.trunc(m); d = Math.trunc(d); if (y >= 0 && y < 1900) y += 1900; if (y < 1900 || y > 9999) fail(E.NUM);
  if (y === 1900 && m === 2 && d === 29) return 60;
  let s = Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 864e5); if (s < 61) s -= 1; if (s < 0 || s > 2958465) fail(E.NUM); return s;
}
const wday = v => ((Math.floor(v) - 1) % 7 + 7) % 7; // 0 = Sunday, follows Excel's serials (so 1900-01-01 is a "Sunday")
function hms(v) { let t = Math.min(86399, Math.round((v - Math.floor(v)) * 86400)); const h = Math.floor(t / 3600); t -= h * 3600; return { h, mi: Math.floor(t / 60), s: t % 60 }; }
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], CWD = '日一二三四五六'; // i18n-ok: date-format output, not UI
const monthOf = name => { const i = MONTHS.findIndex(m => m.toLowerCase().startsWith(name.toLowerCase())); return name.length >= 3 && i >= 0 ? i + 1 : 0; };
const yearOf = y => y == null ? new Date().getFullYear() : +y < 100 ? +y + (+y < 30 ? 2000 : 1900) : +y;
// Text → serial (date and/or time) or null. Accepts 2024-1-1, 2024/1/1, 2024年1月1日, 1/1/2024, 1/1 (this year), 22-Aug-2011, Aug 22, 2011, 12:30[:15] [AM/PM].
function parseDate(s) {
  s = String(s).trim(); if (!s) return null;
  const tm = /(\d{1,2}):(\d{2})(?::(\d{1,2})(?:\.\d+)?)?\s*(am|pm|上午|下午)?$/i.exec(s);
  let date = tm ? s.slice(0, tm.index).replace(/[T\s]+$/, '') : s, y, mo, d, m;
  let t = 0;
  if (tm) {
    let h = +tm[1]; const mi = +tm[2], se = tm[3] ? +tm[3] : 0, ap = tm[4] && tm[4].toLowerCase();
    if (ap) { if (h < 1 || h > 12) return null; h = h % 12 + (/pm|下午/.test(ap) ? 12 : 0); } else if (h > 23) return null;
    if (mi > 59 || se > 59) return null; t = (h * 3600 + mi * 60 + se) / 86400;
  }
  if (date === '') return tm ? t : null;
  if ((m = /^(\d{4})[-/.年](\d{1,2})(?:[-/.月](\d{1,2})日?|月)?$/.exec(date))) { y = +m[1]; mo = +m[2]; d = m[3] ? +m[3] : 1; }
  else if ((m = /^(\d{1,2})[-/](\d{1,2})(?:[-/](\d{2,4}))?$/.exec(date))) { mo = +m[1]; d = +m[2]; y = yearOf(m[3]); }
  else if ((m = /^(\d{1,2})月(\d{1,2})日?$/.exec(date))) { mo = +m[1]; d = +m[2]; y = yearOf(null); }
  else if ((m = /^(\d{1,2})[-\s/]([A-Za-z]{3,9})\.?(?:[-\s/,]+(\d{2,4}))?$/.exec(date))) { d = +m[1]; mo = monthOf(m[2]); y = yearOf(m[3]); }
  else if ((m = /^([A-Za-z]{3,9})\.?[-\s/]+(\d{1,2})(?:(?:,\s*|[-\s/]+)(\d{2,4}))?$/.exec(date))) { mo = monthOf(m[1]); d = +m[2]; y = yearOf(m[3]); }
  else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > (y === 1900 && mo === 2 ? 29 : daysIn(y, mo)) || y > 9999) return null;
  if (y < 1900) { const s = Math.round((Date.UTC(y, mo - 1, d) - EPOCH) / 864e5); return s >= 0 ? s + t : null; } // the engine writes a pure time as 1899-12-30 hh:mm:ss
  return serial(y, mo, d) + t;
}
function dateParts(v) { const p = ymd(v), tt = hms(v); return Object.assign(p, tt, { wd: wday(v) }); }

// ---- rounding & number formatting ----
const shift = (x, d) => { const [m, e] = x.toExponential().split('e'); return +(m + 'e' + (+e + d)); };
const round = (x, d, mode) => { if (!isFinite(x)) return x; d = Math.trunc(d); const y = shift(Math.abs(x), d); const r = mode === 'up' ? Math.ceil(y - 1e-12 * Math.max(1, y)) : mode === 'down' ? Math.floor(y + 1e-12 * Math.max(1, y)) : Math.round(y); return Math.sign(x) * shift(r, -d); };
const snap = q => Math.abs(q - Math.round(q)) < 1e-9 * Math.max(1, Math.abs(q)) ? Math.round(q) : q;
const pad = (n, w = 2) => String(n).padStart(w, '0');
const grp = (v, d) => v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
const general = v => Number.isInteger(v) ? String(v) : String(+v.toPrecision(10));
// Excel-style format codes for TEXT(): number masks (0 # ? , . % E+), literals, [$¥-804], and date/time codes (yyyy m d h s AM/PM aaaa).
function fmtCode(v, code) {
  if (typeof v === 'string') { const n = Number(v.replace(/,/g, '')); if (v.trim() !== '' && !isNaN(n)) v = n; else return v; }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v !== 'number') return String(v);
  if (/^general$/i.test(code) || code === '') return general(v);
  const secs = code.split(';'); let sec = secs[0], neg = v < 0;
  if (neg && secs.length > 1) { sec = secs[1]; v = -v; neg = false; } else if (v === 0 && secs.length > 2) sec = secs[2];
  const toks = []; // {k:'lit'|'num'|'date'|'ap', v}
  for (let i = 0; i < sec.length;) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); toks.push({ k: 'lit', v: sec.slice(i + 1, j < 0 ? sec.length : j) }); i = j < 0 ? sec.length : j + 1; continue; }
    if (ch === '\\') { toks.push({ k: 'lit', v: sec[i + 1] || '' }); i += 2; continue; }
    if (ch === '_') { toks.push({ k: 'lit', v: ' ' }); i += 2; continue; }
    if (ch === '*') { i += 2; continue; }
    if (ch === '[') { const j = sec.indexOf(']', i); const b = sec.slice(i + 1, j < 0 ? sec.length : j); i = j < 0 ? sec.length : j + 1; if (b[0] === '$') toks.push({ k: 'lit', v: b.slice(1).split('-')[0] }); else if (/^(h+|m+|s+)$/i.test(b)) toks.push({ k: 'date', v: '[' + b.toLowerCase() + ']' }); continue; }
    let m;
    if (/^general/i.test(sec.slice(i))) { toks.push({ k: 'num', v: '@' }); i += 7; continue; } // [Red]General, "$"General
    if ((m = /^(AM\/PM|A\/P|上午\/下午)/i.exec(sec.slice(i)))) { toks.push({ k: 'ap', v: m[1] }); i += m[1].length; continue; }
    if ((m = /^(E[+-])/i.exec(sec.slice(i)))) { toks.push({ k: 'num', v: 'E' }); i += 2; continue; }
    if ((m = /^(y+|d+|h+|s+|m+|a+)/i.exec(sec.slice(i)))) { toks.push({ k: 'date', v: m[1].toLowerCase() }); i += m[1].length; continue; }
    if (/[0#?,.%@]/.test(ch)) { toks.push({ k: 'num', v: ch }); i++; continue; }
    toks.push({ k: 'lit', v: ch }); i++;
  }
  const isDate = toks.some(t => t.k === 'ap' || (t.k === 'date' && /^[ydhsa\[]/.test(t.v))) || (toks.some(t => t.k === 'date') && !toks.some(t => t.k === 'num' && /[0#?]/.test(t.v)));
  if (isDate) {
    const p = dateParts(v); const ap = toks.some(t => t.k === 'ap'); let out = '';
    toks.forEach((t, i) => {
      if (t.k === 'lit') { out += t.v; return; }
      if (t.k === 'ap') { out += /上午/.test(t.v) ? (p.h < 12 ? '上午' : '下午') : t.v.length === 3 ? (p.h < 12 ? 'A' : 'P') : (p.h < 12 ? 'AM' : 'PM'); return; } // i18n-ok: date-format output, not UI
      if (t.k !== 'date') { out += t.v; return; }
      const c = t.v, h12 = ap ? (p.h % 12 || 12) : p.h;
      if (c[0] === '[') { const secs = Math.round(v * 86400); out += c[1] === 'h' ? Math.floor(secs / 3600) : c[1] === 'm' ? Math.floor(secs / 60) : secs; return; }
      if (c[0] === 'y') { out += c.length <= 2 ? pad(p.y % 100) : p.y; return; }
      if (c[0] === 'd') { out += c.length === 1 ? p.d : c.length === 2 ? pad(p.d) : c.length === 3 ? WDAYS[p.wd].slice(0, 3) : WDAYS[p.wd]; return; }
      if (c[0] === 'h') { out += c.length === 1 ? h12 : pad(h12); return; }
      if (c[0] === 's') { out += c.length === 1 ? p.s : pad(p.s); return; }
      if (c[0] === 'a') { out += c.length >= 4 ? '星期' + CWD[p.wd] : '周' + CWD[p.wd]; return; } // i18n-ok: date-format output, not UI
      const prev = toks.slice(0, i).reverse().find(x => x.k === 'date'), nxt = toks.slice(i + 1).find(x => x.k === 'date');
      const minute = (prev && /^\[?h/.test(prev.v)) || (nxt && /^s/.test(nxt.v));
      if (minute) out += c.length === 1 ? p.mi : pad(p.mi);
      else out += c.length === 1 ? p.m : c.length === 2 ? pad(p.m) : c.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : c.length === 4 ? MONTHS[p.m - 1] : MONTHS[p.m - 1][0];
    });
    return out;
  }
  const fr = /(\?+)\s*\/\s*(\?+|\d+)/.exec(sec.replace(/"[^"]*"/g, ''));
  if (fr) return fraction(v, /[#0]\s*\?/.test(sec), fr[2]); // # ?/?, # ??/??, ?/4
  if (toks.some(t => t.k === 'num' && t.v === '@')) return toks.map(t => t.k === 'num' && t.v === '@' ? general(v) : t.k === 'lit' ? t.v : '').join('');
  const mask = toks.filter(t => t.k === 'num').map(t => t.v).join('');
  if (!/[0#?]/.test(mask)) return (neg ? '-' : '') + toks.map(t => t.k === 'lit' ? t.v : t.v === '%' ? '%' : '').join('');
  let x = Math.abs(v);
  (mask.match(/%/g) || []).forEach(() => { x *= 100; });
  const sci = mask.includes('E');
  const body = sci ? mask.split('E')[0] : mask;
  const m2 = /^([^.]*)(?:\.(.*))?$/.exec(body.replace(/[%E]/g, '')); const ip = m2[1], dp = m2[2] || '';
  const trail = /,+$/.exec(ip); if (trail) x /= Math.pow(1000, trail[0].length);
  const grouping = /[0#?],[0#?]/.test(ip), minInt = (ip.match(/0/g) || []).length, minDec = (dp.match(/0/g) || []).length, maxDec = (dp.match(/[0#?]/g) || []).length;
  let intS, decS = '', exp = '';
  if (sci) { const ex = mask.split('E')[1] || '00'; let e = x === 0 ? 0 : Math.floor(Math.log10(x)); let mant = x / Math.pow(10, e); if (round(mant, maxDec) >= 10) { mant /= 10; e++; } x = mant; exp = 'E' + (e < 0 ? '-' : '+') + pad(Math.abs(e), (ex.match(/0/g) || []).length); }
  const fixed = round(x, maxDec).toFixed(maxDec); [intS, decS] = fixed.split('.'); decS = decS || '';
  while (decS.length > minDec && decS.endsWith('0')) decS = decS.slice(0, -1);
  if (intS === '0' && minInt === 0) intS = ''; else intS = intS.padStart(minInt, '0');
  if (grouping && intS) intS = intS.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const numS = intS + (decS ? '.' + decS : '') + exp;
  let out = '', placed = false;
  toks.forEach(t => { if (t.k === 'lit') out += t.v; else if (t.k === 'num') { if (t.v === '%') out += '%'; else if (!placed) { out += numS; placed = true; } } });
  return (neg ? '-' : '') + out; // Excel leads with the sign: "$"#,##0 shows -5 as -$5
}
/** A fraction as Excel writes it: a whole part when the code has one, the nearest fraction with at most as many denominator digits as
 *  the code's ?s (or its fixed denominator). */
function fraction(v, whole, den) {
  const x = Math.abs(v); let ip = whole ? Math.floor(x) : 0, n = 0, d = 1;
  const f = x - ip;
  if (/^\d+$/.test(den)) { d = +den; n = Math.round(f * d); }
  else { let err = Infinity; for (let q = 1; q < 10 ** den.length; q++) { const p = Math.round(f * q), e = Math.abs(f - p / q); if (e < err - 1e-12) { err = e; n = p; d = q; } } }
  if (whole && n === d) { ip++; n = 0; }
  const text = whole ? (n ? (ip ? ip + ' ' : '') + n + '/' + d : String(ip)) : n + '/' + d;
  return (v < 0 ? '-' : '') + text;
}
const FMT_COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00' };
/** The colour a format code's section gives a number ([Red] in #,##0;[Red]-#,##0), or null. */
export function fmtColor(v, code) {
  if (typeof v !== 'number' || !code) return null;
  const secs = code.split(';'), sec = v < 0 && secs.length > 1 ? secs[1] : v === 0 && secs.length > 2 ? secs[2] : secs[0];
  const m = /\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(sec);
  return m ? FMT_COLORS[m[1].toLowerCase()] : null;
}
export { fmtCode, parseDate as dateSerial };
export function decimalsOf(v) { if (typeof v !== 'number' || Number.isInteger(v)) return 0; const s = String(+v.toPrecision(10)); return (s.split('.')[1] || '').length; }
export function fmt(v, s) {
  s = s || {};
  if (isErr(v)) return v.err;
  if (v === '' || v == null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    const d = s.dec;
    if (s.code) { try { return fmtCode(v, s.code); } catch (e) { return '#####'; } } // the file's own code, as Excel would show it
    switch (s.fmt) {
      case 'number': return grp(v, d ?? 2);
      case 'money': return (v < 0 ? '-' : '') + '¥' + grp(Math.abs(v), d ?? 2);
      case 'pct': return (v * 100).toFixed(d ?? 0) + '%';
      case 'date': { if (v < 0 || v > 2958465) return '#####'; const p = ymd(v); return `${p.y}/${p.m}/${p.d}`; }
      case 'time': { if (v < 0) return '#####'; const t = hms(v); return `${pad(t.h)}:${pad(t.mi)}:${pad(t.s)}`; }
      case 'datetime': { if (v < 0 || v > 2958465) return '#####'; const p = dateParts(v); return `${p.y}/${p.m}/${p.d} ${pad(p.h)}:${pad(p.mi)}`; }
      default: if (d != null) return v.toFixed(d); return general(v);
    }
  }
  return String(v);
}
// ---- functions ----
// FN[name](argNodes, ev, calc) → value. Scalar functions are wrapped with S() and receive evaluated args.
const FN = {};
const sum = n => n.reduce((s, x) => s + x, 0);
const avg = n => { if (!n.length) fail(E.DIV0); return sum(n) / n.length; };
const vari = (n, pop) => { if (n.length < (pop ? 1 : 2)) fail(E.DIV0); const m = sum(n) / n.length; return sum(n.map(x => (x - m) ** 2)) / (n.length - (pop ? 0 : 1)); };
const asc = n => [...n].sort((x, y) => x - y);
const median = n => { if (!n.length) fail(E.NUM); n = asc(n); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; };
const mode = n => { const cnt = new Map(); let best = null, bc = 1; n.forEach(x => { const c = (cnt.get(x) || 0) + 1; cnt.set(x, c); if (c > bc) { bc = c; best = x; } }); if (best === null) fail(E.NA); return best; };
const kth = (n, k, desc) => { k = Math.trunc(k); if (!n.length || k < 1 || k > n.length) fail(E.NUM); n = asc(n); return desc ? n[n.length - k] : n[k - 1]; };
const pctInc = (n, p) => { if (!n.length || p < 0 || p > 1) fail(E.NUM); n = asc(n); const h = (n.length - 1) * p, lo = Math.floor(h); return lo + 1 < n.length ? n[lo] + (h - lo) * (n[lo + 1] - n[lo]) : n[lo]; };
const pctExc = (n, p) => { if (!n.length || p <= 0 || p >= 1) fail(E.NUM); n = asc(n); const h = (n.length + 1) * p; if (h < 1 || h > n.length) fail(E.NUM); const lo = Math.floor(h); return lo < n.length ? n[lo - 1] + (h - lo) * (n[lo] - n[lo - 1]) : n[lo - 1]; };
// AVERAGEA-style: text in refs → 0, booleans → 1/0, empty ignored; direct args coerced.
const numsA = (a, ev) => { const o = []; flat(a, ev).forEach(({ v, ref }) => { if (isErr(v)) fail(v.err); if (ref) { if (v === '' || v == null) return; o.push(typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0); } else o.push(num(v)); }); return o; };
const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };
const pairs = (ys, xs) => { if (ys.length !== xs.length) fail(E.NA); const p = []; for (let i = 0; i < ys.length; i++) if (typeof ys[i] === 'number' && typeof xs[i] === 'number') p.push([xs[i], ys[i]]); return p; };
const linreg = p => { const n = p.length; if (n < 2) fail(E.DIV0); const mx = sum(p.map(q => q[0])) / n, my = sum(p.map(q => q[1])) / n; let sxy = 0, sxx = 0; p.forEach(([x, y]) => { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; }); if (!sxx) fail(E.DIV0); const b = sxy / sxx; return { b, a: my - b * mx }; };
const fitArgs = (a, ev, ln) => { const y = rngv(a, 0, ev).flat(); const ys = y.map(v => typeof v === 'number' ? (ln ? Math.log(v) : v) : v); const xs = has(a, 1) ? rngv(a, 1, ev).flat() : ys.map((_, i) => i + 1); const nx = has(a, 2) ? rngv(a, 2, ev) : (has(a, 1) ? rngv(a, 1, ev) : [xs]); return { ys, xs, nx }; };
const SUBS = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP', 12: 'MEDIAN', 13: 'MODE', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' };
const normCdf = z => { const az = Math.abs(z); let p; if (az > 37) p = 0; else { const e = Math.exp(-az * az / 2); if (az < 7.07106781186547) p = e * ((((((3.52624965998911e-2 * az + 0.700383064443688) * az + 6.37396220353165) * az + 33.912866078383) * az + 112.079291497871) * az + 221.213596169931) * az + 220.206867912376) / (((((((8.83883476483184e-2 * az + 1.75566716318264) * az + 16.064177579207) * az + 86.7807322029461) * az + 296.564248779674) * az + 637.333633378831) * az + 793.826512519948) * az + 440.413735824752); else p = e / ((az + 1 / (az + 2 / (az + 3 / (az + 4 / (az + 0.65))))) * 2.506628274631); } return z > 0 ? 1 - p : p; };
const normInv = p => { if (p <= 0 || p >= 1) fail(E.NUM); const a = [-39.6968302866538, 220.946098424521, -275.928510446969, 138.357751867269, -30.6647980661472, 2.50662827745924], b = [-54.4760987982241, 161.585836858041, -155.698979859887, 66.8013118877197, -13.2806815528857], c = [-7.78489400243029e-3, -0.322396458041136, -2.40075827716184, -2.54973253934373, 4.37466414146497, 2.93816398269878], d = [7.78469570904146e-3, 0.32246712907004, 2.445134137143, 3.75440866190742]; let x, q, r; if (p < 0.02425) { q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); } else if (p <= 0.97575) { q = p - 0.5; r = q * q; x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); } else { q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); } const e = normCdf(x) - p, u = e * Math.sqrt(2 * Math.PI) * Math.exp(x * x / 2); return x - u / (1 + x * u / 2); };
const rankOf = (x, n, order, avgTies) => { if (!n.includes(x)) fail(E.NA); const gt = n.filter(v => order ? v < x : v > x).length, eq = n.filter(v => v === x).length; return avgTies ? gt + (eq + 1) / 2 : gt + 1; };
Object.assign(FN, {
  // 数学
  SUM: (a, ev) => sum(nums(a, ev)),
  SUMIF: (a, ev) => { const r = rngv(a, 0, ev), f = crit(ev(a[1])), s = has(a, 2) ? rngv(a, 2, ev) : r; let t = 0; r.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') t += x; } })); return t; },
  SUMIFS: (a, ev) => sum(ifsValues(a, ev)),
  SUMPRODUCT: (a, ev) => { const arrs = a.map(n => rng(ev(n))); const [R, C] = dims(arrs[0]); arrs.forEach(m => { const [r, c] = dims(m); if (r !== R || c !== C) fail(E.VALUE); }); let t = 0; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { let p = 1; for (const m of arrs) { const x = m[i][j]; if (isErr(x)) fail(x.err); p *= typeof x === 'number' ? x : 0; } t += p; } return t; },
  PRODUCT: (a, ev) => { const n = nums(a, ev); return n.length ? n.reduce((s, x) => s * x, 1) : 0; },
  ABS: S(x => Math.abs(num(x)), 1),
  ROUND: S((x, d) => round(num(x), oNum(d, 0)), 1),
  ROUNDUP: S((x, d) => round(num(x), oNum(d, 0), 'up'), 1),
  ROUNDDOWN: S((x, d) => round(num(x), oNum(d, 0), 'down'), 1),
  INT: S(x => Math.floor(num(x)), 1),
  TRUNC: S((x, d) => round(num(x), oNum(d, 0), 'down'), 1),
  MOD: S((x, d) => { x = num(x); d = num(d); if (!d) fail(E.DIV0); return P15(x - d * Math.floor(snap(x / d))); }, 2),
  POWER: S((x, y) => { x = num(x); y = num(y); if (x === 0 && y < 0) fail(E.DIV0); const r = Math.pow(x, y); if (isNaN(r) || !isFinite(r)) fail(E.NUM); return r; }, 2),
  SQRT: S(x => { x = num(x); if (x < 0) fail(E.NUM); return Math.sqrt(x); }, 1),
  EXP: S(x => Math.exp(num(x)), 1),
  LN: S(x => { x = num(x); if (x <= 0) fail(E.NUM); return Math.log(x); }, 1),
  LOG: S((x, b) => { x = num(x); b = oNum(b, 10); if (x <= 0 || b <= 0 || b === 1) fail(E.NUM); return P15(Math.log(x) / Math.log(b)); }, 1),
  LOG10: S(x => { x = num(x); if (x <= 0) fail(E.NUM); return Math.log10(x); }, 1),
  CEILING: S((x, s) => { x = num(x); s = num(s); if (s === 0) return 0; if (x * s < 0) fail(E.NUM); return P15(Math.ceil(snap(x / s)) * s); }, 2),
  'CEILING.MATH': S((x, s, m) => { x = num(x); s = Math.abs(oNum(s, 1)); if (!s) return 0; const q = snap(x / s); return P15((x < 0 && oNum(m, 0) ? Math.floor(q) : Math.ceil(q)) * s); }, 1),
  FLOOR: S((x, s) => { x = num(x); s = num(s); if (s === 0) fail(E.DIV0); if (x * s < 0) fail(E.NUM); return P15(Math.floor(snap(x / s)) * s); }, 2),
  'FLOOR.MATH': S((x, s, m) => { x = num(x); s = Math.abs(oNum(s, 1)); if (!s) return 0; const q = snap(x / s); return P15((x < 0 && oNum(m, 0) ? Math.ceil(q) : Math.floor(q)) * s); }, 1),
  MROUND: S((x, m) => { x = num(x); m = num(m); if (!m) return 0; if (x * m < 0) fail(E.NUM); const q = snap(x / m); return P15(Math.sign(q) * Math.round(Math.abs(q)) * m); }, 2),
  SIGN: S(x => Math.sign(num(x)), 1),
  RAND: () => Math.random(),
  RANDBETWEEN: S((lo, hi) => { lo = Math.ceil(num(lo)); hi = Math.floor(num(hi)); if (lo > hi) fail(E.NUM); return lo + Math.floor(Math.random() * (hi - lo + 1)); }, 2),
  PI: () => Math.PI,
  FACT: S(n => { n = Math.floor(num(n)); if (n < 0 || n > 170) fail(E.NUM); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; }, 1),
  COMBIN: S((n, k) => { n = Math.floor(num(n)); k = Math.floor(num(k)); if (n < 0 || k < 0 || k > n) fail(E.NUM); k = Math.min(k, n - k); let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); }, 2),
  GCD: (a, ev) => nums(a, ev).map(x => { if (x < 0) fail(E.NUM); return Math.floor(x); }).reduce(gcd2, 0),
  LCM: (a, ev) => nums(a, ev).map(x => { if (x < 0) fail(E.NUM); return Math.floor(x); }).reduce((l, x) => l === 0 || x === 0 ? 0 : l * x / gcd2(l, x), 1),
  QUOTIENT: S((x, y) => { x = num(x); y = num(y); if (!y) fail(E.DIV0); return Math.trunc(snap(x / y)); }, 2),
  EVEN: S(x => { x = num(x); return x >= 0 ? 2 * Math.ceil(snap(x / 2)) : 2 * Math.floor(snap(x / 2)); }, 1),
  ODD: S(x => { x = num(x); const n = x >= 0 ? Math.ceil(snap(x)) : Math.floor(snap(x)); return n % 2 ? n : x >= 0 ? n + 1 : n - 1; }, 1),
  SUBTOTAL: (a, ev) => { const k = num(ev(a[0])), f = SUBS[k > 100 ? k - 100 : k]; if (!f || k > 11 && k < 101 || k > 111) fail(E.VALUE); return FN[f](a.slice(1), ev); },
  AGGREGATE: (a, ev) => {
    const k = num(ev(a[0])), o = num(ev(a[1])), f = SUBS[k]; if (!f || o < 0 || o > 7) fail(E.VALUE);
    const refs = k >= 14 ? a.slice(2, -1) : a.slice(2), ign = o >= 2 && o !== 4 && o !== 5; // options 2,3,6,7 ignore errors
    const items = flat(refs, ev).filter(({ v }) => !(ign && isErr(v)));
    if (k === 3) return items.filter(({ v }) => v !== '' && v != null).length;
    const n = []; items.forEach(({ v, ref }) => { if (isErr(v)) fail(v.err); if (ref) { if (typeof v === 'number') n.push(v); } else n.push(num(v)); });
    const kk = k >= 14 ? num(ev(a[a.length - 1])) : 0;
    switch (k) { case 1: return avg(n); case 2: return n.length; case 4: return n.length ? Math.max(...n) : 0; case 5: return n.length ? Math.min(...n) : 0; case 6: return n.length ? n.reduce((s, x) => s * x, 1) : 0; case 7: return Math.sqrt(vari(n)); case 8: return Math.sqrt(vari(n, true)); case 9: return sum(n); case 10: return vari(n); case 11: return vari(n, true); case 12: return median(n); case 13: return mode(n); case 14: return kth(n, kk, true); case 15: return kth(n, kk); case 16: return pctInc(n, kk); case 17: if (kk < 0 || kk > 4) fail(E.NUM); return pctInc(n, kk / 4); case 18: return pctExc(n, kk); default: if (kk < 1 || kk > 3) fail(E.NUM); return pctExc(n, kk / 4); }
  },
  // 统计
  AVERAGE: (a, ev) => avg(nums(a, ev)),
  AVERAGEA: (a, ev) => avg(numsA(a, ev)),
  AVERAGEIF: (a, ev) => { const r = rngv(a, 0, ev), f = crit(ev(a[1])), s = has(a, 2) ? rngv(a, 2, ev) : r; const n = []; r.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = s[i] && s[i][j]; if (typeof x === 'number') n.push(x); } })); return avg(n); },
  AVERAGEIFS: (a, ev) => avg(ifsValues(a, ev)),
  COUNT: (a, ev) => flat(a, ev).filter(({ v, ref }) => ref ? typeof v === 'number' : typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && v !== '' && !isNaN(Number(v)))).length,
  COUNTA: (a, ev) => flat(a, ev).filter(({ v }) => v !== '' && v != null).length,
  COUNTBLANK: (a, ev) => flat(a, ev).filter(({ v }) => v === '' || v == null).length,
  COUNTIF: (a, ev) => { const r = rngv(a, 0, ev), f = crit(ev(a[1])); let t = 0; r.forEach(row => row.forEach(v => { if (f(v)) t++; })); return t; },
  COUNTIFS: (a, ev) => ifsHits(a, ev, 0).length,
  MAX: (a, ev) => { const n = nums(a, ev); return n.length ? Math.max(...n) : 0; },
  MAXA: (a, ev) => { const n = numsA(a, ev); return n.length ? Math.max(...n) : 0; },
  MIN: (a, ev) => { const n = nums(a, ev); return n.length ? Math.min(...n) : 0; },
  MINA: (a, ev) => { const n = numsA(a, ev); return n.length ? Math.min(...n) : 0; },
  MAXIFS: (a, ev) => { const n = ifsValues(a, ev); return n.length ? Math.max(...n) : 0; },
  MINIFS: (a, ev) => { const n = ifsValues(a, ev); return n.length ? Math.min(...n) : 0; },
  MEDIAN: (a, ev) => median(nums(a, ev)),
  MODE: (a, ev) => mode(nums(a, ev)),
  'MODE.SNGL': (a, ev) => mode(nums(a, ev)),
  STDEV: (a, ev) => Math.sqrt(vari(nums(a, ev))),
  'STDEV.S': (a, ev) => Math.sqrt(vari(nums(a, ev))),
  'STDEV.P': (a, ev) => Math.sqrt(vari(nums(a, ev), true)),
  STDEVP: (a, ev) => Math.sqrt(vari(nums(a, ev), true)),
  VAR: (a, ev) => vari(nums(a, ev)),
  'VAR.S': (a, ev) => vari(nums(a, ev)),
  'VAR.P': (a, ev) => vari(nums(a, ev), true),
  VARP: (a, ev) => vari(nums(a, ev), true),
  LARGE: (a, ev) => kth(nums([a[0]], ev), num(ev(a[1])), true),
  SMALL: (a, ev) => kth(nums([a[0]], ev), num(ev(a[1]))),
  RANK: (a, ev) => rankOf(num(ev(a[0])), nums([a[1]], ev), has(a, 2) && num(ev(a[2]))),
  'RANK.EQ': (a, ev) => rankOf(num(ev(a[0])), nums([a[1]], ev), has(a, 2) && num(ev(a[2]))),
  'RANK.AVG': (a, ev) => rankOf(num(ev(a[0])), nums([a[1]], ev), has(a, 2) && num(ev(a[2])), true),
  PERCENTILE: (a, ev) => pctInc(nums([a[0]], ev), num(ev(a[1]))),
  'PERCENTILE.INC': (a, ev) => pctInc(nums([a[0]], ev), num(ev(a[1]))),
  'PERCENTILE.EXC': (a, ev) => pctExc(nums([a[0]], ev), num(ev(a[1]))),
  QUARTILE: (a, ev) => { const q = num(ev(a[1])); if (q < 0 || q > 4) fail(E.NUM); return pctInc(nums([a[0]], ev), Math.trunc(q) / 4); },
  'QUARTILE.INC': (a, ev) => { const q = num(ev(a[1])); if (q < 0 || q > 4) fail(E.NUM); return pctInc(nums([a[0]], ev), Math.trunc(q) / 4); },
  'QUARTILE.EXC': (a, ev) => { const q = num(ev(a[1])); if (q < 1 || q > 3) fail(E.NUM); return pctExc(nums([a[0]], ev), Math.trunc(q) / 4); },
  CORREL: (a, ev) => { const p = pairs(rngv(a, 0, ev).flat(), rngv(a, 1, ev).flat()); const n = p.length; if (n < 2) fail(E.DIV0); const mx = sum(p.map(q => q[0])) / n, my = sum(p.map(q => q[1])) / n; let sxy = 0, sxx = 0, syy = 0; p.forEach(([x, y]) => { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; }); if (!sxx || !syy) fail(E.DIV0); return sxy / Math.sqrt(sxx * syy); },
  PEARSON: (a, ev) => FN.CORREL(a, ev),
  SLOPE: (a, ev) => linreg(pairs(rngv(a, 0, ev).flat(), rngv(a, 1, ev).flat())).b,
  INTERCEPT: (a, ev) => linreg(pairs(rngv(a, 0, ev).flat(), rngv(a, 1, ev).flat())).a,
  FORECAST: (a, ev) => { const { a: c, b } = linreg(pairs(rngv(a, 1, ev).flat(), rngv(a, 2, ev).flat())); return c + b * num(ev(a[0])); },
  'FORECAST.LINEAR': (a, ev) => FN.FORECAST(a, ev),
  TREND: (a, ev) => { const { ys, xs, nx } = fitArgs(a, ev); const r = linreg(pairs(ys, xs)); return { range: nx.map(row => row.map(x => r.a + r.b * num(x))) }; },
  GROWTH: (a, ev) => { const { ys, xs, nx } = fitArgs(a, ev, true); const r = linreg(pairs(ys, xs)); return { range: nx.map(row => row.map(x => Math.exp(r.a + r.b * num(x)))) }; },
  FREQUENCY: (a, ev) => { const d = nums([a[0]], ev), bins = asc(nums([a[1]], ev)); const out = bins.map(() => 0).concat([0]); d.forEach(x => { let i = bins.findIndex(b => x <= b); if (i < 0) i = bins.length; out[i]++; }); return { range: out.map(x => [x]) }; },
  GEOMEAN: (a, ev) => { const n = nums(a, ev); if (!n.length || n.some(x => x <= 0)) fail(E.NUM); return Math.exp(sum(n.map(Math.log)) / n.length); },
  HARMEAN: (a, ev) => { const n = nums(a, ev); if (!n.length || n.some(x => x <= 0)) fail(E.NUM); return n.length / sum(n.map(x => 1 / x)); },
  'NORM.DIST': S((x, m, sd, c) => { x = num(x); m = num(m); sd = num(sd); if (sd <= 0) fail(E.NUM); return truthy(c) ? normCdf((x - m) / sd) : Math.exp(-(((x - m) / sd) ** 2) / 2) / (sd * Math.sqrt(2 * Math.PI)); }, 4),
  'NORM.INV': S((p, m, sd) => { p = num(p); sd = num(sd); if (sd <= 0) fail(E.NUM); return num(m) + sd * normInv(p); }, 3),
  'NORM.S.DIST': S((z, c) => oBool(c, true) ? normCdf(num(z)) : Math.exp(-(num(z) ** 2) / 2) / Math.sqrt(2 * Math.PI), 1),
  'NORM.S.INV': S(p => normInv(num(p)), 1),
  NORMDIST: (a, ev) => FN['NORM.DIST'](a, ev), NORMINV: (a, ev) => FN['NORM.INV'](a, ev), NORMSDIST: (a, ev) => FN['NORM.S.DIST'](a, ev), NORMSINV: (a, ev) => FN['NORM.S.INV'](a, ev),
});
// DBCS byte helpers (CJK/full-width = 2 bytes, like Excel in a Chinese locale). A half-cut double-byte char becomes a space.
const bLen = s => { let n = 0; for (const ch of s) n += ch.charCodeAt(0) > 255 ? 2 : 1; return n; };
const bLeft = (s, n) => { let out = '', used = 0; for (const ch of s) { const w = ch.charCodeAt(0) > 255 ? 2 : 1; if (used + w > n) { if (used < n) out += ' '; break; } out += ch; used += w; } return out; };
const bRight = (s, n) => { let out = '', used = 0; for (let i = s.length - 1; i >= 0; i--) { const w = s.charCodeAt(i) > 255 ? 2 : 1; if (used + w > n) { if (used < n) out = ' ' + out; break; } out = s[i] + out; used += w; } return out; };
const bMid = (s, st, n) => { let i = 0, used = 0; while (i < s.length && used < st - 1) { used += s.charCodeAt(i) > 255 ? 2 : 1; i++; } const lead = used > st - 1 ? ' ' : ''; return lead + bLeft(s.slice(i), n - lead.length); };
const nth = (s, d, inst, ci) => { // index of the inst-th (negative: from the end) occurrence of d in s, -1 if none
  const hay = ci ? s.toLowerCase() : s, nd = ci ? d.toLowerCase() : d; const pos = []; let i = hay.indexOf(nd); while (i >= 0) { pos.push(i); i = hay.indexOf(nd, i + Math.max(1, nd.length)); }
  if (!inst || !Number.isInteger(inst)) fail(E.VALUE); const k = inst > 0 ? inst - 1 : pos.length + inst; return k >= 0 && k < pos.length ? pos[k] : -1;
};
// |n| with d decimals (d < 0 keeps zeros), thousands grouped in the integer part only unless noGroup.
const fixedS = (n, d, noGroup) => { const s = Math.abs(n).toFixed(Math.max(0, Math.trunc(d))); if (noGroup) return s; const [i, f] = s.split('.'); return i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? '.' + f : ''); };
const textBA = after => S((t, d, inst, mm, me, nf) => {
  t = str(t); d = str(d); const ci = oNum(mm, 0) !== 0; if (d === '') fail(E.VALUE);
  const i = nth(t, d, oNum(inst, 1), ci); if (i < 0) { if (nf === undefined) fail(E.NA); return nf; }
  return after ? t.slice(i + d.length) : t.slice(0, i);
}, 2);
const logicals = (a, ev) => { const o = []; flat(a, ev).forEach(({ v, ref }) => { if (isErr(v)) fail(v.err); if (ref) { if (typeof v === 'boolean' || typeof v === 'number') o.push(!!v); } else o.push(truthy(v)); }); if (!o.length) fail(E.VALUE); return o; };
Object.assign(FN, {
  // 文本
  LEN: S(x => str(x).length, 1),
  LENB: S(x => bLen(str(x)), 1),
  LEFT: S((s, n) => { n = oNum(n, 1); if (n < 0) fail(E.VALUE); return str(s).slice(0, Math.trunc(n)); }, 1),
  LEFTB: S((s, n) => { n = oNum(n, 1); if (n < 0) fail(E.VALUE); return bLeft(str(s), Math.trunc(n)); }, 1),
  RIGHT: S((s, n) => { n = Math.trunc(oNum(n, 1)); if (n < 0) fail(E.VALUE); s = str(s); return n ? s.slice(-n) : ''; }, 1),
  RIGHTB: S((s, n) => { n = Math.trunc(oNum(n, 1)); if (n < 0) fail(E.VALUE); return bRight(str(s), n); }, 1),
  MID: S((s, st, n) => { st = Math.trunc(num(st)); n = Math.trunc(num(n)); if (st < 1 || n < 0) fail(E.VALUE); return str(s).substr(st - 1, n); }, 3),
  MIDB: S((s, st, n) => { st = Math.trunc(num(st)); n = Math.trunc(num(n)); if (st < 1 || n < 0) fail(E.VALUE); return bMid(str(s), st, n); }, 3),
  UPPER: S(x => str(x).toUpperCase(), 1),
  LOWER: S(x => str(x).toLowerCase(), 1),
  PROPER: S(x => str(x).toLowerCase().replace(/[a-zà-ɏ]+/g, w => w[0].toUpperCase() + w.slice(1)), 1),
  TRIM: S(x => str(x).replace(/^ +| +$/g, '').replace(/ {2,}/g, ' '), 1),
  CLEAN: S(x => str(x).replace(/[\x00-\x1F]/g, ''), 1),
  CONCAT: (a, ev) => flat(a, ev).map(({ v }) => str(v)).join(''),
  CONCATENATE: (a, ev) => flat(a, ev).map(({ v }) => str(v)).join(''),
  TEXTJOIN: (a, ev) => { const d = str(ev(a[0])), ie = truthy(ev(a[1])); return flat(a.slice(2), ev).map(({ v }) => str(v)).filter(s => !ie || s !== '').join(d); },
  TEXT: S((v, f) => fmtCode(v === undefined ? 0 : v, oStr(f, 'General')), 1),
  VALUE: S(x => { if (typeof x === 'number') return x; if (typeof x === 'boolean' || x === undefined || str(x).trim() === '') fail(E.VALUE); return num(x); }, 1),
  NUMBERVALUE: S((t, ds, gs) => { t = oStr(t, ''); ds = oStr(ds, '.'); gs = oStr(gs, ','); if (t.trim() === '') return 0; if (!ds || !gs) fail(E.VALUE); let s = t.split(gs[0]).join('').split(ds[0]).join('.').replace(/\s+/g, ''); let pct = 0; while (s.endsWith('%')) { s = s.slice(0, -1); pct++; } const n = Number(s); if (s === '' || isNaN(n)) fail(E.VALUE); return n / Math.pow(100, pct); }, 1),
  FIND: S((f, w, st) => { f = str(f); w = str(w); st = Math.trunc(oNum(st, 1)); if (st < 1 || st > w.length + 1) fail(E.VALUE); const i = w.indexOf(f, st - 1); if (i < 0) fail(E.VALUE); return i + 1; }, 2),
  FINDB: S((f, w, st) => { f = str(f); w = str(w); st = Math.trunc(oNum(st, 1)); if (st < 1) fail(E.VALUE); const i = w.indexOf(f, Math.max(0, bLeft(w, st - 1).length)); if (i < 0) fail(E.VALUE); return bLen(w.slice(0, i)) + 1; }, 2),
  SEARCH: S((f, w, st) => { f = str(f); w = str(w); st = Math.trunc(oNum(st, 1)); if (st < 1 || st > w.length + 1) fail(E.VALUE); const re = new RegExp(wild(f).source.slice(1, -1), 'i'); const m = re.exec(w.slice(st - 1)); if (!m) fail(E.VALUE); return st + m.index; }, 2),
  REPLACE: S((o, st, n, nw) => { o = str(o); st = Math.trunc(num(st)); n = Math.trunc(num(n)); if (st < 1 || n < 0) fail(E.VALUE); return o.slice(0, st - 1) + str(nw) + o.slice(st - 1 + n); }, 4),
  SUBSTITUTE: S((t, o, n, inst) => { t = str(t); o = str(o); n = str(n); if (o === '') return t; if (inst === undefined) return t.split(o).join(n); inst = Math.trunc(num(inst)); if (inst < 1) fail(E.VALUE); const i = nth(t, o, inst); return i < 0 ? t : t.slice(0, i) + n + t.slice(i + o.length); }, 3),
  REPT: S((t, n) => { t = str(t); n = Math.trunc(num(n)); if (n < 0 || t.length * n > 32767) fail(E.VALUE); return t.repeat(n); }, 2),
  EXACT: S((x, y) => str(x) === str(y), 2),
  CHAR: S(n => { n = Math.trunc(num(n)); if (n < 1 || n > 255) fail(E.VALUE); return String.fromCharCode(n); }, 1),
  CODE: S(s => { s = str(s); if (!s) fail(E.VALUE); return s.charCodeAt(0); }, 1),
  UNICHAR: S(n => { n = Math.trunc(num(n)); if (n < 1 || n > 0x10FFFF) fail(E.VALUE); return String.fromCodePoint(n); }, 1),
  UNICODE: S(s => { s = str(s); if (!s) fail(E.VALUE); return s.codePointAt(0); }, 1),
  T: S(x => typeof x === 'string' ? x : '', 1),
  N: S(x => typeof x === 'number' ? x : typeof x === 'boolean' ? +x : 0, 1),
  FIXED: S((n, d, nc) => { const r = round(num(n), Math.trunc(oNum(d, 2))); return (r < 0 ? '-' : '') + fixedS(r, oNum(d, 2), oBool(nc, false)); }, 1),
  DOLLAR: S((n, d) => { const r = round(num(n), Math.trunc(oNum(d, 2))); const s = '$' + fixedS(r, oNum(d, 2)); return r < 0 ? '(' + s + ')' : s; }, 1),
  RMB: S((n, d) => { const r = round(num(n), Math.trunc(oNum(d, 2))); return (r < 0 ? '-' : '') + '¥' + fixedS(r, oNum(d, 2)); }, 1),
  TEXTBEFORE: textBA(false),
  TEXTAFTER: textBA(true),
  TEXTSPLIT: (a, ev) => {
    const t = str(ev(a[0])), cd = has(a, 1) ? str(ev(a[1])) : '', rd = has(a, 2) ? str(ev(a[2])) : '', ie = has(a, 3) && truthy(ev(a[3])), padv = has(a, 5) ? sc(ev(a[5])) : { err: E.NA };
    if (!cd && !rd) fail(E.VALUE);
    let rows = (rd ? t.split(rd) : [t]).map(r => cd ? r.split(cd) : [r]); if (ie) rows = rows.map(r => r.filter(x => x !== '')).filter(r => r.length);
    if (!rows.length) fail(E.VALUE); const C = Math.max(...rows.map(r => r.length));
    return { range: rows.map(r => r.concat(Array(C - r.length).fill(padv)).map(x => { const n = Number(x); return x !== '' && !isNaN(n) && typeof x === 'string' && x.trim() !== '' ? n : x; })) };
  },
  // 逻辑
  IF: (a, ev) => {
    const c = ev(a[0]); const br = i => a[i] ? (a[i].t === 'empty' ? 0 : ev(a[i])) : i === 1;
    if (isRng(c)) return lift2([c, br(1), br(2)], (x, y, z) => truthy(x) ? y : z);
    return truthy(c) ? br(1) : br(2);
  },
  IFS: (a, ev) => { for (let i = 0; i + 1 < a.length; i += 2) if (truthy(ev(a[i]))) return ev(a[i + 1]); fail(E.NA); },
  IFERROR: (a, ev) => { let v; try { v = ev(a[0]); } catch (e) { return ev(a[1]); } if (isRng(v)) { const alt = has(a, 1) ? ev(a[1]) : ''; return { range: v.range.map(r => r.map(x => isErr(x) ? sc(alt) : x)) }; } return isErr(v) ? ev(a[1]) : v; },
  IFNA: (a, ev) => { let v; try { v = sc(ev(a[0])); } catch (e) { if (e instanceof FErr && e.e === E.NA) return ev(a[1]); throw e; } return v; },
  AND: (a, ev) => logicals(a, ev).every(Boolean),
  OR: (a, ev) => logicals(a, ev).some(Boolean),
  XOR: (a, ev) => logicals(a, ev).filter(Boolean).length % 2 === 1,
  NOT: S(x => !truthy(x), 1),
  TRUE: () => true,
  FALSE: () => false,
  SWITCH: (a, ev) => { const x = sc(ev(a[0])); let i = 1; for (; i + 1 < a.length; i += 2) if (cmpv('=', x, sc(ev(a[i])))) return ev(a[i + 1]); if (i < a.length) return ev(a[i]); fail(E.NA); },
  CHOOSE: (a, ev) => { const i = Math.trunc(num(ev(a[0]))); if (i < 1 || i >= a.length) fail(E.VALUE); return ev(a[i]); },
  LET: (a, ev, calc) => { if (a.length < 3 || a.length % 2 === 0) fail(E.VALUE); const saved = calc.scope; calc.scope = new Map(saved); try { for (let i = 0; i + 1 < a.length - 1; i += 2) { if (a[i].t !== 'name') fail(E.VALUE); calc.scope.set(a[i].v, ev(a[i + 1])); } return ev(a[a.length - 1]); } finally { calc.scope = saved; } },
});
// ---- lookup helpers ----
const blank = v => v === '' || v == null;
// xmatch: index in list per XLOOKUP/XMATCH modes. mm: 0 exact, -1 next smaller, 1 next larger, 2 wildcard. sm: 1 first→last, -1 last→first (2/-2 binary treated as sorted linear).
function xmatch(x, list, mm, sm) {
  const idx = list.map((_, i) => i); if (sm < 0) idx.reverse();
  const same = v => !blank(v) && !isErr(v) && rank(v) === rank(x);
  if (mm === 2) { if (typeof x !== 'string') mm = 0; else { const re = wild(x); for (const i of idx) if (typeof list[i] === 'string' && re.test(list[i])) return i; fail(E.NA); } }
  for (const i of idx) if (same(list[i]) && cmpv('=', list[i], x)) return i;
  if (mm === 0) fail(E.NA);
  let best = -1;
  for (const i of idx) { const v = list[i]; if (!same(v)) continue; if (mm < 0 ? cmpv('<', v, x) && (best < 0 || cmpv('>', v, list[best])) : cmpv('>', v, x) && (best < 0 || cmpv('<', v, list[best]))) best = i; }
  if (best < 0) fail(E.NA); return best;
}
// Classic MATCH (and VLOOKUP/HLOOKUP/LOOKUP): type 0 exact (wildcards), 1 largest ≤ x (ascending), -1 smallest ≥ x (descending).
// Approximate modes binary-search the values of x's type like Excel does, so unsorted data gives Excel's answer, not a scan's. Blanks never match.
function matchIdx(x, list, type) {
  const same = v => !blank(v) && !isErr(v) && rank(v) === rank(x);
  if (type === 0) return xmatch(x, list, typeof x === 'string' && /[*?]/.test(x.replace(/~./g, '')) ? 2 : 0, 1);
  const cand = list.map((_, i) => i).filter(i => same(list[i])), op = type > 0 ? '<=' : '>=';
  let lo = 0, hi = cand.length - 1, k = -1;
  while (lo <= hi) { const mid = (lo + hi) >> 1; if (cmpv(op, list[cand[mid]], x)) { k = cand[mid]; lo = mid + 1; } else hi = mid - 1; }
  if (k < 0) fail(E.NA); return k;
}
const vec = m => { const [R, C] = dims(m); if (R !== 1 && C !== 1) fail(E.VALUE); return R === 1 ? m[0] : m.map(r => r[0]); };
const col = (m, j) => m.map(r => r[j]);
const rowKey = r => r.map(v => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(v))).join('\u0001');
const orderArg = v => { v = v === undefined ? 1 : num(v); if (v !== 1 && v !== -1) fail(E.VALUE); return v; };
// Weekend argument of NETWORKDAYS.INTL / WORKDAY.INTL → [Sun..Sat] booleans: 1 = Sat+Sun … 7 = Fri+Sat, 11 = Sun only … 17 = Sat only, or "0000011" (Mon..Sun).
const wmask = v => {
  if (v === undefined || v === null || v === '') v = 1;
  if (typeof v === 'string') { if (!/^[01]{7}$/.test(v) || v === '1111111') fail(E.VALUE); return [v[6], v[0], v[1], v[2], v[3], v[4], v[5]].map(x => x === '1'); }
  v = Math.trunc(num(v)); const m = Array(7).fill(false);
  if (v >= 1 && v <= 7) { m[(v + 5) % 7] = true; m[(v + 6) % 7] = true; } else if (v >= 11 && v <= 17) m[v - 11] = true; else fail(E.NUM);
  return m;
};
// Holidays: numbers from a range (text cells are ignored, as in Excel); a scalar or an array constant is coerced ({"2006/1/2"} counts).
const holidaySet = (a, i, ev) => { const s = new Set(); if (!has(a, i)) return s; const lit = a[i].t === 'arr'; flat([a[i]], ev).forEach(({ v, ref }) => { if (isErr(v)) fail(v.err); if (blank(v) || (ref && !lit && typeof v !== 'number')) return; s.add(Math.floor(num(v))); }); return s; };
Object.assign(FN, {
  // 查找与引用
  VLOOKUP: (a, ev) => {
    const x = sc(ev(a[0])), t = rngv(a, 1, ev), ci = Math.trunc(num(ev(a[2]))), approx = has(a, 3) ? truthy(ev(a[3])) : true;
    if (ci < 1) fail(E.VALUE); if (ci > t[0].length) fail(E.REF);
    return t[matchIdx(x, col(t, 0), approx ? 1 : 0)][ci - 1];
  },
  HLOOKUP: (a, ev) => {
    const x = sc(ev(a[0])), t = rngv(a, 1, ev), ri = Math.trunc(num(ev(a[2]))), approx = has(a, 3) ? truthy(ev(a[3])) : true;
    if (ri < 1) fail(E.VALUE); if (ri > t.length) fail(E.REF);
    return t[ri - 1][matchIdx(x, t[0], approx ? 1 : 0)];
  },
  LOOKUP: (a, ev) => {
    const x = sc(ev(a[0])), m = rngv(a, 1, ev); const [R, C] = dims(m);
    if (has(a, 2)) { const rv = vec(rngv(a, 2, ev)), lv = vec(m); const i = matchIdx(x, lv, 1); if (i >= rv.length) fail(E.NA); return rv[i]; }
    if (R === 1 || C === 1) { const lv = vec(m); return lv[matchIdx(x, lv, 1)]; }
    return C > R ? m[R - 1][matchIdx(x, m[0], 1)] : m[matchIdx(x, col(m, 0), 1)][C - 1];
  },
  MATCH: (a, ev) => { const x = sc(ev(a[0])), t = has(a, 2) ? num(ev(a[2])) : 1; return matchIdx(x, vec(rngv(a, 1, ev)), t === 0 ? 0 : t > 0 ? 1 : -1) + 1; },
  XMATCH: (a, ev) => xmatch(sc(ev(a[0])), vec(rngv(a, 1, ev)), has(a, 2) ? num(ev(a[2])) : 0, has(a, 3) ? num(ev(a[3])) : 1) + 1,
  XLOOKUP: (a, ev) => {
    const x = sc(ev(a[0])), la = rngv(a, 1, ev), ra = rngv(a, 2, ev), mm = has(a, 4) ? num(ev(a[4])) : 0, sm = has(a, 5) ? num(ev(a[5])) : 1;
    const [R, C] = dims(la); if (R !== 1 && C !== 1) fail(E.VALUE);
    let i; try { i = xmatch(x, vec(la), mm, sm); } catch (e) { if (e instanceof FErr && e.e === E.NA && has(a, 3)) return ev(a[3]); throw e; }
    if (C === 1) { if (i >= ra.length) fail(E.REF); return ra[i].length === 1 ? ra[i][0] : { range: [ra[i]] }; }
    if (i >= ra[0].length) fail(E.REF); return ra.length === 1 ? ra[0][i] : { range: ra.map(r => [r[i]]) };
  },
  INDEX: (a, ev) => {
    const m = rngv(a, 0, ev); const [R, C] = dims(m); let r = has(a, 1) ? Math.trunc(num(ev(a[1]))) : 0, c = has(a, 2) ? Math.trunc(num(ev(a[2]))) : 0;
    if (!has(a, 2) && (R === 1 || C === 1) && r) { const v = vec(m); if (r < 1 || r > v.length) fail(E.REF); return v[r - 1]; }
    if (r < 0 || c < 0 || r > R || c > C) fail(E.REF);
    if (!r && !c) return { range: m }; if (!r) return { range: m.map(row => [row[c - 1]]) }; if (!c) return { range: [m[r - 1]] };
    return m[r - 1][c - 1];
  },
  OFFSET: (a, ev, calc, si) => {
    const b = calc.bounds(a[0], si), dr = Math.trunc(num(ev(a[1]))), dc = Math.trunc(num(ev(a[2])));
    const h = has(a, 3) ? Math.trunc(num(ev(a[3]))) : b.r2 - b.r1 + 1, w = has(a, 4) ? Math.trunc(num(ev(a[4]))) : b.c2 - b.c1 + 1;
    const r1 = b.r1 + dr, c1 = b.c1 + dc; if (r1 < 0 || c1 < 0 || h < 1 || w < 1) fail(E.REF);
    return calc.ev({ t: 'rng', sh: calc.doc.sheets[b.si].name, r1, c1, r2: r1 + h - 1, c2: c1 + w - 1 }, si);
  },
  INDIRECT: (a, ev, calc, si) => { const s = str(ev(a[0])).trim(); const n = calc.nameNode(s) || (REFSTR.test(s) ? refNode(s) : fail(E.REF)); return calc.ev(n, si); },
  ROW: (a, ev, calc, si) => { if (!has(a, 0)) return calc.cur.r + 1; const b = calc.bounds(a[0], si); return b.r1 === b.r2 && b.c1 === b.c2 ? b.r1 + 1 : { range: grid(b.r2 - b.r1 + 1, 1, i => b.r1 + i + 1) }; },
  COLUMN: (a, ev, calc, si) => { if (!has(a, 0)) return calc.cur.c + 1; const b = calc.bounds(a[0], si); return b.r1 === b.r2 && b.c1 === b.c2 ? b.c1 + 1 : { range: grid(1, b.c2 - b.c1 + 1, (i, j) => b.c1 + j + 1) }; },
  ROWS: (a, ev) => rngv(a, 0, ev).length,
  COLUMNS: (a, ev) => rngv(a, 0, ev)[0].length,
  ADDRESS: S((r, c, ab, a1, sh) => {
    r = Math.trunc(num(r)); c = Math.trunc(num(c)); ab = Math.trunc(oNum(ab, 1)); if (r < 1 || c < 1 || ab < 1 || ab > 4) fail(E.VALUE);
    const pre = sh === undefined ? '' : (/^[A-Za-z_一-龥][\w一-龥]*$/.test(str(sh)) ? str(sh) : "'" + str(sh) + "'") + '!';
    if (!oBool(a1, true)) return pre + (ab <= 2 ? 'R' + r : 'R[' + r + ']') + (ab === 1 || ab === 3 ? 'C' + c : 'C[' + c + ']');
    return pre + (ab === 1 || ab === 3 ? '$' : '') + colName(c - 1) + (ab <= 2 ? '$' : '') + r;
  }, 2),
  AREAS: () => 1,
  TRANSPOSE: (a, ev) => { const m = rngv(a, 0, ev); const [R, C] = dims(m); return { range: grid(C, R, (i, j) => m[j][i]) }; },
  UNIQUE: (a, ev) => {
    let m = rngv(a, 0, ev); const byCol = has(a, 1) && truthy(ev(a[1])), once = has(a, 2) && truthy(ev(a[2]));
    if (byCol) m = FN.TRANSPOSE([a[0]], ev).range;
    const cnt = new Map(); m.forEach(r => { const k = rowKey(r); cnt.set(k, (cnt.get(k) || 0) + 1); });
    const seen = new Set(); let out = m.filter(r => { const k = rowKey(r); if (seen.has(k)) return false; seen.add(k); return !once || cnt.get(k) === 1; });
    if (!out.length) fail(E.CALC);
    if (byCol) { const [R, C] = dims(out); out = grid(C, R, (i, j) => out[j][i]); }
    return { range: out };
  },
  FILTER: (a, ev) => {
    const m = rngv(a, 0, ev), inc = rngv(a, 1, ev); const [R, C] = dims(m), [IR, IC] = dims(inc);
    const keep = v => { if (isErr(v)) fail(v.err); return !!num(v); };
    let out;
    if (IC === 1 && IR === R) out = m.filter((r, i) => keep(inc[i][0])); else if (IR === 1 && IC === C) { const cols = inc[0].map(keep); out = m.map(r => r.filter((_, j) => cols[j])); if (!cols.some(Boolean)) out = []; } else fail(E.VALUE);
    if (!out.length) { if (has(a, 2)) return ev(a[2]); fail(E.CALC); }
    return { range: out };
  },
  SORT: (a, ev) => {
    let m = rngv(a, 0, ev); const k = has(a, 1) ? Math.trunc(num(ev(a[1]))) : 1, o = orderArg(evo(a, 2, ev)), byCol = has(a, 3) && truthy(ev(a[3]));
    if (byCol) m = FN.TRANSPOSE([a[0]], ev).range; if (k < 1 || k > m[0].length) fail(E.VALUE);
    const out = m.map((r, i) => [r, i]).sort((x, y) => o * order(x[0][k - 1], y[0][k - 1]) || x[1] - y[1]).map(x => x[0]);
    return { range: byCol ? grid(out[0].length, out.length, (i, j) => out[j][i]) : out };
  },
  SORTBY: (a, ev) => {
    const m = rngv(a, 0, ev); const keys = []; for (let i = 1; i < a.length; i += 2) { const v = vec(rngv(a, i, ev)); if (v.length !== m.length) fail(E.VALUE); keys.push([v, orderArg(evo(a, i + 1, ev))]); }
    if (!keys.length) fail(E.VALUE);
    return { range: m.map((r, i) => i).sort((x, y) => { for (const [v, o] of keys) { const d = o * order(v[x], v[y]); if (d) return d; } return x - y; }).map(i => m[i]) };
  },
  SEQUENCE: S((r, c, st, sp) => { r = Math.trunc(num(r)); c = Math.trunc(oNum(c, 1)); st = oNum(st, 1); sp = oNum(sp, 1); if (r < 1 || c < 1) fail(E.CALC); return { range: grid(r, c, (i, j) => st + (i * c + j) * sp) }; }, 1),
});
// ---- date & time helpers ----
const dser = v => { v = num(v); if (v < 0 || v > 2958465) fail(E.NUM); return Math.floor(v); };
const dimX = (y, m) => y === 1900 && m === 2 ? 29 : daysIn(y, m); // Excel's February 1900 has 29 days
const lastDay = (y, m, d) => d === dimX(y, m);
const addMonths = (s, n) => { const p = ymd(s), t = p.m - 1 + n, y = p.y + Math.floor(t / 12), m = ((t % 12) + 12) % 12 + 1; return { y, m, d: p.d, last: dimX(y, m) }; };
function datedif(s1, s2, u) {
  if (s1 > s2) fail(E.NUM); const a = ymd(s1), b = ymd(s2);
  const before = b.m < a.m || (b.m === a.m && b.d < a.d); // end's month/day comes earlier in the year than start's
  switch (u.toUpperCase()) {
    case 'Y': return b.y - a.y - (before ? 1 : 0);
    case 'M': return (b.y - a.y) * 12 + b.m - a.m - (b.d < a.d ? 1 : 0);
    case 'D': return s2 - s1;
    case 'MD': return s2 - serial(b.y, b.m - (b.d < a.d ? 1 : 0), a.d);
    case 'YM': { const m = b.m - a.m - (b.d < a.d ? 1 : 0); return m < 0 ? m + 12 : m; }
    case 'YD': return s2 - serial(b.y - (before ? 1 : 0), a.m, a.d);
  }
  fail(E.NUM);
}
// DAYS360: US/NASD (a start on the last day of February counts as the 30th; a 31st end stays unless the start is the 30th) or European (31st → 30th).
function days360(s1, s2, eu) {
  const a = ymd(s1), b = ymd(s2); let d1 = a.d, d2 = b.d;
  if (d1 === 31 || (!eu && a.m === 2 && lastDay(a.y, 2, d1))) d1 = 30;
  if (d2 === 31 && (eu || d1 === 30)) d2 = 30;
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
}
function yearfrac(s1, s2, basis) {
  if (s1 > s2) [s1, s2] = [s2, s1]; const a = ymd(s1), b = ymd(s2), n = s2 - s1, d30 = (d1, d2) => ((b.y - a.y) * 360 + (b.m - a.m) * 30 + d2 - d1) / 360;
  switch (basis) {
    case 0: { let d1 = a.d, d2 = b.d; const f1 = a.m === 2 && lastDay(a.y, 2, d1), f2 = b.m === 2 && lastDay(b.y, 2, d2); if (f1 && f2) d2 = 30; if (f1) d1 = 30; if (d2 === 31 && d1 >= 30) d2 = 30; if (d1 === 31) d1 = 30; return d30(d1, d2); }
    case 1: {
      if (a.y === b.y) return n / (leap(a.y) ? 366 : 365);
      if (s2 <= serial(a.y + 1, a.m, a.d)) { const f29 = leap(a.y) ? serial(a.y, 2, 29) : leap(b.y) ? serial(b.y, 2, 29) : -1; return n / (f29 >= s1 && f29 <= s2 ? 366 : 365); }
      let days = 0; for (let y = a.y; y <= b.y; y++) days += leap(y) ? 366 : 365; return n / (days / (b.y - a.y + 1));
    }
    case 2: return n / 360;
    case 3: return n / 365;
    case 4: return d30(a.d === 31 ? 30 : a.d, b.d === 31 ? 30 : b.d);
  }
  fail(E.NUM);
}
// ponytail: working days are counted day by day; a century-long span is still only ~36k steps.
function netdays(s1, s2, mask, hol) { let sign = 1; if (s1 > s2) { [s1, s2] = [s2, s1]; sign = -1; } let n = 0; for (let d = s1; d <= s2; d++) if (!mask[wday(d)] && !hol.has(d)) n++; return sign * n; }
function workday(s, n, mask, hol) { if (mask.every(Boolean)) fail(E.VALUE); const st = n < 0 ? -1 : 1; n = Math.abs(Math.trunc(n)); while (n > 0) { s += st; if (!mask[wday(s)] && !hol.has(s)) n--; } if (s < 0 || s > 2958465) fail(E.NUM); return s; }
const isoweek = s => { const thu = s - (wday(s) + 6) % 7 + 3; return Math.floor((thu - serial(ymd(thu).y, 1, 1)) / 7) + 1; };
function weeknum(s, type) {
  if (type === 21) return isoweek(s);
  if (type !== 1 && type !== 2 && !(type >= 11 && type <= 17)) fail(E.NUM);
  const start = type === 1 ? 0 : type === 2 ? 1 : (type - 10) % 7, j = serial(ymd(s).y, 1, 1);
  return Math.floor((s - j + (wday(j) - start + 7) % 7) / 7) + 1;
}
const nowSerial = time => { const d = new Date(); return serial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (time ? (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400 : 0); };
Object.assign(FN, {
  // 日期与时间
  TODAY: () => nowSerial(false),
  NOW: () => nowSerial(true),
  DATE: S((y, m, d) => serial(num(y), num(m), num(d)), 3),
  TIME: S((h, m, s) => { h = num(h); m = num(m); s = num(s); if (h < 0 || m < 0 || s < 0) fail(E.NUM); return ((Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s)) % 86400) / 86400; }, 3),
  YEAR: S(v => ymd(num(v)).y, 1), MONTH: S(v => ymd(num(v)).m, 1), DAY: S(v => ymd(num(v)).d, 1),
  HOUR: S(v => hms(num(v)).h, 1), MINUTE: S(v => hms(num(v)).mi, 1), SECOND: S(v => hms(num(v)).s, 1),
  DATEVALUE: S(v => { const d = parseDate(str(v)); if (d == null) fail(E.VALUE); return Math.floor(d); }, 1),
  TIMEVALUE: S(v => { const d = parseDate(str(v)); if (d == null) fail(E.VALUE); return d - Math.floor(d); }, 1),
  WEEKDAY: S((v, t) => { const wd = wday(dser(v)); t = Math.trunc(oNum(t, 1)); if (t === 1) return wd + 1; if (t === 2) return (wd + 6) % 7 + 1; if (t === 3) return (wd + 6) % 7; if (t >= 11 && t <= 17) return (wd - (t - 10) % 7 + 7) % 7 + 1; fail(E.NUM); }, 1),
  WEEKNUM: S((v, t) => weeknum(dser(v), Math.trunc(oNum(t, 1))), 1),
  ISOWEEKNUM: S(v => isoweek(dser(v)), 1),
  DATEDIF: S((a, b, u) => datedif(dser(a), dser(b), str(u)), 3),
  EDATE: S((v, n) => { const p = addMonths(dser(v), Math.trunc(num(n))); return serial(p.y, p.m, Math.min(p.d, p.last)); }, 2),
  EOMONTH: S((v, n) => { const p = addMonths(dser(v), Math.trunc(num(n))); return serial(p.y, p.m, p.last); }, 2),
  DAYS: S((b, a) => dser(b) - dser(a), 2),
  DAYS360: S((a, b, eu) => days360(dser(a), dser(b), oBool(eu, false)), 2),
  NETWORKDAYS: (a, ev) => netdays(dser(ev(a[0])), dser(ev(a[1])), wmask(1), holidaySet(a, 2, ev)),
  'NETWORKDAYS.INTL': (a, ev) => netdays(dser(ev(a[0])), dser(ev(a[1])), wmask(evo(a, 2, ev)), holidaySet(a, 3, ev)),
  WORKDAY: (a, ev) => workday(dser(ev(a[0])), num(ev(a[1])), wmask(1), holidaySet(a, 2, ev)),
  'WORKDAY.INTL': (a, ev) => workday(dser(ev(a[0])), num(ev(a[1])), wmask(evo(a, 2, ev)), holidaySet(a, 3, ev)),
  YEARFRAC: S((a, b, basis) => yearfrac(dser(a), dser(b), Math.trunc(oNum(basis, 0))), 2),
});
// ---- information ----
const ERRNO = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 14, '#CALC!': 19, '#CIRC!': 19 };
const tryVal = (a, ev) => { try { return sc(ev(a[0])); } catch (e) { if (e instanceof FErr) return { err: e.e }; throw e; } }; // first argument, errors as values
const cellAt = (calc, n, si) => { const b = calc.bounds(n, si); return { b, cell: calc.doc.sheets[b.si].cells[A(b.r1, b.c1)] }; };
Object.assign(FN, {
  ISBLANK: (a, ev) => blank(tryVal(a, ev)),
  ISERROR: (a, ev) => isErr(tryVal(a, ev)),
  ISERR: (a, ev) => { const v = tryVal(a, ev); return isErr(v) && v.err !== E.NA; },
  ISNA: (a, ev) => { const v = tryVal(a, ev); return isErr(v) && v.err === E.NA; },
  ISNUMBER: (a, ev) => typeof tryVal(a, ev) === 'number',
  ISTEXT: (a, ev) => { const v = tryVal(a, ev); return typeof v === 'string' && v !== ''; },
  ISNONTEXT: (a, ev) => { const v = tryVal(a, ev); return typeof v !== 'string' || v === ''; },
  ISLOGICAL: (a, ev) => typeof tryVal(a, ev) === 'boolean',
  ISEVEN: S(x => Math.trunc(num(x)) % 2 === 0, 1),
  ISODD: S(x => Math.trunc(num(x)) % 2 !== 0, 1),
  ISFORMULA: (a, ev, calc, si) => { const { cell } = cellAt(calc, a[0], si); return !!cell && String(cell.v)[0] === '=' && String(cell.v).length > 1; },
  ISREF: (a, ev, calc) => !!a[0] && (a[0].t === 'ref' || a[0].t === 'rng' || (a[0].t === 'name' && !!calc.nameNode(a[0].v))),
  TYPE: (a, ev) => { let v; try { v = ev(a[0]); } catch (e) { return 16; } if (isRng(v)) return 64; if (isErr(v)) return 16; return typeof v === 'number' || blank(v) ? 1 : typeof v === 'string' ? 2 : 4; },
  NA: () => fail(E.NA),
  'ERROR.TYPE': (a, ev) => { const v = tryVal(a, ev); if (!isErr(v)) fail(E.NA); return ERRNO[v.err] || fail(E.NA); },
  CELL: (a, ev, calc, si) => {
    const t = str(ev(a[0])).toLowerCase(), n = has(a, 1) ? a[1] : { t: 'ref', sh: null, r: calc.cur.r, c: calc.cur.c }, { b, cell } = cellAt(calc, n, si), v = calc.value(b.si, b.r1, b.c1);
    switch (t) {
      case 'address': return '$' + colName(b.c1) + '$' + (b.r1 + 1); case 'row': return b.r1 + 1; case 'col': return b.c1 + 1;
      case 'contents': return blank(v) ? 0 : v; case 'type': return blank(v) ? 'b' : typeof v === 'string' ? 'l' : 'v';
      case 'prefix': return typeof v === 'string' && v !== '' ? ({ center: '^', right: '"' }[cell?.s?.align] || "'") : '';
      case 'format': return 'G'; case 'width': return 8.43; case 'protect': return 1; case 'color': return 0; case 'parentheses': return 0; case 'filename': return '';
    }
    fail(E.VALUE);
  },
  INFO: (a, ev, calc) => { const t = str(ev(a[0])).toLowerCase(); const o = { directory: '', numfile: calc.doc.sheets.length, osversion: 'Web', recalc: 'Automatic', release: '16.0', system: 'pcdos', origin: '$A:$A$1', memavail: 0, memused: 0, totmem: 0 }; if (!(t in o)) fail(E.VALUE); return o[t]; },
});
// ---- financial helpers (Excel sign convention: money paid out is negative) ----
const pmtOf = (r, n, pv, fv = 0, t = 0) => { if (!n) fail(E.NUM); if (!r) return -(pv + fv) / n; const k = Math.pow(1 + r, n); return -(pv * k + fv) * r / ((1 + r * t) * (k - 1)); };
const fvOf = (r, n, p, pv = 0, t = 0) => { if (!r) return -(pv + p * n); const k = Math.pow(1 + r, n); return -(pv * k + p * (1 + r * t) * (k - 1) / r); };
const pvOf = (r, n, p, fv = 0, t = 0) => { if (!r) return -(fv + p * n); const k = Math.pow(1 + r, n); return -(fv + p * (1 + r * t) * (k - 1) / r) / k; };
const nperOf = (r, p, pv, fv = 0, t = 0) => { if (!r) { if (!p) fail(E.NUM); return -(pv + fv) / p; } const a = p * (1 + r * t) - fv * r, b = p * (1 + r * t) + pv * r; if (!b || a / b <= 0) fail(E.NUM); return Math.log(a / b) / Math.log(1 + r); };
const ipmtOf = (r, per, n, pv, fv, t) => { per = Math.trunc(per); if (per < 1 || per > n) fail(E.NUM); const p = pmtOf(r, n, pv, fv, t); const bal = per === 1 ? (t ? 0 : -pv) : t ? fvOf(r, per - 2, p, pv, 1) - p : fvOf(r, per - 1, p, pv, 0); return bal * r; };
const fin = (a, ev, n) => Array.from({ length: n }, (_, i) => has(a, i) ? num(ev(a[i])) : 0); // positional numeric args, missing → 0
// Newton's method with a numeric derivative; a few restarts from other guesses before giving up with #NUM!.
function solve(f, guess) {
  const F = x => { try { const y = f(x); return isFinite(y) ? y : NaN; } catch (e) { if (e instanceof FErr) return NaN; throw e; } };
  for (const g of [guess, 0.1, 0.01, -0.5, 0.5, 1]) {
    let x = g;
    for (let i = 0; i < 100; i++) {
      const y = F(x); if (isNaN(y)) break; if (Math.abs(y) < 1e-10) return x;
      const h = Math.abs(x) * 1e-6 + 1e-8, d = (F(x + h) - y) / h; if (!d || isNaN(d)) break;
      const nx = x - y / d; if (Math.abs(nx - x) < 1e-12) return nx; x = nx;
      if (i === 99 && Math.abs(F(x)) < 1e-6) return x;
    }
  }
  fail(E.NUM);
}
const cashflows = (a, ev, from) => flat(a.slice(from), ev).map(({ v, ref }) => { if (isErr(v)) fail(v.err); return ref ? (typeof v === 'number' ? v : null) : num(v); }).filter(x => x !== null);
const npvOf = (r, vs, from = 1) => { if (r <= -1) fail(E.NUM); return vs.reduce((s, v, i) => s + v / Math.pow(1 + r, i + from), 0); };
const xnpvOf = (r, vs, ds) => { if (r <= -1) fail(E.NUM); const d0 = ds[0]; return vs.reduce((s, v, i) => s + v / Math.pow(1 + r, (ds[i] - d0) / 365), 0); };
const dbOf = (cost, sal, life, per, month) => {
  if (cost < 0 || sal < 0 || life <= 0 || per <= 0 || month < 1 || month > 12 || per > life + 1 || (month === 12 && per > life)) fail(E.NUM); if (!cost) return 0;
  const rate = Math.round((1 - Math.pow(sal / cost, 1 / life)) * 1000) / 1000; let total = cost * rate * month / 12; if (per === 1) return total;
  for (let i = 2; i < per; i++) total += (cost - total) * rate;
  return per > life ? (cost - total) * rate * (12 - month) / 12 : (cost - total) * rate;
};
const ddbOf = (cost, sal, life, per, f) => { if (cost < 0 || sal < 0 || life <= 0 || per <= 0 || f <= 0 || per > life) fail(E.NUM); const rate = Math.min(f / life, 1); let book = cost, dep = 0; for (let i = 1; i <= Math.ceil(per); i++) { dep = Math.min(book * rate, Math.max(0, book - sal)); book -= dep; } return dep; };
const cumOf = (r, n, pv, s, e, t, principal) => { if (r <= 0 || n <= 0 || pv <= 0 || s < 1 || e < s || e > n || (t !== 0 && t !== 1)) fail(E.NUM); let tot = 0; for (let i = s; i <= e; i++) { const ip = ipmtOf(r, i, n, pv, 0, t); tot += principal ? pmtOf(r, n, pv, 0, t) - ip : ip; } return tot; };
Object.assign(FN, {
  // 财务
  PMT: (a, ev) => { const [r, n, pv, fv, t] = fin(a, ev, 5); if (a.length < 3) fail(E.VALUE); return pmtOf(r, n, pv, fv, t ? 1 : 0); },
  PV: (a, ev) => { const [r, n, p, fv, t] = fin(a, ev, 5); if (a.length < 3) fail(E.VALUE); return pvOf(r, n, p, fv, t ? 1 : 0); },
  FV: (a, ev) => { const [r, n, p, pv, t] = fin(a, ev, 5); if (a.length < 3) fail(E.VALUE); return fvOf(r, n, p, pv, t ? 1 : 0); },
  NPER: (a, ev) => { const [r, p, pv, fv, t] = fin(a, ev, 5); if (a.length < 3) fail(E.VALUE); return nperOf(r, p, pv, fv, t ? 1 : 0); },
  RATE: (a, ev) => { const [n, p, pv, fv, t, g] = fin(a, ev, 6); if (a.length < 3) fail(E.VALUE); const ty = t ? 1 : 0; return solve(r => r === 0 ? pv + p * n + fv : pv * Math.pow(1 + r, n) + p * (1 + r * ty) * (Math.pow(1 + r, n) - 1) / r + fv, has(a, 5) ? g : 0.1); },
  IPMT: (a, ev) => { const [r, per, n, pv, fv, t] = fin(a, ev, 6); if (a.length < 4) fail(E.VALUE); return ipmtOf(r, per, n, pv, fv, t ? 1 : 0); },
  PPMT: (a, ev) => { const [r, per, n, pv, fv, t] = fin(a, ev, 6); if (a.length < 4) fail(E.VALUE); const ty = t ? 1 : 0; return pmtOf(r, n, pv, fv, ty) - ipmtOf(r, per, n, pv, fv, ty); },
  NPV: (a, ev) => npvOf(num(ev(a[0])), cashflows(a, ev, 1)),
  IRR: (a, ev) => { const vs = cashflows([a[0]], ev, 0); if (!vs.some(v => v > 0) || !vs.some(v => v < 0)) fail(E.NUM); return solve(r => npvOf(r, vs, 0), has(a, 1) ? num(ev(a[1])) : 0.1); },
  XNPV: (a, ev) => { const vs = cashflows([a[1]], ev, 0), ds = cashflows([a[2]], ev, 0).map(Math.floor); if (vs.length !== ds.length || !vs.length || ds.some(d => d < ds[0])) fail(E.NUM); return xnpvOf(num(ev(a[0])), vs, ds); },
  XIRR: (a, ev) => { const vs = cashflows([a[0]], ev, 0), ds = cashflows([a[1]], ev, 0).map(Math.floor); if (vs.length !== ds.length || !vs.length || ds.some(d => d < ds[0]) || !vs.some(v => v > 0) || !vs.some(v => v < 0)) fail(E.NUM); return solve(r => xnpvOf(r, vs, ds), has(a, 2) ? num(ev(a[2])) : 0.1); },
  SLN: S((c, s, l) => { l = num(l); if (!l) fail(E.DIV0); return (num(c) - num(s)) / l; }, 3),
  DB: S((c, s, l, p, m) => dbOf(num(c), num(s), num(l), Math.trunc(num(p)), Math.trunc(oNum(m, 12))), 4),
  DDB: S((c, s, l, p, f) => ddbOf(num(c), num(s), num(l), num(p), oNum(f, 2)), 4),
  SYD: S((c, s, l, p) => { c = num(c); s = num(s); l = num(l); p = num(p); if (l <= 0 || p < 1 || p > l) fail(E.NUM); return (c - s) * (l - p + 1) * 2 / (l * (l + 1)); }, 4),
  EFFECT: S((r, n) => { r = num(r); n = Math.trunc(num(n)); if (r <= 0 || n < 1) fail(E.NUM); return Math.pow(1 + r / n, n) - 1; }, 2),
  NOMINAL: S((r, n) => { r = num(r); n = Math.trunc(num(n)); if (r <= 0 || n < 1) fail(E.NUM); return n * (Math.pow(1 + r, 1 / n) - 1); }, 2),
  CUMIPMT: S((r, n, pv, s, e, t) => cumOf(num(r), num(n), num(pv), Math.trunc(num(s)), Math.trunc(num(e)), num(t), false), 6),
  CUMPRINC: S((r, n, pv, s, e, t) => cumOf(num(r), num(n), num(pv), Math.trunc(num(s)), Math.trunc(num(e)), num(t), true), 6),
});
// ---- engineering ----
// Base conversion with Excel's two's-complement widths: 10 binary / 30 octal / 40 hex bits, at most 10 output characters.
const toBase = (n, base, bits, places) => {
  n = Math.trunc(num(n)); const lim = Math.pow(2, bits - 1); if (n < -lim || n >= lim) fail(E.NUM);
  let s = (n < 0 ? n + Math.pow(2, bits) : n).toString(base).toUpperCase();
  if (places !== undefined) { places = Math.trunc(num(places)); if (places < 0 || places > 10 || s.length > places) fail(E.NUM); s = s.padStart(places, '0'); }
  return s;
};
const fromBase = (v, base, bits) => { const s = str(v).trim().toUpperCase(); if (!({ 2: /^[01]{1,10}$/, 8: /^[0-7]{1,10}$/, 16: /^[0-9A-F]{1,10}$/ })[base].test(s)) fail(E.NUM); const n = parseInt(s, base); return s.length === 10 && n >= Math.pow(2, bits - 1) ? n - Math.pow(2, bits) : n; };
const UNITS = { // category → unit → factor to the category's base unit (g, m, s, Pa, N, J, W, T, m3, m2, bit, m/s)
  weight: { g: 1, sg: 14593.9029372064, lbm: 453.59237, u: 1.66053886282828e-24, ozm: 28.349523125, grain: 0.06479891, cwt: 45359.237, shweight: 45359.237, uk_cwt: 50802.34544, lcwt: 50802.34544, hweight: 50802.34544, stone: 6350.29318, ton: 907184.74, brton: 1016046.9088 },
  distance: { m: 1, mi: 1609.344, Nmi: 1852, in: 0.0254, ft: 0.3048, yd: 0.9144, ang: 1e-10, ell: 1.143, ly: 9.46073047258e15, parsec: 3.08567758128e16, pc: 3.08567758128e16, Pica: 0.000352777777777778, pica: 0.000352777777777778, survey_mi: 1609.34721869444 },
  time: { yr: 31557600, day: 86400, d: 86400, hr: 3600, mn: 60, min: 60, sec: 1, s: 1 },
  pressure: { Pa: 1, p: 1, atm: 101325, at: 101325, mmHg: 133.322, psi: 6894.75729316836, Torr: 133.322368421053 },
  force: { N: 1, dyn: 1e-5, dy: 1e-5, lbf: 4.4482216152605, pond: 0.00980665 },
  energy: { J: 1, e: 1e-7, c: 4.184, cal: 4.1868, eV: 1.602176462e-19, ev: 1.602176462e-19, HPh: 2684519.5376962, hh: 2684519.5376962, Wh: 3600, wh: 3600, flb: 1.3558179483314, BTU: 1055.05585262, btu: 1055.05585262 },
  power: { W: 1, w: 1, HP: 745.69987158227, h: 745.69987158227, PS: 735.49875 },
  magnetism: { T: 1, ga: 1e-4 },
  volume: { l: 1e-3, L: 1e-3, lt: 1e-3, tsp: 4.92892159375e-6, tspm: 5e-6, tbs: 1.478676478125e-5, oz: 2.95735295625e-5, cup: 2.365882365e-4, pt: 4.73176473e-4, us_pt: 4.73176473e-4, uk_pt: 5.6826125e-4, qt: 9.46352946e-4, uk_qt: 1.1365225e-3, gal: 3.785411784e-3, uk_gal: 4.54609e-3, ang3: 1e-30, barrel: 0.158987294928, bushel: 0.03523907016688, ft3: 0.028316846592, in3: 1.6387064e-5, ly3: 8.46786664623715e47, m3: 1, mi3: 4168181825.44058, yd3: 0.764554857984, Nmi3: 6352182208, Pica3: 4.39039566186557e-11, GRT: 2.8316846592, regton: 2.8316846592, MTON: 1.13267386368 },
  area: { m2: 1, uk_acre: 4046.8564224, us_acre: 4046.87260987425, ang2: 1e-20, ar: 100, ft2: 0.09290304, ha: 1e4, in2: 6.4516e-4, ly2: 8.95054210748189e31, Morgen: 2500, mi2: 2589988.110336, Nmi2: 3429904, Pica2: 1.24452160493827e-7, yd2: 0.83612736 },
  information: { bit: 1, byte: 8 },
  speed: { 'm/s': 1, 'm/sec': 1, 'm/h': 1 / 3600, 'm/hr': 1 / 3600, mph: 0.44704, kn: 0.514444444444444, admkn: 0.514773333333333 },
};
const PREFIXABLE = new Set(['g', 'u', 'm', 'ang', 'ly', 'parsec', 'pc', 'sec', 's', 'Pa', 'p', 'atm', 'at', 'mmHg', 'N', 'dyn', 'dy', 'pond', 'J', 'e', 'c', 'cal', 'eV', 'ev', 'Wh', 'wh', 'W', 'w', 'T', 'ga', 'l', 'L', 'lt', 'ang3', 'm3', 'ang2', 'm2', 'bit', 'byte', 'm/s', 'm/sec', 'm/h', 'm/hr']);
const PREFIX = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, e: 10, d: 0.1, c: 0.01, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 };
const BINPREFIX = { ki: 2 ** 10, Mi: 2 ** 20, Gi: 2 ** 30, Ti: 2 ** 40, Pi: 2 ** 50, Ei: 2 ** 60, Zi: 2 ** 70, Yi: 2 ** 80 };
const TEMP = { C: [1, 273.15], cel: [1, 273.15], K: [1, 0], kel: [1, 0], F: [5 / 9, 459.67 * 5 / 9], fah: [5 / 9, 459.67 * 5 / 9], Rank: [5 / 9, 0], Reau: [1.25, 273.15] }; // unit → kelvin = v * a + b
const own = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
function unitOf(u) {
  for (const cat in UNITS) if (own(UNITS[cat], u)) return { cat, k: UNITS[cat][u] };
  for (const pre in Object.assign({}, PREFIX, BINPREFIX)) if (u.startsWith(pre) && u.length > pre.length) {
    const base = u.slice(pre.length), bin = own(BINPREFIX, pre); if (!PREFIXABLE.has(base) || (bin && base !== 'bit' && base !== 'byte')) continue;
    for (const cat in UNITS) if (own(UNITS[cat], base)) return { cat, k: UNITS[cat][base] * (bin ? BINPREFIX[pre] : PREFIX[pre]) };
  }
  return null;
}
Object.assign(FN, {
  // 工程
  DEC2BIN: S((n, p) => toBase(n, 2, 10, p), 1), BIN2DEC: S(v => fromBase(v, 2, 10), 1),
  DEC2OCT: S((n, p) => toBase(n, 8, 30, p), 1), OCT2DEC: S(v => fromBase(v, 8, 30), 1),
  DEC2HEX: S((n, p) => toBase(n, 16, 40, p), 1), HEX2DEC: S(v => fromBase(v, 16, 40), 1),
  CONVERT: S((v, from, to) => {
    v = num(v); from = str(from); to = str(to);
    if (own(TEMP, from) && own(TEMP, to)) { const [a, b] = TEMP[from], [c, d] = TEMP[to]; return P15(((v * a + b) - d) / c); }
    const f = unitOf(from), t = unitOf(to); if (!f || !t || f.cat !== t.cat) fail(E.NA); return P15(v * f.k / t.k);
  }, 3),
  DELTA: S((a, b) => P15(num(a)) === P15(oNum(b, 0)) ? 1 : 0, 1),
});
export const FUNCS = Object.keys(FN);
export const REF_SRC = REFSRC; // the reference grammar, for editors that colour or point at references
// Groups for the function picker: [group, [[name, 中文说明], …]].
export const FUNC_GROUPS = [
  ['数学', [['SUM', '求和'], ['SUMIF', '条件求和'], ['SUMIFS', '多条件求和'], ['SUMPRODUCT', '乘积之和'], ['PRODUCT', '乘积'], ['ABS', '绝对值'], ['ROUND', '四舍五入'], ['ROUNDUP', '向上舍入'], ['ROUNDDOWN', '向下舍入'], ['INT', '向下取整'], ['TRUNC', '截尾取整'], ['MOD', '求余数'], ['POWER', '乘幂'], ['SQRT', '平方根'], ['EXP', 'e 的乘幂'], ['LN', '自然对数'], ['LOG', '指定底数的对数'], ['LOG10', '常用对数'], ['CEILING', '向上舍入到倍数'], ['CEILING.MATH', '向上舍入到倍数'], ['FLOOR', '向下舍入到倍数'], ['FLOOR.MATH', '向下舍入到倍数'], ['MROUND', '舍入到最近倍数'], ['SIGN', '正负号'], ['RAND', '0–1 随机数'], ['RANDBETWEEN', '区间随机整数'], ['PI', '圆周率'], ['FACT', '阶乘'], ['COMBIN', '组合数'], ['GCD', '最大公约数'], ['LCM', '最小公倍数'], ['QUOTIENT', '整除的商'], ['EVEN', '舍入到偶数'], ['ODD', '舍入到奇数'], ['SUBTOTAL', '分类汇总'], ['AGGREGATE', '聚合（可忽略错误）']]],
  ['统计', [['AVERAGE', '平均值'], ['AVERAGEA', '平均值（含文本和逻辑值）'], ['AVERAGEIF', '条件平均值'], ['AVERAGEIFS', '多条件平均值'], ['COUNT', '数值个数'], ['COUNTA', '非空个数'], ['COUNTBLANK', '空白个数'], ['COUNTIF', '条件计数'], ['COUNTIFS', '多条件计数'], ['MAX', '最大值'], ['MAXA', '最大值（含文本和逻辑值）'], ['MIN', '最小值'], ['MINA', '最小值（含文本和逻辑值）'], ['MAXIFS', '条件最大值'], ['MINIFS', '条件最小值'], ['MEDIAN', '中位数'], ['MODE', '众数'], ['MODE.SNGL', '众数'], ['STDEV', '样本标准差'], ['STDEV.S', '样本标准差'], ['STDEV.P', '总体标准差'], ['STDEVP', '总体标准差'], ['VAR', '样本方差'], ['VAR.S', '样本方差'], ['VAR.P', '总体方差'], ['VARP', '总体方差'], ['LARGE', '第 k 大的值'], ['SMALL', '第 k 小的值'], ['RANK', '排名'], ['RANK.EQ', '排名'], ['RANK.AVG', '平均排名'], ['PERCENTILE', '百分位数'], ['PERCENTILE.INC', '百分位数（含端点）'], ['PERCENTILE.EXC', '百分位数（不含端点）'], ['QUARTILE', '四分位数'], ['QUARTILE.INC', '四分位数（含端点）'], ['QUARTILE.EXC', '四分位数（不含端点）'], ['CORREL', '相关系数'], ['PEARSON', '皮尔逊相关系数'], ['SLOPE', '回归直线斜率'], ['INTERCEPT', '回归直线截距'], ['FORECAST', '线性预测'], ['FORECAST.LINEAR', '线性预测'], ['TREND', '线性趋势值'], ['GROWTH', '指数趋势值'], ['FREQUENCY', '频率分布'], ['GEOMEAN', '几何平均值'], ['HARMEAN', '调和平均值'], ['NORM.DIST', '正态分布'], ['NORM.INV', '正态分布反函数'], ['NORM.S.DIST', '标准正态分布'], ['NORM.S.INV', '标准正态分布反函数']]],
  ['文本', [['LEN', '字符数'], ['LENB', '字节数'], ['LEFT', '左侧字符'], ['LEFTB', '左侧字节'], ['RIGHT', '右侧字符'], ['RIGHTB', '右侧字节'], ['MID', '中间字符'], ['MIDB', '中间字节'], ['UPPER', '转大写'], ['LOWER', '转小写'], ['PROPER', '单词首字母大写'], ['TRIM', '删除多余空格'], ['CLEAN', '删除控制字符'], ['CONCAT', '连接文本'], ['CONCATENATE', '连接文本'], ['TEXTJOIN', '用分隔符连接'], ['TEXT', '按格式转为文本'], ['VALUE', '文本转数值'], ['NUMBERVALUE', '按分隔符转数值'], ['FIND', '查找位置（区分大小写）'], ['FINDB', '查找字节位置'], ['SEARCH', '查找位置（支持通配符）'], ['REPLACE', '按位置替换'], ['SUBSTITUTE', '替换指定文本'], ['REPT', '重复文本'], ['EXACT', '是否完全相同'], ['CHAR', '编码转字符'], ['CODE', '字符转编码'], ['UNICHAR', 'Unicode 转字符'], ['UNICODE', '字符转 Unicode'], ['T', '仅保留文本'], ['N', '转为数值'], ['FIXED', '固定小数位的文本'], ['DOLLAR', '美元格式文本'], ['RMB', '人民币格式文本'], ['TEXTBEFORE', '分隔符之前的文本'], ['TEXTAFTER', '分隔符之后的文本'], ['TEXTSPLIT', '拆分文本']]],
  ['逻辑', [['IF', '条件判断'], ['IFS', '多条件判断'], ['IFERROR', '出错时返回指定值'], ['IFNA', '#N/A 时返回指定值'], ['AND', '全部为真'], ['OR', '任一为真'], ['NOT', '取反'], ['XOR', '异或'], ['TRUE', '逻辑值 TRUE'], ['FALSE', '逻辑值 FALSE'], ['SWITCH', '按值匹配返回'], ['CHOOSE', '按序号选择'], ['LET', '定义变量']]],
  ['查找引用', [['VLOOKUP', '纵向查找'], ['HLOOKUP', '横向查找'], ['XLOOKUP', '查找（可指定未找到值）'], ['LOOKUP', '向量查找'], ['INDEX', '按行列取值'], ['MATCH', '查找位置'], ['XMATCH', '查找位置'], ['OFFSET', '偏移引用'], ['INDIRECT', '文本转引用'], ['ROW', '行号'], ['ROWS', '行数'], ['COLUMN', '列号'], ['COLUMNS', '列数'], ['ADDRESS', '单元格地址文本'], ['AREAS', '区域个数'], ['TRANSPOSE', '转置'], ['UNIQUE', '去重'], ['FILTER', '按条件筛选'], ['SORT', '排序'], ['SORTBY', '按指定列排序'], ['SEQUENCE', '生成序列']]],
  ['日期时间', [['TODAY', '今天的日期'], ['NOW', '当前日期时间'], ['DATE', '构造日期'], ['TIME', '构造时间'], ['YEAR', '年'], ['MONTH', '月'], ['DAY', '日'], ['HOUR', '时'], ['MINUTE', '分'], ['SECOND', '秒'], ['WEEKDAY', '星期几'], ['WEEKNUM', '第几周'], ['ISOWEEKNUM', 'ISO 周数'], ['DATEDIF', '两日期之差'], ['DATEVALUE', '文本转日期'], ['TIMEVALUE', '文本转时间'], ['EDATE', '若干月后的日期'], ['EOMONTH', '若干月后的月末'], ['DAYS', '相差天数'], ['DAYS360', '按 360 天计的天数'], ['NETWORKDAYS', '工作日天数'], ['NETWORKDAYS.INTL', '工作日天数（自定义周末）'], ['WORKDAY', '若干工作日后'], ['WORKDAY.INTL', '若干工作日后（自定义周末）'], ['YEARFRAC', '相差的年分数']]],
  ['信息', [['ISBLANK', '是否为空'], ['ISERROR', '是否为错误'], ['ISERR', '是否为 #N/A 以外的错误'], ['ISNA', '是否为 #N/A'], ['ISNUMBER', '是否为数值'], ['ISTEXT', '是否为文本'], ['ISNONTEXT', '是否非文本'], ['ISLOGICAL', '是否为逻辑值'], ['ISEVEN', '是否为偶数'], ['ISODD', '是否为奇数'], ['ISFORMULA', '是否含公式'], ['ISREF', '是否为引用'], ['TYPE', '值的类型'], ['NA', '返回 #N/A'], ['ERROR.TYPE', '错误类型编号'], ['CELL', '单元格信息'], ['INFO', '运行环境信息']]],
  ['财务', [['PMT', '每期付款额'], ['PV', '现值'], ['FV', '终值'], ['NPER', '付款期数'], ['RATE', '每期利率'], ['IPMT', '某期的利息'], ['PPMT', '某期的本金'], ['NPV', '净现值'], ['IRR', '内部收益率'], ['XNPV', '不定期净现值'], ['XIRR', '不定期内部收益率'], ['SLN', '直线折旧'], ['DB', '固定余额递减折旧'], ['DDB', '双倍余额递减折旧'], ['SYD', '年数总和折旧'], ['EFFECT', '实际年利率'], ['NOMINAL', '名义年利率'], ['CUMIPMT', '累计利息'], ['CUMPRINC', '累计本金']]],
  ['工程', [['DEC2BIN', '十进制转二进制'], ['BIN2DEC', '二进制转十进制'], ['DEC2HEX', '十进制转十六进制'], ['HEX2DEC', '十六进制转十进制'], ['DEC2OCT', '十进制转八进制'], ['OCT2DEC', '八进制转十进制'], ['CONVERT', '单位换算'], ['DELTA', '两值是否相等']]]
];

export class Calc {
  constructor(doc) { this.doc = doc; this.cache = new Map(); this.stack = new Set(); this.scope = new Map(); this.cur = { r: 0, c: 0 }; }
  sheetIdx(name, cur) { if (name == null) return cur; let i = this.doc.sheets.findIndex(s => s.name === name); if (i < 0) i = this.doc.sheets.findIndex(s => sameSheet(s.name, name)); if (i < 0) fail(E.REF); return i; }
  nameNode(name) {
    const raw = this.doc.names && (this.doc.names[name] ?? this.doc.names[name.toUpperCase()]);
    if (typeof raw !== 'string') return null;
    const ref = raw.replace(/^=/, ''); return REFSTR.test(ref) ? refNode(ref) : ast(ref);
  }
  bounds(n, si) {
    if (n.t === 'name') { const named = this.nameNode(n.v); if (!named) fail(E.NAME); return this.bounds(named, si); }
    if (n.t !== 'ref' && n.t !== 'rng') fail(E.VALUE);
    const s2 = this.sheetIdx(n.sh, si), open = n.t === 'rng' && (n.r2 == null || n.c2 == null);
    const u = open ? (this.used || (this.used = new Map())).get(s2) ?? this.used.set(s2, usedRange(this.doc.sheets[s2])).get(s2) : null; // only whole rows/columns need the used range; one scan per sheet per Calc
    const r1 = n.t === 'ref' ? n.r : n.r1, c1 = n.t === 'ref' ? n.c : n.c1;
    const r2 = n.t === 'ref' ? n.r : n.r2 == null ? Math.max(NR - 1, u?.r2 ?? 0) : n.r2;
    const c2 = n.t === 'ref' ? n.c : n.c2 == null ? Math.max(NC - 1, u?.c2 ?? 0) : n.c2;
    if (r1 < 0 || c1 < 0 || r2 < r1 || c2 < c1 || (r2 - r1 + 1) * (c2 - c1 + 1) > MAXCELLS) fail(E.REF);
    return { si: s2, r1, c1, r2, c2 };
  }
  value(si, r, c) {
    const key = `${si}:${r}:${c}`; if (this.cache.has(key)) return this.cache.get(key);
    const sh = this.doc.sheets[si], cell = sh && sh.cells[A(r, c)]; let v = '';
    if (cell && cell.v !== '' && cell.v != null) {
      const raw = String(cell.v);
      if (raw[0] === '=' && raw.length > 1) {
        if (this.stack.has(key)) return { err: E.CIRC };
        this.stack.add(key); const prior = this.cur; this.cur = { r, c };
        // A formula whose result is a blank cell shows 0 (=A1 with A1 empty); a text result stays "" (=MID(...), =A1&"").
        // ponytail: IF/CHOOSE returning a blank reference show "" rather than Excel's 0, reference semantics would be needed.
        try { const root = ast(raw.slice(1)); v = this.ev(root, si); if (isRng(v)) v = v.range[0]?.[0] ?? ''; if (v == null || (v === '' && (root.t === 'ref' || root.t === 'rng' || root.t === 'name'))) v = 0; if (typeof v === 'number' && !isFinite(v)) v = { err: E.NUM }; }
        catch (e) { v = { err: e instanceof FErr ? e.e : E.VALUE }; }
        finally { this.cur = prior; this.stack.delete(key); }
      } else if (cell.s && cell.s.fmt === 'text') v = raw;
      else if (cell.s && /^(date|time|datetime)$/.test(cell.s.fmt) && isNaN(Number(raw))) v = parseDate(raw) ?? raw; // "2024-01-01" typed as a date is its serial
      else { const n = Number(raw.replace(/,/g, '')); v = raw.trim() !== '' && !isNaN(n) ? n : /^(true|false)$/i.test(raw) ? raw.toUpperCase() === 'TRUE' : raw; }
    }
    this.cache.set(key, v); return v;
  }
  ev(n, si) {
    const ev = x => this.ev(x, si);
    switch (n.t) {
      case 'err': fail(n.e);
      case 'empty': return undefined;
      case 'n': case 's': case 'b': return n.v;
      case 'name': {
        if (this.scope.has(n.v)) return this.scope.get(n.v);
        const named = this.nameNode(n.v); if (!named) fail(E.NAME); return this.ev(named, si);
      }
      case 'ref': { const v = this.value(this.sheetIdx(n.sh, si), n.r, n.c); if (isErr(v)) fail(v.err); return v; }
      case 'rng': { const b = this.bounds(n, si); return { range: grid(b.r2 - b.r1 + 1, b.c2 - b.c1 + 1, (i, j) => this.value(b.si, b.r1 + i, b.c1 + j)) }; }
      case 'arr': return { range: n.v.map(row => row.map(x => { try { return sc(ev(x)); } catch (e) { return { err: e instanceof FErr ? e.e : E.VALUE }; } })) };
      case 'neg': return lift2([ev(n.a)], x => -num(x));
      case 'pct': return lift2([ev(n.a)], x => num(x) / 100);
      case 'op': {
        const a = ev(n.a), b = ev(n.b);
        return lift2([a, b], (x, y) => {
          switch (n.op) {
            case '+': return P15(num(x) + num(y)); case '-': return P15(num(x) - num(y)); case '*': return P15(num(x) * num(y));
            case '/': { const d = num(y); if (d === 0) fail(E.DIV0); return P15(num(x) / d); }
            case '^': { const v = Math.pow(num(x), num(y)); if (!isFinite(v)) fail(E.NUM); return P15(v); }
            case '&': return str(x) + str(y);
            default: return cmpv(n.op, x, y);
          }
        });
      }
      case 'fn': { const f = FN[n.name]; if (!f) fail(E.NAME); return f(n.args, ev, this, si); }
    }
    fail(E.VALUE);
  }
}

// References inside string literals and function names must not be rewritten. Cells: A1 / $A$1; whole columns A:C and whole rows 3:5 in a second pass.
const SHEETRE = "((?:'[^']+'|[A-Za-z_\\u4e00-\\u9fa5][\\w\\u4e00-\\u9fa5]*)!)?";
const REFRE = new RegExp('("(?:[^"]|"")*")|' + SHEETRE + '(\\$?)([A-Za-z]{1,2})(\\$?)(\\d+)(?![\\w(])', 'g');
const WHOLERE = new RegExp('("(?:[^"]|"")*")|' + SHEETRE + '(?:(\\$?)([A-Za-z]{1,2}):(\\$?)([A-Za-z]{1,2})|(\\$?)(\\d+):(\\$?)(\\d+))(?![\\w(:])', 'g');
const nameChar = ch => ch && /[A-Za-z0-9_.\u4e00-\u9fa5]/.test(ch);
// Rewrites every reference: cells through `cell(sh, abs1, col, abs2, row)`, whole columns/rows through `whole(sh, axis, abs1, i1, abs2, i2)` (0-based indexes).
function mapRefs(f, cell, whole) {
  const out = f.replace(REFRE, (m, s, sh, d1, col, d2, row, off, full) => s || (!sh && nameChar(full[off - 1])) ? m : cell(sh || '', d1, colIdx(col), d2, +row - 1));
  return out.replace(WHOLERE, (m, s, sh, c1, ca, c2, cb, r1, ra, r2, rb, off, full) => {
    if (s || (!sh && nameChar(full[off - 1]))) return m;
    return ca !== undefined ? whole(sh || '', 'c', c1, colIdx(ca), c2, colIdx(cb)) : whole(sh || '', 'r', r1, +ra - 1, r2, +rb - 1);
  });
}
const wholeStr = (sh, axis, d1, i1, d2, i2) => sh + d1 + (axis === 'c' ? colName(i1) : i1 + 1) + ':' + d2 + (axis === 'c' ? colName(i2) : i2 + 1);
export function shiftF(f, dr, dc) {
  return mapRefs(f,
    (sh, d1, c, d2, r) => { if (!d1) c += dc; if (!d2) r += dr; return c < 0 || r < 0 ? E.REF : sh + d1 + colName(c) + d2 + (r + 1); },
    (sh, axis, d1, i1, d2, i2) => { const d = axis === 'c' ? dc : dr; if (!d1) i1 += d; if (!d2) i2 += d; return i1 < 0 || i2 < 0 ? E.REF : wholeStr(sh, axis, d1, i1, d2, i2); });
}
export function adjF(f, curName, target, axis, at, n) {
  const onTarget = sh => sameSheet(sh ? sh.slice(0, -1).replace(/^'|'$/g, '') : curName, target);
  const move = idx => { if (n > 0) return idx >= at ? idx + n : idx; if (idx >= at && idx < at - n) return null; return idx >= at - n ? idx + n : idx; }; // null: deleted
  return mapRefs(f,
    (sh, d1, c, d2, r) => { if (!onTarget(sh)) return sh + d1 + colName(c) + d2 + (r + 1); const i = move(axis === 'r' ? r : c); if (i == null) return E.REF; return sh + d1 + colName(axis === 'r' ? c : i) + d2 + ((axis === 'r' ? i : r) + 1); },
    (sh, ax, d1, i1, d2, i2) => { if (!onTarget(sh) || ax !== axis) return wholeStr(sh, ax, d1, i1, d2, i2); const a = move(i1), b = move(i2); return a == null && b == null ? E.REF : wholeStr(sh, ax, d1, a ?? at, d2, b ?? at - 1); });
}
// Upper-cases function names and references as typed, leaving string literals and sheet names ('My Sheet'!, Other!) as written.
export function upperF(f) { return f.split(/("(?:[^"]|"")*"|'[^']*'!|[A-Za-z_一-龥][\w.一-龥]*!)/).map((p, i) => i % 2 ? p : p.toUpperCase()).join(''); }
const sameSheet = (a, b) => a === b || (a != null && b != null && a.toLowerCase() === b.toLowerCase()); // sheet names compare case-insensitively, as in Excel
export function usedRange(sh) {
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  Object.keys(sh.cells).forEach(a => { const cell = sh.cells[a]; if (!cell || cell.v === '' || cell.v == null) return; const p = parseA(a); if (!p) return; r1 = Math.min(r1, p.r); c1 = Math.min(c1, p.c); r2 = Math.max(r2, p.r); c2 = Math.max(c2, p.c); });
  return r2 < 0 ? null : { r1, c1, r2, c2 };
}
