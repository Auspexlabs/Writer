// Immutable workbook edits. Cell dictionaries use shared buckets: changing one cell never copies a million keys.
const WIDTH = 1024, META = new WeakMap();
const bucket = key => { let h = 0; for (let i = 0; i < key.length; i++) h = (h * 31 + key.charCodeAt(i)) | 0; return h & (WIDTH - 1); };
function buckets(cells) {
  if (META.has(cells)) return META.get(cells);
  const parts = Array.from({ length: WIDTH }, () => Object.create(null));
  for (const key of Object.keys(cells)) parts[bucket(key)][key] = cells[key];
  return parts;
}
function mapFrom(parts) {
  const out = new Proxy({}, {
    get: (_, key) => typeof key === 'string' ? parts[bucket(key)][key] : undefined,
    has: (_, key) => typeof key === 'string' && Object.hasOwn(parts[bucket(key)], key),
    ownKeys: () => parts.flatMap(p => Object.keys(p)),
    getOwnPropertyDescriptor: (_, key) => typeof key === 'string' && Object.hasOwn(parts[bucket(key)], key) ? { enumerable: true, configurable: true, value: parts[bucket(key)][key], writable: false } : undefined,
    set: () => { throw new TypeError('Edit workbook cells through editWorkbook'); },
    deleteProperty: () => { throw new TypeError('Edit workbook cells through editWorkbook'); }
  });
  META.set(out, parts); return out;
}
export function changedCellKeys(a, b) {
  if (a === b) return [];
  const aa = buckets(a), bb = buckets(b), out = [];
  for (let i = 0; i < WIDTH; i++) if (aa[i] !== bb[i]) for (const k of new Set([...Object.keys(aa[i]), ...Object.keys(bb[i])])) if (aa[i][k] !== bb[i][k]) out.push(k);
  return out;
}
export function shareCells(cells) { return META.has(cells) ? cells : mapFrom(buckets(cells)); }
export function editWorkbook(base, recipe) {
  const states = new WeakMap();
  const draft = (value, cells = false) => {
    const state = { value, cells, writes: new Map(), removed: new Set(), kids: new Map() };
    const current = k => state.writes.has(k) ? state.writes.get(k) : state.removed.has(k) ? undefined : value[k];
    const proxy = new Proxy(Array.isArray(value) ? [] : {}, {
      get(_, k) {
        const v = current(k); if (!v || typeof v !== 'object') return v;
        if (!state.kids.has(k) || state.kids.get(k).value !== v) state.kids.set(k, { value: v, proxy: draft(v, k === 'cells') });
        return state.kids.get(k).proxy;
      },
      set(_, k, v) { state.kids.delete(k); state.removed.delete(k); state.writes.set(k, v); return true; },
      deleteProperty(_, k) { state.kids.delete(k); state.writes.delete(k); state.removed.add(k); return true; },
      has: (_, k) => !state.removed.has(k) && (state.writes.has(k) || k in value),
      ownKeys: () => [...new Set([...Reflect.ownKeys(value), ...state.writes.keys()])].filter(k => !state.removed.has(k)),
      getOwnPropertyDescriptor(_, k) {
        if (Array.isArray(value) && k === 'length') return { value: current(k), writable: true, enumerable: false, configurable: false };
        return !state.removed.has(k) && (state.writes.has(k) || Object.hasOwn(value, k)) ? { value: current(k), writable: true, enumerable: true, configurable: true } : undefined;
      }
    });
    states.set(proxy, state); return proxy;
  };
  const done = proxy => {
    const s = states.get(proxy);
    if (!s) {
      if (!proxy || typeof proxy !== 'object' || META.has(proxy)) return proxy;
      let out = proxy;
      for (const k of Object.keys(proxy)) { const v = done(proxy[k]); if (v !== proxy[k]) { if (out === proxy) out = Array.isArray(proxy) ? proxy.slice() : { ...proxy }; out[k] = v; } }
      return out;
    }
    const changes = new Map(s.writes);
    for (const [k, child] of s.kids) { const result = done(child.proxy); if (result !== child.value) changes.set(k, result); }
    for (const [k, v] of changes) changes.set(k, done(v));
    if (!changes.size && !s.removed.size) return s.value;
    if (s.cells) {
      const parts = buckets(s.value).slice(), copied = new Set();
      const edit = k => { const i = bucket(k); if (!copied.has(i)) { parts[i] = Object.assign(Object.create(null), parts[i]); copied.add(i); } return parts[i]; };
      for (const k of s.removed) delete edit(k)[k];
      for (const [k, v] of changes) edit(k)[k] = v;
      return mapFrom(parts);
    }
    const out = Array.isArray(s.value) ? s.value.slice() : { ...s.value };
    for (const k of s.removed) delete out[k];
    for (const [k, v] of changes) out[k] = v;
    return out;
  };
  const value = draft(base); recipe(value.sheets[Math.min(value.active || 0, value.sheets.length - 1)], value); return done(value);
}
// Copies metadata for native ids written by the save; immutable cells can be shared with history/AI snapshots.
export function cloneWorkbook(doc) {
  const clone = (v, key) => key === 'cells' ? shareCells(v) : Array.isArray(v) ? v.map(x => clone(x)) : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, clone(x, k)])) : v;
  return clone(doc);
}

// Small clipboard/style values may be drafts; structuredClone rejects proxies.
export const cloneValue = v => Array.isArray(v) ? v.map(cloneValue) : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, cloneValue(x)])) : v;
export function isSpillFormula(text) {
  if(typeof text!=='string'||text[0]!=='=')return false;
  // A table's current-row references are scalars, even in a million-row calculated column.
  // They must not make every blank viewport cell pre-evaluate that entire column.
  if(text.includes('[@')||/#This Row/i.test(text)){
    let probe='',start=0,i=0;while(i<text.length){if(text[i]!=='['){i++;continue;}let end=i,depth=0;for(;end<text.length;end++){if(text[end]==="'"&&/[\[\]#@']/.test(text[end+1]||'')){end++;continue;}if(text[end]==='[')depth++;else if(text[end]===']'&&--depth===0)break;}const token=text.slice(i,end+1);probe+=text.slice(start,i)+(/\[@|#This Row/i.test(token)?'1':token);i=end+1;start=i;}text=probe+text.slice(start);
  }
  return /[\[{:;]|\b(?:MAP|SCAN|BYROW|BYCOL|MAKEARRAY|TOCOL|TOROW|WRAPROWS|WRAPCOLS|EXPAND|SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|TRANSPOSE|TEXTSPLIT|HSTACK|VSTACK|TAKE|DROP|CHOOSECOLS|CHOOSEROWS|REGEXEXTRACT|LINEST|LOGEST|TREND|GROWTH|FREQUENCY)\s*\(|^=[A-Za-z_]+$/i.test(text);
}
const PART_STATS = new WeakMap(), MAP_STATS = new WeakMap();
export function cellStats(cells) {
  if (META.has(cells) && MAP_STATS.has(cells)) return MAP_STATS.get(cells);
  const result = { count: 0, rows: 0, cols: 0, used: null, spills: [], wrapped: [] };
  const include = (u, r1, c1, r2, c2) => u ? { r1: Math.min(u.r1, r1), c1: Math.min(u.c1, c1), r2: Math.max(u.r2, r2), c2: Math.max(u.c2, c2) } : { r1, c1, r2, c2 };
  for (const part of META.has(cells) ? META.get(cells) : [cells]) {
    let stat = META.has(cells) && PART_STATS.get(part);
    if (!stat) {
      stat = { count: 0, rows: 0, cols: 0, used: null, spills: [], wrapped: [] };
      for (const a of Object.keys(part)) {
        const m = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(a); if (!m) continue;
        const r = +m[2] - 1; let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; c--;
        if (r >= 1048576 || c >= 16384) continue;
        stat.count++; stat.rows = Math.max(stat.rows, r + 1); stat.cols = Math.max(stat.cols, c + 1);
        const x = part[a]; if (x?.s?.wrap) stat.wrapped.push(a); if (x && !x.literal && isSpillFormula(x.v)) stat.spills.push(a); if (x && x.v !== '' && x.v != null) stat.used = include(stat.used, r, c, r, c);
      }
      if (META.has(cells)) PART_STATS.set(part, stat);
    }
    result.spills.push(...stat.spills); result.wrapped.push(...stat.wrapped); result.count += stat.count; result.rows = Math.max(result.rows, stat.rows); result.cols = Math.max(result.cols, stat.cols);
    if (stat.used) { const u = stat.used; result.used = include(result.used, u.r1, u.c1, u.r2, u.c2); }
  }
  if (META.has(cells)) MAP_STATS.set(cells, result);
  return result;
}
