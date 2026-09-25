// mindmap.js — pure mind-map helpers shared by MindMapEditor and engine.js: tree walking, auto layout,
// the save planner (model diff → writer commands) and change marking. No DOM: runs in node for tests.
// Model: { id, text, children: [...], collapsed?, side?: 'left'|'right', note?, link?, color?: 'RRGGBB', fill?: 'RRGGBB', icon? }

export const ICONS = [['idea', '💡'], ['flag', '🚩'], ['button_ok', '✅'], ['button_cancel', '❌'], ['help', '❓'], ['info', 'ℹ️'], ['messagebox_warning', '⚠️'], ['stop', '⛔'], ['bookmark', '🔖'], ['attach', '📎'], ['calendar', '📅'], ['clock', '⏰']];
export const glyph = name => (ICONS.find(i => i[0] === name) || [])[1] || '';
export const BRANCH = ['#3F7D5C', '#2F5D8A', '#B5563A', '#8E5AB8', '#C28A1A', '#4A8F9C'];
export const FONT = "'IBM Plex Sans','Noto Sans SC',sans-serif";
// per depth: font size / weight, padding, line height, max text width, corner radius
const LEVEL = [{ fs: 18, fw: 600, px: 22, py: 12, lh: 26, max: 300, rx: 16 }, { fs: 15, fw: 500, px: 16, py: 9, lh: 22, max: 240, rx: 12 }, { fs: 13.5, fw: 400, px: 12, py: 6, lh: 19, max: 220, rx: 9 }];
const GX = [70, 48, 36], GY = [26, 16, 10]; // gap from a parent at that depth to its children / between its children
export const styleOf = depth => LEVEL[Math.min(depth, 2)];
let seq = 0;
export const newId = () => 'tmp_' + Date.now().toString(36) + '_' + (seq++).toString(36);

export function walk(root, fn, parent = null, depth = 0, index = 0) {
  if (fn(root, parent, depth, index) === false) return;
  (root.children || []).forEach((c, i) => walk(c, fn, root, depth + 1, i));
}
export function find(root, id) { let out = null; walk(root, n => { if (n.id === id) { out = n; return false; } }); return out; }
export function parentOf(root, id) { let out = null; walk(root, (n, p) => { if (n.id === id) out = p; }); return out; }
export function count(root) { let n = 0; walk(root, () => { n++; }); return n; }
export function norm(n) { return { text: n.text || '', note: n.note || '', collapsed: !!n.collapsed, side: n.side || '', link: n.link || '', color: n.color || '', fill: n.fill || '', icon: n.icon || '' }; }
/** id → normalized props + parentId / index / depth, in document order. */
export function flatten(root) { const out = {}; walk(root, (n, p, d, i) => { out[n.id] = Object.assign(norm(n), { parentId: p ? p.id : null, index: i, depth: d }); }); return out; }

// ---------- layout ----------
/** Width estimate for node tests and as a fallback: CJK a full em, everything else a bit over half. */
export const estimate = text => { let w = 0; for (const ch of text) w += ch.codePointAt(0) > 0x2E7F ? 1 : 0.56; return w; };
const TOK = /[\u2E80-\u{10FFFF}]|[^\s\u2E80-\u{10FFFF}]+\s*|\s+/gu;
/** Greedy wrap: CJK breaks per character, Latin per word. measure(text) → width at the given font size. */
export function wrap(text, maxW, measure) {
  const lines = [];
  for (const para of String(text || '').split('\n')) {
    let cur = '';
    for (const t of para.match(TOK) || []) {
      const cand = cur + t;
      if (cur && measure(cand.trimEnd()) > maxW) { lines.push(cur.trimEnd()); cur = t.trimStart(); } else cur = cand;
    }
    lines.push(cur.trimEnd());
  }
  return lines;
}

/** Positions every visible node. measure(text, fs, fw) → px. Returns { nodes, edges, bounds }; nodes carry the model node as .n. */
export function layout(root, measure) {
  const M = measure || ((t, fs) => estimate(t) * fs);
  const nodes = [], edges = [];
  const box = (n, depth) => {
    const s = styleOf(depth), m = t => M(t, s.fs, s.fw);
    const lines = wrap(n.text, s.max, m), tw = lines.reduce((a, l) => Math.max(a, m(l)), 0), iw = n.icon ? s.fs * 1.4 + 4 : 0;
    const kids = n.collapsed ? [] : (n.children || []).map(c => box(c, depth + 1));
    const b = { id: n.id, n, depth, s, lines, iw, w: Math.max(depth ? 44 : 96, Math.ceil(tw + iw + s.px * 2)), h: Math.ceil(lines.length * s.lh + s.py * 2), kids };
    b.badge = n.collapsed && (n.children || []).length ? count(n) - 1 : 0;
    b.sh = kids.length ? Math.max(b.h, kids.reduce((a, k) => a + k.sh, 0) + GY[Math.min(depth, 2)] * (kids.length - 1)) : b.h;
    return b;
  };
  const place = (b, x, cy, dir, color, parentId) => {
    b.x = dir > 0 ? x : x - b.w; b.y = cy - b.h / 2; b.dir = dir; b.color = color; b.parentId = parentId;
    nodes.push(b);
    const gy = GY[Math.min(b.depth, 2)], gx = GX[Math.min(b.depth, 2)];
    let start = cy - (b.kids.reduce((a, k) => a + k.sh, 0) + gy * (b.kids.length - 1)) / 2;
    for (const k of b.kids) {
      place(k, dir > 0 ? b.x + b.w + gx : b.x - gx, start + k.sh / 2, dir, color, b.id);
      const x1 = dir > 0 ? b.x + b.w : b.x, y1 = b.y + b.h / 2, x2 = dir > 0 ? k.x : k.x + k.w, y2 = k.y + k.h / 2, dx = (x2 - x1) / 2;
      edges.push({ from: b.id, to: k.id, color, sw: b.depth ? 1.8 : 2.6, d: `M${x1} ${y1}C${x1 + dx} ${y1},${x2 - dx} ${y2},${x2} ${y2}` });
      start += k.sh + gy;
    }
  };
  const R = box(root, 0);
  R.x = -R.w / 2; R.y = -R.h / 2; R.dir = 1; R.color = '#1D1D1F'; R.parentId = null; nodes.push(R);
  const sides = { right: [], left: [] }; // explicit side wins; the rest go to the emptier side, so a fresh map alternates right/left
  R.kids.forEach((k, i) => { k.color = BRANCH[i % BRANCH.length]; sides[k.n.side === 'left' || k.n.side === 'right' ? k.n.side : (sides.left.length < sides.right.length ? 'left' : 'right')].push(k); });
  for (const side of ['right', 'left']) {
    const list = sides[side], dir = side === 'right' ? 1 : -1, gy = GY[0];
    let start = -(list.reduce((a, k) => a + k.sh, 0) + gy * (list.length - 1)) / 2;
    for (const k of list) {
      place(k, dir > 0 ? R.w / 2 + GX[0] : -R.w / 2 - GX[0], start + k.sh / 2, dir, k.color, root.id);
      const x1 = dir * R.w / 2, x2 = dir > 0 ? k.x : k.x + k.w, y2 = k.y + k.h / 2, dx = (x2 - x1) / 2;
      edges.push({ from: root.id, to: k.id, color: k.color, sw: 2.6, d: `M${x1} 0C${x1 + dx} 0,${x2 - dx} ${y2},${x2} ${y2}` });
      start += k.sh + gy;
    }
  }
  const bounds = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
  for (const b of nodes) { bounds.minX = Math.min(bounds.minX, b.x); bounds.minY = Math.min(bounds.minY, b.y); bounds.maxX = Math.max(bounds.maxX, b.x + b.w); bounds.maxY = Math.max(bounds.maxY, b.y + b.h); }
  return { nodes, edges, bounds };
}

// ---------- save planner ----------
const propArgs = d => Object.entries(d).flatMap(([k, v]) => ['--prop', k + '=' + v]);
/** The props that changed from o (normalized) to c (normalized), as command values. Side only matters right under the root. */
export function diffProps(o, c, depth) {
  const d = {};
  for (const k of ['text', 'note', 'link']) if (o[k] !== c[k]) d[k] = c[k];
  if (o.collapsed !== c.collapsed) d.collapsed = c.collapsed ? 'true' : 'false';
  for (const k of ['color', 'fill', 'icon']) if (o[k] !== c[k]) d[k] = c[k] || 'none';
  if (depth === 1 && c.side && o.side !== c.side) d.side = c.side;
  return d;
}

/** Turns the difference between the flat snapshot `orig` (see flatten) and the model `root` into ordered steps.
 * Each step has argv (resolved lazily, so a child of a new node uses the real id its parent got back) and an optional
 * onResult(json) that records the engine id of an added node in `ids` (old id → real id) and on the node itself.
 * Order: survivors whose original parent goes away are moved out first, then removals (topmost only), then a pre-order
 * pass adding / setting / moving so every parent exists before its children. */
export function plan(orig, root, file) {
  orig = orig || {};
  const ids = {}, steps = [];
  const P = id => '//topic[@id=' + (ids[id] || id) + ']';
  const S = (argv, onResult) => ({ get argv() { return argv(); }, onResult });
  const nw = flatten(root);
  const removed = new Set(Object.keys(orig).filter(id => !nw[id]));
  // the engine's children lists as the commands will leave them: the file as opened, minus what goes away
  const sim = {};
  Object.keys(orig).filter(id => !removed.has(id) && orig[id].parentId && !removed.has(orig[id].parentId)).sort((a, b) => orig[a].index - orig[b].index).forEach(id => (sim[orig[id].parentId] = sim[orig[id].parentId] || []).push(id));
  walk(root, n => {
    const o = orig[n.id]; if (!o || !o.parentId || !removed.has(o.parentId)) return;
    const np = nw[n.id].parentId, dest = orig[np] && !removed.has(np) ? np : root.id;
    steps.push(S(() => ['move', file, P(n.id), '--to', P(dest)]));
    (sim[dest] = sim[dest] || []).push(n.id);
  });
  for (const id of removed) if (!removed.has(orig[id].parentId)) steps.push(S(() => ['remove', file, P(id)]));
  // cursor[parent] = position in the engine's list where the next child in model order must sit; entries that will
  // leave for another parent later are skipped over rather than pushed around
  const cursor = {};
  walk(root, (n, parent, depth) => {
    const o = orig[n.id], c = norm(n);
    if (!parent) { const d = o ? diffProps(o, c, 0) : {}; if (Object.keys(d).length) steps.push(S(() => ['set', file, P(n.id), ...propArgs(d)])); return; }
    const list = sim[parent.id] = sim[parent.id] || [];
    let pos = cursor[parent.id] || 0;
    while (pos < list.length && nw[list[pos]] && nw[list[pos]].parentId !== parent.id) pos++;
    if (!o) {
      steps.push(S(() => ['add', file, P(parent.id), '--type', 'topic', '--prop', 'text=' + c.text, '--index', String(pos + 1)], r => { const real = r && r.props && r.props.id; if (real) { ids[n.id] = real; n.id = real; } }));
      const extra = diffProps(norm({}), c, depth); delete extra.text;
      if (Object.keys(extra).length) steps.push(S(() => ['set', file, P(n.id), ...propArgs(extra)]));
      list.splice(pos, 0, n.id);
    } else {
      const d = diffProps(o, c, depth);
      if (Object.keys(d).length) steps.push(S(() => ['set', file, P(n.id), ...propArgs(d)]));
      if (list[pos] !== n.id) {
        for (const k in sim) { const i = sim[k].indexOf(n.id); if (i >= 0) sim[k].splice(i, 1); }
        list.splice(pos, 0, n.id);
        steps.push(S(() => ['move', file, P(n.id), '--to', P(parent.id), '--index', String(pos + 1)]));
      }
    }
    cursor[parent.id] = pos + 1;
  });
  return { steps, ids };
}

// No $t under node (see ui/tests/mindmap.test.mjs): fall back to the exact Chinese this always produced, context
// dropped like i18n.js's own $t does in Chinese. '删除@@diff' needs its own English ("Deleted N topics") distinct
// from the Delete button MindMapEditor.dc.html translates on its own.
function tr(s, v) {
  const at = s.indexOf('@@'), bare = at < 0 ? s : s.slice(0, at);
  if (globalThis.$t) return globalThis.$t(s, v);
  return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare;
}

/** Marks nodes of `after` that differ from `before` with ai = true and returns [[kind, label], ...] for the change card. */
export function markAi(before, after) {
  const items = [];
  if (!before || !after) return items;
  const b = flatten(before), a = flatten(after);
  const label = n => tr('主题 · {text}', { text: (n.text || '').slice(0, 24) || tr('（空）') });
  walk(after, n => {
    const o = b[n.id]; if (!o) { n.ai = true; items.push([tr('新增'), label(n)]); return; }
    const c = a[n.id];
    if (['text', 'note', 'link', 'collapsed', 'color', 'fill', 'icon'].some(k => o[k] !== c[k])) { n.ai = true; items.push([tr('修改'), label(n)]); }
    else if (o.parentId !== c.parentId) { n.ai = true; items.push([tr('移动'), label(n)]); }
  });
  const gone = Object.keys(b).filter(id => !a[id] && !(b[id].parentId && !a[b[id].parentId])).length; // topmost removed only
  if (gone) items.push([tr('删除@@diff'), tr('{n} 个主题', { n: gone })]);
  return items;
}
export function clearAi(root) { walk(root, n => { delete n.ai; }); }
