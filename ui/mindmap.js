// mindmap.js — pure mind-map helpers shared by MindMapEditor and engine.js: tree walking, auto layout, the save planner
// (model diff → writer commands), outlines for the clipboard and change marking. No DOM: runs in node for tests.
// Model: { id, text, children: [...], collapsed?, side?: 'left'|'right', note?, link?, color?: 'RRGGBB', fill?: 'RRGGBB',
//          icon?: 'name,name' (markers and icons, FreeMind builtin names), labels?: 'a,b', image?: data URI, imageSize?: 'w,h',
//          bold?, italic?, strike?, size?: px, font?, free?: 'x,y' }

export const ICONS = [['idea', '💡'], ['button_ok', '✅'], ['button_cancel', '❌'], ['help', '❓'], ['info', 'ℹ️'], ['messagebox_warning', '⚠️'], ['stop', '⛔'], ['bookmark', '🔖'], ['attach', '📎'], ['calendar', '📅'], ['clock', '⏰'], ['pencil', '✏️']];
export const glyph = name => (ICONS.find(i => i[0] === name) || [])[1] || '';
// Markers drawn as shapes, by their FreeMind / Freeplane builtin names (the coloured stars are ours): priority discs
// full-1…full-9, progress pies 0%…100%, flags and stars. Each is drawn in a 16 × 16 box.
const PRIO = ['#E5484D', '#F76B15', '#E5B800', '#3F7D5C', '#2F5D8A', '#8E5AB8', '#8E8E93', '#8E8E93', '#8E8E93'];
const FLAGS = { flag: '#E5484D', 'flag-black': '#1D1D1F', 'flag-blue': '#2F5D8A', 'flag-green': '#3F7D5C', 'flag-orange': '#F76B15', 'flag-pink': '#E0609A', 'flag-yellow': '#E5B800' };
const STARS = { star: '#E5B800', 'star-red': '#E5484D', 'star-green': '#3F7D5C', 'star-blue': '#2F5D8A', 'star-purple': '#8E5AB8' };
const FLAG_D = 'M3 15.5V1.5h1.4v1h9.1l-2.6 3.4 2.6 3.4H4.4v6.2z', STAR_D = 'M8 1.2l2.1 4.4 4.8.6-3.5 3.3.9 4.8L8 12 3.7 14.3l.9-4.8L1.1 6.2l4.8-.6z';
export const MARKERS = {
  priority: PRIO.map((c, i) => 'full-' + (i + 1)), progress: ['0%', '25%', '50%', '75%', '100%'], flags: Object.keys(FLAGS), stars: Object.keys(STARS), icons: ICONS.map(i => i[0])
};
/** The icons and markers of a topic, in order. */
export const iconList = n => String((n && n.icon) || '').split(',').map(s => s.trim()).filter(Boolean);
/** A marker as { path, fill } shapes plus optional text, in 16 × 16 units; null for an emoji icon (see glyph). */
export function marker(name) {
  let m;
  if ((m = /^full-([1-9])$/.exec(name))) return { shapes: [{ d: 'M8 .5a7.5 7.5 0 1 1 0 15a7.5 7.5 0 1 1 0-15z', fill: PRIO[m[1] - 1] }], text: m[1], fg: '#FFFFFF' };
  if ((m = /^(0|25|50|75|100)%$/.exec(name))) {
    const f = +m[1] / 100, a = f * Math.PI * 2, x = 8 + 7 * Math.sin(a), y = 8 - 7 * Math.cos(a), sector = f >= 1 ? 'M8 1a7 7 0 1 1 0 14a7 7 0 1 1 0-14z' : f > 0 ? `M8 8V1A7 7 0 ${f > 0.5 ? 1 : 0} 1 ${x.toFixed(2)} ${y.toFixed(2)}z` : '';
    return { shapes: [{ d: 'M8 .75a7.25 7.25 0 1 1 0 14.5a7.25 7.25 0 1 1 0-14.5z', fill: 'none', stroke: '#2F5D8A' }].concat(sector ? [{ d: sector, fill: '#2F5D8A' }] : []) };
  }
  if (name in FLAGS) return { shapes: [{ d: FLAG_D, fill: FLAGS[name] }] };
  if (name in STARS) return { shapes: [{ d: STAR_D, fill: STARS[name] }] };
  return null;
}
/** SVG for one icon or marker with its top-left at (x, y), d pixels square. */
export function markerSvg(name, x, y, d) {
  const mk = marker(name), esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  if (!mk) return `<text x="${x}" y="${(y + d * 0.86).toFixed(1)}" style="font-size:${(d * 0.92).toFixed(1)}px">${esc(glyph(name) || '•')}</text>`;
  const k = d / 16;
  let o = `<g transform="translate(${x} ${y}) scale(${k.toFixed(4)})">`;
  for (const s of mk.shapes) o += `<path d="${s.d}" style="fill:${s.fill};stroke:${s.stroke || 'none'};stroke-width:1.5"/>`;
  if (mk.text) o += `<text x="8" y="11.4" style="font-size:9.5px;font-weight:700;fill:${mk.fg};text-anchor:middle;font-family:${FONT}">${mk.text}</text>`;
  return o + '</g>';
}
/** The same marker on a 2D canvas (PNG export). */
export function markerCanvas(g, name, x, y, d) {
  const mk = marker(name);
  if (!mk) { g.font = `${d * 0.92}px ${FONT}`; g.textBaseline = 'alphabetic'; g.textAlign = 'left'; g.fillStyle = '#1D1D1F'; g.fillText(glyph(name) || '•', x, y + d * 0.86); return; }
  g.save(); g.translate(x, y); g.scale(d / 16, d / 16);
  for (const s of mk.shapes) { const p = new Path2D(s.d); if (s.fill !== 'none') { g.fillStyle = s.fill; g.fill(p); } if (s.stroke) { g.strokeStyle = s.stroke; g.lineWidth = 1.5; g.stroke(p); } }
  if (mk.text) { g.fillStyle = mk.fg; g.font = `700 9.5px ${FONT}`; g.textAlign = 'center'; g.textBaseline = 'alphabetic'; g.fillText(mk.text, 8, 11.4); g.textAlign = 'left'; }
  g.restore();
}
export const labelList = n => String((n && n.labels) || '').split(/[,，]/).map(s => s.trim()).filter(Boolean);
export const LABEL = { fs: 11, fw: 500, h: 18, px: 7, gap: 5 };
export const BRANCH = ['#3F7D5C', '#2F5D8A', '#B5563A', '#8E5AB8', '#C28A1A', '#4A8F9C'];
export const FONT = "'IBM Plex Sans','Noto Sans SC',sans-serif";
// per depth: font size / weight, padding, line height, max text width, corner radius
const LEVEL = [{ fs: 18, fw: 600, px: 22, py: 12, lh: 26, max: 300, rx: 16 }, { fs: 15, fw: 500, px: 16, py: 9, lh: 22, max: 240, rx: 12 }, { fs: 13.5, fw: 400, px: 12, py: 6, lh: 19, max: 220, rx: 9 }];
const GX = [70, 48, 36], GY = [26, 16, 10]; // gap from a parent at that depth to its children / between its children
export const styleOf = depth => LEVEL[Math.min(depth, 2)];
// FreeMind writes Java's logical font names; they mean "the default" here
const JAVA = { sansserif: '', dialog: '', dialoginput: '', serif: 'serif', monospaced: 'monospace' };
export const fontOf = f => { if (!f) return ''; const k = String(f).toLowerCase().replace(/[\s_-]/g, ''); return k in JAVA ? JAVA[k] : f; };
/** A topic's look: its level's metrics with the topic's own font size, weight, slant, strike and family. */
export function topicStyle(n, depth) {
  const L = styleOf(depth), fs = +n.size > 0 ? +n.size : L.fs;
  return Object.assign({}, L, { fs, lh: +n.size > 0 ? Math.round(fs * 1.42) : L.lh, fw: n.bold ? 700 : L.fw, fi: !!n.italic, strike: !!n.strike, family: fontOf(n.font) });
}
/** The CSS / canvas font shorthand for a topic style. */
export const cssFont = s => `${s.fi ? 'italic ' : ''}${s.fw || 400} ${s.fs}px ${s.family ? JSON.stringify(s.family) + ',' : ''}${FONT}`;
let seq = 0;
export const newId = () => 'tmp_' + Date.now().toString(36) + '_' + (seq++).toString(36);

export function walk(root, fn, parent = null, depth = 0, index = 0) {
  if (fn(root, parent, depth, index) === false) return;
  (root.children || []).forEach((c, i) => walk(c, fn, root, depth + 1, i));
}
export function find(root, id) { let out = null; walk(root, n => { if (n.id === id) { out = n; return false; } }); return out; }
export function parentOf(root, id) { let out = null; walk(root, (n, p) => { if (n.id === id) out = p; }); return out; }
export function count(root) { let n = 0; walk(root, () => { n++; }); return n; }
/** True when `id` is `anc` or sits somewhere under it. */
export function within(root, anc, id) { const a = find(root, anc); return !!a && !!find(a, id); }

// Every topic property the planner writes: model key and how it travels ('s' text, 'b' flag, 'c' colour or none, 'n' number or 0).
const PROPS = [['text', 's'], ['note', 's'], ['link', 's'], ['collapsed', 'b'], ['color', 'c'], ['fill', 'c'], ['icon', 'c'], ['bold', 'b'], ['italic', 'b'], ['strike', 'b'], ['size', 'n'], ['font', 's'], ['free', 's'], ['labels', 's'], ['image', 's'], ['imageSize', 's'], ['cloud', 'c'], ['summary', 's'], ['rels', 'j'], ['structure', 's'], ['theme', 's'], ['lines', 's'], ['mono', 'b']];
/** A topic's relationship lines, each { to, label, color, arrows: 'end' | 'both' | 'start' | 'none' } with the defaults filled in. */
export const relList = n => ((n && n.rels) || []).filter(r => r && r.to).map(r => ({ to: String(r.to), label: r.label || '', color: r.color || '', arrows: r.arrows || 'end' }));
/** Ids of fresh topics inside a topic's relationships and summary range → the engine's ids, once saved. */
export function retarget(n, ids) {
  const real = id => { for (let i = 0; ids[id] && i < 20; i++) id = ids[id]; return id; };
  if (n.rels) n.rels.forEach(r => { r.to = real(r.to); });
  if (n.summary) n.summary = n.summary.split(':').map(real).join(':');
}
/** Where a floating topic sits: free = "x,y", its top-left in map coordinates (the centre of the root is 0,0). */
export const freeAt = n => { const [x, y] = String(n.free || '').split(',').map(Number); return { x: x || 0, y: y || 0 }; };
export const freeOf = (x, y) => Math.round(x) + ',' + Math.round(y);
export function norm(n) {
  const o = { side: n.side || '' };
  for (const [k, t] of PROPS) o[k] = t === 'b' ? !!n[k] : t === 'n' ? (+n[k] > 0 ? String(+n[k]) : '') : t === 'j' ? (relList(n).length ? JSON.stringify(relList(n)) : '') : String(n[k] == null ? '' : n[k]);
  return o;
}
/** id → normalized props + parentId / index / depth, in document order. */
export function flatten(root) { const out = {}; walk(root, (n, p, d, i) => { out[n.id] = Object.assign(norm(n), { parentId: p ? p.id : null, index: i, depth: d }); }); return out; }

// ---------- layout ----------
/** Width estimate for node tests and as a fallback: CJK a full em, everything else a bit over half. */
export const estimate = text => { let w = 0; for (const ch of text) w += ch.codePointAt(0) > 0x2E7F ? 1 : 0.56; return w; };
const TOK = /[⺀-\u{10FFFF}]|[^\s⺀-\u{10FFFF}]+\s*|\s+/gu;
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

export const STRUCTURES = ['map', 'logic', 'logic-left', 'org', 'tree', 'timeline'], LINES = ['curve', 'straight', 'elbow'];
// branch palettes and the centre's colours; the map's theme is the root's theme prop
export const THEMES = {
  classic: { branches: BRANCH, root: 'var(--k7, #1D1D1F)', fg: 'var(--kinv, #FFFFFF)' }, // follows light / dark like the rest of the window
  ocean: { branches: ['#2F5D8A', '#4A8F9C', '#2A9D8F', '#6A7FDB', '#457B9D', '#3F7D5C'], root: '#1B3A5C', fg: '#FFFFFF' },
  sunset: { branches: ['#B5563A', '#C28A1A', '#E07A5F', '#9C4F2E', '#D4A03A', '#B8860B'], root: '#7A2E1F', fg: '#FFFFFF' },
  forest: { branches: ['#3F7D5C', '#5B8C5A', '#2E6B4F', '#8AA35B', '#4F7942', '#6B8E23'], root: '#24483A', fg: '#FFFFFF' },
  candy: { branches: ['#E0609A', '#8E5AB8', '#F76B15', '#3F7D5C', '#2F5D8A', '#E5B800'], root: '#8E5AB8', fg: '#FFFFFF' },
  ink: { branches: ['#3A3A3C', '#5A5A5E', '#6E6E73', '#48484A', '#8E8E93', '#2C2C2E'], root: '#1D1D1F', fg: '#FFFFFF' }
};
export const themeOf = root => THEMES[root && root.theme] || THEMES.classic;
const TREE = { ind: 28 }; // a tree child's indent from its parent's left edge
const union = (e, k) => { e.x0 = Math.min(e.x0, k.dx + k.ext.x0); e.y0 = Math.min(e.y0, k.dy + k.ext.y0); e.x1 = Math.max(e.x1, k.dx + k.ext.x1); e.y1 = Math.max(e.y1, k.dy + k.ext.y1); };
/** Children stacked top to bottom beside the parent (dir 1: right, -1: left), the stack centred on the parent. Offsets are from the parent's top-left. */
function stackV(b, kids, dir, gx, gy) {
  const hs = kids.map(k => k.ext.y1 - k.ext.y0);
  let y = b.h / 2 - (hs.reduce((a, h) => a + h, 0) + gy * Math.max(0, kids.length - 1)) / 2;
  kids.forEach((k, i) => { k.dy = y - k.ext.y0; k.dx = dir > 0 ? b.w + gx : -gx - k.w; y += hs[i] + gy; });
}

/** Children side by side under the parent (org chart), the row centred on it. */
function stackH(b, kids, gx, gy) {
  const ws = kids.map(k => k.ext.x1 - k.ext.x0);
  let x = b.w / 2 - (ws.reduce((a, w) => a + w, 0) + gx * Math.max(0, kids.length - 1)) / 2;
  kids.forEach((k, i) => { k.dx = x - k.ext.x0; k.dy = b.h + gy; x += ws[i] + gx; });
}
/** Children one under the other, indented from the parent (tree); vs -1 grows upwards instead. */
function stackT(b, kids, vs, gy0, gy) {
  let y = vs > 0 ? b.h + gy0 : -gy0;
  for (const k of kids) { k.dx = TREE.ind; const h = k.ext.y1 - k.ext.y0; if (vs > 0) { k.dy = y - k.ext.y0; y += h + gy; } else { k.dy = y - k.ext.y1; y -= h + gy; } }
}

/** Positions every visible topic. measure(text, s) → px width of text in topic style s (see topicStyle); opts.edit =
 * { id, text } lays that topic out with the text being typed. Returns { nodes, edges, bounds }; each box carries the
 * model node as .n, its place (x, y, w, h), .dir (1: its children grow right, -1: left), its branch .color and .parentId. */
export function layout(root, measure, opts = {}) {
  const M = measure || ((t, s) => estimate(t) * s.fs), edit = opts.edit;
  const nodes = [], edges = [];
  const build = (n, depth) => {
    const s = topicStyle(n, depth), m = t => M(t, s), text = edit && edit.id === n.id ? edit.text : n.text;
    const icons = iconList(n), id = Math.round(s.fs * 1.25), iw = icons.length ? icons.length * (id + 3) + 3 : 0;
    const lines = wrap(text, s.max, m), tw = lines.reduce((a, l) => Math.max(a, m(l)), 0);
    // a picture sits above the text, shrunk to the level's text width; labels are a row of pills under it
    const [pw0, ph0] = String(n.imageSize || '').split(',').map(Number), img = n.image ? { w: Math.min(pw0 || 160, s.max), h: 0 } : null;
    if (img) img.h = Math.round(Math.max(24, (ph0 || 120) * img.w / (pw0 || 160)));
    const ls = Object.assign({}, s, { fs: LABEL.fs, fw: LABEL.fw, fi: false, strike: false }), labels = labelList(n).map(t => ({ t, w: Math.ceil(M(t, ls)) + LABEL.px * 2 }));
    const lw = labels.reduce((a, l) => a + l.w + LABEL.gap, -LABEL.gap);
    const ty = s.py + (img ? img.h + 6 : 0);
    const b = { id: n.id, n, depth, s, lines, icons, id2: id, iw, img, labels, ty, w: Math.max(depth ? 44 : 96, Math.ceil(Math.max(tw + iw, img ? img.w : 0, lw) + s.px * 2)), h: Math.ceil(ty + lines.length * s.lh + (labels.length ? LABEL.h + 4 : 0) + s.py), dx: 0, dy: 0 };
    const all = n.children || [];
    b.kids = n.collapsed ? [] : all.map(c => build(c, depth + 1));
    b.badge = n.collapsed && all.length ? count(n) - 1 : 0;
    return b;
  };
  // how a topic's children are arranged: its own structure, else its parent's; a branch of the two-sided map grows one way
  const modeOf = (n, inherited) => STRUCTURES.includes(n.structure) ? n.structure : inherited;
  const arrange = (b, dir, mode, vs) => {
    b.dir = dir; b.vs = vs || 1;
    const mine = modeOf(b.n, mode); b.mode = mine === 'map' || mine === 'logic-left' ? 'logic' : mine === 'timeline' ? 'tree' : mine;
    b.kids.forEach(k => arrange(k, dir, b.mode, b.vs));
    const sums = b.mode === 'logic' ? summaries(b.kids) : [], summed = new Set(sums.map(x => x.s)), kids = b.kids.filter(k => !summed.has(k)), d = Math.min(b.depth, 2);
    if (b.mode === 'org') stackH(b, kids, GX[d], 36);
    else if (b.mode === 'tree') stackT(b, kids, b.vs, 26, 12);
    else stackV(b, kids, dir, GX[d], GY[d]);
    for (const { s, range } of sums) {
      // the brace spans the subtrees it sums up; the summary topic sits beyond them, centred on the brace
      const y0 = Math.min(...range.map(k => k.dy + k.ext.y0)), y1 = Math.max(...range.map(k => k.dy + k.ext.y1));
      const edge = dir > 0 ? Math.max(...range.map(k => k.dx + k.ext.x1)) : Math.min(...range.map(k => k.dx + k.ext.x0));
      s.dy = (y0 + y1) / 2 - (s.ext.y0 + s.ext.y1) / 2;
      s.dx = dir > 0 ? edge + SUM.gap * 2 + SUM.w : edge - SUM.gap * 2 - SUM.w - s.w;
      s.brace = { x: edge + dir * SUM.gap, y0, y1 }; s.isSummary = true;
    }
    b.ext = { x0: 0, y0: 0, x1: b.w, y1: b.h }; b.kids.forEach(k => union(b.ext, k));
    if (b.n.cloud && b.depth) { b.cloud = Object.assign({}, b.ext); b.ext = { x0: b.ext.x0 - CLOUD, y0: b.ext.y0 - CLOUD, x1: b.ext.x1 + CLOUD, y1: b.ext.y1 + CLOUD }; }
  };
  const R = build(root, 0), rootMode = modeOf(root, 'map'), th = themeOf(root), colors = root.mono ? [th.branches[0]] : th.branches, lines = LINES.includes(root.lines) ? root.lines : 'curve';
  R.mode = rootMode; R.dir = 1; R.vs = 1;
  // explicit side wins; the rest go to the emptier side, so a fresh map alternates right/left; floating topics sit where they were dropped
  const sides = { right: [], left: [] }, free = [], main = [];
  R.kids.forEach((k, i) => { k.color = colors[i % colors.length]; if (k.n.free) return free.push(k); main.push(k); if (rootMode === 'map') sides[k.n.side === 'left' || k.n.side === 'right' ? k.n.side : (sides.left.length < sides.right.length ? 'left' : 'right')].push(k); });
  free.forEach(k => arrange(k, 1, 'logic'));
  if (rootMode === 'map') { sides.right.forEach(k => arrange(k, 1, 'logic')); sides.left.forEach(k => arrange(k, -1, 'logic')); stackV(R, sides.right, 1, GX[0], GY[0]); stackV(R, sides.left, -1, GX[0], GY[0]); }
  else if (rootMode === 'logic' || rootMode === 'logic-left') { const dir = rootMode === 'logic' ? 1 : -1; main.forEach(k => arrange(k, dir, 'logic')); stackV(R, main, dir, GX[0], GY[0]); }
  else if (rootMode === 'org') { main.forEach(k => arrange(k, 1, 'org')); stackH(R, main, GX[1], 48); }
  else if (rootMode === 'tree') { main.forEach(k => arrange(k, 1, 'tree')); stackT(R, main, 1, 30, 14); }
  else { // timeline: milestones along an axis to the right of the centre, below and above it in turn, each with its own tree growing away from the axis
    main.forEach((k, i) => arrange(k, 1, 'tree', i % 2 ? -1 : 1));
    let x = R.w + 70;
    main.forEach((k, i) => { k.dx = x - k.ext.x0; k.dy = i % 2 ? R.h / 2 - 34 - k.h : R.h / 2 + 34; x += k.ext.x1 - k.ext.x0 + 40; });
  }
  const clouds = [], braces = [];
  /** The branch line from b to its child k, in b's arrangement and the map's line style. */
  const edgeD = (b, k) => {
    const m = b.mode;
    if (m === 'org') { const x1 = b.x + b.w / 2, y1 = b.y + b.h, x2 = k.x + k.w / 2, ym = (y1 + k.y) / 2; return `M${x1} ${y1}V${ym}H${x2}V${k.y}`; }
    if (m === 'tree') return `M${b.x + TREE.ind / 2} ${b.vs > 0 ? b.y + b.h : b.y}V${k.y + k.h / 2}H${k.x}`;
    if (m === 'timeline') { const ay = b.y + b.h / 2, x2 = k.x + k.w / 2; return `M${x2} ${ay}V${k.y + k.h / 2 > ay ? k.y : k.y + k.h}`; }
    const right = k.x > b.x, x1 = right ? b.x + b.w : b.x, y1 = b.y + b.h / 2, x2 = right ? k.x : k.x + k.w, y2 = k.y + k.h / 2;
    if (lines === 'straight') return `M${x1} ${y1}L${x2} ${y2}`;
    if (lines === 'elbow') { const xm = (x1 + x2) / 2; return `M${x1} ${y1}H${xm}V${y2}H${x2}`; }
    const dx = (x2 - x1) / 2; return `M${x1} ${y1}C${x1 + dx} ${y1},${x2 - dx} ${y2},${x2} ${y2}`;
  };
  const place = (b, x, y, color, parentId) => {
    b.x = x; b.y = y; b.color = color; b.parentId = parentId; nodes.push(b);
    if (b.cloud) clouds.push({ id: b.id, x: x + b.cloud.x0 - CLOUD, y: y + b.cloud.y0 - CLOUD, w: b.cloud.x1 - b.cloud.x0 + CLOUD * 2, h: b.cloud.y1 - b.cloud.y0 + CLOUD * 2, color: '#' + b.n.cloud });
    for (const k of b.kids) {
      if (b === R && k.n.free) continue;
      place(k, x + k.dx, y + k.dy, b.depth ? color : k.color, b.id);
      if (k.brace) { const br = k.brace; braces.push({ id: k.id, x: x + br.x, y0: y + br.y0, y1: y + br.y1, dir: k.dir, tx: k.dir > 0 ? k.x : k.x + k.w, ty: k.y + k.h / 2, color: k.color }); continue; }
      edges.push({ from: b.id, to: k.id, color: k.color, sw: b.depth ? 1.8 : 2.6, d: edgeD(b, k) });
    }
  };
  place(R, -R.w / 2, -R.h / 2, th.root, null);
  if (rootMode === 'timeline' && main.length) edges.unshift({ from: R.id, to: '', color: th.root, sw: 2.6, d: `M${R.x + R.w} ${R.y + R.h / 2}H${Math.max(...main.map(k => k.x + k.w / 2))}` });
  for (const k of free) { const at = freeAt(k.n); k.free = true; place(k, at.x, at.y, k.color, root.id); }
  // relationship lines between any two visible topics
  const byId = new Map(nodes.map(b => [b.id, b])), rels = [];
  for (const a of nodes) relList(a.n).forEach((r, i) => { const t = byId.get(r.to); if (t && t !== a) rels.push(Object.assign({ from: a.id, i }, r, relGeom(a, t))); });
  const bounds = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
  for (const b of nodes.concat(clouds)) { bounds.minX = Math.min(bounds.minX, b.x); bounds.minY = Math.min(bounds.minY, b.y); bounds.maxX = Math.max(bounds.maxX, b.x + b.w); bounds.maxY = Math.max(bounds.maxY, b.y + b.h); }
  return { nodes, edges, bounds, clouds, braces, rels };
}
const CLOUD = 12, SUM = { w: 10, gap: 8 }; // a boundary's padding; a summary brace's width and its gaps
/** The summary topics among kids, each with the regular siblings its range (summary = "firstId:lastId") covers. */
function summaries(kids) {
  const regular = kids.filter(k => !k.n.summary), out = [];
  for (const s of kids) {
    if (!s.n.summary) continue;
    const [a, b] = String(s.n.summary).split(':'), i0 = regular.findIndex(k => k.id === a), i1 = regular.findIndex(k => k.id === b), lo = i0 < 0 ? i1 : i0, hi = i1 < 0 ? i0 : i1;
    if (lo >= 0) out.push({ s, range: regular.slice(Math.min(lo, hi), Math.max(lo, hi) + 1) }); // both topics gone: it stacks like any child
  }
  return out;
}
/** The brace of a summary as SVG path data: a curly bracket from y0 to y1 whose tip points at the summary topic. */
export function bracePath(br) {
  const w = SUM.w * br.dir, r = Math.min(6, (br.y1 - br.y0) / 4), mid = (br.y0 + br.y1) / 2, x = br.x;
  return `M${x} ${br.y0}q${w} 0 ${w} ${r}V${mid - r}q0 ${r} ${w} ${r}q${-w} 0 ${-w} ${r}V${br.y1 - r}q0 ${r} ${-w} ${r}M${x + 2 * w} ${mid}H${br.tx}`;
}
const rnd = v => Math.round(v * 10) / 10;
/** Where the line from box b's centre towards point t leaves the box. */
function edgePoint(b, t) {
  const cx = b.x + b.w / 2, cy = b.y + b.h / 2, dx = t.x - cx, dy = t.y - cy;
  if (!dx && !dy) return { x: cx, y: cy };
  const s = Math.min(dx ? b.w / 2 / Math.abs(dx) : Infinity, dy ? b.h / 2 / Math.abs(dy) : Infinity);
  return { x: cx + dx * s, y: cy + dy * s };
}
/** A relationship line between two boxes: a cubic curve bowing upwards, its middle point (for the label) and filled
 *  arrowhead paths at either end. */
export function relGeom(a, b) {
  const p1 = edgePoint(a, { x: b.x + b.w / 2, y: b.y + b.h / 2 }), p2 = edgePoint(b, { x: a.x + a.w / 2, y: a.y + a.h / 2 });
  const dx = p2.x - p1.x, dy = p2.y - p1.y, len = Math.hypot(dx, dy) || 1, k = Math.min(90, len * 0.28);
  let nx = -dy / len, ny = dx / len; if (ny > 0) { nx = -nx; ny = -ny; }
  const c1 = { x: p1.x + dx / 3 + nx * k, y: p1.y + dy / 3 + ny * k }, c2 = { x: p1.x + dx * 2 / 3 + nx * k, y: p1.y + dy * 2 / 3 + ny * k };
  const head = (p, c) => { const ang = Math.atan2(p.y - c.y, p.x - c.x), L = 11, s = 0.4; return `M${rnd(p.x - L * Math.cos(ang - s))} ${rnd(p.y - L * Math.sin(ang - s))}L${rnd(p.x)} ${rnd(p.y)}L${rnd(p.x - L * Math.cos(ang + s))} ${rnd(p.y - L * Math.sin(ang + s))}z`; };
  return {
    d: `M${rnd(p1.x)} ${rnd(p1.y)}C${rnd(c1.x)} ${rnd(c1.y)},${rnd(c2.x)} ${rnd(c2.y)},${rnd(p2.x)} ${rnd(p2.y)}`,
    mid: { x: (p1.x + 3 * c1.x + 3 * c2.x + p2.x) / 8, y: (p1.y + 3 * c1.y + 3 * c2.y + p2.y) / 8 }, heads: { start: head(p1, c1), end: head(p2, c2) }
  };
}

/** What dropping the topics `ids` at map point p would do, given the layout L of `root`:
 *  { kind: 'into', id, side? } on a topic outside the dragged branches (side by the pointer when it is the root),
 *  { kind: 'between', parent, before, after, side?, y, x0, x1, color } in the gap between two siblings (before / after are
 *  the neighbours' ids, so the caller can turn them into a model index), or { kind: 'free', x, y } on empty canvas, the
 *  topic's top-left once grab (the pointer's offset inside the pressed topic) is taken off. */
export function dropAt(L, root, p, ids, grab = { dx: 0, dy: 0 }) {
  const inDrag = id => ids.some(d => within(root, d, id));
  const hit = L.nodes.find(b => !inDrag(b.id) && p.x >= b.x && p.x <= b.x + b.w && p.y >= b.y && p.y <= b.y + b.h);
  if (hit) return { kind: 'into', id: hit.id, side: hit.id === root.id ? (p.x < 0 ? 'left' : 'right') : null };
  let best = null;
  for (const b of L.nodes) {
    if (inDrag(b.id) || !b.kids.length) continue;
    const regular = b.kids.filter(k => !k.isSummary && !(b.depth === 0 && k.n.free)), horiz = b.mode === 'org' || b.mode === 'timeline';
    const groups = b.depth === 0 && b.mode === 'map' ? [regular.filter(k => k.dir > 0), regular.filter(k => k.dir < 0)] : [regular];
    for (const kids of groups) {
      if (!kids.length || kids.every(k => inDrag(k.id))) continue;
      const last = kids[kids.length - 1], side = b.depth || b.mode !== 'map' ? null : (kids[0].dir > 0 ? 'right' : 'left'), color = kids[0].color;
      if (horiz) { // a row: the gap is between neighbours left and right of the pointer
        const y0 = Math.min(...kids.map(k => k.y)) - 10, y1 = Math.max(...kids.map(k => k.y + k.h)) + 10, x0 = kids[0].x - 18, x1 = last.x + last.w + 18;
        if (p.x < x0 || p.x > x1 || p.y < y0 || p.y > y1) continue;
        const area = (x1 - x0) * (y1 - y0); if (best && best.area <= area) continue;
        const i = kids.filter(k => k.x + k.w / 2 < p.x).length, l = kids[i - 1], r = kids[i];
        best = { kind: 'between', horiz: true, area, parent: b.id, before: r ? r.id : null, after: l ? l.id : null, side, x: l && r ? (l.x + l.w + r.x) / 2 : l ? l.x + l.w + 9 : r.x - 9, y0: y0 + 10, y1: y1 - 10, color };
        continue;
      }
      const x0 = Math.min(...kids.map(k => k.x)) - 10, x1 = Math.max(...kids.map(k => k.x + k.w)) + 10, y0 = kids[0].y - 18, y1 = last.y + last.h + 18;
      if (p.x < x0 || p.x > x1 || p.y < y0 || p.y > y1) continue;
      const area = (x1 - x0) * (y1 - y0); if (best && best.area <= area) continue;
      const i = kids.filter(k => k.y + k.h / 2 < p.y).length, above = kids[i - 1], below = kids[i];
      const y = above && below ? (above.y + above.h + below.y) / 2 : above ? above.y + above.h + 9 : below.y - 9;
      best = { kind: 'between', area, parent: b.id, before: below ? below.id : null, after: above ? above.id : null, side, y, x0: x0 + 10, x1: x1 - 10, color };
    }
  }
  return best || { kind: 'free', x: p.x - grab.dx, y: p.y - grab.dy };
}

// ---------- save planner ----------
// ids of topics added in this save appear inside rels / summary values too: they are mapped when the step's argv is built
const propArgs = (d, ids) => Object.entries(d).flatMap(([k, v]) => ['--prop', k + '=' + (ids && (k === 'rels' || k === 'summary') ? String(v).replace(/tmp_[\w]+/g, t => { for (let i = 0; ids[t] && i < 20; i++) t = ids[t]; return t; }) : v)]);
/** The props that changed from o (normalized) to c (normalized), as command values. Side only matters right under the root. */
export function diffProps(o, c, depth) {
  const d = {};
  for (const [k, t] of PROPS) if (o[k] !== c[k]) d[k] = t === 'b' ? (c[k] ? 'true' : 'false') : t === 'c' ? c[k] || 'none' : t === 'n' ? c[k] || '0' : t === 'j' ? c[k] || '[]' : c[k];
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
  const cursor = {}, late = [];
  // rels and summary name other topics, some of them added in this very save: those sets run last, once every id is known
  const set = (n, d) => { const l = {}; for (const k of ['rels', 'summary']) if (k in d) { l[k] = d[k]; delete d[k]; } if (Object.keys(d).length) steps.push(S(() => ['set', file, P(n.id), ...propArgs(d, ids)])); if (Object.keys(l).length) late.push(S(() => ['set', file, P(n.id), ...propArgs(l, ids)])); };
  walk(root, (n, parent, depth) => {
    const o = orig[n.id], c = norm(n);
    if (!parent) { set(n, o ? diffProps(o, c, 0) : {}); return; }
    const list = sim[parent.id] = sim[parent.id] || [];
    let pos = cursor[parent.id] || 0;
    while (pos < list.length && nw[list[pos]] && nw[list[pos]].parentId !== parent.id) pos++;
    if (!o) {
      steps.push(S(() => ['add', file, P(parent.id), '--type', 'topic', '--prop', 'text=' + c.text, '--index', String(pos + 1)], r => { const real = r && r.props && r.props.id; if (real) { ids[n.id] = real; n.id = real; } }));
      const extra = diffProps(norm({}), c, depth); delete extra.text;
      set(n, extra);
      list.splice(pos, 0, n.id);
    } else {
      set(n, diffProps(o, c, depth));
      if (list[pos] !== n.id) {
        for (const k in sim) { const i = sim[k].indexOf(n.id); if (i >= 0) sim[k].splice(i, 1); }
        list.splice(pos, 0, n.id);
        steps.push(S(() => ['move', file, P(n.id), '--to', P(parent.id), '--index', String(pos + 1)]));
      }
    }
    cursor[parent.id] = pos + 1;
  });
  steps.push(...late);
  return { steps, ids };
}

// ---------- clipboard ----------
/** Branches as an indented outline, one topic per line and children one tab deeper: what other apps paste as a list. */
export function toOutline(branches) {
  const out = [], rec = (n, d) => { out.push('\t'.repeat(d) + String(n.text || '').replace(/\s*\n\s*/g, ' ')); (n.children || []).forEach(c => rec(c, d + 1)); };
  branches.forEach(n => rec(n, 0));
  return out.join('\n');
}
/** Text pasted from anywhere → topic trees ({ text, children }): one topic per non-empty line, deeper indentation nests;
 * list bullets, numbers, checkboxes and heading marks are dropped. */
export function parseOutline(text) {
  const roots = [], stack = [];
  for (const raw of String(text || '').replace(/\r\n?/g, '\n').split('\n')) {
    const m = /^([ \t　]*)(?:[-*+•·]\s+(?:\[[ xX]\]\s+)?|\d+[.)、]\s*|#{1,6}\s+)?(.*)$/.exec(raw), t = m[2].trim();
    if (!t) continue;
    const ind = m[1].replace(/\t/g, '    ').replace(/　/g, '  ').length, node = { text: t, children: [] };
    while (stack.length && stack[stack.length - 1].ind >= ind) stack.pop();
    (stack.length ? stack[stack.length - 1].node.children : roots).push(node);
    stack.push({ ind, node });
  }
  return roots;
}
/** The map rewritten from an edited outline (the first line is the centre, deeper lines its branches, further top lines
 *  more branches). A topic whose text is still there keeps its id and looks — matched under the same parent first, then
 *  anywhere — so moves and renames-by-line both work; lines with new text are new topics, missing lines are deletions. */
export function fromOutline(root, text) {
  const roots = parseOutline(text); if (!roots.length) return null;
  const top = roots[0]; top.children.push(...roots.slice(1));
  const key = t => String(t || '').replace(/\s*\n\s*/g, ' ').trim(), pool = new Map();
  walk(root, n => { if (n !== root) (pool.get(key(n.text)) || pool.set(key(n.text), []).get(key(n.text))).push(n); });
  const build = (line, oldParent) => {
    const list = pool.get(line.text) || [], near = list.find(n => oldParent && (oldParent.children || []).includes(n)), old = list.length ? list.splice(list.indexOf(near || list[0]), 1)[0] : null;
    const n = Object.assign({}, old || { id: newId() }, { text: old && key(old.text) === line.text ? old.text : line.text, children: [] }); delete n.ai;
    n.children = line.children.map(c => build(c, old));
    return n;
  };
  const out = Object.assign({}, root, { text: key(root.text) === top.text ? root.text : top.text }); delete out.ai;
  out.children = top.children.map(c => build(c, root));
  return out;
}
/** A deep copy of a branch with fresh ids everywhere; flags that only mean something in the original place go. */
export function cloneBranch(n) {
  const c = JSON.parse(JSON.stringify(n));
  walk(c, x => { x.id = newId(); delete x.ai; });
  delete c.side;
  return c;
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
    if (PROPS.some(([k]) => o[k] !== c[k])) { n.ai = true; items.push([tr('修改'), label(n)]); }
    else if (o.parentId !== c.parentId) { n.ai = true; items.push([tr('移动'), label(n)]); }
  });
  const gone = Object.keys(b).filter(id => !a[id] && !(b[id].parentId && !a[b[id].parentId])).length; // topmost removed only
  if (gone) items.push([tr('删除@@diff'), tr('{n} 个主题', { n: gone })]);
  return items;
}
export function clearAi(root) { walk(root, n => { delete n.ai; }); }
