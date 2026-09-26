// Shared slide kit + file import/export/print for 素笺 Office.
import { pictureView, picSrc } from './picture.js';
// $t under node (this module is node-tested): falls back to the Chinese, vars filled the same way. Only for text a new
// slide/table is created with — never for existing content, which engine.js reads from the file as it is.
const T = (s, v) => { if (globalThis.$t) return globalThis.$t(s, v); const b = String(s).split('@@')[0]; return v ? b.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : b; };
// The 设计 tab's palettes: the same keys and colours as the engine's palette= (src/Writer.Formats/Pptx/PptxTemplate.cs Palettes),
// which a save writes into the file's theme and an open reads back. Keep the two in step.
export const THEMES = {
  ink: { name: '墨色', bg: '#1D1D1F', fg: '#FFFFFF', sub: '#BDB7AA', acc: '#E3B25A', card: '#2C2C2E', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  paper: { name: '素白', bg: '#FFFFFF', fg: '#1D1D1F', sub: '#6E6E73', acc: '#1D1D1F', card: '#F5F5F7', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  sea: { name: '海蓝', bg: '#1F3550', fg: '#FFFFFF', sub: '#C5D2E0', acc: '#6CC6D9', card: '#2A4666', hf: 'Noto Sans SC', bf: 'Noto Sans SC' },
  clay: { name: '陶土', bg: '#F3E6DA', fg: '#3A2618', sub: '#7A5A45', acc: '#B5563A', card: '#EAD5C3', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  mist: { name: '雾灰', bg: '#F4F6F8', fg: '#1F2A37', sub: '#5B6573', acc: '#3E6FB0', card: '#E4E9EF', hf: 'Noto Sans SC', bf: 'Noto Sans SC' },
  sand: { name: '暖沙', bg: '#F7F3EC', fg: '#2B2620', sub: '#7A6E5F', acc: '#C2833A', card: '#ECE4D6', hf: 'Noto Serif SC', bf: 'Noto Sans SC' },
  rose: { name: '玫红', bg: '#FFFFFF', fg: '#1D1D1F', sub: '#6E6E73', acc: '#C4383C', card: '#F7ECEC', hf: 'Noto Sans SC', bf: 'Noto Sans SC' },
  night: { name: '夜蓝', bg: '#0F1B2D', fg: '#F5F7FA', sub: '#9FB0C7', acc: '#5AC8FA', card: '#1B2A40', hf: 'Noto Sans SC', bf: 'Noto Sans SC' }
};
export const SW = 1600;
export const slideH = ratio => ratio === '4:3' ? 1200 : 900;
let uid = Date.now() % 100000;
export const oid = () => 'o' + (uid++).toString(36);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
export const escHtml = esc;
export function txt(o) { return Object.assign({ id: oid(), t: 'text', x: 128, y: 100, w: 1344, h: 120, rot: 0, html: '<p></p>', fs: 32, color: null, font: null, bold: false, italic: false, underline: false, align: 'left', va: 'top', lh: 1.35, fill: '', stroke: '', sw: 0, op: 1 }, o); }
export function shape(o) { return Object.assign(txt({ t: 'shape', shape: 'rect', fill: null, html: '', align: 'center', va: 'middle', fs: 28 }), o); }
/** A line or connector: its box runs from the start to the end point, flipH / flipV when the line goes left or up (as the
 *  file's p:cxnSp has it); head / tail: the arrowheads; bent: an elbow; start / end: { id, idx } of the shape an end sticks to. */
export function line(o) { return Object.assign({ id: oid(), t: 'line', x: 500, y: 450, w: 600, h: 0, rot: 0, flipH: false, flipV: false, bent: false, stroke: 'acc', sw: 3, dash: 'solid', head: 'none', tail: 'none', start: null, end: null, shadow: false, op: 1 }, o); }
/** A group: its kids in slide coordinates (moving the group moves them); its box is the box around them. */
export function group(kids) { const b = union(kids); return Object.assign({ id: oid(), t: 'group', rot: 0, op: 1, kids }, b); }
export function union(objs) { const x = Math.min(...objs.map(o => o.x)), y = Math.min(...objs.map(o => o.y)); return { x, y, w: Math.max(...objs.map(o => o.x + o.w)) - x, h: Math.max(...objs.map(o => o.y + o.h)) - y }; }
// 插入 › 形状: PowerPoint's gallery by family. Keys are DrawingML presets (rect, roundRect as round, ellipse, pill: a roundRect
// with round ends), which the file keeps; lines are connectors.
export const SHAPE_GALLERY = [
  ['基本形状', [['rect', '矩形'], ['round', '圆角矩形'], ['ellipse', '椭圆'], ['pill', '胶囊'], ['triangle', '三角形'], ['rtTriangle', '直角三角形'], ['diamond', '菱形'], ['parallelogram', '平行四边形'], ['trapezoid', '梯形'], ['pentagon', '五边形'], ['hexagon', '六边形'], ['octagon', '八边形']]],
  ['箭头', [['rightArrow', '右箭头'], ['leftArrow', '左箭头'], ['upArrow', '上箭头'], ['downArrow', '下箭头'], ['leftRightArrow', '左右箭头'], ['chevron', '燕尾形']]],
  ['星形与标注', [['star4', '四角星'], ['star5', '五角星'], ['star6', '六角星'], ['wedgeRectCallout', '矩形标注'], ['wedgeRoundRectCallout', '圆角矩形标注'], ['wedgeEllipseCallout', '椭圆标注']]],
  ['线条', [['line', '直线'], ['arrowLine', '箭头'], ['doubleArrowLine', '双箭头'], ['bentLine', '肘形连接符']]]
];
export const LINE_KINDS = { line: {}, arrowLine: { tail: 'triangle' }, doubleArrowLine: { head: 'triangle', tail: 'triangle' }, bentLine: { bent: true, tail: 'triangle' } };
const star = (n, r) => { const pts = []; for (let i = 0; i < 2 * n; i++) { const a = -Math.PI / 2 + i * Math.PI / n, k = i % 2 ? r : 1; pts.push([k * Math.cos(a), k * Math.sin(a)]); } const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]), x0 = Math.min(...xs), y0 = Math.min(...ys), sx = 100 / (Math.max(...xs) - x0), sy = 100 / (Math.max(...ys) - y0); return pts.map(([x, y]) => [+((x - x0) * sx).toFixed(1), +((y - y0) * sy).toFixed(1)]); };
const ellipseCallout = () => { const pts = []; for (let i = 0; i < 32; i++) { const a = i / 32 * 2 * Math.PI; pts.push([+(50 + 50 * Math.cos(a)).toFixed(1), +(36 + 36 * Math.sin(a)).toFixed(1)]); if (i === 10) pts.push([12, 100]); } return pts; };
/** The polygon shapes as points on a 0–100 square (drawn as SVG, so their outline can be dashed and their fill a gradient). */
export const POLY = {
  triangle: [[50, 0], [100, 100], [0, 100]], rtTriangle: [[0, 0], [100, 100], [0, 100]], diamond: [[50, 0], [100, 50], [50, 100], [0, 50]],
  parallelogram: [[25, 0], [100, 0], [75, 100], [0, 100]], trapezoid: [[25, 0], [75, 0], [100, 100], [0, 100]], pentagon: [[50, 0], [100, 38], [81, 100], [19, 100], [0, 38]],
  hexagon: [[25, 0], [75, 0], [100, 50], [75, 100], [25, 100], [0, 50]], octagon: [[29, 0], [71, 0], [100, 29], [100, 71], [71, 100], [29, 100], [0, 71], [0, 29]],
  rightArrow: [[0, 25], [60, 25], [60, 0], [100, 50], [60, 100], [60, 75], [0, 75]], leftArrow: [[100, 25], [40, 25], [40, 0], [0, 50], [40, 100], [40, 75], [100, 75]],
  upArrow: [[25, 100], [25, 40], [0, 40], [50, 0], [100, 40], [75, 40], [75, 100]], downArrow: [[25, 0], [25, 60], [0, 60], [50, 100], [100, 60], [75, 60], [75, 0]],
  leftRightArrow: [[0, 50], [25, 0], [25, 25], [75, 25], [75, 0], [100, 50], [75, 100], [75, 75], [25, 75], [25, 100]], chevron: [[0, 0], [75, 0], [100, 50], [75, 100], [0, 100], [25, 50]],
  star4: star(4, 0.38), star5: star(5, 0.382), star6: star(6, 0.5),
  wedgeRectCallout: [[0, 0], [100, 0], [100, 72], [30, 72], [10, 100], [18, 72], [0, 72]], wedgeRoundRectCallout: [[6, 0], [94, 0], [100, 8], [100, 64], [94, 72], [30, 72], [10, 100], [18, 72], [6, 72], [0, 64], [0, 8]], wedgeEllipseCallout: ellipseCallout()
};
POLY.arrow = POLY.rightArrow; // the key older decks in memory carry
/** stroke-dasharray for a dash preset at a line width, px. */
export const dashArray = (dash, sw) => ({ dash: [4, 3], dot: [1, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1] }[dash] || []).map(n => Math.max(1, n * sw)).join(' ');
/** A fill value as CSS: a colour, or grad:A,B,angle (angle as the file has it: 0 left to right, 90 top to bottom). */
export const fillCss = f => f && f.startsWith('grad:') ? (g => `linear-gradient(${(+g[2] || 0) + 90}deg,${g[0]},${g[1]})`)(f.slice(5).split(',')) : f;
/** The two points of a line in its own box: [start, end]. */
export function lineEnds(o) { const p1 = [o.flipH ? o.w : 0, o.flipV ? o.h : 0]; return [p1, [o.w - p1[0], o.h - p1[1]]]; }
/** Where a connector sticks to a shape: PowerPoint's rectangle sites 0 top, 1 left, 2 bottom, 3 right. */
export const site = (s, idx) => [[s.x + s.w / 2, s.y], [s.x, s.y + s.h / 2], [s.x + s.w / 2, s.y + s.h], [s.x + s.w, s.y + s.h / 2]][idx] || [s.x + s.w / 2, s.y + s.h / 2];
/** A line's box from two points on the slide, with the flips that say which corner it starts at. */
export function lineBox(p1, p2) { return { x: Math.round(Math.min(p1[0], p2[0])), y: Math.round(Math.min(p1[1], p2[1])), w: Math.round(Math.abs(p2[0] - p1[0])), h: Math.round(Math.abs(p2[1] - p1[1])), flipH: p2[0] < p1[0], flipV: p2[1] < p1[1] }; }
/** Lines stuck to shapes follow them: every connected end takes its shape's site again. Mutates the slide's lines. */
export const flatObjs = objs => objs.flatMap(o => o.t === 'group' ? [o, ...flatObjs(o.kids)] : [o]);
export function relinkLines(s) {
  const all = flatObjs(s.objs);
  for (const o of s.objs) {
    if (o.t !== 'line' || !(o.start || o.end)) continue;
    const at = ref => { const t = ref && all.find(x => x.id === ref.id); return t ? site(t, ref.idx) : null; };
    const [a, b] = lineEnds(o), p1 = at(o.start) || [o.x + a[0], o.y + a[1]], p2 = at(o.end) || [o.x + b[0], o.y + b[1]];
    Object.assign(o, lineBox(p1, p2));
  }
}
// ----- tables: PowerPoint's styles drawn, merges, and the row / column edits on the model -----
/** 表格样式: PowerPoint's built-in styles the engine names (PptxTable.Styles); an unknown style id draws as the first. */
export const TABLE_STYLES = [['MediumStyle2Accent1', '中等样式 2'], ['LightStyle1', '浅色样式 1'], ['LightStyle2Accent1', '浅色样式 2'], ['DarkStyle1', '深色样式 1'], ['TableGrid', '网格'], ['NoStyle', '无样式']];
const mixHex = (a, b, t) => { const h = x => x.replace('#', ''); if (h(a).length !== 6 || h(b).length !== 6) return a; return '#' + [0, 2, 4].map(i => Math.round(parseInt(h(a).slice(i, i + 2), 16) * (1 - t) + parseInt(h(b).slice(i, i + 2), 16) * t).toString(16).padStart(2, '0')).join(''); };
/** A table's column widths in slide units: its own, else equal shares of its width. */
export const colWidths = o => { const n = Math.max(1, ...(o.rows || []).map(r => r.length)); return o.colW && o.colW.length === n ? o.colW : Array.from({ length: n }, () => Math.round(o.w / n)); };
/** The cells a merge hides, "r:c". */
export const coveredCells = merges => { const s = new Set(); for (const m of merges || []) for (let i = m.r; i < m.r + m.rs; i++) for (let j = m.c; j < m.c + m.cs; j++) if (i !== m.r || j !== m.c) s.add(i + ':' + j); return s; };
/** The visible cells of a table drawn in its style: grid placement (gc / gr), fill, colour, weight, lines. */
export function tableCells(o, th) {
  const rows = o.rows || [], nc = Math.max(0, ...rows.map(r => r.length)), covered = coveredCells(o.merges), span = Object.fromEntries((o.merges || []).map(m => [m.r + ':' + m.c, m]));
  const st = TABLE_STYLES.some(x => x[0] === o.tstyle) ? o.tstyle : 'MediumStyle2Accent1', header = o.header !== false, banded = o.banded !== false, firstCol = !!o.firstCol;
  const dark = st === 'DarkStyle1', light1 = st === 'LightStyle1', light2 = st === 'LightStyle2Accent1', grid = st === 'TableGrid', none = st === 'NoStyle', gray = 'rgba(128,128,128,0.35)';
  const out = [];
  for (let r = 0; r < rows.length; r++) for (let c = 0; c < nc; c++) {
    const key = r + ':' + c; if (covered.has(key)) continue;
    const m = span[key], own = (o.cells || {})[key] || {}, isH = header && r === 0, band = banded && !isH && (header ? r % 2 === 1 : r % 2 === 0); // the first body row is the tinted band
    let bg = 'transparent', color = th.fg, border = 'none', bb = 'none';
    if (dark) { bg = isH ? th.acc : band ? mixHex(th.fg, th.bg, 0.18) : th.fg; color = th.bg; border = `1px solid ${th.bg}`; }
    else if (light1) { bb = `1px solid ${gray}`; if (isH) bb = `2px solid ${th.fg}`; if (band) bg = mixHex(th.card, th.bg, 0.4); }
    else if (light2) { bb = `1px solid ${mixHex(th.acc, th.bg, 0.6)}`; if (isH) { color = th.acc; bb = `2px solid ${th.acc}`; } if (band) bg = mixHex(th.acc, th.bg, 0.9); }
    else if (grid) border = `1px solid ${gray}`;
    else if (!none) { bg = isH ? th.acc : band ? th.card : 'transparent'; color = isH ? '#FFFFFF' : th.fg; border = `1px solid ${th.bg}`; }
    if (own.fill) bg = own.fill; if (own.line) border = own.line === 'none' ? 'none' : `1px solid ${own.line}`;
    out.push({ key: o.id + ':' + key, r, c, text: rows[r][c] || '', bg, color, fw: isH || (firstCol && c === 0) ? 700 : 400, border, bb, jc: own.align === 'center' ? 'center' : own.align === 'right' ? 'flex-end' : 'flex-start', align: own.align || 'left',
      gc: `${c + 1} / span ${m ? m.cs : 1}`, gr: `${r + 1} / span ${m ? m.rs : 1}`, cs: m ? m.cs : 1, rs: m ? m.rs : 1 });
  }
  return out;
}
const shiftMerges = (merges, axis, at, delta) => (merges || []).map(m => { const p = m[axis], n = axis === 'r' ? m.rs : m.cs, o = Object.assign({}, m); if (at <= p) o[axis] = p + delta; else if (at < p + n) { if (axis === 'r') o.rs = n + delta; else o.cs = n + delta; } return o; }).filter(m => m.rs > 0 && m.cs > 0 && (m.rs > 1 || m.cs > 1));
const shiftCells = (cells, axis, at, delta) => Object.fromEntries(Object.entries(cells || {}).flatMap(([k, v]) => { const [r, c] = k.split(':').map(Number), p = axis === 'r' ? r : c; if (delta < 0 && p === at) return []; const q = at <= p ? p + delta : p; return [[axis === 'r' ? q + ':' + c : r + ':' + q, v]]; }));
/** A row inserted before index at (at = rows.length appends). */
export function tableInsertRow(o, at) { const nc = Math.max(1, ...o.rows.map(r => r.length)); o.rows.splice(at, 0, Array.from({ length: nc }, () => '')); o.merges = shiftMerges(o.merges, 'r', at, 1); o.cells = shiftCells(o.cells, 'r', at, 1); o.h = Math.round(o.h * o.rows.length / (o.rows.length - 1)); return o; }
export function tableDeleteRow(o, r) { if (o.rows.length < 2) return o; o.rows.splice(r, 1); o.merges = shiftMerges(o.merges, 'r', r, -1); o.cells = shiftCells(o.cells, 'r', r, -1); o.h = Math.round(o.h * o.rows.length / (o.rows.length + 1)); return o; }
export function tableInsertCol(o, at) { const w = colWidths(o), nw = w[Math.min(at, w.length - 1)]; o.rows.forEach(r => r.splice(at, 0, '')); w.splice(at, 0, nw); o.colW = w; o.w = w.reduce((a, b) => a + b, 0); o.merges = shiftMerges(o.merges, 'c', at, 1); o.cells = shiftCells(o.cells, 'c', at, 1); return o; }
export function tableDeleteCol(o, c) { const w = colWidths(o); if (w.length < 2) return o; o.rows.forEach(r => r.splice(c, 1)); w.splice(c, 1); o.colW = w; o.w = w.reduce((a, b) => a + b, 0); o.merges = shiftMerges(o.merges, 'c', c, -1); o.cells = shiftCells(o.cells, 'c', c, -1); return o; }
/** 合并单元格: the range becomes one cell (the texts join it), merges it touches go. */
export function tableMerge(o, r1, c1, r2, c2) {
  const [ra, rb, ca, cb] = [Math.min(r1, r2), Math.max(r1, r2), Math.min(c1, c2), Math.max(c1, c2)];
  if (ra === rb && ca === cb) return o;
  o.merges = (o.merges || []).filter(m => m.r + m.rs <= ra || m.r > rb || m.c + m.cs <= ca || m.c > cb);
  const texts = []; for (let r = ra; r <= rb; r++) for (let c = ca; c <= cb; c++) { if (o.rows[r][c]) texts.push(o.rows[r][c]); if (r !== ra || c !== ca) o.rows[r][c] = ''; }
  o.rows[ra][ca] = texts.join(' '); o.merges.push({ r: ra, c: ca, rs: rb - ra + 1, cs: cb - ca + 1 }); return o;
}
/** 拆分单元格: the merge anchored at r, c goes. */
export function tableSplit(o, r, c) { o.merges = (o.merges || []).filter(m => !(m.r === r && m.c === c)); return o; }

// ----- animations and transitions: PowerPoint's presets (the engine's PptxAnim), the show's clicks, what each looks like on screen -----
/** 动画: the presets by class as the engine names them, [key, label, default ms]; fly and wipe come from the bottom, as in PowerPoint. */
export const FX = [
  ['entr', '进入', [['appear', '出现', 1], ['fade', '淡入', 500], ['fly', '飞入', 500], ['float', '浮入', 1000], ['zoom', '缩放', 500], ['wipe', '擦除', 500]]],
  ['emph', '强调', [['grow', '放大/缩小', 2000], ['spin', '陀螺旋', 2000], ['transparency', '透明', 2000]]],
  ['exit', '退出@@fx', [['disappear', '消失', 1], ['fadeOut', '淡出', 500], ['flyOut', '飞出', 500], ['zoomOut', '收缩', 500], ['wipeOut', '擦除', 500]]]
];
const FX_BY = Object.fromEntries(FX.flatMap(([c, , list]) => list.map(([k, l, d]) => [k, { c, l, d }])));
/** An effect's class: entr, emph, exit, or other (an effect the engine keeps as it was: its class as the file names it). */
export const fxClass = a => FX_BY[a.fx] ? FX_BY[a.fx].c : ({ entrance: 'entr', emphasis: 'emph', exit: 'exit' }[a.cls] || 'other');
export const fxLabel = a => FX_BY[a.fx] ? FX_BY[a.fx].l : '其他效果';
export const fxDefault = key => FX_BY[key] ? FX_BY[key].d : 500;
export const fxDur = a => a.fx === 'appear' || a.fx === 'disappear' ? 1 : a.dur || fxDefault(a.fx);
/** The show's clicks: each with its effects and when each starts (ms after the click). A click starts a step; with the previous
 *  starts at the same time as the effect before it, after the previous when the effects before it are done. Effects before the first
 *  click play as the slide appears (the first step, auto). The engine groups p:timing the same way. */
export function fxSteps(anims) {
  const steps = []; let cur = null, at = 0, end = 0;
  for (const a of anims || []) {
    const start = a.start || 'click';
    if (!cur || start === 'click') { cur = []; steps.push({ auto: !steps.length && start !== 'click', fx: cur }); at = end = 0; }
    else if (start === 'after') at = end;
    const t = at + (a.delay || 0); cur.push({ a, at: t }); end = Math.max(end, t + fxDur(a));
  }
  return steps;
}
/** The objects hidden once `played` steps have played: an entrance hides its object until it plays, an exit hides it after. While
 *  the last step is live (its effects running), its exits still show: the animation takes them out. */
export function fxHidden(anims, played, live) {
  const vis = {};
  for (const a of anims || []) if (a.id && !(a.id in vis)) vis[a.id] = fxClass(a) !== 'entr';
  fxSteps(anims).slice(0, played).forEach((st, k) => { for (const { a } of st.fx) { if (!a.id) continue; const c = fxClass(a); if (c === 'entr') vis[a.id] = true; else if (c === 'exit' && !(live && k === played - 1)) vis[a.id] = false; } });
  return Object.keys(vis).filter(id => !vis[id]);
}
/** What an effect does to its object on screen: keyframes for Element.animate, run with fill both (an exit stays out, 放大 stays big).
 *  translate / scale / rotate leave the object's own rotation alone. H: the slide's height, which fly measures from. */
export function fxFrames(a, o, H) {
  const below = `0 ${Math.round(H - o.y)}px`, c = fxClass(a);
  return {
    appear: [{ opacity: 1 }, { opacity: 1 }], fade: [{ opacity: 0 }, { opacity: 1 }], fly: [{ translate: below }, { translate: '0 0' }],
    float: [{ opacity: 0, translate: `0 ${Math.round(H * 0.1)}px` }, { opacity: 1, translate: '0 0' }], zoom: [{ opacity: 0, scale: '0' }, { opacity: 1, scale: '1' }],
    wipe: [{ clipPath: 'inset(100% 0 0 0)' }, { clipPath: 'inset(0 0 0 0)' }],
    grow: [{ scale: '1' }, { scale: '1.5' }], spin: [{ rotate: '0deg' }, { rotate: '360deg' }], transparency: [{ opacity: 1 }, { opacity: 0.5 }],
    disappear: [{ opacity: 0 }, { opacity: 0 }], fadeOut: [{ opacity: 1 }, { opacity: 0 }], flyOut: [{ translate: '0 0' }, { translate: below }],
    zoomOut: [{ opacity: 1, scale: '1' }, { opacity: 0, scale: '0' }], wipeOut: [{ clipPath: 'inset(0 0 0 0)' }, { clipPath: 'inset(100% 0 0 0)' }]
  }[a.fx] || (c === 'exit' ? [{ opacity: 1 }, { opacity: 0 }] : c === 'emph' ? [{ scale: '1' }, { scale: '1.1' }, { scale: '1' }] : [{ opacity: 0 }, { opacity: 1 }]);
}
/** 切换: the transitions the engine writes, with labels; morph (平滑) moves what the slide shares with the one before. */
export const TRANS = [['none', '无'], ['fade', '淡入淡出'], ['push', '推入'], ['wipe', '擦除'], ['split', '分割'], ['cover', '覆盖'], ['zoom', '缩放'], ['morph', '平滑']];
/** The incoming slide's keyframes per transition (push and cover come from the right, as the engine's dir="l" does). */
export const TRANS_FRAMES = {
  fade: [{ opacity: 0 }, { opacity: 1 }], push: [{ translate: '100% 0' }, { translate: '0 0' }], cover: [{ translate: '100% 0' }, { translate: '0 0' }],
  wipe: [{ clipPath: 'inset(0 0 0 100%)' }, { clipPath: 'inset(0 0 0 0)' }], split: [{ clipPath: 'inset(50% 0 50% 0)' }, { clipPath: 'inset(0 0 0 0)' }],
  zoom: [{ opacity: 0, scale: '0.92' }, { opacity: 1, scale: '1' }]
};
/** 平滑 (morph, lite): which object of the slide before each object of this one continues — the same picture, the same text, the same
 *  table, else the same kind of shape in the same fill, first come first served — so the show can move it from where it was. */
export function morphPairs(prev, next) {
  const key = o => o.t === 'image' ? 'i|' + o.src : o.t === 'table' ? 't|' + JSON.stringify(o.rows) : textOf(o.html) ? 'x|' + textOf(o.html) : 's|' + (o.shape || o.t) + '|' + (o.fill || '');
  const pool = new Map(); for (const o of prev || []) { const k = key(o); pool.set(k, (pool.get(k) || []).concat(o)); }
  const out = {}; for (const o of next || []) { const l = pool.get(key(o)); if (l && l.length) out[o.id] = l.shift(); }
  return out;
}
/** The keyframes that carry object o from where its pair p was. */
export const morphFrames = (p, o) => [{ translate: `${Math.round(p.x + p.w / 2 - o.x - o.w / 2)}px ${Math.round(p.y + p.h / 2 - o.y - o.h / 2)}px`, scale: `${o.w ? p.w / o.w : 1} ${o.h ? p.h / o.h : 1}` }, { translate: '0 0', scale: '1 1' }];

// ----- sections (节): a slide's sec names the section it starts, as the engine's section prop does -----
/** Sections follow their first slide, as the file's do: a deleted first slide hands its section to the next slide of that section, and
 *  the first section always starts at slide 1 (the engine lays p14:sectionLst out the same way). Works on after in place. */
export function keepSections(before, after) {
  const alive = new Set(after.map(s => s.id));
  before.forEach((s, i) => {
    if (!s.sec || alive.has(s.id)) return;
    const n = before.slice(i + 1).find(x => alive.has(x.id) || x.sec), t = n && !n.sec && after.find(x => x.id === n.id);
    if (t) t.sec = s.sec;
  });
  const first = after.findIndex(s => s.sec);
  if (first > 0) { after[0].sec = after[first].sec; after[first].sec = ''; }
  return after;
}

// ----- smart guides: what a dragged box snaps to, and the gaps it shows -----
/** The lines a box snaps to: the slide's edges and centre and every other object's edges and centres, per axis. */
export function snapCands(others, W, H) { const v = [0, W / 2, W], h = [0, H / 2, H]; for (const o of others) { v.push(o.x, o.x + o.w / 2, o.x + o.w); h.push(o.y, o.y + o.h / 2, o.y + o.h); } return { v, h }; }
const nearest = (edges, list, th) => { let d = th, r = null; for (const e of edges) for (const c of list) { const dd = Math.abs(e - c); if (dd < d) { d = dd; r = { move: c - e, at: c }; } } return r; };
/** A moved box b snapped within th: left / centre / right to a vertical line, top / middle / bottom to a horizontal one.
 *  Returns the snapped x, y and the guide lines to draw ({ axis: 'v' | 'h', at }). */
export function snapMove(b, cands, th) {
  const v = nearest([b.x, b.x + b.w / 2, b.x + b.w], cands.v, th), h = nearest([b.y, b.y + b.h / 2, b.y + b.h], cands.h, th), guides = [];
  if (v) guides.push({ axis: 'v', at: v.at }); if (h) guides.push({ axis: 'h', at: h.at });
  return { x: Math.round(b.x + (v ? v.move : 0)), y: Math.round(b.y + (h ? h.move : 0)), guides };
}
/** A resized box: the edges being dragged (dir: n, s, e, w or a corner) snap; the opposite edges stay. */
export function snapResize(b, dir, cands, th) {
  const guides = []; let { x, y, w, h } = b;
  if (dir.includes('e')) { const r = nearest([x + w], cands.v, th); if (r) { w += r.move; guides.push({ axis: 'v', at: r.at }); } }
  if (dir.includes('w')) { const r = nearest([x], cands.v, th); if (r) { x += r.move; w -= r.move; guides.push({ axis: 'v', at: r.at }); } }
  if (dir.includes('s')) { const r = nearest([y + h], cands.h, th); if (r) { h += r.move; guides.push({ axis: 'h', at: r.at }); } }
  if (dir.includes('n')) { const r = nearest([y], cands.h, th); if (r) { y += r.move; h -= r.move; guides.push({ axis: 'h', at: r.at }); } }
  return { x, y, w, h, guides };
}
/** Distance hints: the gap between b and its nearest neighbour on each side that it overlaps with, as { axis, from, to, at, text }
 *  (a line from `from` to `to` along the axis at the cross position `at`, labelled with the gap). */
export function gapHints(b, others) {
  const hints = [], overlap = (a0, a1, b0, b1) => Math.min(a1, b1) - Math.max(a0, b0) > 0;
  let L = null, R = null, T = null, B = null;
  for (const o of others) {
    if (overlap(b.y, b.y + b.h, o.y, o.y + o.h)) { if (o.x + o.w <= b.x && (!L || o.x + o.w > L.x + L.w)) L = o; if (o.x >= b.x + b.w && (!R || o.x < R.x)) R = o; }
    if (overlap(b.x, b.x + b.w, o.x, o.x + o.w)) { if (o.y + o.h <= b.y && (!T || o.y + o.h > T.y + T.h)) T = o; if (o.y >= b.y + b.h && (!B || o.y < B.y)) B = o; }
  }
  const mid = (a0, a1, b0, b1) => (Math.max(a0, b0) + Math.min(a1, b1)) / 2;
  if (L) hints.push({ axis: 'h', from: L.x + L.w, to: b.x, at: mid(b.y, b.y + b.h, L.y, L.y + L.h), text: String(Math.round(b.x - L.x - L.w)) });
  if (R) hints.push({ axis: 'h', from: b.x + b.w, to: R.x, at: mid(b.y, b.y + b.h, R.y, R.y + R.h), text: String(Math.round(R.x - b.x - b.w)) });
  if (T) hints.push({ axis: 'v', from: T.y + T.h, to: b.y, at: mid(b.x, b.x + b.w, T.x, T.x + T.w), text: String(Math.round(b.y - T.y - T.h)) });
  if (B) hints.push({ axis: 'v', from: b.y + b.h, to: B.y, at: mid(b.x, b.x + b.w, B.x, B.x + B.w), text: String(Math.round(B.y - b.y - b.h)) });
  return hints.filter(h => h.to > h.from);
}
function svgOf(o, fill, stroke, sw, dash, gid) {
  const grad = o.fill && String(o.fill).startsWith('grad:') ? o.fill.slice(5).split(',') : null;
  const defs = grad ? `<defs><linearGradient id="${gid}" gradientTransform="rotate(${+grad[2] || 0} .5 .5)"><stop offset="0" stop-color="${grad[0]}"/><stop offset="1" stop-color="${grad[1]}"/></linearGradient></defs>` : '';
  const paint = grad ? `url(#${gid})` : fill || 'none', da = dashArray(dash, sw);
  const attrs = `fill="${paint}" stroke="${sw ? stroke : 'none'}" stroke-width="${sw}"${da ? ` stroke-dasharray="${da}"` : ''} stroke-linejoin="round" vector-effect="non-scaling-stroke"`;
  return `<svg viewBox="0 0 100 100" preserveAspectRatio="none" style="position:absolute;inset:0;width:100%;height:100%;overflow:visible">${defs}<polygon points="${POLY[o.shape].map(p => p.join(',')).join(' ')}" ${attrs}/></svg>`;
}
function lineSvg(o, stroke) {
  const sw = o.sw || 1, [p1, p2] = lineEnds(o), w = Math.max(o.w, 1), h = Math.max(o.h, 1), da = dashArray(o.dash, sw);
  const mid = (p1[0] + p2[0]) / 2, d = o.bent ? `M${p1[0]},${p1[1]} H${mid} V${p2[1]} H${p2[0]}` : `M${p1[0]},${p1[1]} L${p2[0]},${p2[1]}`;
  const head = (from, to, kind) => { // an arrowhead at `to`, pointing away from `from`
    if (!kind || kind === 'none') return ''; const L = 4 * sw + 4, ang = Math.atan2(to[1] - from[1], to[0] - from[0]), c = Math.cos(ang), s = Math.sin(ang);
    if (kind === 'oval') return `<circle cx="${to[0]}" cy="${to[1]}" r="${L / 2.5}" fill="${stroke}"/>`;
    const pts = kind === 'diamond' ? [[0, 0], [-L / 2, L / 3], [-L, 0], [-L / 2, -L / 3]] : kind === 'stealth' ? [[0, 0], [-L, L / 2.2], [-L * 0.7, 0], [-L, -L / 2.2]] : [[0, 0], [-L, L / 2.2], [-L, -L / 2.2]];
    const at = pts.map(([x, y]) => `${(to[0] + x * c - y * s).toFixed(1)},${(to[1] + x * s + y * c).toFixed(1)}`).join(' ');
    return kind === 'arrow' ? `<polyline points="${at}" fill="none" stroke="${stroke}" stroke-width="${sw}" stroke-linejoin="round"/>` : `<polygon points="${at}" fill="${stroke}"/>`;
  };
  const beforeStart = o.bent ? [mid, p1[1]] : p2, beforeEnd = o.bent ? [mid, p2[1]] : p1; // the segment each arrowhead points along
  return `<svg width="${w}" height="${h}" style="position:absolute;left:0;top:0;overflow:visible"><path d="${d}" fill="none" stroke="transparent" stroke-width="${Math.max(14, sw + 10)}"/><path d="${d}" fill="none" stroke="${stroke}" stroke-width="${sw}"${da ? ` stroke-dasharray="${da}"` : ''} stroke-linecap="${o.dash === 'dot' || o.dash === 'sysDot' ? 'round' : 'butt'}"/>${head(beforeStart, p1, o.head)}${head(beforeEnd, p2, o.tail)}</svg>`;
}
/** Slide units per point: a 16:9 deck is 960 pt wide, a 4:3 one 720 (the editor's grid is 1600 wide either way). */
export const ptPxOf = ratio => SW / (ratio === '4:3' ? 720 : 960);
/** The text of some html without markup or spaces: '' for an empty placeholder, which shows its hint instead. */
export const textOf = html => String(html || '').replace(/<[^>]*>|&nbsp;/g, '').replace(/\s+/g, '');
const HINTS = { title: '单击此处添加标题', sub: '单击此处添加副标题', body: '单击此处添加文本', text: '单击此处添加文本', pic: '单击此处添加图片' };
/** What a slide's placeholders stand in for when it takes another layout (the engine binds the same way): a title for a title,
 *  a picture for a picture, and any text placeholder for any other. */
export const phFamily = ph => ph === 'title' || ph === 'pic' ? ph : 'body';
// PowerPoint's layouts on the 1600 × 900 grid (a 4:3 deck stretches y to 1200), placeholders in drawing order: the boxes
// src/Writer.Formats/Pptx/PptxTemplate.cs Layouts writes into the file, so a new slide shows what its save keeps — keep the
// two in step. ph: title; sub (text in the secondary colour); text (text in the text colour, no bullets); body (content, with
// bullets); pic. pt: the size where the layout sets one, else the master's (36 for a title, 24 for text). decor: the layout's
// own graphics, drawn under the slide and never saved — the file's layout carries them.
const TITLE = { ph: 'title', x: 128, y: 56, w: 1344, h: 140, va: 'bottom' };
const BAR = y => ({ t: 'shape', x: 128, y, w: 96, h: 6 });
export const LAYOUT_SPECS = {
  title: { name: '标题幻灯片', ph: [{ ph: 'title', x: 128, y: 200, w: 1344, h: 260, pt: 48, va: 'bottom' }, { ph: 'sub', x: 128, y: 508, w: 1344, h: 120, pt: 24 }], decor: [BAR(482)] },
  content: { name: '标题和内容', ph: [TITLE, { ph: 'body', x: 128, y: 236, w: 1344, h: 584 }] },
  section: { name: '节标题', ph: [{ ph: 'title', x: 128, y: 250, w: 1344, h: 250, pt: 44, va: 'bottom' }, { ph: 'sub', x: 128, y: 548, w: 1344, h: 120, pt: 20 }], decor: [BAR(522)] },
  two: { name: '两栏内容', ph: [TITLE, { ph: 'body', x: 128, y: 236, w: 640, h: 584, pt: 20 }, { ph: 'body', x: 832, y: 236, w: 640, h: 584, pt: 20 }] },
  comparison: { name: '比较', ph: [TITLE, { ph: 'text', x: 128, y: 236, w: 640, h: 72, pt: 24, bold: true, va: 'bottom' }, { ph: 'body', x: 128, y: 324, w: 640, h: 496, pt: 20 }, { ph: 'text', x: 832, y: 236, w: 640, h: 72, pt: 24, bold: true, va: 'bottom' }, { ph: 'body', x: 832, y: 324, w: 640, h: 496, pt: 20 }] },
  titleOnly: { name: '仅标题', ph: [TITLE] },
  blank: { name: '空白', ph: [] },
  caption: { name: '内容与标题', ph: [{ ph: 'title', x: 128, y: 96, w: 480, h: 220, pt: 28, va: 'bottom' }, { ph: 'body', x: 672, y: 96, w: 800, h: 724 }, { ph: 'text', x: 128, y: 348, w: 480, h: 472, pt: 16 }] },
  picture: { name: '图片与标题', ph: [{ ph: 'title', x: 128, y: 96, w: 480, h: 220, pt: 28, va: 'bottom' }, { ph: 'pic', x: 672, y: 96, w: 800, h: 724 }, { ph: 'text', x: 128, y: 348, w: 480, h: 472, pt: 16 }] },
  quote: { name: '引用', ph: [{ ph: 'title', x: 192, y: 300, w: 1216, h: 260, pt: 36, bold: false }, { ph: 'sub', x: 192, y: 600, w: 1216, h: 80, pt: 20 }], decor: [{ t: 'text', x: 168, y: 130, w: 240, h: 200, html: '<p>“</p>', pt: 120, color: 'acc', font: 'head', lh: 1 }] }
};
export const LAYOUTS = Object.entries(LAYOUT_SPECS).map(([key, l]) => [key, l.name]);
/** An empty placeholder for a layout's slot (its hint shows until the user types), sized as the file sizes it. */
export function placeholder(slot, ratio) {
  const ky = slideH(ratio) / 900, k = ptPxOf(ratio), pic = slot.ph === 'pic';
  return txt({ ph: slot.ph, x: slot.x, y: Math.round(slot.y * ky), w: slot.w, h: Math.round(slot.h * ky), fs: Math.round((slot.pt || (slot.ph === 'title' ? 36 : 24)) * k), bold: slot.bold ?? slot.ph === 'title', va: pic ? 'middle' : slot.va || 'top', align: pic ? 'center' : 'left', lh: slot.ph === 'body' ? 1.5 : 1.35, html: '' });
}
/** A new slide of a layout: the layout's placeholders, empty, over its graphics. */
export function makeSlide(layout, ratio) {
  const spec = LAYOUT_SPECS[layout] || LAYOUT_SPECS.blank, ky = slideH(ratio) / 900, k = ptPxOf(ratio);
  const decor = (spec.decor || []).map(d => Object.assign(d.t === 'shape' ? shape({ fill: null, html: '' }) : txt({ fs: Math.round(d.pt * k), color: d.color, font: d.font, html: d.html, lh: d.lh }), { x: d.x, y: Math.round(d.y * ky), w: d.w, h: Math.round(d.h * ky) }));
  return { id: oid(), layout, decor, objs: spec.ph.map(slot => placeholder(slot, ratio)), notes: '', trans: 'fade', hidden: false, bg: null };
}

export function resolveColor(v, th, fallback) { if (v === 'acc') return th.acc; if (v === 'card') return th.card; if (v === 'sub') return th.sub; if (v === 'fg') return th.fg; if (v == null) return fallback; return v; }
/** How a slide object is drawn. ptPx: slide units per point (16:9 decks are 960 pt wide), for picture borders and shadows. */
export function objView(o, th, ptPx) {
  const isShape = o.t === 'shape', isLine = o.t === 'line', isGroup = o.t === 'group';
  const pic = o.t === 'image' ? pictureView(o.look, o.w, o.h, ptPx || SW / 960) : null;
  const fill = isShape ? resolveColor(o.fill, th, th.acc) : resolveColor(o.fill, th, '');
  const color = resolveColor(o.color, th, o.ph === 'sub' ? th.sub : (isShape && (o.fill == null) ? '#FFFFFF' : th.fg));
  const font = o.font === 'head' || (!o.font && o.ph === 'title') ? th.hf : (o.font || th.bf);
  const stroke = resolveColor(o.stroke, th, isLine ? th.acc : ''), sw = o.sw || 0, poly = isShape && POLY[o.shape];
  let radius = '0';
  if (o.shape === 'round') radius = Math.min(o.w, o.h) * 0.18 + 'px';
  if (o.shape === 'ellipse') radius = '50%';
  if (o.shape === 'pill') radius = Math.min(o.w, o.h) / 2 + 'px';
  const gid = 'g' + String(o.id).replace(/\W/g, '_');
  const svg = isLine ? lineSvg(o, stroke) : poly ? svgOf(o, fill, stroke, sw, o.dash, gid) : '';
  const fs = Math.round(o.fs * (o.autofit === 'shrink' && o.fit ? o.fit : 1) * 10) / 10; // 溢出时缩排文字: the scale the box needs
  return {
    left: o.x + 'px', top: o.y + 'px', width: o.w + 'px', height: o.h + 'px', tf: `rotate(${o.rot || 0}deg)`, op: o.op ?? 1,
    flt: o.shadow ? 'drop-shadow(0 3px 6px rgba(0,0,0,0.35))' : 'none',
    bg: poly || isLine ? 'transparent' : fillCss(fill) || 'transparent', radius,
    border: !poly && !isLine && sw && stroke ? `${sw}px ${o.dash === 'dot' || o.dash === 'sysDot' ? 'dotted' : o.dash && o.dash !== 'solid' ? 'dashed' : 'solid'} ${stroke}` : 'none',
    svg: { __html: svg }, isLine, isGroup, kids: isGroup ? o.kids.map(k => Object.assign(objView(k, th, ptPx), { left: (k.x - o.x) + 'px', top: (k.y - o.y) + 'px' })) : [],
    tx: textStyle(o, fs, th),
    color, fs: fs + 'px', font: `'${font}','Noto Sans SC',sans-serif`, fw: o.bold ? 700 : 400, fst: o.italic ? 'italic' : 'normal', td: o.underline ? 'underline' : 'none',
    align: o.align || 'left', va: o.va === 'middle' ? 'center' : o.va === 'bottom' ? 'flex-end' : 'flex-start', lh: o.lh || 1.35,
    pad: o.t === 'text' ? '8px 12px' : '16px 24px', isImg: o.t === 'image', src: picSrc(o.src || ''), picFrame: pic ? pic.frame : '', picImage: pic ? pic.image : '', isTable: o.t === 'table', hasText: o.t === 'text' || o.t === 'shape',
    tcells: o.t === 'table' ? tableCells(o, th) : [], gtc: o.t === 'table' ? colWidths(o).map(w => w + 'px').join(' ') : '', gtr: o.t === 'table' ? `repeat(${Math.max(1, (o.rows || []).length)}, 1fr)` : '',
    inner: { __html: o.html || '' },
    hint: o.ph && HINTS[o.ph] && !textOf(o.html) ? T(HINTS[o.ph]) : '' // an empty placeholder's prompt, drawn by CSS and never part of the text
  };
}
/** The text container's extra style: paragraph spacing (as variables the .sv-t rules read), character spacing, columns, direction
 *  and the WordArt effects — an outline around the letters, a shadow behind them, a gradient through them. */
export function textStyle(o, fs, th) {
  let s = '';
  if (o.sb) s += `--sb:${o.sb}px;`; if (o.sa) s += `--sa:${o.sa}px;`;
  if (o.cs) s += `letter-spacing:${o.cs}px;`;
  if (o.cols > 1) s += `column-count:${o.cols};column-gap:${Math.round((fs || 24) * 1.2)}px;`;
  if (o.vert === 'eaVert') s += 'writing-mode:vertical-rl;text-orientation:mixed;';
  else if (o.vert === 'vert' || o.vert === 'wordArtVert') s += 'writing-mode:vertical-rl;text-orientation:sideways;';
  else if (o.vert === 'vert270') s += 'writing-mode:vertical-lr;text-orientation:sideways;transform:rotate(180deg);';
  if (o.tOutline) s += `-webkit-text-stroke:${Math.max(0.5, (fs || 24) / 40).toFixed(1)}px ${resolveColor(o.tOutline, th, th.fg)};paint-order:stroke fill;`;
  if (o.tShadow) s += `text-shadow:${Math.max(1, (fs || 24) / 16)}px ${Math.max(1, (fs || 24) / 16)}px ${Math.max(2, (fs || 24) / 8)}px rgba(0,0,0,0.4);`;
  if (o.tGrad) { const g = o.tGrad.split(','); s += `background:linear-gradient(${(+g[2] || 0) + 90}deg,${resolveColor(g[0], th, th.fg)},${resolveColor(g[1], th, th.acc)});-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;`; }
  return s;
}
/** Bullet levels (Tab / ⇧Tab): li and p carry data-lvl 1–4; both editors and the print page use this CSS. */
export const LVL_CSS = '.sv-t p,.sv-t li{margin:var(--sb,0) 0 var(--sa,0)}.sv-t li{margin-left:0}.sv-t [data-lvl="1"]{margin-left:1.6em}.sv-t [data-lvl="2"]{margin-left:3.2em}.sv-t [data-lvl="3"]{margin-left:4.8em}.sv-t [data-lvl="4"]{margin-left:6.4em}.sv-t ol,.sv-t ul{margin:0;padding-left:1.1em}';
function slideHtml(s, th, H, scale) {
  const bg = s.bg || th.bg;
  let h = `<div class="sl" style="width:${SW}px;height:${H}px;position:relative;overflow:hidden;background:${bg};zoom:${scale}">`;
  const one = (o, v) => {
    let h = `<div style="position:absolute;left:${v.left};top:${v.top};width:${v.width};height:${v.height};transform:${v.tf};opacity:${v.op};filter:${v.flt}">`;
    h += `<div style="position:absolute;inset:0;background:${v.bg};border-radius:${v.radius};border:${v.border};box-sizing:border-box">${v.svg.__html}</div>`;
    if (v.isImg) h += `<div style="position:absolute;inset:0;${v.picFrame}"><img src="${esc(v.src)}" style="${v.picImage}"></div>`;
    if (v.hasText) h += `<div style="position:absolute;inset:0;padding:${v.pad};display:flex;flex-direction:column;justify-content:${v.va};color:${v.color};font-size:${v.fs};font-family:${v.font};font-weight:${v.fw};font-style:${v.fst};text-decoration:${v.td};text-align:${v.align};line-height:${v.lh};${v.tx}"><div class="sv-t t">${o.html || ''}</div></div>`;
    if (v.isTable) h += `<div style="position:absolute;inset:0;display:grid;grid-template-columns:${v.gtc};grid-template-rows:${v.gtr};font-size:${v.fs};font-family:${v.font}">` + v.tcells.map(c => `<div style="grid-column:${c.gc};grid-row:${c.gr};display:flex;align-items:center;justify-content:${c.jc};padding:0 16px;background:${c.bg};color:${c.color};font-weight:${c.fw};border:${c.border};border-bottom:${c.bb};overflow:hidden;box-sizing:border-box">${esc(c.text)}</div>`).join('') + '</div>';
    if (v.isGroup) h += o.kids.map((k, i) => one(k, v.kids[i])).join('');
    return h + '</div>';
  };
  (s.decor || []).concat(s.objs).forEach(o => { h += one(o, objView(o, th)); });
  return h + '</div>';
}

/** 演示者视图 (a second window, or over the show when there is none): the current slide, the next, the notes and the clock
 *  ([data-clock], which the show ticks). Its buttons carry data-act: prev, next, black, reset, end. */
export function presenterHtml(doc, i, clock) {
  const th = THEMES[doc.theme] || THEMES.paper, H = slideH(doc.ratio), s = doc.slides[i], next = doc.slides.find((x, j) => j > i && !x.hidden);
  const btn = (act, label) => `<button data-act="${act}" style="height:32px;padding:0 14px;border:1px solid rgba(255,255,255,0.18);border-radius:999px;background:rgba(255,255,255,0.08);color:#F5F5F7;font:inherit;font-size:13px;cursor:pointer">${esc(T(label))}</button>`;
  const lbl = t => `<div style="font-size:13px;color:#A1A1A6">${esc(T(t))}</div>`;
  return `<div style="position:absolute;inset:0;overflow:hidden;background:#141414;color:#F5F5F7;font-family:'IBM Plex Sans','Noto Sans SC',sans-serif;display:flex;flex-direction:column;gap:16px;padding:18px 22px;box-sizing:border-box"><style>${LVL_CSS}</style>`
    + `<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap"><span data-clock style="font-size:30px;font-variant-numeric:tabular-nums;margin-right:8px">${clock}</span><span style="color:#A1A1A6;font-size:14px">${esc(T('幻灯片 {i} / {n}', { i: i + 1, n: doc.slides.length }))}</span><span style="flex:1"></span>`
    + btn('prev', '上一张') + btn('next', '下一张') + btn('black', '黑屏') + btn('reset', '重置计时') + btn('end', '结束放映') + '</div>'
    + `<div style="flex:1;min-height:0;display:flex;gap:24px"><div style="flex:none;box-shadow:0 0 0 1px rgba(255,255,255,0.12)">${slideHtml(s, th, H, 0.45)}</div>`
    + `<div style="flex:1;min-width:0;display:flex;flex-direction:column;gap:10px">${lbl('下一张')}<div style="flex:none">${next ? slideHtml(next, th, H, 0.2) : `<div style="font-size:14px;color:#A1A1A6">${esc(T('放映结束'))}</div>`}</div>`
    + `${lbl('备注@@notes')}<div style="flex:1;min-height:0;overflow:auto;font-size:20px;line-height:1.6;white-space:pre-wrap">${esc(s.notes || '')}</div></div></div></div>`;
}

// ----- 导出: each slide as a PNG (drawn on a canvas, lite: a paragraph takes its box's font, a picture fills its box) and the outline as Markdown -----
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', nbsp: ' ', '#39': "'" };
const plainText = h => String(h).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&(amp|lt|gt|quot|nbsp|#39);/g, (x, e) => ENT[e]);
/** A text box's paragraphs: text (hard breaks as \n), level, and the bullet or number it carries. */
export function htmlParas(html) {
  const out = []; let list = null, n = 0;
  String(html || '').replace(/<(ul|ol)\b[^>]*>|<\/(?:ul|ol)>|<(p|li|div|h\d)\b([^>]*)>([\s\S]*?)<\/\2>/gi, (m, open, tag, attrs, inner) => {
    if (open) { list = open.toLowerCase(); n = 0; return ''; }
    if (!tag) { list = null; return ''; }
    const text = plainText(inner);
    const lvl = +((/data-lvl="(\d)"/.exec(attrs || '') || [])[1] || 0);
    out.push({ text, lvl, bullet: tag.toLowerCase() === 'li' ? (list === 'ol' ? ++n + '.' : '•') : '' });
    return '';
  });
  if (!out.length && textOf(html)) out.push({ text: plainText(html), lvl: 0, bullet: '' });
  return out;
}
/** One slide on a 2D context in slide units (1600 wide); pics: the decoded pictures by src. */
export function paintSlide(g, s, th, H, pics) {
  const col = (c, f) => { const v = resolveColor(c, th, f); return String(v || '').startsWith('grad:') ? v.slice(5).split(',')[0] : v; };
  const text = (o, paras, box, font, color, align) => {
    const fs = o.fs || 32, lh = fs * (o.lh || 1.35), lines = [];
    g.font = font; g.fillStyle = color; g.textBaseline = 'middle';
    for (const p of paras) for (const hard of p.text.split('\n')) {
      const ind = p.lvl * fs * 1.6 + (p.bullet ? fs * 1.1 : 0), max = box.w - ind; let cur = '', first = true;
      for (const tok of hard.match(/[\u3000-\u9fff\uff00-\uffef]|[^\s\u3000-\u9fff\uff00-\uffef]+|\s+/g) || ['']) {
        if (cur && g.measureText(cur + tok).width > max) { lines.push({ t: cur, ind, b: first ? p.bullet : '' }); cur = tok.trim() ? tok : ''; first = false; } else cur += tok;
      }
      lines.push({ t: cur, ind, b: first ? p.bullet : '' });
    }
    const y0 = box.y + (o.va === 'middle' ? (box.h - lines.length * lh) / 2 : o.va === 'bottom' ? box.h - lines.length * lh : 0);
    lines.forEach((l, i) => {
      const y = y0 + lh * (i + 0.5); g.textAlign = align === 'center' ? 'center' : align === 'right' ? 'right' : 'left';
      if (l.b) { g.textAlign = 'left'; g.fillText(l.b, box.x + l.ind - fs * 1.1, y); g.textAlign = align === 'center' ? 'center' : align === 'right' ? 'right' : 'left'; }
      g.fillText(l.t, align === 'center' ? box.x + l.ind + (box.w - l.ind) / 2 : align === 'right' ? box.x + box.w : box.x + l.ind, y);
    });
  };
  const one = o => {
    if (o.t === 'group') return (o.kids || []).forEach(one);
    g.save(); g.globalAlpha = o.op ?? 1; g.translate(o.x + o.w / 2, o.y + o.h / 2); g.rotate((o.rot || 0) * Math.PI / 180); g.translate(-o.w / 2, -o.h / 2);
    if (o.t === 'line') { const [a, b] = lineEnds(o); g.beginPath(); g.moveTo(a[0], a[1]); g.lineTo(b[0], b[1]); g.strokeStyle = col(o.stroke, th.acc); g.lineWidth = o.sw || 2; g.stroke(); }
    else if (o.t === 'image') { const im = pics && pics.get(o.src); if (im) g.drawImage(im, 0, 0, o.w, o.h); }
    else if (o.t === 'table') {
      const ws = colWidths(o), xs = ws.map((w, i) => ws.slice(0, i).reduce((a, b) => a + b, 0)), rh = o.h / Math.max(1, (o.rows || []).length);
      for (const c of tableCells(o, th)) {
        const x = xs[c.c], y = c.r * rh, w = ws.slice(c.c, c.c + c.cs).reduce((a, b) => a + b, 0), h = rh * c.rs;
        if (c.bg !== 'transparent') { g.fillStyle = c.bg; g.fillRect(x, y, w, h); }
        if (c.border !== 'none') { g.strokeStyle = c.border.split(' ').slice(2).join(' '); g.lineWidth = 1; g.strokeRect(x, y, w, h); }
        text({ fs: o.fs, va: 'middle', lh: 1.2 }, [{ text: c.text, lvl: 0, bullet: '' }], { x: x + 16, y, w: w - 32, h }, `${c.fw} ${o.fs || 24}px ${th.bf}, sans-serif`, c.color, c.align);
      }
    } else {
      const shape = o.shape, w = o.w, h = o.h, p = new Path2D();
      if (shape === 'ellipse') p.ellipse(w / 2, h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
      else if (POLY[shape]) POLY[shape].forEach(([x, y], i) => i ? p.lineTo(x * w / 100, y * h / 100) : p.moveTo(x * w / 100, y * h / 100));
      else p.roundRect(0, 0, w, h, shape === 'pill' ? Math.min(w, h) / 2 : shape === 'round' ? Math.min(w, h) * 0.12 : 0);
      p.closePath();
      const fill = col(o.fill, o.t === 'shape' ? th.acc : ''); if (fill) { g.fillStyle = fill; g.fill(p); }
      if (o.sw && o.stroke) { g.strokeStyle = col(o.stroke, th.fg); g.lineWidth = o.sw; g.stroke(p); }
      const pad = o.t === 'text' ? [12, 8] : [24, 16], v = objView(o, th);
      text(o, htmlParas(o.html), { x: pad[0], y: pad[1], w: w - pad[0] * 2, h: h - pad[1] * 2 }, `${o.italic ? 'italic ' : ''}${v.fw} ${o.fs || 32}px ${v.font}`, v.color, o.align);
    }
    g.restore();
  };
  g.fillStyle = s.bg || th.bg; g.fillRect(0, 0, SW, H);
  (s.decor || []).concat(s.objs).forEach(one);
}
/** 导出每页为 PNG: every shown slide at 1920 px wide, one file each. */
export async function exportSlidesPng(doc) {
  const th = THEMES[doc.theme] || THEMES.paper, H = slideH(doc.ratio), shown = doc.slides.filter(s => !s.hidden), pics = new Map();
  const imgs = shown.flatMap(s => flatObjs((s.decor || []).concat(s.objs))).filter(o => o.t === 'image');
  await Promise.all(imgs.map(o => new Promise(res => { const im = new Image(); im.onload = () => { pics.set(o.src, im); res(); }; im.onerror = res; im.src = picSrc(o.src); })));
  for (let i = 0; i < shown.length; i++) {
    const c = document.createElement('canvas'), k = 1920 / SW; c.width = 1920; c.height = Math.round(H * k);
    const g = c.getContext('2d'); g.scale(k, k); paintSlide(g, shown[i], th, H, pics);
    download(await new Promise(r => c.toBlob(r, 'image/png')), `${doc.title || 'slides'}-${String(i + 1).padStart(2, '0')}.png`);
    await new Promise(r => setTimeout(r, 300)); // one download at a time
  }
}
/** 导出大纲: each shown slide a heading (its title, else its first line), the rest of its text as bullets, its notes quoted. */
export function slidesMarkdown(doc) {
  const out = doc.title ? ['# ' + doc.title, ''] : [];
  doc.slides.forEach((s, i) => {
    if (s.hidden) return;
    const objs = flatObjs(s.objs).filter(o => o.t === 'text' || o.t === 'shape'), title = objs.find(o => o.ph === 'title');
    const lines = o => htmlParas(o.html).flatMap(p => p.text.split('\n').map(t => ({ t: t.trim(), lvl: p.lvl }))).filter(l => l.t);
    const all = (title ? [title] : []).concat(objs.filter(o => o !== title)).flatMap(lines), head = all.shift();
    out.push('## ' + (head ? head.t : T('第 {n} 页', { n: i + 1 })), '', ...all.map(l => '  '.repeat(l.lvl) + '- ' + l.t));
    for (const t of flatObjs(s.objs).filter(o => o.t === 'table')) out.push('', ...(t.rows || []).flatMap((r, j) => ['| ' + r.join(' | ') + ' |'].concat(j ? [] : ['|' + r.map(() => ' --- |').join('')])));
    if (s.notes) out.push('', ...s.notes.split('\n').map(l => '> ' + l));
    out.push('');
  });
  return out.join('\n').replace(/\n{3,}/g, '\n\n').replace(/\n+$/, '\n');
}
export function exportOutline(doc) { download(new Blob([slidesMarkdown(doc)], { type: 'text/markdown' }), (doc.title || 'slides') + '.md'); }

export const DOC_CSS = `h1{font-size:26px;font-weight:700;margin:18px 0 8px;line-height:1.4}h2{font-size:20px;font-weight:600;margin:16px 0 6px}h3{font-size:16px;font-weight:600;margin:14px 0 4px}p{margin:0 0 8px}blockquote{margin:8px 0;padding:4px 14px;border-left:3px solid #D1D1D6;color:#6E6E73}pre{font-family:'IBM Plex Mono',monospace;background:#F5F5F7;padding:10px 12px;font-size:13px;white-space:pre-wrap}ul,ol{margin:0 0 8px;padding-left:1.6em}img{max-width:100%}a{color:#2F5D8A}ins{color:#3F7D5C;text-decoration:underline}del{color:#B5563A}[data-cid]{background:#F3E6C4}hr[data-pb]{border:none;border-top:1px dashed #C7C7CC;margin:24px 0;break-after:page}`;
const PAGES = { A4: ['210mm', '297mm'], Letter: ['8.5in', '11in'], A5: ['148mm', '210mm'], B5: ['176mm', '250mm'], A3: ['297mm', '420mm'], Legal: ['8.5in', '14in'] };
// 页边距: the engine's presets (DocxSection.Margins) in cm, top / sides; any other margins are the four lengths the engine prints
const MARGIN_CM = { normal: [2.54, 2.54], narrow: [1.27, 1.27], moderate: [2.54, 1.91], wide: [2.54, 5.08] };
/** A Word page's margins in cm, [top, right, bottom, left], from a preset's name or "2.54cm 3.18cm 2.54cm 3.18cm". */
export function marginsCm(margin) {
  const p = MARGIN_CM[margin]; if (p) return [p[0], p[1], p[0], p[1]];
  const v = String(margin || '').trim().split(/[\s,]+/).map(x => { const m = /^(-?[\d.]+)\s*(cm|mm|in|pt)?$/i.exec(x); return m ? +m[1] * ({ mm: 0.1, in: 2.54, pt: 2.54 / 72 }[(m[2] || '').toLowerCase()] || 1) : NaN; });
  const r = v.length === 4 ? v : v.length === 2 ? [v[0], v[1], v[0], v[1]] : v.length === 1 ? [v[0], v[0], v[0], v[0]] : [];
  return r.length && r.every(x => x >= 0) ? r : marginsCm('normal');
}
/** The Noto aliases of assets/fonts/fonts.css (local() fonts only, nothing to fetch), for the pages the app writes out: prints and HTML exports. */
export const fontFaces = () => [...document.styleSheets].filter(s => /\/fonts\.css$/.test(s.href)).flatMap(s => [...s.cssRules].map(r => r.cssText)).filter(t => t.includes('local(')).join('');
export function printDoc(doc, ctx) {
  let css = '', body = '';
  const fonts = `<style>${fontFaces()}</style>`;
  if (doc.type === 'docx') {
    const pg = doc.page || {}; const sz = PAGES[pg.size] || PAGES.A4; const [w, h] = pg.orient === 'landscape' ? [sz[1], sz[0]] : sz;
    css = `@page{size:${w} ${h};margin:${marginsCm(pg.margin).map(x => +x.toFixed(2) + 'cm').join(' ')}}body{margin:0;font-family:'Noto Serif SC',serif;font-size:11pt;line-height:1.8;color:#1D1D1F}.hd,.ft{font-size:9pt;color:#8E8E93}.ed{column-count:${pg.cols || 1};column-gap:32px}` + DOC_CSS;
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
    css = `@page{size:${SW * 0.6}px ${H * 0.6}px;margin:0}body{margin:0}.sl{break-after:page}` + LVL_CSS;
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
  return { theme: 'paper', ratio, slides: slides.length ? slides : [makeSlide('title', ratio)] };
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
  const pg = doc.page || {}; const sizes = { A4: [11906, 16838], Letter: [12240, 15840], A5: [8391, 11906], B5: [10319, 14571], A3: [16838, 23811], Legal: [12240, 20160] }; let [w, h] = sizes[pg.size] || sizes.A4; if (pg.orient === 'landscape') [w, h] = [h, w];
  const [mt, mr, mb, ml] = marginsCm(pg.margin).map(x => Math.round(x * 1440 / 2.54));
  z.file('[Content_Types].xml', CT + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  z.file('_rels/.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file('word/_rels/document.xml.rels', CT + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  const hs = (id, name, sz) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${id.slice(-1) - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
  z.file('word/styles.xml', CT + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:eastAsia="SimSun" w:hAnsi="Georgia"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${hs('Heading1', 'heading 1', 40)}${hs('Heading2', 'heading 2', 32)}${hs('Heading3', 'heading 3', 26)}<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="5C5850"/></w:rPr></w:style></w:styles>`);
  z.file('word/document.xml', CT + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${docxBody(doc.html || '')}<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${pg.orient === 'landscape' ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${mt}" w:right="${mr}" w:bottom="${mb}" w:left="${ml}" w:header="708" w:footer="708" w:gutter="0"/>${(pg.cols || 1) > 1 ? `<w:cols w:num="${pg.cols}" w:space="720"/>` : ''}</w:sectPr></w:body></w:document>`);
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
