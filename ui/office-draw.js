// What Office draws that the editors show but do not edit — a chart from the values it caches, a SmartArt graphic from the shapes
// Office laid out for it, an embedded object by its preview or a box that names it — as SVG or HTML strings of a given size in px.
// The engine reads them (OfficeGraphics.cs, DocxObject.cs); nothing here writes to a file.

const _t = (zh, v) => globalThis.$t ? globalThis.$t(zh, v) : (v ? String(zh).replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : zh);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300'];
const INK = '#595959', GRID = '#D9D9D9', FONT = "'Noto Sans SC','IBM Plex Sans',sans-serif";
const hex = c => c && c !== 'none' ? '#' + c : 'none';
const fmtNum = (v, pct, step) => {
  if (v == null || !isFinite(v)) return '';
  if (pct) return Math.round(v * 100 * 10) / 10 + '%';
  const d = step && step < 1 ? Math.min(4, Math.ceil(-Math.log10(step))) : 0;
  return Number(v.toFixed(d)).toLocaleString('en-US', { maximumFractionDigits: d });
};
/** Nice axis ticks over [lo, hi]: about five steps of 1, 2 or 5 times a power of ten. */
export function ticks(lo, hi, fixedMin, fixedMax) {
  if (!(hi > lo)) { hi = lo + 1; }
  const raw = (hi - lo) / 5, p = Math.pow(10, Math.floor(Math.log10(raw))), m = raw / p, step = (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p;
  const a = fixedMin != null ? fixedMin : Math.floor(lo / step) * step, b = fixedMax != null ? fixedMax : Math.ceil(hi / step) * step;
  const out = []; for (let v = a; v <= b + step / 1e6 && out.length < 40; v += step) out.push(Math.round(v / step) * step);
  return { min: a, max: b, step, list: out };
}
const text = (x, y, s, o = {}) => `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" font-size="${o.size || 12}" fill="${o.fill || INK}" text-anchor="${o.anchor || 'middle'}"${o.bold ? ' font-weight="600"' : ''}${o.rot ? ` transform="rotate(${o.rot} ${x.toFixed(1)} ${y.toFixed(1)})"` : ''} dominant-baseline="${o.base || 'middle'}">${esc(s)}</text>`;
const clip = (s, n) => { s = String(s); return s.length > n ? s.slice(0, Math.max(1, n - 1)) + '…' : s; };

/** A chart (DocxObject's chart prop) as SVG, w × h px. */
export function chartSvg(c, w, h) {
  const series = (c.series || []).filter(s => s && (s.values || []).length), cats = c.cats || [];
  const kind = c.kind || 'bar', title = c.title || '';
  let out = '', top = 8, bottom = h - 8, left = 8, right = w - 8;
  if (title) { out += text(w / 2, 18, clip(title, Math.floor(w / 9)), { size: 14, fill: '#404040' }); top = 34; }
  const pieLike = kind === 'pie' || kind === 'doughnut' || kind === 'treemap' || kind === 'funnel';
  const names = pieLike && kind !== 'funnel' ? cats : series.map(s => s.name);
  const colorAt = (s, i, j) => (s.points && s.points[j]) || (pieLike && kind !== 'funnel' ? PALETTE[j % PALETTE.length] : s.color && s.color !== 'none' ? s.color : PALETTE[i % PALETTE.length]);
  // the legend: its entries in a row under (or over) the chart, or in a column beside it
  const legend = c.legend && names.length && kind !== 'funnel' ? c.legend : '';
  if (legend) {
    const items = names.map((n, i) => ({ n: clip(n || _t('系列 {n}', { n: i + 1 }), 18), c: pieLike ? colorAt(series[0] || {}, 0, i) : (series[i].color === 'none' ? null : colorAt(series[i], i, 0)) })).filter(x => x.c);
    if (legend === 'r' || legend === 'l') {
      const lw = Math.min(w * 0.3, 16 + Math.max(...items.map(x => x.n.length)) * 7.5), x0 = legend === 'r' ? w - lw : 8, y0 = (top + bottom) / 2 - items.length * 9;
      items.forEach((x, i) => { out += `<rect x="${x0}" y="${(y0 + i * 18 - 4).toFixed(1)}" width="8" height="8" fill="${hex(x.c)}"/>` + text(x0 + 12, y0 + i * 18, x.n, { size: 11, anchor: 'start' }); });
      if (legend === 'r') right = w - lw - 8; else left = 8 + lw + 4;
    } else {
      const widths = items.map(x => 20 + x.n.length * 7.5), total = widths.reduce((a, b) => a + b, 0);
      let x = Math.max(8, (w - total) / 2); const y = legend === 't' ? top + 8 : h - 12;
      items.forEach((it, i) => { out += `<rect x="${x.toFixed(1)}" y="${y - 4}" width="8" height="8" fill="${hex(it.c)}"/>` + text(x + 12, y, it.n, { size: 11, anchor: 'start' }); x += widths[i]; });
      if (legend === 't') top += 22; else bottom = h - 28;
    }
  }
  if (kind === 'pie' || kind === 'doughnut') return svg(w, h, out + pie(c, series[0] || { values: [] }, colorAt, left, top, right, bottom));
  if (kind === 'treemap') return svg(w, h, out + treemap(series[0] || { values: [] }, cats, colorAt, left, top, right, bottom));
  if (kind === 'funnel') return svg(w, h, out + funnel(series[0] || { values: [] }, cats, left, top, right, bottom));
  if (kind === 'radar') return svg(w, h, out + radar(series, cats, colorAt, left, top, right, bottom));
  if (!['bar', 'line', 'area', 'scatter', 'waterfall', 'histogram'].includes(kind) || !series.length)
    return svg(w, h, out + text(w / 2, (top + bottom) / 2, _t('图表'), { size: 13 }));
  return svg(w, h, out + axes(c, series, cats, colorAt, left, top, right, bottom));
}
const svg = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" font-family="${FONT}" style="display:block">${body}</svg>`;

/** Bars, lines, areas and points over a category (or value) axis and a value axis, the combo chart's second axis on the right. */
function axes(c, series, cats, colorAt, L, T, R, B) {
  const kind = c.kind, horiz = kind === 'bar' && c.dir === 'bar', grp = c.grouping || 'clustered', pct = grp === 'percentStacked';
  const n = Math.max(cats.length, ...series.map(s => s.values.length));
  const main = series.filter(s => !s.y2), second = series.filter(s => s.y2);
  const stacked = s => (s.kind || kind) !== 'line' && (grp === 'stacked' || pct);
  // waterfall: each point rises or falls from the running total; subtotals stand on zero
  let wf = null;
  if (kind === 'waterfall') {
    const s = series[0], tot = new Set(s.totals || []); let run = 0; wf = s.values.map((v, i) => { v = v || 0; if (tot.has(i)) { run = v; return [0, v, 'total']; } const a = run; run += v; return [a, run, v < 0 ? 'down' : 'up']; });
  }
  const range = list => {
    let lo = 0, hi = 0;
    if (wf) { for (const [a, b] of wf) { lo = Math.min(lo, a, b); hi = Math.max(hi, a, b); } return [lo, hi]; }
    for (let j = 0; j < n; j++) {
      let pos = 0, neg = 0;
      for (const s of list) { const v = s.values[j]; if (v == null) continue; if (stacked(s)) { if (v >= 0) pos += v; else neg += v; } else { hi = Math.max(hi, v); lo = Math.min(lo, v); } }
      hi = Math.max(hi, pos); lo = Math.min(lo, neg);
    }
    if (pct) return [lo < 0 ? -1 : 0, 1];
    if (kind === 'line' || kind === 'scatter') { const all = list.flatMap(s => s.values).filter(v => v != null); if (all.length && Math.min(...all) > 0 && Math.min(...all) > Math.max(...all) * 0.5) lo = Math.min(...all); }
    return [lo, hi];
  };
  const [lo, hi] = range(main.length ? main : series), t = ticks(lo, hi, c.min, c.max);
  const t2 = second.length ? (() => { const [a, b] = range(second); return ticks(a, b); })() : null;
  const labelW = Math.max(...t.list.map(v => fmtNum(v, pct || /%/.test(c.format || ''), t.step).length)) * 6.5 + 10;
  let pl = L + (horiz ? Math.min(120, 10 + Math.max(0, ...cats.map(x => String(x).length)) * 7) : labelW) + (c.yTitle ? 18 : 0), pr = R - (t2 ? labelW : 4), pt = T + 4, pb = B - 20 - (c.xTitle ? 18 : 0);
  if (pr - pl < 40 || pb - pt < 30) return '';
  let out = '';
  // the scatter's x is a value axis too
  const xs = kind === 'scatter' ? (() => { const all = series.flatMap(s => s.x || s.values.map((v, i) => i + 1)).filter(v => v != null); return ticks(Math.min(...all), Math.max(...all)); })() : null;
  const V = (v, tk = t) => horiz ? pl + (v - tk.min) / (tk.max - tk.min) * (pr - pl) : pb - (v - tk.min) / (tk.max - tk.min) * (pb - pt);
  const slot = (horiz ? pb - pt : pr - pl) / Math.max(1, n), C = j => horiz ? pt + slot * (j + 0.5) : pl + slot * (j + 0.5);
  const X = v => pl + (v - xs.min) / (xs.max - xs.min) * (pr - pl);
  for (const v of t.list) {
    const p = V(v);
    if (horiz) out += `<line x1="${p.toFixed(1)}" y1="${pt}" x2="${p.toFixed(1)}" y2="${pb}" stroke="${GRID}" stroke-width="0.75"/>` + text(p, pb + 11, fmtNum(v, pct || /%/.test(c.format || ''), t.step), { size: 10 });
    else out += `<line x1="${pl}" y1="${p.toFixed(1)}" x2="${pr}" y2="${p.toFixed(1)}" stroke="${GRID}" stroke-width="0.75"/>` + text(pl - 5, p, fmtNum(v, pct || /%/.test(c.format || ''), t.step), { size: 10, anchor: 'end' });
  }
  if (t2) for (const v of t2.list) out += text(pr + 5, V(v, t2), fmtNum(v, false, t2.step), { size: 10, anchor: 'start' });
  if (xs) for (const v of xs.list) out += text(X(v), pb + 11, fmtNum(v, false, xs.step), { size: 10 });
  else {
    const every = Math.max(1, Math.ceil(n * 7 * 3 / Math.max(1, horiz ? (pb - pt) * 3 : pr - pl)));
    for (let j = 0; j < n; j += horiz ? 1 : every) out += horiz ? text(pl - 5, C(j), clip(cats[j] ?? j + 1, 16), { size: 10, anchor: 'end' }) : text(C(j), pb + 11, clip(cats[j] ?? j + 1, Math.max(3, Math.floor(slot * every / 6.5))), { size: 10 });
  }
  const axis0 = V(Math.max(t.min, Math.min(t.max, 0)));
  out += horiz ? `<line x1="${axis0.toFixed(1)}" y1="${pt}" x2="${axis0.toFixed(1)}" y2="${pb}" stroke="#BFBFBF"/>` : `<line x1="${pl}" y1="${axis0.toFixed(1)}" x2="${pr}" y2="${axis0.toFixed(1)}" stroke="#BFBFBF"/>`;
  if (c.xTitle) out += text((pl + pr) / 2, B - 6, clip(c.xTitle, 60), { size: 11 });
  if (c.yTitle) out += text(L + 8, (pt + pb) / 2, clip(c.yTitle, 40), { size: 11, rot: -90 });
  // bars first (clustered side by side, stacked on one another), then areas, then lines and points over them
  const bars = series.filter(s => (s.kind || kind) === 'bar' || kind === 'waterfall' || kind === 'histogram');
  const clustered = bars.filter(s => !stacked(s)), stackedBars = bars.filter(stacked);
  const groups = clustered.length + (stackedBars.length ? 1 : 0), bw = slot * 0.64 / Math.max(1, groups);
  const bar = (j, gi, a, b, fill) => {
    if (fill === 'none') return '';
    const c0 = C(j) - slot * 0.32 + gi * bw, p0 = V(a), p1 = V(b);
    return horiz ? `<rect x="${Math.min(p0, p1).toFixed(1)}" y="${c0.toFixed(1)}" width="${Math.abs(p1 - p0).toFixed(1)}" height="${(bw * 0.92).toFixed(1)}" fill="${hex(fill)}"/>`
      : `<rect x="${c0.toFixed(1)}" y="${Math.min(p0, p1).toFixed(1)}" width="${(bw * 0.92).toFixed(1)}" height="${Math.abs(p1 - p0).toFixed(1)}" fill="${hex(fill)}"/>`;
  };
  if (wf) wf.forEach(([a, b, k], j) => { out += bar(j, 0, a, b, k === 'total' ? '4472C4' : k === 'down' ? 'ED7D31' : '70AD47'); });
  else {
    clustered.forEach((s, gi) => s.values.forEach((v, j) => { if (v != null) out += bar(j, gi, 0, v, colorAt(s, series.indexOf(s), j)); }));
    if (stackedBars.length) for (let j = 0; j < n; j++) {
      let pos = 0, neg = 0; const sum = pct ? stackedBars.reduce((a, s) => a + Math.abs(s.values[j] || 0), 0) || 1 : 1;
      for (const s of stackedBars) { let v = s.values[j]; if (v == null) continue; v /= sum; const a = v >= 0 ? pos : neg; if (v >= 0) pos += v; else neg += v; out += bar(j, clustered.length, a, a + v, colorAt(s, series.indexOf(s), j)); }
    }
  }
  const areaS = series.filter(s => (s.kind || kind) === 'area');
  let base = Array(n).fill(0);
  for (const s of areaS) {
    const top = s.values.map((v, j) => (stacked(s) ? base[j] : 0) + (v || 0));
    const pts = top.map((v, j) => `${C(j).toFixed(1)},${V(v).toFixed(1)}`), back = (stacked(s) ? base : Array(n).fill(0)).map((v, j) => `${C(j).toFixed(1)},${V(v).toFixed(1)}`).reverse();
    out += `<polygon points="${pts.concat(back).join(' ')}" fill="${hex(colorAt(s, series.indexOf(s), 0))}" fill-opacity="0.85"/>`;
    if (stacked(s)) base = top;
  }
  for (const s of series.filter(s => (s.kind || kind) === 'line' || kind === 'scatter')) {
    const tk = s.y2 && t2 ? t2 : t, col = hex(colorAt(s, series.indexOf(s), 0));
    const pts = s.values.map((v, j) => v == null ? null : [kind === 'scatter' ? X(s.x ? s.x[j] : j + 1) : C(j), V(v, tk)]).filter(Boolean);
    const lined = kind !== 'scatter' || /line|smooth/i.test(c.scatterStyle || 'lineMarker');
    if (lined && pts.length > 1) out += `<polyline points="${pts.map(p => p.map(x => x.toFixed(1)).join(',')).join(' ')}" fill="none" stroke="${col}" stroke-width="2.25" stroke-linejoin="round"/>`;
    for (const p of pts) out += `<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="${lined ? 2.5 : 3.5}" fill="${col}"/>`;
  }
  if (c.labels) for (const s of series) s.values.forEach((v, j) => { if (v != null && kind !== 'scatter') out += text(horiz ? V(v) + 12 : C(j), horiz ? C(j) : V(v) - 8, fmtNum(v, false, t.step), { size: 9 }); });
  return out;
}

function pie(c, s, colorAt, L, T, R, B) {
  const vals = s.values.map(v => Math.max(0, v || 0)), sum = vals.reduce((a, b) => a + b, 0) || 1;
  const cx = (L + R) / 2, cy = (T + B) / 2, r = Math.max(10, Math.min(R - L, B - T) / 2 - 6), hole = c.kind === 'doughnut' ? r * (c.hole || 50) / 100 : 0;
  let a = -Math.PI / 2, out = '';
  vals.forEach((v, j) => {
    const b = a + v / sum * Math.PI * 2, large = b - a > Math.PI ? 1 : 0, P = (ang, rr) => `${(cx + rr * Math.cos(ang)).toFixed(1)},${(cy + rr * Math.sin(ang)).toFixed(1)}`;
    const d = v >= sum - 1e-9 ? `M${P(0, r)}A${r},${r} 0 1 1 ${P(Math.PI, r)}A${r},${r} 0 1 1 ${P(0, r)}Z`
      : hole ? `M${P(a, r)}A${r},${r} 0 ${large} 1 ${P(b, r)}L${P(b, hole)}A${hole},${hole} 0 ${large} 0 ${P(a, hole)}Z` : `M${cx},${cy}L${P(a, r)}A${r},${r} 0 ${large} 1 ${P(b, r)}Z`;
    out += `<path d="${d}" fill="${hex(colorAt(s, 0, j))}" stroke="#FFFFFF" stroke-width="1"/>`;
    if (v / sum > 0.04) { const m = (a + b) / 2, rr = hole ? (r + hole) / 2 : r * 0.62; out += text(cx + rr * Math.cos(m), cy + rr * Math.sin(m), Math.round(v / sum * 100) + '%', { size: 10, fill: '#FFFFFF' }); }
    a = b;
  });
  if (hole && vals.length) out += `<circle cx="${cx}" cy="${cy}" r="${hole.toFixed(1)}" fill="#FFFFFF"/>`;
  return out;
}

function radar(series, cats, colorAt, L, T, R, B) {
  const n = Math.max(3, cats.length, ...series.map(s => s.values.length)), cx = (L + R) / 2, cy = (T + B) / 2 + 4, r = Math.min(R - L, B - T) / 2 - 16;
  const hi = Math.max(1, ...series.flatMap(s => s.values.filter(v => v != null))), t = ticks(0, hi);
  const P = (j, v) => { const a = -Math.PI / 2 + j / n * Math.PI * 2, rr = v / t.max * r; return [cx + rr * Math.cos(a), cy + rr * Math.sin(a)]; };
  let out = '';
  for (const v of t.list.slice(1)) out += `<polygon points="${Array.from({ length: n }, (_, j) => P(j, v).map(x => x.toFixed(1)).join(',')).join(' ')}" fill="none" stroke="${GRID}" stroke-width="0.75"/>`;
  for (let j = 0; j < n; j++) { const [x, y] = P(j, t.max); out += `<line x1="${cx}" y1="${cy}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="${GRID}" stroke-width="0.75"/>`; const [lx, ly] = P(j, t.max * 1.12); out += text(lx, ly, clip(cats[j] ?? j + 1, 12), { size: 10 }); }
  series.forEach((s, i) => { const col = hex(colorAt(s, i, 0)); out += `<polygon points="${s.values.map((v, j) => P(j, v || 0).map(x => x.toFixed(1)).join(',')).join(' ')}" fill="${col}" fill-opacity="0.2" stroke="${col}" stroke-width="2"/>`; });
  return out;
}

function funnel(s, cats, L, T, R, B) {
  const vals = s.values.map(v => Math.max(0, v || 0)), hi = Math.max(1, ...vals), h = (B - T) / Math.max(1, vals.length), cx = (L + R) / 2;
  return vals.map((v, j) => { const w = v / hi * (R - L - 90); return `<rect x="${(cx - w / 2 + 40).toFixed(1)}" y="${(T + j * h + h * 0.1).toFixed(1)}" width="${w.toFixed(1)}" height="${(h * 0.8).toFixed(1)}" fill="${hex(s.color || PALETTE[0])}"/>`
    + text(L + 40, T + j * h + h / 2, clip(cats[j] ?? '', 10), { size: 10, anchor: 'end' }) + text(cx + 40, T + j * h + h / 2, fmtNum(v), { size: 10, fill: '#FFFFFF' }); }).join('');
}

/** Treemap: the values as rectangles with the areas they have, cut along the longer side each time. */
function treemap(s, cats, colorAt, L, T, R, B) {
  const items = s.values.map((v, j) => ({ v: Math.max(0, v || 0), j })).filter(x => x.v > 0).sort((a, b) => b.v - a.v);
  let out = '';
  const cut = (list, x, y, w, h) => {
    if (!list.length) return;
    if (list.length === 1) { const { j } = list[0]; out += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" fill="${hex(colorAt(s, 0, j))}" stroke="#FFFFFF" stroke-width="1.5"/>` + (w > 30 && h > 16 ? text(x + 5, y + 11, clip(cats[j] ?? '', Math.floor(w / 7)), { size: 10, fill: '#FFFFFF', anchor: 'start' }) : ''); return; }
    const total = list.reduce((a, b) => a + b.v, 0); let acc = 0, k = 0; while (k < list.length - 1 && acc + list[k].v <= total / 2) acc += list[k++].v; if (k === 0) acc = list[k++].v;
    const f = acc / total; if (w >= h) { cut(list.slice(0, k), x, y, w * f, h); cut(list.slice(k), x + w * f, y, w * (1 - f), h); } else { cut(list.slice(0, k), x, y, w, h * f); cut(list.slice(k), x, y + h * f, w, h * (1 - f)); }
  };
  cut(items, L, T, R - L, B - T);
  return out;
}

/** The outline of a preset geometry in a box, as SVG: the shapes SmartArt uses most, else a rectangle. */
function geom(g, x, y, w, h) {
  const P = pts => `<polygon points="${pts.map(([a, b]) => `${(x + a * w).toFixed(1)},${(y + b * h).toFixed(1)}`).join(' ')}"`;
  switch (g) {
    case 'ellipse': case 'donut': case 'pie': case 'blockArc': return `<ellipse cx="${(x + w / 2).toFixed(1)}" cy="${(y + h / 2).toFixed(1)}" rx="${(w / 2).toFixed(1)}" ry="${(h / 2).toFixed(1)}"`;
    case 'roundRect': case 'round2SameRect': case 'snipRoundRect': case 'flowChartAlternateProcess': return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" rx="${(Math.min(w, h) * 0.1667).toFixed(1)}"`;
    case 'rightArrow': return P([[0, 0.25], [0.6, 0.25], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.75], [0, 0.75]]);
    case 'leftArrow': return P([[1, 0.25], [0.4, 0.25], [0.4, 0], [0, 0.5], [0.4, 1], [0.4, 0.75], [1, 0.75]]);
    case 'upArrow': return P([[0.25, 1], [0.25, 0.4], [0, 0.4], [0.5, 0], [1, 0.4], [0.75, 0.4], [0.75, 1]]);
    case 'downArrow': return P([[0.25, 0], [0.25, 0.6], [0, 0.6], [0.5, 1], [1, 0.6], [0.75, 0.6], [0.75, 0]]);
    case 'chevron': return P([[0, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0, 1], [0.25, 0.5]]);
    case 'homePlate': return P([[0, 0], [0.8, 0], [1, 0.5], [0.8, 1], [0, 1]]);
    case 'triangle': return P([[0.5, 0], [1, 1], [0, 1]]);
    case 'diamond': return P([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]);
    case 'hexagon': return P([[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]]);
    case 'parallelogram': return P([[0.25, 0], [1, 0], [0.75, 1], [0, 1]]);
    case 'trapezoid': return P([[0.2, 0], [0.8, 0], [1, 1], [0, 1]]);
    case 'pentagon': return P([[0.5, 0], [1, 0.38], [0.81, 1], [0.19, 1], [0, 0.38]]);
    case 'line': case 'straightConnector1': return `<line x1="${x.toFixed(1)}" y1="${y.toFixed(1)}" x2="${(x + w).toFixed(1)}" y2="${(y + h).toFixed(1)}"`;
    default: return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}"`;
  }
}

/** A SmartArt graphic (DocxObject's smartart prop) as SVG, w × h px: the shapes Office laid out (EMU from its corner), their text in
 *  the text's own box; without them, the words of its nodes in a row of boxes. */
export function smartArtSvg(sa, w, h) {
  const E = 9525; // EMU per px at 96 dpi
  if (!sa.shapes) {
    const texts = sa.texts || [], n = Math.max(1, texts.length), bw = (w - 16) / n;
    return svg(w, h, texts.map((t, i) => `<rect x="${(8 + i * bw + 4).toFixed(1)}" y="${(h / 2 - 28).toFixed(1)}" width="${(bw - 8).toFixed(1)}" height="56" rx="8" fill="#4472C4"/>`
      + text(8 + i * bw + bw / 2, h / 2, clip(t, Math.floor(bw / 8)), { size: 12, fill: '#FFFFFF' })).join(''));
  }
  let out = '';
  for (const s of sa.shapes) {
    const x = s.x / E, y = s.y / E, sw = s.w / E, sh = s.h / E;
    if (sw <= 0 && sh <= 0) continue;
    const rot = s.rot ? ` transform="rotate(${s.rot} ${(x + sw / 2).toFixed(1)} ${(y + sh / 2).toFixed(1)})"` : '';
    const flips=s.flipH||s.flipV?`translate(${x+sw/2} ${y+sh/2}) scale(${s.flipH?-1:1} ${s.flipV?-1:1}) translate(${-x-sw/2} ${-y-sh/2})`:'';
    if(flips)out+=`<g transform="${flips}">`;
    out += `${geom(s.geom, x, y, sw, sh)} fill="${hex(s.fill || 'none')}" stroke="${hex(s.line || 'none')}" stroke-width="${s.lw ? Math.max(0.5, s.lw / 12700 * 4 / 3).toFixed(2) : 1}"${rot}/>`;
    if(s.endArrow&&['line','straightConnector1'].includes(s.geom)){const length=Math.hypot(sw,sh)||1,ux=sw/length,uy=sh/length,size=Math.max(5,(s.lw||19050)/E*3),ex=x+sw,ey=y+sh;out+=`<polygon points="${ex},${ey} ${ex-ux*size-uy*size/2},${ey-uy*size+ux*size/2} ${ex-ux*size+uy*size/2},${ey-uy*size-ux*size/2}" fill="${hex(s.line||'4472C4')}"${rot}/>`;}
    if(flips)out+='</g>';
    if (s.text) {
      const tb = s.tx || s, tx = tb.x / E, ty = tb.y / E, tw = tb.w / E, th = tb.h / E, size = (s.size || 12) * 4 / 3;
      const justify = s.anchor === 't' ? 'flex-start' : s.anchor === 'b' ? 'flex-end' : 'center', align = s.align === 'l' ? 'left' : s.align === 'r' ? 'right' : 'center';
      out += `<foreignObject x="${tx.toFixed(1)}" y="${ty.toFixed(1)}" width="${Math.max(1, tw).toFixed(1)}" height="${Math.max(1, th).toFixed(1)}"${rot}><div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:${justify};text-align:${align};font-size:${size.toFixed(1)}px;line-height:1.15;color:${hex(s.color || '000000')};${s.bold ? 'font-weight:600;' : ''}overflow:hidden;word-break:break-word;padding:2px 4px;box-sizing:border-box">${esc(s.text).replace(/\n/g, '<br>')}</div></foreignObject>`;
    }
  }
  return svg(w, h, sa.width>0&&sa.height>0?`<g transform="scale(${w/(sa.width/E)} ${h/(sa.height/E)})">${out}</g>`:out);
}

/** What an embedded object is, by its program: 公式（MathType）, Excel 工作表… */
export function objectLabel(o) {
  const p = o.progId || '';
  if (o.type === 'ole') return /^Equation\.DSMT/i.test(p) ? _t('公式（MathType）') : /^Equation\./i.test(p) ? _t('公式（公式 3.0）') : /^Excel\.Chart/i.test(p) ? _t('Excel 图表')
    : /^Excel\./i.test(p) ? _t('Excel 工作表') : /^Word\./i.test(p) ? _t('Word 文档') : /^PowerPoint\./i.test(p) ? _t('PowerPoint 演示文稿') : /^Visio\./i.test(p) ? _t('Visio 绘图')
    : /^AcroExch\./i.test(p) ? _t('PDF 文档') : /^Package/i.test(p) ? _t('嵌入的文件') : _t('嵌入的对象');
  return { chart: _t('图表'), smartart: 'SmartArt', group: _t('组合图形'), canvas: _t('绘图画布'), vml: _t('图形（旧格式）') }[o.type] || _t('图形');
}

/** The inside of an object's box, w × h px: the chart or SmartArt drawn, the preview picture (src: its url), a VML text box's text, or
 *  the box's name. */
export function objectInner(o, w, h, src) {
  const parse = x => { if (!x) return null; if (typeof x === 'object') return x; try { return JSON.parse(x); } catch (e) { return null; } };
  const chart = parse(o.chart), sa = parse(o.smartart);
  if (chart) return chartSvg(chart, w, h);
  if (sa) return smartArtSvg(sa, w, h);
  if (src && o.previewFormat === 'wmf') return `<span data-wmf-src="${esc(src)}" data-wmf-key="${esc(o.src || '')}" data-wmf-w="${w}" data-wmf-h="${h}" style="display:block;width:100%;height:100%">${objectInner({ ...o, previewFormat: '' }, w, h, null)}</span>`;
  if (src) return `<img src="${esc(src)}" alt="${esc(o.alt || objectLabel(o))}" draggable="false" style="display:block;width:100%;height:100%;object-fit:fill">`;
  if (o.text) return `<span style="display:block;width:100%;height:100%;box-sizing:border-box;border:1px solid #7F7F7F;padding:4px 7px;font-size:10.5pt;white-space:pre-wrap;overflow:hidden">${esc(o.text)}</span>`;
  const small = h < 40;
  return `<span class="wd-obj-ph" style="width:100%;height:100%;box-sizing:border-box;border:1px dashed #A6A6A6;background:repeating-linear-gradient(135deg,rgba(0,0,0,0.025) 0 6px,transparent 6px 12px);display:flex;align-items:center;justify-content:center;gap:6px;font:${small ? 10 : 12}px ${FONT};color:#7F7F7F;overflow:hidden;white-space:nowrap">${esc(objectLabel(o))}${o.title ? '：' + esc(clip(o.title, 30)) : ''}</span>`;
}
