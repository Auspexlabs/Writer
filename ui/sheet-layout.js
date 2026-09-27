import { cellStats } from './sheet-model.js';
let canvas;
export function textWidth(text, style = {}, fs = 11) {
  if (!canvas && typeof document !== 'undefined' && document.createElement) canvas = document.createElement('canvas').getContext('2d');
  const size = (style.fs || fs) * 1.2;
  if (canvas) { canvas.font = `${style.i ? 'italic ' : ''}${style.b ? 'bold ' : ''}${size}px ${style.font || 'Arial'}`; return canvas.measureText(String(text)).width; }
  return [...String(text)].reduce((w, c) => w + size * (c.codePointAt(0) > 255 ? 1 : .55), 0);
}
const cache = new WeakMap();
export function autoRowHeights(sh, parseA, colName, value, fs = 11, context = sh) {
  const old = cache.get(sh.cells); if (old && old.widths === sh.colW && old.merges === sh.merges && old.fs === fs && old.context === context) return old.rows;
  const rows = {};
  for (const a of cellStats(sh.cells).wrapped) {
    const { r, c } = parseA(a), cell = sh.cells[a], s = cell.s, m = sh.merges?.find(m => r >= m.r && r < m.r + m.rs && c >= m.c && c < m.c + m.cs);
    if (m && (m.rs > 1 || m.c !== c)) continue;
    let width = 0; for (let i = c; i < c + (m?.cs || 1); i++) width += sh.colW?.[colName(i)] || 100;
    width = Math.max(8, width - 12 - (s.indent || 0) * 10);
    const text = String(cell.v?.[0] === '=' && value ? value(r, c) : cell.v ?? '');
    let lines = 0;
    for (const line of text.split('\n')) { let count = 1, used = 0; for (const char of line) { const w = textWidth(char, s, fs); if (used && used + w > width) { count++; used = 0; } used += w; } lines += count; }
    rows[r + 1] = Math.min(546, Math.max(rows[r + 1] || 26, Math.ceil(lines * (s.fs || fs) * 1.2 * 1.35 + 5)));
  }
  cache.set(sh.cells, { widths: sh.colW, merges: sh.merges, fs, context, rows }); return rows;
}
