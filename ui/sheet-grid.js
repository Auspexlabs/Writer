// A complete coordinate space, with only visible rows/columns mounted in the DOM.
export const MAX_ROWS = 1048576, MAX_COLS = 16384;
export function axisCoordinates(total, size) {
  const positions = new Float64Array(total + 1);
  for (let i = 0; i < total; i++) positions[i + 1] = positions[i] + size(i);
  return positions;
}
export function axisIndex(positions, pixel) {
  let lo = 0, hi = positions.length - 1;
  while (lo < hi) { const mid = Math.ceil((lo + hi) / 2); if (positions[mid] <= pixel) lo = mid; else hi = mid - 1; }
  return Math.min(positions.length - 2, lo);
}
export function visibleAxis(positions, start, span, frozen = 0) {
  const total = positions.length - 1, indices = new Set();
  const a = Math.max(0, axisIndex(positions, Math.max(0, start)) - 3), b = Math.min(total - 1, axisIndex(positions, start + span) + 3);
  for (let i = a; i <= b; i++) if (positions[i + 1] > positions[i]) indices.add(i);
  const end = Math.min(frozen, axisIndex(positions, span) + 1);
  for (let i = 0; i < end; i++) if (positions[i + 1] > positions[i]) indices.add(i);
  return indices;
}
export function gridWindow(xs, ys, view, merges = []) {
  const cols = visibleAxis(xs, view.x, view.width, view.frC), rows = visibleAxis(ys, view.y, view.height, view.frR);
  const intersect = (set, start, count) => { for (const i of set) if (i >= start && i < start + count) return true; return false; };
  const ms = merges.filter(m => m.r >= 0 && m.c >= 0 && m.r + m.rs < ys.length && m.c + m.cs < xs.length && intersect(rows, m.r, m.rs) && intersect(cols, m.c, m.cs));
  for (const m of ms) { rows.add(m.r); cols.add(m.c); }
  const tracks = (coords, indices, extra) => {
    const edges = new Set([0, coords.length - 1]);
    for (const i of indices) { edges.add(i); edges.add(i + 1); }
    for (const n of extra) edges.add(n);
    const sorted = [...edges].sort((a, b) => a - b), lines = new Map(sorted.map((n, i) => [n, i + 2]));
    return { indices: [...indices].sort((a, b) => a - b), line: n => lines.get(n), css: sorted.slice(1).map((n, i) => (coords[n] - coords[sorted[i]]) + 'px').join(' ') };
  };
  return { rows: tracks(ys, rows, ms.flatMap(m => [m.r, m.r + m.rs])), cols: tracks(xs, cols, ms.flatMap(m => [m.c, m.c + m.cs])), merges: ms };
}
