import { Calc, parseA, A, isErr } from './sheet-engine.js';

/** A bounded single-variable solve. Trials use an overlay; the caller alone commits a successful value. */
export function goalSeek(doc, si, targetRef, changingRef, desired, options = {}) {
  const sh = doc.sheets[si], target = parseA(targetRef), changing = parseA(changingRef);
  if (!sh || !target || !changing || !Number.isFinite(desired)) return { ok: false, error: 'invalid' };
  const key = A(changing.r, changing.c), cell = sh.cells[key] || { v: '' };
  if (sh.cells[A(target.r,target.c)]?.literal || !String(sh.cells[A(target.r, target.c)]?.v || '').startsWith('=') || (!cell.literal && String(cell.v).startsWith('=')) || cell.spill || targetRef.toUpperCase() === changingRef.toUpperCase()) return { ok: false, error: 'formula' };
  const start = Number(cell.v || 0);
  if (!Number.isFinite(start)) return { ok: false, error: 'invalid' };
  const max = Math.max(10, Math.min(300, options.maxIterations || 150)), tolerance = options.tolerance || 1e-8;
  let trialValue = start, iterations = 0, best = null;
  const cells = new Proxy(sh.cells, { get: (base, k) => k === key ? { ...cell, v: trialValue } : base[k], has: (base, k) => k === key || k in base,
    ownKeys: base => Object.hasOwn(base, key) ? Reflect.ownKeys(base) : Reflect.ownKeys(base).concat(key),
    getOwnPropertyDescriptor: (base, k) => k === key ? { configurable: true, enumerable: true } : Object.getOwnPropertyDescriptor(base, k) });
  const sheets = doc.sheets.slice(); sheets[si] = { ...sh, cells }; const work = { ...doc, sheets };
  const sample = x => {
    if (!Number.isFinite(x) || iterations >= max) return null;
    trialValue = x; iterations++;
    const value = new Calc(work).value(si, target.r, target.c);
    if (isErr(value) || typeof value !== 'number' || !Number.isFinite(value)) return null;
    const point = { x, y: value - desired, value };
    if (!best || Math.abs(point.y) < Math.abs(best.y)) best = point;
    return point;
  };
  const solved = point => point && Math.abs(point.y) <= tolerance * Math.max(1, Math.abs(desired));
  let a = sample(start), b = sample(start + Math.max(0.001, Math.abs(start) * 0.01)), bracket = null;
  if (!a && b) { a = b; b = sample(b.x + Math.max(1, Math.abs(b.x) * 0.1)); }
  else if (a && !b) b = sample(start - Math.max(0.001, Math.abs(start) * 0.01));
  const history = [a, b].filter(Boolean);
  const record = point => {
    if (!point) return;
    const other = history.find(p => p.y * point.y < 0);
    if (other) bracket = other.x < point.x ? [other, point] : [point, other];
    history.push(point);
  };
  if (a && b && a.y * b.y < 0) bracket = a.x < b.x ? [a, b] : [b, a];
  // Secant steps converge quickly for smooth financial formulas, including roots tangent to zero.
  for (let i = 0; i < 24 && !solved(best) && a && b && !bracket; i++) {
    const slope = (b.y - a.y) / (b.x - a.x); if (!Number.isFinite(slope) || Math.abs(slope) < 1e-16) break;
    const next = sample(b.x - b.y / slope); if (!next || next.x === b.x) break;
    record(next); a = b; b = next;
  }
  // Search both directions when the initial slope is flat or a trial crosses a formula domain boundary.
  let step = Math.max(1, Math.abs(start) * 0.1);
  for (let i = 0; i < 32 && !bracket && !solved(best) && iterations + 2 < max; i++, step *= 2) {
    record(sample(start + step)); if (!bracket) record(sample(start - step));
  }
  while (bracket && iterations < max && !solved(best)) {
    const [lo, hi] = bracket, x = lo.x + (hi.x - lo.x) / 2;
    if (x === lo.x || x === hi.x) break;
    const point = sample(x); if (!point) break;
    bracket = lo.y * point.y <= 0 ? [lo, point] : [point, hi];
  }
  if (solved(best) && iterations < max) sample(Number(best.x.toPrecision(12)));
  return solved(best) ? { ok: true, value: best.x, result: best.value, iterations } : { ok: false, error: 'convergence', iterations, result: best?.value };
}
