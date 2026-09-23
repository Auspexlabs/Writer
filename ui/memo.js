// Render memoisation for the dc runtime. support.js skips re-rendering a Design Component whose props are
// unchanged (propsEqual), and a component hands its children callbacks that keep their identity across renders
// (stable / stableAll), so a parent's setState no longer re-renders every editor below it.
// A classic script: index.dc.html loads it before support.js; ui/tests/perf.test.mjs imports it for its side effect.
(function (g) {
  const sameKeys = (a, b) => {
    const ka = Object.keys(a);
    return ka.length === Object.keys(b).length && ka.every(k => Object.is(a[k], b[k]));
  };
  /** Props are equal when every value is identical. The runtime rebuilds the host-style object on every render, so that one is compared by value. */
  function propsEqual(a, b) {
    if (a === b) return true;
    const ka = Object.keys(a);
    if (ka.length !== Object.keys(b).length) return false;
    for (const k of ka) {
      const x = a[k], y = b[k];
      if (Object.is(x, y)) continue;
      if (k === '__hostStyle' && x && y && sameKeys(x, y)) continue;
      return false;
    }
    return true;
  }
  /** A function that keeps its identity across renders and always calls the latest `fn` registered for `key` on `owner`. */
  function stable(owner, key, fn) {
    const cache = owner.__stable || (owner.__stable = new Map());
    let c = cache.get(key);
    if (!c) cache.set(key, c = { fn, call: (...args) => c.fn(...args) });
    c.fn = fn;
    return c.call;
  }
  /** `vals` with every top-level function swapped for its stable counterpart (keyed by property name). */
  function stableAll(owner, vals) {
    for (const k in vals) if (typeof vals[k] === 'function') vals[k] = stable(owner, k, vals[k]);
    return vals;
  }
  g.dcMemo = { propsEqual, stable, stableAll };
})(globalThis);
