// Frame-time diagnostic. index.dc.html imports this when the page is opened with ?perf=1 (HUD only),
// ?perf=auto (HUD + the scripted scenarios below, results PUT to <workspace>/perf/) or localStorage.writerPerf=1.
// The pure helpers are exported for ui/tests/perf.test.mjs; everything else needs a window.
const LONG = 50; // a frame longer than this is a visible hitch

/** Frame deltas (ms between consecutive rAF callbacks) → fps, mean, worst, hitches over 50 ms and dropped frames:
 *  vsyncs missed against `base`, the display's refresh interval (default: the median delta). */
export function frameStats(deltas, base) {
  let sum = 0, worst = 0, long = 0, dropped = 0;
  const n = deltas.length;
  base = base || (n ? [...deltas].sort((a, b) => a - b)[n >> 1] : 0);
  for (const d of deltas) { sum += d; if (d > worst) worst = d; if (d > LONG) long++; dropped += Math.max(0, Math.round(d / base) - 1); }
  return { frames: n, fps: n ? Math.round(1000 * n / sum) : 0, avg: n ? Math.round(sum / n * 10) / 10 : 0, worst: Math.round(worst), long, dropped };
}

/** Stats for the frames whose rAF timestamps fall in (t0, t1]; the refresh interval is the 10th percentile of the last 600 deltas. */
export function windowStats(times, t0, t1) {
  const deltas = [], recent = [];
  for (let i = Math.max(1, times.length - 600); i < times.length; i++) recent.push(times[i] - times[i - 1]);
  for (let i = 1; i < times.length; i++) if (times[i] > t0 && times[i] <= t1) deltas.push(times[i] - times[i - 1]);
  recent.sort((a, b) => a - b);
  return Object.assign(frameStats(deltas, recent[Math.floor(recent.length / 10)]), { ms: Math.round(t1 - t0) });
}

const times = []; // rAF timestamps, most recent last
let tasks = 0, taskMs = 0, hides = 0; // a window that was hidden (occluded) meanwhile has no rAF frames to judge
const results = [];
const sleep = ms => new Promise(r => setTimeout(r, ms));
const frame = () => new Promise(r => requestAnimationFrame(t => r(t)));
const visible = (sel, text) => Array.from(document.querySelectorAll(sel)).find(el => el.offsetParent !== null && (!text || el.textContent.trim() === text));

async function measure(name, fn) {
  await frame(); const t0 = performance.now(), h0 = hides;
  let more; try { more = await fn(); } catch (e) { results.push({ name, error: e.message }); return null; } // a missing control skips one row, not the run
  const t1 = await frame();
  const r = Object.assign({ name }, windowStats(times, t0, t1), more || {}, hides !== h0 || document.hidden ? { hidden: 1 } : {});
  results.push(r); console.log('[perf] ' + JSON.stringify(r));
  return r;
}
async function click(sel, text, settle) {
  const el = visible(sel, text); if (!el) throw new Error('perf: no element ' + (text || sel));
  el.click(); await sleep(settle);
}
/** React only sees a programmatic input/textarea change through the native value setter plus an input event. */
function setValue(el, v) { Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), 'value').set.call(el, String(v)); el.dispatchEvent(new Event('input', { bubbles: true })); }
const setRange = setValue;
async function openDoc(path) { // that file, else the first document of the same type
  const sh = window.__shell, d = sh.state.docs.find(x => x.path === path) || sh.state.docs.find(x => x.type === path.split('.').pop()); if (!d) return false;
  if (sh.doc && sh.doc.id === d.id) { sh.setState({ view: 'create' }); await sleep(400); } // measure a real switch, not a no-op
  const t0 = performance.now(); sh.open(d.id);
  for (let i = 0; i < 400 && !(sh.doc && sh.doc.id === d.id && sh.state.view === 'doc' && document.querySelector('[data-edroot]')); i++) await sleep(10);
  await frame(); results.push({ name: d.type + ' open', ms: Math.round(performance.now() - t0) });
  await sleep(1200); return true;
}
const L = s => (globalThis.$t ? globalThis.$t(s) : s); // UI text as the page shows it (English or Chinese)
const S = {
  async panels(pre) {
    await measure(pre + ' ai-close', () => click('[data-ai-btn]', null, 700)); await measure(pre + ' ai-open', () => click('[data-ai-btn]', null, 700));
    if (visible(`[title="${L('显示/隐藏侧边栏')}"]`)) { await measure(pre + ' left-close', () => click(`[title="${L('显示/隐藏侧边栏')}"]`, null, 700)); await measure(pre + ' left-open', () => click(`[title="${L('显示/隐藏侧边栏')}"]`, null, 900)); }
    await measure(pre + ' bubble-open', () => click('[data-glass-bar] button', L('插入'), 600)); await measure(pre + ' bubble-close', () => click('[data-glass-bar] button', L('插入'), 600));
  },
  async type(pre, zoom) {
    const ed = visible('.wd-ed[contenteditable="true"]'); if (!ed) return;
    const range = visible('input[type="range"]'); if (range) { setRange(range, zoom); await sleep(600); }
    ed.focus(); const s = getSelection(), r = document.createRange(); r.selectNodeContents(ed); r.collapse(false); s.removeAllRanges(); s.addRange(r);
    await measure(pre + ' type100@' + zoom, async () => {
      let sum = 0, max = 0;
      for (let i = 0; i < 100; i++) { const k0 = performance.now(); document.execCommand('insertText', false, '测试文字 abcd '[i % 13]); void ed.offsetHeight; const k = performance.now() - k0; sum += k; if (k > max) max = k; await frame(); } // i18n-ok: text typed into the document
      return { keyAvg: Math.round(sum) / 100, keyMax: Math.round(max * 10) / 10 };
    });
    await measure(pre + ' after-typing@' + zoom, () => sleep(1600)); // the debounced flush, thumbnails and autosave land here
    if (range) { setRange(range, 100); await sleep(400); }
  },
  async chat(pre) { // typing into the assistant box re-renders the shell on every character
    const ta = visible(`textarea[placeholder="${L('让助手修改…')}"]`); if (!ta) return;
    ta.focus();
    await measure(pre + ' ai-input30', async () => { for (let i = 0; i < 30; i++) { setValue(ta, ta.value + '把表格按第二列从高到低排序'[i % 13]); await frame(); } }); // i18n-ok: text typed into the document
    setValue(ta, ''); ta.blur(); await sleep(300);
  },
  async scroll(pre) {
    const z = visible('[data-edroot] [style*="zoom"]'), sc = z && z.parentElement; if (!sc) return;
    await measure(pre + ' scroll60', async () => { for (let i = 0; i < 60; i++) { sc.scrollTop += 24; await frame(); } });
    sc.scrollTop = 0; await sleep(200);
  },
  async rerender(pre) { // what one render of the open editor costs: renderVals + template + React commit, no layout
    const host = document.querySelector('[data-edroot]'), k = host && Object.keys(host).find(x => x.startsWith('__reactFiber$'));
    let f = k && host[k]; while (f && !(f.stateNode && f.stateNode.logic)) f = f.return;
    if (!f || !window.ReactDOM || !ReactDOM.flushSync) return;
    const ms = []; for (let i = 0; i < 7; i++) { const t = performance.now(); ReactDOM.flushSync(() => f.stateNode.forceUpdate()); ms.push(performance.now() - t); await frame(); }
    ms.sort((a, b) => a - b); results.push({ name: pre + ' rerender', ms: Math.round(ms[3] * 10) / 10 });
  },
  async drag(pre) {
    const el = visible('[data-soid]'); if (!el) return; const rc = el.getBoundingClientRect(); let x = rc.left + rc.width / 2, y = rc.top + rc.height / 2;
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: x, clientY: y, button: 0 })); await frame();
    await measure(pre + ' drag60', async () => { for (let i = 0; i < 60; i++) { x += 2; y += 1; window.dispatchEvent(new MouseEvent('mousemove', { clientX: x, clientY: y })); await frame(); } });
    window.dispatchEvent(new MouseEvent('mouseup')); await sleep(300); window.__shell.undo(false); await sleep(300);
  }
};
async function run() {
  results.length = 0;
  if (await openDoc('long.docx')) { await S.rerender('docx'); await S.panels('docx'); await S.chat('docx'); await S.type('docx', 100); await S.type('docx', 200); await S.scroll('docx'); }
  if (await openDoc('Mars-Settlement-Guide.pptx')) { await S.rerender('pptx'); await S.panels('pptx'); await S.chat('pptx'); await S.drag('pptx'); }
  if (await openDoc('big.xlsx')) { await S.rerender('xlsx'); await S.panels('xlsx'); await S.chat('xlsx'); await S.scroll('xlsx'); }
  return results;
}
async function upload() {
  const name = 'perf/' + (navigator.userAgent.includes('Chrome') ? 'chrome' : 'webkit') + (/[?&](desktop|native)=/.test(location.search) ? '-desktop' : '') + '-' + Date.now() + '.json';
  const body = { ua: navigator.userAgent, url: location.href, memo: !!window.dcMemo, errors: document.querySelectorAll('.sc-logic-error').length, results };
  try { await fetch('/file?file=' + encodeURIComponent(name), { method: 'PUT', body: JSON.stringify(body, null, 1) }); console.log('[perf] uploaded ' + name); } catch (e) { console.log('[perf] upload failed', e); }
}

if (typeof window !== 'undefined' && typeof document !== 'undefined') {
  const tick = t => { times.push(t); if (times.length > 20000) times.splice(0, 10000); requestAnimationFrame(tick); };
  requestAnimationFrame(tick);
  document.addEventListener('visibilitychange', () => { if (document.hidden) hides++; });
  try { new PerformanceObserver(l => l.getEntries().forEach(e => { tasks++; taskMs += e.duration; })).observe({ type: 'longtask', buffered: true }); } catch (e) { /* WebKit has no Long Tasks API */ }
  const hud = document.createElement('div');
  hud.style.cssText = 'position:fixed;left:10px;bottom:10px;z-index:99999;padding:4px 8px;border-radius:6px;background:rgba(0,0,0,.72);color:#fff;font:11px/1.4 ui-monospace,monospace;pointer-events:none;white-space:pre';
  document.body.appendChild(hud);
  const log = window.__perfLog = [];
  setInterval(() => {
    const t = performance.now(), s = windowStats(times, t - 1000, t);
    log.push(Object.assign({ t: Math.round(t), tasks, taskMs: Math.round(taskMs) }, s)); if (log.length > 900) log.shift();
    hud.textContent = `${s.fps} fps  worst ${s.worst}ms  dropped ${s.dropped}  long ${s.long}  tasks ${tasks}/${Math.round(taskMs)}ms` + (hud.dataset.done ? '\ndone' : '');
  }, 1000);
  window.__perf = { times, results, stats: (t0, t1) => windowStats(times, t0, t1 ?? performance.now()), run, upload };
  if (new URLSearchParams(location.search).get('perf') === 'auto') {
    (async () => { while (!(window.__shell && window.__shell.state.ready)) await sleep(100); await sleep(1500); await run(); await upload(); hud.dataset.done = '1'; })();
  }
}
