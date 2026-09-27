import symbol from './vendor/wmf/symbol.js';
let ready;
const cache = new Map();
const pending = new WeakMap();
function library() {
  return ready ||= new Promise((resolve, reject) => {
    if (globalThis.WMFJS) return resolve(globalThis.WMFJS);
    const script = document.createElement('script'); script.src = new URL('./vendor/wmf/WMFJS.bundle.js', import.meta.url).href;
    script.onload = () => { globalThis.WMFJS.loggingEnabled(false); resolve(globalThis.WMFJS); };
    script.onerror = () => { ready = null; script.remove(); reject(new Error('WMF renderer could not load')); };
    document.head.appendChild(script);
  });
}
export async function renderWmf(bytes, width, height) {
  const WMF = await library(), w = Math.max(1, +width || 240), h = Math.max(1, +height || 120);
  const svg = new WMF.Renderer(bytes).render({ width: '100%', height: '100%', xExt: w, yExt: h, mapMode: 8 });
  // Legacy Symbol encodes alpha as the byte for 'a'. Modern SVG text needs Unicode,
  // including when Symbol/MathType fonts are not installed on this computer.
  for (const text of svg.querySelectorAll('text')) {
    const font = text.getAttribute('font-family') || '';
    if (/^symbol$/i.test(font.trim())) { text.textContent = Array.from(text.textContent, ch => symbol[ch.charCodeAt(0) & 255] || ch).join(''); text.setAttribute('font-family', 'STIX Two Math, Cambria Math, STIXGeneral, serif'); }
  }
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg'); svg.style.display = 'block';
  return svg;
}
export async function paintMetafiles(root) {
  return Promise.all(Array.from(root.querySelectorAll('[data-wmf-src]:not([data-wmf-ready="yes"]):not([data-wmf-ready="error"])'), el => {
    if (pending.has(el)) return pending.get(el);
    const task = (async () => {
      el.setAttribute('data-wmf-ready', 'pending');
      try {
        const src = el.getAttribute('data-wmf-src'), w = el.getAttribute('data-wmf-w'), h = el.getAttribute('data-wmf-h'), key = [src, el.getAttribute('data-wmf-key'), w, h].join('|');
        if (!cache.has(key)) {
          if (cache.size >= 64) cache.delete(cache.keys().next().value);
          cache.set(key, fetch(src).then(r => { if (!r.ok) throw new Error('WMF preview could not load'); return r.arrayBuffer(); }).then(bytes => renderWmf(bytes, w, h)).then(svg => svg.outerHTML).catch(e => { cache.delete(key); throw e; }));
        }
        const html = await cache.get(key); el.innerHTML = html; el.setAttribute('data-wmf-ready', 'yes');
      } catch (error) { el.setAttribute('data-wmf-ready', 'error'); el.title = error.message; }
    })(); pending.set(el, task); return task.finally(() => pending.delete(el));
  }));
}
