// Liquid-glass motion: spring-in for any element carrying data-pop, plus outside-click helper.
const SPRING = 'cubic-bezier(0.2, 1.28, 0.32, 1)';
const FROM = {
  right: 'translateX(18px) scale(0.97)', left: 'translateX(-18px) scale(0.97)', hud: 'translateY(12px) scale(0.94)', menu: 'translateY(-4px) scale(0.9)', center: 'scale(0.9)'
};
const calm = () => document.documentElement.getAttribute('data-motion') === 'reduce' || matchMedia('(prefers-reduced-motion: reduce)').matches;
function run(el) {
  if (!el.animate || el.__popT && performance.now() - el.__popT < 60) return; el.__popT = performance.now();
  const k = el.getAttribute('data-pop'); const from = FROM[k] || 'translateY(-8px) scale(0.93)';
  // transform + opacity only: both stay on the compositor. A filter: blur() keyframe re-rasterises the element and
  // everything under its backdrop-filter on every frame, which WebKit cannot keep at 60 fps.
  if (calm()) el.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 140, easing: 'ease-out' });
  else {
    el.style.willChange = 'transform, opacity';
    const a = el.animate([{ opacity: 0, transform: from }, { opacity: 1, transform: 'none' }], { duration: k === 'menu' ? 340 : 480, easing: SPRING });
    a.onfinish = a.oncancel = () => { el.style.willChange = ''; };
  }
  const sync = () => el.querySelectorAll && el.querySelectorAll('select[data-val]').forEach(s => { const v = s.getAttribute('data-val'); if (s.value !== v && Array.from(s.options).some(o => o.value === v)) s.value = v; });
  requestAnimationFrame(sync); setTimeout(sync, 60); setTimeout(sync, 250);
  if (el.hasAttribute('data-bubble')) { el.style.marginLeft = '0px'; setTimeout(() => { const r = el.getBoundingClientRect(), W = window.innerWidth; let dx = 0; if (r.left < 10) dx = 10 - r.left; else if (r.right > W - 10) dx = W - 10 - r.right; if (dx) el.style.marginLeft = Math.round(dx) + 'px'; }, 500); }
}
if (!window.__glassObs) {
  window.__glassObs = new MutationObserver(list => {
    list.forEach(m => {
      if (m.type === 'attributes') { if (m.target.hasAttribute && m.target.hasAttribute('data-pop')) run(m.target); return; }
      m.addedNodes.forEach(n => { if (n.nodeType !== 1) return; if (n.hasAttribute('data-pop')) run(n); n.querySelectorAll && n.querySelectorAll('[data-pop]').forEach(run); });
    });
  });
  window.__glassObs.observe(document.documentElement, { subtree: true, childList: true, attributes: true, attributeFilter: ['data-pop'] });
}
const EASE = 'cubic-bezier(.2,.8,.2,1)';
/** FLIP for a layout change the user asked for (a side panel opening, closing or taking the other's place): note where the
 *  [data-flip] elements under root are now; the play() it returns, called once the DOM has its new layout, slides each one
 *  from there to where it is now as a transform, on the compositor, so nothing is laid out again frame by frame.
 *  data-flip="scale" also scales (a page whose fit zoom changed); a [data-flip] inside another moves with it. */
export function flip(root, ms = 300) {
  if (!root || !root.querySelectorAll || calm()) return () => { };
  const els = Array.from(root.querySelectorAll('[data-flip]')).filter(el => !(el.parentElement && el.parentElement.closest('[data-flip]')));
  const from = els.map(el => el.getBoundingClientRect());
  return () => els.forEach((el, i) => {
    const a = from[i]; if (!el.isConnected || !el.animate || !a.width) return;
    const b = el.getBoundingClientRect(); if (!b.width) return;
    const scale = el.getAttribute('data-flip') === 'scale', z = parseFloat(getComputedStyle(el).zoom) || 1; // a zoomed element moves by z per px
    const sx = scale ? a.width / b.width : 1, sy = scale ? a.height / b.height : 1;
    const dx = (scale ? a.left - b.left : a.left + a.width / 2 - b.left - b.width / 2) / z, dy = (a.top - b.top) / z;
    if (Math.abs(dx) < 0.5 && Math.abs(dy) < 0.5 && Math.abs(sx - 1) < 0.002 && Math.abs(sy - 1) < 0.002) return;
    el.animate([{ transform: `translate(${dx}px,${dy}px) scale(${sx},${sy})`, transformOrigin: '0 0' }, { transform: 'none', transformOrigin: '0 0' }], { duration: ms, easing: EASE });
  });
}
/** A side panel that is about to close: a copy of it, taken now while it is still there; the play() it returns (called once the
 *  panel is gone) fades and slides the copy out where the panel was. */
export function ghost(el, ms = 220) {
  if (!el || !el.isConnected || !el.animate || calm()) return () => { };
  const r = el.getBoundingClientRect(); if (!r.width) return () => { };
  const g = el.cloneNode(true);
  ['data-enter', 'data-format-panel', 'data-ai-panel', 'id'].forEach(a => g.removeAttribute(a)); g.setAttribute('aria-hidden', 'true');
  g.style.inset = 'auto';
  Object.assign(g.style, { position: 'fixed', left: r.left + 'px', top: r.top + 'px', width: r.width + 'px', height: r.height + 'px', margin: '0', zIndex: '46', pointerEvents: 'none', animation: 'none' });
  return () => {
    document.body.appendChild(g);
    const a = g.animate([{ opacity: 1, transform: 'none' }, { opacity: 0, transform: 'translateX(16px)' }], { duration: ms, easing: 'cubic-bezier(.4,0,1,1)', fill: 'forwards' });
    a.onfinish = a.oncancel = () => g.remove();
  };
}
export function outside(cb, sel) {
  sel = sel || '[data-bubble],[data-glass-bar],[data-menu],[data-dlg]';
  const h = e => { const t = e.target; if (!(t && t.closest && t.closest(sel))) cb(); };
  document.addEventListener('mousedown', h, true);
  return () => document.removeEventListener('mousedown', h, true);
}
