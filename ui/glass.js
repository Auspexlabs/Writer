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
export function outside(cb, sel) {
  sel = sel || '[data-bubble],[data-glass-bar],[data-menu],[data-dlg]';
  const h = e => { const t = e.target; if (!(t && t.closest && t.closest(sel))) cb(); };
  document.addEventListener('mousedown', h, true);
  return () => document.removeEventListener('mousedown', h, true);
}
