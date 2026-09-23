// Keeps every <select data-val> showing its intended value after options render.
let queued = false;
export function syncSelects() {
  if (queued) return; queued = true;
  requestAnimationFrame(() => { queued = false; document.querySelectorAll('select[data-val]').forEach(el => { const v = el.getAttribute('data-val'); if (el.value !== v && Array.from(el.options).some(o => o.value === v)) el.value = v; }); });
}
if (!window.__selSync) {
  window.__selSync = new MutationObserver(syncSelects);
  window.__selSync.observe(document.documentElement, { subtree: true, childList: true, attributes: true, attributeFilter: ['data-val'] });
  syncSelects();
}
