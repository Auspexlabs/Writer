// tabs.js — a document's tab dragged off the title strip (ui/mac.dc.html, ui/win.dc.html) tears off into a window of its
// own, as in a browser. tornOff is pure and runs in node for tests.

/** Whether a tab let go at client point p (px) leaves the strip for a window of its own: outside the window (view: its
 *  inner width and height), or more than 40px below the strip (strip: its client rect). A single tab never tears off. */
export function tornOff(p, strip, view, tabs) {
  return tabs > 1 && (p.x < 0 || p.y < 0 || p.x >= view.w || p.y >= view.h || p.y > strip.bottom + 40);
}

/** Follows the tab pressed in pointerdown `e` (not from its close button): onMove(dx, dy) while it moves and (0, 0) when it
 *  is let go, then onTear() when it was let go off the strip. With `dock` (the desktop app's bridge to main.rs tab_drag and
 *  dock_tab): a window's only tab carries its window along, the tab strip of another window under the pointer shows where
 *  the tab would go, and a tab let go there moves into that window (dock.to) instead of into a window of its own. */
export function dragTab(e, tabs, onMove, onTear, dock) {
  if (e.button !== 0 || e.target.closest('button')) return;
  // No text selection and no native drag while a tab moves: WebKit arms a drag-selection on the press even over
  // user-select:none and extends it into the page below once the tab leaves the strip. The click still comes.
  e.preventDefault();
  const quiet = ev => ev.preventDefault();
  document.addEventListener('selectstart', quiet, true); document.addEventListener('dragstart', quiet, true);
  const el = e.currentTarget, strip = el.parentElement.getBoundingClientRect(), x0 = e.clientX, y0 = e.clientY, carry = !!dock && tabs === 1;
  if (dock) dock.start(carry);
  let queued = false; // one native step per frame
  const move = ev => {
    if (!carry) onMove(ev.clientX - x0, ev.clientY - y0); // a carried window moves under the pointer: its tab stays put
    if (dock && !queued) { queued = true; requestAnimationFrame(() => { queued = false; dock.move(); }); }
  };
  const end = async ev => {
    window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', end); window.removeEventListener('pointercancel', end);
    document.removeEventListener('selectstart', quiet, true); document.removeEventListener('dragstart', quiet, true);
    onMove(0, 0);
    const target = dock ? await dock.end() : null; // another window's tab strip under the pointer
    if (ev.type !== 'pointerup') return;
    const eat = () => { const c = x => x.stopPropagation(); el.addEventListener('click', c, true); setTimeout(() => el.removeEventListener('click', c, true)); }; // the click that may follow would bring the tab to the front here again
    if (target) { eat(); dock.to(target); return; }
    if (!tornOff({ x: ev.clientX, y: ev.clientY }, strip, { w: innerWidth, h: innerHeight }, tabs)) return;
    eat(); onTear();
  };
  try { el.setPointerCapture(e.pointerId); } catch (x) { } // the moves and the release keep coming outside the window
  window.addEventListener('pointermove', move); window.addEventListener('pointerup', end); window.addEventListener('pointercancel', end);
}
