// tabs.js — a document's tab dragged off the title strip (ui/mac.dc.html, ui/win.dc.html) tears off into a window of its
// own, as in a browser. tornOff is pure and runs in node for tests.

/** Whether a tab let go at client point p (px) leaves the strip for a window of its own: outside the window (view: its
 *  inner width and height), or more than 40px below the strip (strip: its client rect). A single tab never tears off. */
export function tornOff(p, strip, view, tabs) {
  return tabs > 1 && (p.x < 0 || p.y < 0 || p.x >= view.w || p.y >= view.h || p.y > strip.bottom + 40);
}

/** Follows the tab pressed in pointerdown `e` (not from its close button): onMove(dx, dy) while it moves and (0, 0) when it
 *  is let go, then onTear() when it was let go off the strip. */
export function dragTab(e, tabs, onMove, onTear) {
  if (e.button !== 0 || e.target.closest('button')) return;
  // No text selection and no native drag while a tab moves: WebKit arms a drag-selection on the press even over
  // user-select:none and extends it into the page below once the tab leaves the strip. The click still comes.
  e.preventDefault();
  const quiet = ev => ev.preventDefault();
  document.addEventListener('selectstart', quiet, true); document.addEventListener('dragstart', quiet, true);
  const el = e.currentTarget, strip = el.parentElement.getBoundingClientRect(), x0 = e.clientX, y0 = e.clientY;
  const move = ev => onMove(ev.clientX - x0, ev.clientY - y0);
  const end = ev => {
    window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', end); window.removeEventListener('pointercancel', end);
    document.removeEventListener('selectstart', quiet, true); document.removeEventListener('dragstart', quiet, true);
    onMove(0, 0);
    if (ev.type !== 'pointerup' || !tornOff({ x: ev.clientX, y: ev.clientY }, strip, { w: innerWidth, h: innerHeight }, tabs)) return;
    const eat = c => c.stopPropagation(); // the click that may follow would bring the tab to the front here again
    el.addEventListener('click', eat, true); setTimeout(() => el.removeEventListener('click', eat, true));
    onTear();
  };
  try { el.setPointerCapture(e.pointerId); } catch (x) { } // the moves and the release keep coming outside the window
  window.addEventListener('pointermove', move); window.addEventListener('pointerup', end); window.addEventListener('pointercancel', end);
}
