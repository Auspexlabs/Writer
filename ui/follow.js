// follow.js — the view follows the caret ("镜头跟随"). While the user types in a zoomed-in editor the scroll
// container eases so the caret stays inside a comfort box; a line wrap brings the view back to the line start;
// mouse selections and manual scrolling (wheel / scrollbar / touch) are never fought.
// Shared by the Word, Markdown and Slide editors. planScroll is pure and runs in node for tests.

const DUR = 140, PAUSE = 800; // easing constants match MindMapEditor's camera pan
const ease = p => 1 - Math.pow(1 - p, 3);
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

/**
 * Scroll offset that puts `caret` (client rect) back inside the comfort box of `view` (the scroller's client rect).
 * Visible box = view minus `insets` (overlays); comfort box = the visible box inset by padX / padY on each side.
 * Leaving the box targets the box edge, clamped to [0, max]; a line wrap (caret x moved left by more than half
 * the visible width since the previous update, `prevX` in content coordinates) targets the left box edge.
 * An axis whose content fits (max ≤ 0) is left untouched.
 */
export function planScroll(caret, view, cur, max, opts = {}) {
  const { padX = 0.12, padY = 0.2, insets = {}, prevX } = opts;
  const il = insets.left || 0, it = insets.top ?? 56, ir = insets.right || 0, ib = insets.bottom ?? 48;
  const vl = view.left + il, vt = view.top + it, vw = view.width - il - ir, vh = view.height - it - ib;
  const cl = caret.left, ct = caret.top, cr = caret.right ?? cl + (caret.width || 0), cb = caret.bottom ?? ct + (caret.height || 0);
  let { x, y } = cur;
  if (max.x > 0 && vw > 0) {
    const L = vl + vw * padX, R = vl + vw * (1 - padX), cx = cl - view.left + cur.x;
    if ((prevX != null && prevX - cx > vw / 2) || cl < L) x = cur.x + cl - L;
    else if (cr > R) x = cur.x + Math.min(cr - R, cl - L); // wider than the box: keep its start visible
    x = clamp(x, 0, max.x);
  }
  if (max.y > 0 && vh > 0) {
    const T = vt + vh * padY, B = vt + vh * (1 - padY);
    if (ct < T) y = cur.y + ct - T; else if (cb > B) y = cur.y + Math.min(cb - B, ct - T);
    y = clamp(y, 0, max.y);
  }
  return { x, y };
}

/** Client rect of the collapsed selection when it sits inside `root`, else null. Falls back to the node beside the caret where a range has no rects (empty lines, element boundaries). */
export function caretRect(root) {
  const s = window.getSelection(); if (!root || !s || !s.rangeCount || !s.isCollapsed || !root.contains(s.anchorNode)) return null;
  const r = s.getRangeAt(0), rc = r.getClientRects()[0]; if (rc && rc.height) return rc;
  const c = r.startContainer, after = c.childNodes[r.startOffset]; let n = after || c.childNodes[r.startOffset - 1] || c;
  while (n.nodeType === 1 && n.childNodes.length) n = after ? n.firstChild : n.lastChild; // an element boundary: the text beside the caret
  if (n.nodeType === 3) { const rr = document.createRange(); rr.selectNodeContents(n); rr.collapse(!!after); const x = rr.getClientRects()[0]; if (x && x.height) return x; }
  return (n.nodeType === 1 ? n : n.parentNode).getBoundingClientRect();
}

const MIRROR = ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'lineHeight', 'letterSpacing', 'wordSpacing', 'tabSize', 'textIndent', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft', 'wordBreak'];
/** Client rect of the caret in a focused textarea, measured with a mirror div (same font, padding and content width). */
export function textareaCaret(ta) {
  if (!ta || document.activeElement !== ta || ta.selectionStart !== ta.selectionEnd) return null;
  const cs = getComputedStyle(ta), rc = ta.getBoundingClientRect(), d = document.createElement('div'), m = document.createElement('span'), i = ta.selectionStart;
  MIRROR.forEach(k => { d.style[k] = cs[k]; });
  Object.assign(d.style, { position: 'absolute', left: '-9999px', top: '0', visibility: 'hidden', whiteSpace: 'pre-wrap', overflowWrap: 'break-word', boxSizing: 'border-box', width: ta.clientWidth + 'px' });
  d.textContent = ta.value.slice(0, i); m.textContent = /^\S*/.exec(ta.value.slice(i))[0] || '\u200b'; d.appendChild(m); // rest of the word keeps a wrapped word on its real line
  document.body.appendChild(d);
  const left = rc.left + ta.clientLeft + m.offsetLeft - ta.scrollLeft, top = rc.top + ta.clientTop + m.offsetTop - ta.scrollTop, h = m.offsetHeight;
  d.remove();
  return { left, top, right: left, bottom: top + h, width: 0, height: h };
}

/**
 * Follows the caret inside `scrollEl`. `getCaretRect(reason)` returns the caret's client rect or null (default: the
 * collapsed selection inside scrollEl). opts: padX, padY, insets, scroller() → the element to scroll (default scrollEl,
 * for a textarea that scrolls on its own). Returns { el, update(reason), detach() }.
 * Triggers wired here: input / beforeinput / compositionupdate → 'type'; selectionchange → 'caret'; a mouse press in the
 * editor suppresses 'caret' updates until the next key; wheel / scrollbar / touch pauses everything for 800 ms.
 */
export function attach(scrollEl, getCaretRect, opts = {}) {
  const o = Object.assign({ scroller: () => scrollEl }, opts), rect = getCaretRect || (() => caretRect(scrollEl));
  let raf = 0, anim = 0, pauseUntil = 0, mouse = false, prevX = null, reason = '';
  const stop = () => { if (anim) cancelAnimationFrame(anim); anim = 0; };
  const pause = () => { pauseUntil = performance.now() + PAUSE; stop(); };
  /** Eased scroll over ~140 ms; a call mid-flight retargets from wherever the view is now; beyond 1.5 viewports it jumps. */
  const go = (el, to) => {
    stop(); const from = { x: el.scrollLeft, y: el.scrollTop }, dx = to.x - from.x, dy = to.y - from.y;
    if (Math.abs(dx) > 1.5 * el.clientWidth || Math.abs(dy) > 1.5 * el.clientHeight) { el.scrollLeft = to.x; el.scrollTop = to.y; return; }
    const t0 = performance.now();
    const step = now => { const e = ease(clamp((now - t0) / DUR, 0, 1)); el.scrollLeft = from.x + dx * e; el.scrollTop = from.y + dy * e; anim = e < 1 ? requestAnimationFrame(step) : 0; };
    anim = requestAnimationFrame(step);
  };
  const measure = () => {
    raf = 0; const el = o.scroller(); const rc = el && rect(reason); if (!rc) return;
    const view = el.getBoundingClientRect(), cur = { x: el.scrollLeft, y: el.scrollTop }, max = { x: el.scrollWidth - el.clientWidth, y: el.scrollHeight - el.clientHeight };
    const to = planScroll(rc, view, cur, max, Object.assign({ prevX }, o));
    prevX = rc.left - view.left + cur.x;
    if (Math.abs(to.x - cur.x) > 0.5 || Math.abs(to.y - cur.y) > 0.5) go(el, to);
  };
  const update = r => {
    if (r !== 'caret') mouse = false; if (r === 'zoom') stop(); // the zoom anchor set the scroll: never continue an older glide
    if (mouse || performance.now() < pauseUntil) return;
    reason = r; if (!raf) raf = requestAnimationFrame(measure);
  };
  const H = [
    [scrollEl, 'beforeinput', () => update('type')], [scrollEl, 'input', () => update('type')], [scrollEl, 'compositionupdate', () => update('type')],
    [scrollEl, 'keydown', () => { mouse = false; }],
    [scrollEl, 'mousedown', e => { mouse = true; prevX = null; if (e.target === scrollEl && (e.offsetX >= scrollEl.clientWidth || e.offsetY >= scrollEl.clientHeight)) pause(); }],
    [scrollEl, 'wheel', e => { if (!(e.ctrlKey || e.metaKey)) pause(); }], [scrollEl, 'touchmove', pause],
    [document, 'selectionchange', () => update('caret')]
  ];
  H.forEach(([t, k, f]) => t.addEventListener(k, f, { passive: true }));
  return { el: scrollEl, update, detach: () => { H.forEach(([t, k, f]) => t.removeEventListener(k, f)); stop(); if (raf) cancelAnimationFrame(raf); raf = 0; } };
}
