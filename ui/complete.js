// complete.js — AI 自动补全, everywhere a document is typed: after a pause, the words the user is likely to type next appear in
// grey after the caret; Tab takes them, Esc or typing something else drops them, and typing the same letters walks through
// them. The shell installs it once per page (install(document, …)); an editor opts a typing surface in with data-complete:
//   a contenteditable root (the Word body, a slide's text box, a PDF note) — rich: the caret's block gets the grey text as a
//   ::after rule in a <style> of its own, never a node of the document (the editors save their roots' innerHTML);
//   an <input> or <textarea> (a spreadsheet cell and the formula bar, speaker notes, a mind map topic, Markdown source) — field:
//   the grey text is drawn by a transparent copy of the field laid over it.
// Either way a suggestion shows only with the caret at the end of the text it continues (a block, a line, the value). The
// attribute's value is a short description for the model ("a spreadsheet cell"); a surface may set el.__complete(el, at) to
// give the context itself ({ before, after, hint }), or null for "not here" (a formula).

export const DELAY = 600; // ms without a keystroke before asking
const BLOCK = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,pre,td,th,dd,dt,figcaption,div';
const NESTED = 'p,ul,ol,table,li,div,blockquote,pre,h1,h2,h3,h4,h5,h6'; // a ::after goes after these, not at the caret
const ZW = /\u200B/g;
const FIELD_STYLE = ['fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'fontVariant', 'fontStretch', 'letterSpacing', 'wordSpacing', 'lineHeight', 'textTransform',
  'textIndent', 'textAlign', 'direction', 'tabSize', 'paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft', 'borderTopWidth', 'borderRightWidth', 'borderBottomWidth', 'borderLeftWidth', 'boxSizing'];
const GREY = 'var(--ink-3, rgba(29,29,31,0.38))';

/** A CSS string literal for content: */
export function cssString(text) { return '"' + String(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\n\r\f]/g, ' ') + '"'; }

/** The selector of el from root, one :nth-child() per level (">:nth-child(3)>:nth-child(1)"; "" for root itself), or null when el
 *  is not inside root. */
export function pathSelector(root, el) {
  const parts = [];
  let n = el;
  for (; n && n !== root; n = n.parentElement) {
    const p = n.parentElement; if (!p) return null;
    parts.unshift('>:nth-child(' + (Array.prototype.indexOf.call(p.children, n) + 1) + ')');
  }
  return n === root ? parts.join('') : null;
}

/** The text a completion sees with the caret at the end of block: the blocks before it, newest last, and after it (each line a
 *  block), at most max and maxAfter characters. */
export function context(root, block, max = 2000, maxAfter = 400) {
  const text = n => n.textContent.replace(ZW, '');
  let before = text(block), after = '';
  for (let n = block; n && n !== root && before.length < max; n = n.parentElement)
    for (let s = n.previousElementSibling; s && before.length < max; s = s.previousElementSibling) before = text(s) + '\n' + before;
  for (let n = block; n && n !== root && after.length < maxAfter; n = n.parentElement)
    for (let s = n.nextElementSibling; s && after.length < maxAfter; s = s.nextElementSibling) after += text(s) + '\n';
  return { before: before.slice(-max), after: after.slice(0, maxAfter) };
}

/** A field's caret, when it is collapsed at the end of its line (a textarea) or of the value (an input), else -1. */
export function fieldCaret(value, start, end, multiline) {
  if (start !== end || start == null) return -1;
  const rest = value.slice(start);
  return (multiline ? rest.split('\n')[0] : rest).trim() === '' ? start : -1;
}

/** Typing into a suggestion: what is left of it once typed is typed, or null when the typing does not follow it. */
export function advance(text, typed) { return typed && text.length > typed.length && text.startsWith(typed) ? text.slice(typed.length) : null; }

/** The typing surface an event target belongs to: an element with data-complete that can be typed in now, or null. */
export function surfaceOf(t) {
  const el = t && (t.nodeType === 1 ? t : t.parentElement), s = el && el.closest && el.closest('[data-complete]');
  if (!s) return null;
  if (s.tagName === 'INPUT' || s.tagName === 'TEXTAREA') return /^(text|search|)$/i.test(s.getAttribute('type') || '') && !s.readOnly && !s.disabled ? s : null;
  return s.isContentEditable ? s : null;
}

/** Installs the autocomplete on a page. enabled() is asked before each request (the setting, a model set up); request(before,
 *  after, hint, signal) → the suggestion; onError(error) hears a failed request once per message. Returns { clear, state, remove }. */
export function install(doc, { enabled, request, onError, delay = DELAY }) {
  // once per page: 设置 shown inside the shell's page finds the shell's already there (its own window installs its own)
  if (doc.__writerComplete) { const h = doc.__writerComplete; return { clear: h.clear, state: h.state, remove() { } }; }
  const win = doc.defaultView;
  const style = doc.head.appendChild(doc.createElement('style'));
  let mirror = null, timer = 0, ctl = null, ghost = null, composing = false, quietUntil = 0, failures = 0, told = '', seq = 0;
  // ghost: { surface, field, block, at, text, expect, checking } — field: the surface is an input or textarea; block (rich) the
  // caret's block; at (field) the caret; expect: the block's text or the field's value when shown

  const field = s => s.tagName === 'INPUT' || s.tagName === 'TEXTAREA';
  const rootId = s => s.getAttribute('data-ghost-root') || (s.setAttribute('data-ghost-root', 'ac' + (++seq) + Math.random().toString(36).slice(2, 6)), s.getAttribute('data-ghost-root'));

  const draw = () => {
    style.textContent = '';
    if (mirror) mirror.style.display = 'none';
    if (!ghost) return;
    if (!ghost.field) {
      const sel = pathSelector(ghost.surface, ghost.block); if (sel === null) { ghost = null; return; }
      style.textContent = `[data-ghost-root="${rootId(ghost.surface)}"]${sel}::after{content:${cssString(ghost.text)};color:${GREY};white-space:pre-wrap;pointer-events:none;-webkit-user-select:none;user-select:none}`;
      return;
    }
    // a transparent copy of the field's text up to the caret, laid out as the field lays it out (scrolled as it is), and the grey
    // text at the caret's place: absolutely positioned with no offsets, it sits where it would have flowed and runs on to the
    // right on one line, past the field's edge when the field is narrow (a mind map topic, a spreadsheet cell)
    const el = ghost.surface, cs = win.getComputedStyle(el), r = el.getBoundingClientRect(), area = el.tagName === 'TEXTAREA';
    const m = mirror || (mirror = doc.body.appendChild(doc.createElement('div')));
    m.setAttribute('aria-hidden', 'true');
    m.style.cssText = 'position:fixed;pointer-events:none;overflow:visible;z-index:2147483000;margin:0;background:transparent;color:transparent;border-style:solid;border-color:transparent;-webkit-user-select:none;user-select:none';
    for (const k of FIELD_STYLE) m.style[k] = cs[k];
    Object.assign(m.style, { left: r.left + 'px', top: r.top + 'px', width: r.width + 'px', height: r.height + 'px', display: 'block' });
    const inner = doc.createElement('div');
    inner.style.cssText = `width:100%;transform:translate(${-el.scrollLeft}px,${-el.scrollTop}px);white-space:${area ? 'pre-wrap' : 'pre'};overflow-wrap:${area ? 'break-word' : 'normal'}`;
    if (!area) { // an input centres its one line: so does the copy
      const h = r.height - ['paddingTop', 'paddingBottom', 'borderTopWidth', 'borderBottomWidth'].reduce((a, k) => a + (parseFloat(cs[k]) || 0), 0);
      inner.style.lineHeight = Math.max(0, h) + 'px';
    }
    const g = doc.createElement('span'); g.textContent = ghost.text; g.style.cssText = `position:absolute;color:${GREY};white-space:pre`;
    inner.append(doc.createTextNode(el.value.slice(0, ghost.at)), g);
    m.replaceChildren(inner);
  };
  const hide = () => { if (ghost) { ghost = null; draw(); } };
  const clear = () => { clearTimeout(timer); timer = 0; if (ctl) { ctl.abort(); ctl = null; } hide(); };

  /** Where the caret is in a surface, when a suggestion can go there: { block } (rich) or { at } (field), else null. */
  const caretIn = s => {
    if (!s) return null;
    if (field(s)) {
      if (doc.activeElement !== s) return null;
      const at = fieldCaret(s.value, s.selectionStart, s.selectionEnd, s.tagName === 'TEXTAREA');
      return at < 0 ? null : { at };
    }
    const sel = doc.getSelection && doc.getSelection();
    if (!sel || !sel.rangeCount || !sel.isCollapsed) return null;
    const r = sel.getRangeAt(0), node = r.startContainer;
    if (!s.contains(node)) return null;
    const el = node.nodeType === 1 ? node : node.parentElement;
    let block = el && el.closest(BLOCK);
    if (!block || !s.contains(block)) block = s;
    if (block !== s && block.querySelector(NESTED)) return null;
    const last = block.lastChild; if (last && last.nodeName === 'BR' && block.textContent.replace(ZW, '') !== '') return null;
    const tail = doc.createRange(); tail.setStart(r.startContainer, r.startOffset); tail.setEnd(block, block.childNodes.length);
    return tail.toString().replace(ZW, '') === '' ? { block } : null;
  };
  const textOf = (s, at) => field(s) ? s.value : at.block.textContent;
  const same = (s, at) => ghost && ghost.surface === s && at && (ghost.field ? at.at === ghost.at : at.block === ghost.block) && textOf(s, at) === ghost.expect;

  /** The context of a request: the surface's own (el.__complete), or the text around the caret with the surface's description. */
  const contextOf = (s, at) => {
    const hook = s.__complete; let x;
    if (typeof hook === 'function') { x = hook(s, at); if (x === null) return null; }
    const hint = s.getAttribute('data-complete'); const base = hint && hint !== '1' && hint !== 'true' ? hint : '';
    if (!x) x = field(s) ? { before: s.value.slice(Math.max(0, at.at - 2000), at.at), after: s.value.slice(at.at, at.at + 400) } : context(s, at.block);
    return { before: x.before || '', after: x.after || '', hint: x.hint != null ? x.hint : base };
  };

  let target = null; // the surface typed in last
  const schedule = s => { clear(); target = s; if (s && !composing && enabled()) timer = setTimeout(ask, delay); };
  const ask = async () => {
    timer = 0;
    const s = target;
    if (!s || composing || !enabled() || Date.now() < quietUntil || !s.isConnected) return;
    if (!field(s) && doc.activeElement !== s && !s.contains(doc.activeElement)) return;
    const at = caretIn(s); if (!at) return;
    const expect = textOf(s, at);
    if (expect.replace(ZW, '').trim().length < 2 && !s.__complete) return; // an empty line: nothing to go on
    const cx = contextOf(s, at); if (!cx || !cx.before.trim()) return;
    const c = ctl = new AbortController();
    let text = '';
    try { text = await request(cx.before, cx.after, cx.hint, c.signal); failures = 0; }
    catch (e) {
      if (ctl === c) ctl = null;
      if (e && e.name === 'AbortError') return;
      failures++; quietUntil = Date.now() + Math.min(120000, 10000 * 2 ** (failures - 1)); // a failing service is asked less and less
      const msg = String((e && e.message) || e); if (msg !== told) { told = msg; onError && onError(e); }
      return;
    }
    if (ctl !== c) return;
    ctl = null;
    const now = caretIn(s);
    if (!text || !enabled() || !now || textOf(s, now) !== expect || (field(s) ? now.at !== at.at : now.block !== at.block)) return;
    ghost = { surface: s, field: field(s), block: at.block, at: at.at, text, expect, checking: false }; draw();
  };

  const accept = () => {
    const s = ghost.surface, text = ghost.text; clear();
    if (field(s)) s.focus();
    doc.execCommand('insertText', false, text); // through the editor's own input path: its undo, tracked changes and save
  };
  const onKey = e => {
    if (!ghost || e.isComposing || e.keyCode === 229 || !ghost.surface.contains(e.target)) return;
    const plain = !e.shiftKey && !e.altKey && !e.metaKey && !e.ctrlKey;
    if (e.key === 'Tab' && plain) { e.preventDefault(); e.stopPropagation(); accept(); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); clear(); }
  };
  const onBeforeInput = e => {
    const s = surfaceOf(e.target);
    if (!s) { if (ghost) clear(); return; }
    if (composing || e.isComposing) return;
    const g = ghost, rest = g && g.surface === s && e.inputType === 'insertText' ? advance(g.text, e.data) : null;
    if (g && rest !== null) {
      const want = g.field ? g.expect.slice(0, g.at) + e.data + g.expect.slice(g.at) : g.expect + e.data; g.checking = true;
      setTimeout(() => { // once the letters are in (by the browser, or by the editor itself when it tracks changes)
        if (ghost !== g) return;
        g.checking = false;
        const at = caretIn(s);
        if (at && textOf(s, at) === want && (g.field ? at.at === g.at + e.data.length : at.block === g.block)) { g.text = rest; g.expect = want; if (g.field) g.at = at.at; draw(); }
        else schedule(s);
      }, 0);
      return;
    }
    hide();
    setTimeout(() => schedule(s), 0);
  };
  const check = () => { if (ghost && !ghost.checking && !same(ghost.surface, caretIn(ghost.surface))) clear(); };
  const onCompStart = e => { if (surfaceOf(e.target)) { composing = true; clear(); } };
  const onCompEnd = e => { composing = false; const s = surfaceOf(e.target); if (s) setTimeout(() => schedule(s), 0); };
  const onOut = e => { if (ghost && !ghost.surface.contains(e.relatedTarget)) clear(); };
  const onScroll = () => { if (ghost && ghost.field) clear(); };

  const on = [['keydown', onKey, true], ['beforeinput', onBeforeInput, true], ['compositionstart', onCompStart, true], ['compositionend', onCompEnd, true],
    ['focusout', onOut, true], ['selectionchange', check, false], ['keyup', check, true], ['pointerup', check, true], ['scroll', onScroll, true]];
  for (const [t, f, c] of on) doc.addEventListener(t, f, c);
  win.addEventListener('resize', clear);
  return doc.__writerComplete = {
    clear,
    state: () => ghost ? { text: ghost.text, surface: ghost.surface, block: ghost.block || null, at: ghost.field ? ghost.at : null } : null,
    remove() {
      clear();
      for (const [t, f, c] of on) doc.removeEventListener(t, f, c);
      win.removeEventListener('resize', clear);
      style.remove(); if (mirror) mirror.remove();
      delete doc.__writerComplete;
    }
  };
}
