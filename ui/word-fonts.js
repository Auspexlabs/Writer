const family = value => String(value || '').split(',')[0].replace(/["']/g, '').trim();
export function scriptFontsAt(el, root) {
  const west = el.closest('[data-font-west]')?.getAttribute('data-font-west'), east = el.closest('[data-font-ea]')?.getAttribute('data-font-ea');
  const base = family(getComputedStyle(el).fontFamily), fallback = /^wd-font-/.test(base) ? family(getComputedStyle(root).fontFamily) : base;
  return { west: west || fallback || 'Calibri', ea: east || fallback || 'Noto Serif SC' };
}
function setFonts(el, fonts) {
  el.setAttribute('data-font-west', fonts.west); el.setAttribute('data-font-ea', fonts.ea);
  el.style.fontFamily = [fonts.west, fonts.ea].map(f => JSON.stringify(f)).join(',');
}
/** Wrap only the selected pieces of text. Formatting, links, paragraph boundaries and fields stay in place. */
export function setScriptFont(root, range, kind, font) {
  if (!range || !font) return null;
  const doc = root.ownerDocument, selected = [];
  if (range.collapsed) {
    const at = range.startContainer, host = at.nodeType === 3 ? at.parentElement : at, fonts = { ...scriptFontsAt(host, root), [kind]: font };
    const span = doc.createElement('span'); setFonts(span, fonts); span.textContent = '\u200b'; range.insertNode(span);
    const next = doc.createRange(); next.setStart(span.firstChild, 1); next.collapse(true); return next;
  }
  const walk = doc.createTreeWalker(root, 4); let n;
  while ((n = walk.nextNode())) if (!n.parentElement.closest('[contenteditable="false"],script,style') && range.intersectsNode(n)) {
    const a = n === range.startContainer ? range.startOffset : 0, b = n === range.endContainer ? range.endOffset : n.nodeValue.length;
    if (b > a) selected.push({ n, a, b, fonts: { ...scriptFontsAt(n.parentElement, root), [kind]: font } });
  }
  for (const piece of selected) {
    let text = piece.n; if (piece.b < text.length) text.splitText(piece.b); if (piece.a) text = text.splitText(piece.a);
    const span = doc.createElement('span'); setFonts(span, piece.fonts); text.replaceWith(span); span.appendChild(text); piece.text = text;
  }
  if (!selected.length) return range;
  const next = doc.createRange(); next.setStart(selected[0].text, 0); next.setEnd(selected.at(-1).text, selected.at(-1).text.length); return next;
}
/** Local font aliases restrict each face to its script without splitting text or moving the caret. */
export function scriptFontCss(root) {
  const faces = new Map(), rules = new Map();
  const supplied = new Map();
  for (const sheet of root.ownerDocument.styleSheets) {
    let entries; try { entries = sheet.cssRules; } catch { continue; }
    for (const rule of entries) if (rule.type === 5 && !family(rule.style.fontFamily).startsWith('wd-font-')) {
      const name = family(rule.style.fontFamily), list = supplied.get(name) || [];
      const src = rule.style.getPropertyValue('src').replace(/url\(["']?([^"')]+)["']?\)/g, (_, url) => `url(${JSON.stringify(new URL(url, sheet.href || root.ownerDocument.baseURI).href)})`);
      list.push({ src, weight: rule.style.fontWeight || 'normal', style: rule.style.fontStyle || 'normal' }); supplied.set(name, list);
    }
  }
  const alias = (font, script) => {
    const key = script + ':' + font; let hash = 0; for (const c of key) hash = (Math.imul(hash, 31) + c.codePointAt(0)) >>> 0;
    const name = 'wd-font-' + hash.toString(36), range = script === 'ea' ? 'U+2E80-9FFF,U+AC00-D7FF,U+F900-FAFF,U+FF00-FFFF,U+20000-3FFFF' : 'U+0000-2E7F,U+A000-ABFF,U+D800-F8FF,U+FB00-FEFF,U+10000-1FFFF';
    faces.set(key, (supplied.get(font) || [{ src: `local(${JSON.stringify(font)})`, weight: 'normal', style: 'normal' }]).map(f => `@font-face{font-family:${name};src:${f.src};font-weight:${f.weight};font-style:${f.style};unicode-range:${range};font-display:swap}`).join('\n')); return name;
  };
  for (const el of root.querySelectorAll('[data-font-west][data-font-ea]')) {
    const west = el.getAttribute('data-font-west'), ea = el.getAttribute('data-font-ea'), key = JSON.stringify([west, ea]);
    if (!rules.has(key)) rules.set(key, `.wd-ed [data-font-west="${CSS.escape(west)}"][data-font-ea="${CSS.escape(ea)}"],.wd-hf [data-font-west="${CSS.escape(west)}"][data-font-ea="${CSS.escape(ea)}"]{font-family:${alias(west, 'west')},${alias(ea, 'ea')},serif!important}`);
  }
  return [...faces.values(), ...rules.values()].join('\n');
}
