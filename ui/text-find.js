// Find across formatting runs while preserving document structure and atomic fields/objects.
export function editableText(root) {
  const runs = []; let text = '';
  const walk = n => {
    if (n.nodeType === 3) { runs.push({ node: n, start: text.length, end: text.length + n.nodeValue.length }); text += n.nodeValue; return; }
    if (n.nodeType !== 1) return;
    if (n.matches('script,style,del,[contenteditable="false"],[data-toc],[data-bib]')) { text += '\ufffc'; return; }
    if (n.tagName === 'BR') { text += '\n'; return; }
    for (const c of n.childNodes) walk(c);
    if (n !== root && /^(P|DIV|LI|H[1-6]|BLOCKQUOTE|PRE|TD|TH)$/.test(n.tagName)) text += '\n';
  };
  walk(root); return { runs, text };
}
export function textMatches(text, query, options = {}) {
  if (!query) return [];
  const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.caseSensitive ? 'gu' : 'giu');
  const word = s => /[\p{L}\p{N}_]/u.test(s || '');
  const at = i => i < 0 || i >= text.length ? '' : String.fromCodePoint(text.codePointAt(i));
  const before = i => { const c = text.charCodeAt(i - 1); return at(c >= 0xdc00 && c <= 0xdfff ? i - 2 : i - 1); };
  return [...text.matchAll(re)].filter(m => !options.wholeWord || (!word(before(m.index)) && !word(at(m.index + m[0].length))))
    .map(m => ({ index: m.index, length: m[0].length }));
}
export function findTextRanges(root, query, options) {
  const { runs, text } = editableText(root);
  return textMatches(text, query, options).map(hit => {
    const touched = runs.filter(r => r.end > hit.index && r.start < hit.index + hit.length);
    if (!touched.length) return null;
    const first = touched[0], last = touched.at(-1), range = root.ownerDocument.createRange();
    range.setStart(first.node, Math.max(0, hit.index - first.start)); range.setEnd(last.node, Math.min(last.end, hit.index + hit.length) - last.start);
    return { ...hit, range, runs: touched.map(r => ({ node: r.node, start: Math.max(0, hit.index - r.start), end: Math.min(r.end, hit.index + hit.length) - r.start })) };
  }).filter(Boolean);
}
export function replaceTextRanges(hits, replacement) {
  for (const hit of hits.slice().reverse()) hit.runs.forEach((r, i) => { r.node.nodeValue = r.node.nodeValue.slice(0, r.start) + (i === 0 ? replacement : '') + r.node.nodeValue.slice(r.end); });
  return hits.length;
}
