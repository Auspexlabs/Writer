// Slide HTML uses canvas pixels. DrawingML uses points; translate only CSS lengths,
// never text, URLs or data attributes, and keep enough precision for repeated saves.
export function slideHtmlUnits(html, ptPx, toCanvas = true) {
  return String(html || '').replace(/\bstyle=("[^"]*"|'[^']*')/gi, (attr, quoted) => {
    const style = quoted.slice(1, -1).replace(/(^|;)\s*(font-size|letter-spacing)\s*:\s*(-?[\d.]+)(pt|px)\b/gi, (m, sep, key, n, unit) => {
      const value = Number(n); if (!Number.isFinite(value)) return m;
      const size = toCanvas ? value * (unit.toLowerCase() === 'pt' ? ptPx : ptPx * .75) : value * (unit.toLowerCase() === 'px' ? 1 / ptPx : 1);
      return `${sep}${key}:${Number(size.toFixed(5))}${toCanvas ? 'px' : 'pt'}`;
    });
    return 'style=' + quoted[0] + style + quoted[0];
  });
}

// A whole-box format overrides only that property on its runs. Other rich text,
// paragraph structure, links and fields survive. Partial selections use the DOM editor.
export function formatSlideText(html, patch) {
  const css = { fs: 'font-size', color: 'color', font: 'font-family', bold: 'font-weight', italic: 'font-style', underline: 'text-decoration', align: 'text-align' };
  const keys = Object.keys(css).filter(k => k in patch);
  if (!keys.length || !html) return html;
  const root = new DOMParser().parseFromString(html, 'text/html').body;
  for (const el of root.querySelectorAll('*')) {
    for (const k of keys) {
      if (k === 'underline') { for (const prop of ['text-decoration', 'text-decoration-line']) { const v = el.style.getPropertyValue(prop); if (v) el.style.setProperty(prop, v.replace(/\bunderline\b/g, '').trim() || 'none'); } }
      else el.style.removeProperty(css[k]);
      if (el.tagName === 'FONT') el.removeAttribute({ fs: 'size', color: 'color', font: 'face' }[k] || 'data-unused');
      if (k === 'align') el.removeAttribute('align');
    }
    if ((keys.includes('bold') && /^(B|STRONG)$/.test(el.tagName)) || (keys.includes('italic') && /^(I|EM)$/.test(el.tagName)) || (keys.includes('underline') && el.tagName === 'U')) el.replaceWith(...el.childNodes);
  }
  return root.innerHTML;
}

export function listMarker(n, format) {
  let value = String(n);
  if (/^alpha/.test(format)) { value = ''; for (let i = n; i > 0; i = Math.floor((i - 1) / 26)) value = String.fromCharCode(65 + (i - 1) % 26) + value; }
  if (/^roman/.test(format)) { value = ''; let i = n; for (const [v, s] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']]) while (i >= v) { value += s; i -= v; } }
  if (/Lc/.test(format)) value = value.toLowerCase();
  return /ParenBoth$/.test(format) ? '(' + value + ')' : /ParenR$/.test(format) ? value + ')' : /Period$/.test(format) ? value + '.' : value;
}
