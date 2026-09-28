// Excel uses quoted fields when a cell includes tabs, newlines, or quotes.
export const writeTSV = rows => rows.map(row => row.map(v => { const s = String(v ?? ''); return /[\t\r\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join('\t')).join('\n');
export function readTSV(text, separator = '\t') {
  const rows = [], row = []; let value = '', quoted = false, start = true;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) { if (c === '"') { if (text[i + 1] === '"') { value += '"'; i++; } else quoted = false; } else value += c; continue; }
    if (c === '"' && start) { quoted = true; start = false; continue; }
    if (c === separator) { row.push(value); value = ''; start = true; continue; }
    if (c === '\r' || c === '\n') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(value); rows.push(row.splice(0)); value = ''; start = true; continue; }
    value += c; start = false;
  }
  if (row.length || value || !rows.length || !/[\r\n]$/.test(text)) { row.push(value); rows.push(row); }
  return rows;
}

export function readTableHTML(html) {
  if (!html || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(html, 'text/html'), table = doc.querySelector('table'); if (!table) return null;
  const rules = [];
  for (const style of doc.querySelectorAll('style')) for (const m of style.textContent.matchAll(/([^{}]+)\{([^{}]+)\}/g)) {
    const node = doc.createElement('span'); node.setAttribute('style', m[2]); rules.push([m[1].trim(), node.style]);
  }
  const color = v => { const rgb = /^rgba?\((\d+)[, ]+(\d+)[, ]+(\d+)/.exec(v); return rgb ? '#' + rgb.slice(1).map(n => (+n).toString(16).padStart(2, '0')).join('') : /^#[0-9a-f]{3,8}$/i.test(v) ? v : null; };
  const cells = [], merges = [], occupied = new Set();
  for (const [r, row] of [...table.rows].entries()) {
    const line = cells[r] || (cells[r] = []); let c = 0;
    for (const el of row.cells) {
      while (occupied.has(r + ':' + c)) c++;
      const style = doc.createElement('span').style;
      for (const [selector, decl] of rules) { try { if (el.matches(selector)) style.cssText += ';' + decl.cssText; } catch { /* unsupported Office selectors */ } }
      style.cssText += ';' + el.style.cssText;
      const font = el.querySelector('font'); if (font) { if (font.color) style.color = font.color; if (font.face) style.fontFamily = font.face; }
      const s = {}, read = (key, value) => { if (value) s[key] = value; };
      read('font', style.fontFamily.replace(/^['"]|['"]$/g, '')); if (style.fontSize) s.fs = parseFloat(style.fontSize) * (/px$/.test(style.fontSize) ? 0.75 : 1);
      if (style.fontWeight === 'bold' || +style.fontWeight >= 600 || el.querySelector('b,strong')) s.b = true;
      if (style.fontStyle === 'italic' || el.querySelector('i,em')) s.i = true;
      if (/underline/.test(style.textDecoration) || el.querySelector('u')) s.u = true;
      read('color', color(style.color)); read('fill', color(style.backgroundColor)); read('align', style.textAlign); read('va', style.verticalAlign === 'middle' ? 'middle' : /^(top|bottom)$/.test(style.verticalAlign) ? style.verticalAlign : null);
      if (/normal|pre-wrap/.test(style.whiteSpace)) s.wrap = true;
      const link = el.querySelector('a[href]'); if (link && /^(https?:|mailto:|#)/i.test(link.getAttribute('href'))) s.link = link.getAttribute('href');
      const bd = {}; for (const side of ['Top', 'Right', 'Bottom', 'Left']) if (!['', 'none', 'hidden'].includes(style['border' + side + 'Style'])) { bd[side.toLowerCase()] = style['border' + side + 'Style'] === 'double' ? 'double' : parseFloat(style['border' + side + 'Width']) > 1.5 ? 'medium' : 'thin'; read('bdc', color(style['border' + side + 'Color'])); } if (Object.keys(bd).length) s.bd = bd;
      const content = el.cloneNode(true); content.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
      line[c] = { v: content.textContent.replace(/\r\n?/g, '\n'), ...(Object.keys(s).length ? { s } : {}) };
      const rs = Math.max(1, el.rowSpan), cs = Math.max(1, el.colSpan);
      if (rs > 1 || cs > 1) merges.push({ r, c, rs, cs });
      for (let y = r; y < r + rs; y++) for (let x = c; x < c + cs; x++) { occupied.add(y + ':' + x); if (y !== r || x !== c) (cells[y] || (cells[y] = []))[x] = null; }
      c += cs;
    }
  }
  return { cells, merges };
}
