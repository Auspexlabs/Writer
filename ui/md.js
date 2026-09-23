// Minimal Markdown → HTML (GFM-ish): headings, emphasis, code, links, images, lists, tasks, tables, quotes, hr.
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = s => esc(s).replace(/"/g, '&quot;');
export function inline(s) {
  const codes = [];
  s = s.replace(/`([^`\n]+)`/g, (m, c) => { codes.push(c); return '\u0000' + (codes.length - 1) + '\u0000'; });
  s = esc(s);
  s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+&quot;([^&]*)&quot;)?\)/g, (m, a, u) => `<img alt="${a}" src="${u}">`);
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
  s = s.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
  s = s.replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>').replace(/__(?=\S)([\s\S]*?\S)__/g, '<strong>$1</strong>');
  s = s.replace(/(^|[^*])\*(?=\S)([^*\n]*?\S)\*(?!\*)/g, '$1<em>$2</em>').replace(/(^|[^\w])_(?=\S)([^_\n]*?\S)_(?!\w)/g, '$1<em>$2</em>');
  s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<del>$1</del>').replace(/==(?=\S)([\s\S]*?\S)==/g, '<mark>$1</mark>');
  s = s.replace(/ {2,}\n/g, '<br>').replace(/\n/g, ' ');
  return s.replace(/\u0000(\d+)\u0000/g, (m, i) => '<code>' + esc(codes[+i]) + '</code>');
}
const RE = {
  fence: /^\s*(```|~~~)\s*([\w+-]*)\s*$/, head: /^(#{1,6})\s+(.*?)\s*#*\s*$/, hr: /^\s*([-*_])(\s*\1){2,}\s*$/,
  quote: /^\s*>\s?(.*)$/, list: /^(\s*)([-*+]|\d+[.)])\s+(.*)$/, tsep: /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/
};
const splitRow = l => l.trim().replace(/^\||\|$/g, '').split(/(?<!\\)\|/).map(c => c.trim().replace(/\\\|/g, '|'));
let hid = 0;
function blocks(lines, base, top) {
  let out = '', i = 0;
  const L = n => top ? ` data-line="${base + n}"` : '';
  while (i < lines.length) {
    const ln = lines[i];
    if (!ln.trim()) { i++; continue; }
    let m;
    if ((m = RE.fence.exec(ln))) { const start = i, fence = m[1], lang = m[2]; const body = []; i++; while (i < lines.length && !lines[i].trim().startsWith(fence)) body.push(lines[i++]); i++; out += `<pre${L(start)}><code${lang ? ` data-lang="${attr(lang)}"` : ''}>${esc(body.join('\n'))}</code></pre>`; continue; }
    if ((m = RE.head.exec(ln))) { const n = m[1].length; out += `<h${n}${L(i)} id="h-${hid++}">${inline(m[2])}</h${n}>`; i++; continue; }
    if (RE.hr.test(ln)) { out += `<hr${L(i)}>`; i++; continue; }
    if (RE.quote.test(ln)) { const start = i, body = []; while (i < lines.length && (RE.quote.test(lines[i]) || (lines[i].trim() && body.length && !RE.list.test(lines[i]) && !RE.head.test(lines[i])))) { const q = RE.quote.exec(lines[i]); body.push(q ? q[1] : lines[i]); i++; } out += `<blockquote${L(start)}>${blocks(body, 0, false)}</blockquote>`; continue; }
    if (ln.includes('|') && i + 1 < lines.length && RE.tsep.test(lines[i + 1])) {
      const start = i, head = splitRow(ln), al = splitRow(lines[i + 1]).map(c => c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : ''); i += 2;
      const rows = []; while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(splitRow(lines[i++]));
      const td = (t, c, j) => `<${t}${al[j] ? ` style="text-align:${al[j]}"` : ''}>${inline(c || '')}</${t}>`;
      out += `<table${L(start)}><thead><tr>${head.map((c, j) => td('th', c, j)).join('')}</tr></thead><tbody>${rows.map(r => '<tr>' + head.map((_, j) => td('td', r[j], j)).join('') + '</tr>').join('')}</tbody></table>`;
      continue;
    }
    if (RE.list.test(ln)) {
      const items = []; const start = i;
      while (i < lines.length) {
        const l = lines[i]; const lm = RE.list.exec(l);
        if (lm) { items.push({ ind: lm[1].replace(/\t/g, '    ').length, ord: /\d/.test(lm[2]), num: parseInt(lm[2]) || 1, text: lm[3], line: base + i }); i++; continue; }
        if (l.trim() && /^\s+/.test(l) && items.length) { items[items.length - 1].text += '\n' + l.trim(); i++; continue; }
        if (!l.trim() && i + 1 < lines.length && RE.list.test(lines[i + 1])) { i++; continue; }
        break;
      }
      const build = (arr, lvl) => {
        if (!arr.length) return '';
        const ord = arr[0].ord; let h = `<${ord ? 'ol' : 'ul'}${lvl === 0 ? L(start) : ''}${ord && arr[0].num !== 1 ? ` start="${arr[0].num}"` : ''}>`;
        let k = 0;
        while (k < arr.length) {
          const it = arr[k], kids = []; k++;
          while (k < arr.length && arr[k].ind > it.ind) kids.push(arr[k++]);
          const tm = /^\[([ xX])\]\s+(.*)$/s.exec(it.text);
          const body = tm ? `<input type="checkbox" data-task="${it.line}"${tm[1] !== ' ' ? ' checked' : ''}> <span>${inline(tm[2])}</span>` : inline(it.text);
          h += `<li${tm ? ' class="task"' : ''} data-li="${it.line}">${body}${build(kids, lvl + 1)}</li>`;
        }
        return h + `</${ord ? 'ol' : 'ul'}>`;
      };
      out += build(items, 0); continue;
    }
    const start = i, para = [];
    while (i < lines.length && lines[i].trim() && !RE.fence.test(lines[i]) && !RE.head.test(lines[i]) && !RE.hr.test(lines[i]) && !RE.quote.test(lines[i]) && !RE.list.test(lines[i]) && !(lines[i].includes('|') && i + 1 < lines.length && RE.tsep.test(lines[i + 1]))) para.push(lines[i++]);
    out += `<p${L(start)}>${inline(para.join('\n'))}</p>`;
  }
  return out;
}
export function mdToHtml(src) { hid = 0; return blocks(String(src || '').replace(/\r\n?/g, '\n').split('\n'), 0, true); }
export function headings(src) {
  const out = []; let inFence = false;
  String(src || '').split('\n').forEach((l, i) => { if (RE.fence.test(l)) inFence = !inFence; if (inFence) return; const m = RE.head.exec(l); if (m) out.push({ lvl: m[1].length, text: m[2].replace(/[*_`~]/g, ''), line: i }); });
  return out;
}
export function wordStats(src) {
  const t = String(src || '').replace(/```[\s\S]*?```/g, '').replace(/[#>*_`~\-|[\]()!]/g, ' ');
  const cjk = (t.match(/[\u4e00-\u9fa5]/g) || []).length, en = (t.replace(/[\u4e00-\u9fa5]/g, ' ').match(/[A-Za-z0-9]+/g) || []).length;
  return { words: cjk + en, minutes: Math.max(1, Math.round((cjk + en) / 400)) };
}
// Rich DOM → Markdown
export function htmlToMd(root) {
  const inl = n => {
    let s = '';
    n.childNodes.forEach(c => {
      if (c.nodeType === 3) { s += c.nodeValue.replace(/\u200B/g, '').replace(/\u00A0/g, ' '); return; }
      if (c.nodeType !== 1) return;
      const t = c.tagName.toLowerCase(), st = c.style || {};
      if (t === 'br') { s += '  \n'; return; }
      if (t === 'input') return;
      if (t === 'img') { s += `![${c.getAttribute('alt') || ''}](${c.getAttribute('src') || ''})`; return; }
      const x = inl(c); if (!x.trim() && t !== 'code') { s += x; return; }
      const wrap = m => { const lead = x.match(/^\s*/)[0], trail = x.match(/\s*$/)[0]; return lead + m + x.trim() + m + trail; };
      if (t === 'strong' || t === 'b' || +st.fontWeight >= 600 || st.fontWeight === 'bold') s += wrap('**');
      else if (t === 'em' || t === 'i' || st.fontStyle === 'italic') s += wrap('*');
      else if (t === 'del' || t === 's' || t === 'strike' || (st.textDecoration || '').includes('line-through')) s += wrap('~~');
      else if (t === 'mark') s += wrap('==');
      else if (t === 'code') s += '`' + c.textContent + '`';
      else if (t === 'a') s += `[${x}](${c.getAttribute('href') || ''})`;
      else s += x;
    });
    return s;
  };
  const list = (el, depth) => {
    const ord = el.tagName.toLowerCase() === 'ol'; let i = parseInt(el.getAttribute('start') || '1');
    const out = [];
    Array.from(el.children).forEach(li => {
      if (li.tagName.toLowerCase() !== 'li') return;
      const cb = li.querySelector(':scope > input[type=checkbox]');
      const clone = li.cloneNode(true); clone.querySelectorAll(':scope > ul, :scope > ol').forEach(x => x.remove());
      let text = inl(clone).replace(/^\s+/, '').replace(/\s+$/, '');
      const mark = ord ? (i++) + '. ' : '- ';
      out.push('  '.repeat(depth) + mark + (cb ? (cb.checked ? '[x] ' : '[ ] ') : '') + text);
      li.querySelectorAll(':scope > ul, :scope > ol').forEach(sub => out.push(list(sub, depth + 1)));
    });
    return out.join('\n');
  };
  const block = el => {
    if (el.nodeType === 3) { const t = el.nodeValue.replace(/\u200B/g, ''); return t.trim() ? t.trim() : null; }
    if (el.nodeType !== 1) return null;
    const t = el.tagName.toLowerCase();
    const m = /^h([1-6])$/.exec(t);
    if (m) return '#'.repeat(+m[1]) + ' ' + inl(el).trim();
    if (t === 'ul' || t === 'ol') return list(el, 0);
    if (t === 'blockquote') { const inner = Array.from(el.childNodes).map(block).filter(x => x != null); const body = inner.length ? inner.join('\n\n') : inl(el); return body.split('\n').map(l => '> ' + l).join('\n'); }
    if (t === 'pre') { const code = el.querySelector('code'); const lang = code && code.getAttribute('data-lang') || ''; return '```' + lang + '\n' + (el.innerText || el.textContent).replace(/\n$/, '') + '\n```'; }
    if (t === 'hr') return '---';
    if (t === 'table') {
      const rows = Array.from(el.querySelectorAll('tr')).map(tr => Array.from(tr.children).map(c => inl(c).trim().replace(/\|/g, '\\|')));
      if (!rows.length) return null; const w = Math.max(...rows.map(r => r.length));
      const line = r => '| ' + Array.from({ length: w }, (_, i) => r[i] || '').join(' | ') + ' |';
      return [line(rows[0]), '| ' + Array.from({ length: w }, () => '---').join(' | ') + ' |'].concat(rows.slice(1).map(line)).join('\n');
    }
    if (t === 'img') return `![${el.getAttribute('alt') || ''}](${el.getAttribute('src') || ''})`;
    if (t === 'div' && el.querySelector('p,h1,h2,h3,ul,ol,pre,blockquote,table')) return Array.from(el.childNodes).map(block).filter(x => x != null).join('\n\n');
    const s = inl(el).replace(/\s+$/, '');
    return s.trim() ? s : '';
  };
  const parts = []; let buf = '';
  root.childNodes.forEach(n => {
    const inline = n.nodeType === 3 || (n.nodeType === 1 && /^(b|strong|i|em|a|code|span|mark|del|s|img|br)$/i.test(n.tagName));
    if (inline) { const d = document.createElement('div'); d.appendChild(n.cloneNode(true)); buf += inl(d); return; }
    if (buf.trim()) parts.push(buf.trim()); buf = '';
    const b = block(n); if (b != null) parts.push(b);
  });
  if (buf.trim()) parts.push(buf.trim());
  return parts.join('\n\n').replace(/\n{3,}/g, '\n\n').replace(/^\n+/, '') + '\n';
}
export function wordHtml(html) {
  return html.replace(/<table/g, '<table style="border-collapse:collapse;width:100%;margin:8px 0"').replace(/<(t[hd])( style="[^"]*")?>/g, (m, t, st) => `<${t} style="border:1px solid #C7C7CC;padding:6px 8px;${t === 'th' ? 'font-weight:600;background:#F5F5F7;' : ''}${st ? st.slice(8, -1) : ''}">`).replace(/<input type="checkbox"[^>]*checked>/g, '☑').replace(/<input type="checkbox"[^>]*>/g, '☐').replace(/ data-(line|li|task)="\d+"/g, '');
}
