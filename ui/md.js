// Minimal Markdown → HTML (GFM-ish): headings, emphasis, code, links, images, lists, tasks, tables, quotes, hr, math (KaTeX).
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = s => esc(s).replace(/"/g, '&quot;');

// KaTeX, vendored under ./vendor/katex (no network at run time). It's a UMD build: a dynamic import() still runs it as
// plain script, so with no "exports"/"module" globals (real ESM, browser or Node) it falls through to the browser-global
// branch and assigns self.katex — and under Node's CJS interop it lands on the default export instead. Either way one
// of the two is the library.
let KX = null;
try { const m = await import('./vendor/katex/katex.min.js'); KX = (m && (m.default || globalThis.katex)) || null; } catch (e) { KX = null; }

/** LaTeX -> HTML for one formula (inline, or display when display is true). `trust` stays false: without it
 *  \href/\includegraphics/\url could smuggle an arbitrary (e.g. javascript:) URL or a foreign image into the page.
 *  `throwOnError` stays false: KaTeX then renders bad LaTeX as "$src$"-ish text in red with the parse error as its
 *  title, instead of throwing — one malformed formula must never take the rest of the document down with it. */
export function renderMath(src, display) {
  const raw = () => esc(display ? '$$' + src + '$$' : '$' + src + '$');
  if (!KX) return raw();
  try { return KX.renderToString(src, { throwOnError: false, trust: false, displayMode: !!display }); }
  catch (e) { return `<span class="katex-error" title="${attr(String((e && e.message) || e))}">${raw()}</span>`; }
}
/** The exact markup mdToHtml would produce for one formula — the rich editor calls these again once the user closes
 *  the source box, so a re-render always matches what a fresh parse of the saved Markdown would show. */
export const mathBlockHtml = src => `<div class="md-math-block" contenteditable="false" data-src="${attr(src)}">${renderMath(src, true)}</div>`;
export const mathInlineHtml = src => `<span class="md-math" contenteditable="false" data-src="${attr(src)}">${renderMath(src, false)}</span>`;

let SAFE = false; // set for the run by mdToHtml(src, {safe:true}) — the AI chat panel, where the text is untrusted
const schemeOf = u => { const m = /^([a-z][a-z0-9+.-]*):/i.exec(String(u).trim()); return m ? m[1].toLowerCase() : ''; };

export function inline(s) {
  const codes = [], maths = [];
  s = s.replace(/`([^`\n]+)`/g, (m, c) => { codes.push(c); return '\u0000' + (codes.length - 1) + '\u0000'; });
  // math is pulled out (and rendered to trusted HTML) before escaping, like code spans above — its LaTeX may contain
  // < > & that must reach KaTeX untouched, and the HTML KaTeX hands back must not be escaped a second time
  // (with the data-src wrapper, so saving the file writes $…$ again rather than the rendered text)
  s = s.replace(/\$(?=\S)([^$\n]*?\S)\$/g, (m, expr) => { maths.push(mathInlineHtml(expr)); return '\u0001' + (maths.length - 1) + '\u0001'; });
  s = esc(s);
  // esc leaves quotes alone, so every value that goes inside an attribute gets q: a " in a URL or alt text would
  // otherwise close the attribute and open a new one (onmouseover=…), in the chat panel and in any .md file opened
  const q = v => v.replace(/"/g, '&quot;');
  s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"([^"]*)")?\)/g, (m, a, u) => {
    const sc = schemeOf(u); // http(s)/data cover every real use (remote images, pasted/local images); anything else (e.g. javascript:) is dropped
    return (!sc || sc === 'http' || sc === 'https' || sc === 'data') ? `<img alt="${q(a)}" src="${q(u)}">` : a;
  });
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (m, label, u) => {
    const sc = schemeOf(u); // web and mail links anywhere; a relative link only in the editor; never javascript: and the like
    return (sc ? ['http', 'https', 'mailto'].includes(sc) : !SAFE) ? `<a href="${q(u)}" target="_blank" rel="noopener">${label}</a>` : label;
  });
  s = s.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, (m, u) => `<a href="${q(u)}" target="_blank" rel="noopener">${u}</a>`);
  s = s.replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>').replace(/__(?=\S)([\s\S]*?\S)__/g, '<strong>$1</strong>');
  s = s.replace(/(^|[^*])\*(?=\S)([^*\n]*?\S)\*(?!\*)/g, '$1<em>$2</em>').replace(/(^|[^\w])_(?=\S)([^_\n]*?\S)_(?!\w)/g, '$1<em>$2</em>');
  s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<del>$1</del>').replace(/==(?=\S)([\s\S]*?\S)==/g, '<mark>$1</mark>');
  s = s.replace(/ {2,}\n/g, '<br>').replace(/\n/g, ' ');
  s = s.replace(/\u0000(\d+)\u0000/g, (m, i) => '<code>' + esc(codes[+i]) + '</code>');
  return s.replace(/\u0001(\d+)\u0001/g, (m, i) => maths[+i]);
}
const RE = {
  fence: /^\s*(```|~~~)\s*([\w+-]*)\s*$/, mathfence: /^\s*\$\$\s*$/, head: /^(#{1,6})\s+(.*?)\s*#*\s*$/, hr: /^\s*([-*_])(\s*\1){2,}\s*$/,
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
    if (RE.mathfence.test(ln)) { const start = i; const body = []; i++; while (i < lines.length && !RE.mathfence.test(lines[i])) body.push(lines[i++]); i++; const src = body.join('\n'); out += `<div class="md-math-block"${L(start)} contenteditable="false" data-src="${attr(src)}">${renderMath(src, true)}</div>`; continue; }
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
    while (i < lines.length && lines[i].trim() && !RE.fence.test(lines[i]) && !RE.head.test(lines[i]) && !RE.hr.test(lines[i]) && !RE.quote.test(lines[i]) && !RE.list.test(lines[i]) && !RE.mathfence.test(lines[i]) && !(lines[i].includes('|') && i + 1 < lines.length && RE.tsep.test(lines[i + 1]))) para.push(lines[i++]);
    out += `<p${L(start)}>${inline(para.join('\n'))}</p>`;
  }
  return out;
}
export function mdToHtml(src, opts) { hid = 0; SAFE = !!(opts && opts.safe); return blocks(String(src || '').replace(/\r\n?/g, '\n').split('\n'), 0, true); }
export function headings(src) {
  const out = []; let inFence = false;
  String(src || '').split('\n').forEach((l, i) => { if (RE.fence.test(l)) inFence = !inFence; if (inFence) return; const m = RE.head.exec(l); if (m) out.push({ lvl: m[1].length, text: m[2].replace(/[*_`~]/g, ''), line: i }); });
  return out;
}
export function wordStats(src) {
  const t = String(src || '').replace(/```[\s\S]*?```/g, '').replace(/[#>*_`~\-|[\]()!]/g, ' ');
  const cjk = (t.match(/[一-龥]/g) || []).length, en = (t.replace(/[一-龥]/g, ' ').match(/[A-Za-z0-9]+/g) || []).length;
  return { words: cjk + en, minutes: Math.max(1, Math.round((cjk + en) / 400)) };
}

// ---- live "Typora-style" shortcuts (MarkdownEditor's rich view): pure text matching, so the DOM glue that turns a
// match into real formatting can stay a thin wrapper, and the matching itself is testable without a browser. ----

/** The inline shortcut that was just completed at the end of `before` (the current paragraph's text from its start up
 *  to the caret), if any. Longer/more specific markers are tried first so e.g. "**x**" doesn't read as "*" + literal.
 *  For the two markers that need a non-word/non-* character before them (so `*`/`_` don't fire inside `snake_case` or
 *  `a*b`), that guard character rides along in the match; `len` strips it back off before the caller replaces text. */
export function inlineTrigger(before) {
  const P = [
    ['**', /\*\*(?=\S)([^*\n]*?\S)\*\*$/], ['__', /__(?=\S)([^_\n]*?\S)__$/],
    ['~~', /~~(?=\S)([^~\n]*?\S)~~$/], ['==', /==(?=\S)([^=\n]*?\S)==$/],
    ['`', /`(?=\S)([^`\n]*?\S)`$/], ['$', /\$(?=\S)([^$\n]*?\S)\$$/],
    ['*', /(?:^|[^*])\*(?=\S)([^*\n]*?\S)\*$/], ['_', /(?:^|[^\w])_(?=\S)([^_\n]*?\S)_$/]
  ];
  const TAG = { '**': 'strong', '__': 'strong', '~~': 'del', '==': 'mark', '`': 'code', $: 'math', '*': 'em', _: 'em' };
  for (const [marker, re] of P) {
    const m = re.exec(before); if (!m) continue;
    const full = m[0].startsWith(marker) ? m[0] : m[0].slice(1); // drop the borrowed guard character, if any
    return { tag: TAG[marker], text: m[1], len: full.length };
  }
  return null;
}
/** The block shortcut for a paragraph whose entire text so far (start to caret) is `before`, just after its trailing
 *  trigger character (a space, normally). `len` is how much of `before` is the marker, to strip before applying it. */
export function blockTrigger(before) {
  before = before.replace(/\u00a0/g, ' '); // a space typed at the end of a text node arrives as a no-break space
  let m;
  if ((m = /^(#{1,6}) $/.exec(before))) return { type: 'h' + m[1].length, len: m[0].length };
  if (/^[-*+] $/.test(before)) return { type: 'ul', len: 2 };
  if (/^1[.)] $/.test(before)) return { type: 'ol', len: 3 };
  if (/^> $/.test(before)) return { type: 'quote', len: 2 };
  if (/^\[ ?\] $/.test(before)) return { type: 'task', len: before.length };
  return null;
}
export const MATH_BLOCK_MARK = '$$'; // a paragraph whose entire text is exactly this opens the math block editor

// Rich DOM → Markdown
export function htmlToMd(root) {
  const inl = n => {
    let s = '';
    n.childNodes.forEach(c => {
      if (c.nodeType === 3) { s += c.nodeValue.replace(/​/g, '').replace(/ /g, ' '); return; }
      if (c.nodeType !== 1) return;
      const t = c.tagName.toLowerCase(), st = c.style || {};
      if (t === 'br') { s += '  \n'; return; }
      if (t === 'input') return;
      if (t === 'img') { s += `![${c.getAttribute('alt') || ''}](${c.getAttribute('src') || ''})`; return; }
      if (t === 'span' && c.classList && c.classList.contains('md-math')) { s += '$' + (c.getAttribute('data-src') || '') + '$'; return; }
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
    if (el.nodeType === 3) { const t = el.nodeValue.replace(/​/g, ''); return t.trim() ? t.trim() : null; }
    if (el.nodeType !== 1) return null;
    const t = el.tagName.toLowerCase();
    if (el.classList && el.classList.contains('md-math-block')) {
      const ta = el.querySelector('textarea'); // still open for editing: save its live value, not the last-rendered data-src
      return '$$\n' + (ta ? ta.value : (el.getAttribute('data-src') || '')) + '\n$$';
    }
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
