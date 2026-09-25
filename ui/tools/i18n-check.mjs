// node ui/tools/i18n-check.mjs <files…>
// The translators' checklist: Chinese UI text an English user would still see, as file:line and the text.
//  - <x-dc> templates: text between tags and title / placeholder / aria-label / alt values (what $tTemplate translates)
//    with no entry in ui/i18n/en-*.js; the text printed is the key to add.
//  - JS (.js/.mjs, and <script> blocks in .html): string and template literals with Chinese that are neither the first
//    argument of $t( nor a dictionary key.
// Comments are skipped. A line containing i18n-ok (a trailing // i18n-ok) is intentional Chinese: document samples,
// font family names, CJK punctuation tables. Approximate on purpose. Exits 1 while anything is left.
import { readFileSync, readdirSync } from 'node:fs';

const dir = new URL('../i18n/', import.meta.url), win = {};
for (const f of readdirSync(dir).filter(f => /^en-.*\.js$/.test(f))) new Function('window', readFileSync(new URL(f, dir), 'utf8'))(win);
const keys = new Set(Object.keys(win.I18N_EN || {}));
const CJK = /[　-〿㐀-鿿豈-﫿＀-￯]/;
const norm = s => s.replace(/\s+/g, ' ').trim();
const blank = s => s.replace(/[^\n]/g, ' ');

/** String and template literals in src[from, to) as { at, text }; comments and regex literals skipped. Stops after an
 *  unmatched } when inBrace (the end of a template literal's ${ }). */
function scanJs(src, from, to, out, inBrace) {
  let i = from, depth = 0;
  while (i < to) {
    const c = src[i], n = src[i + 1];
    if (c === '/' && n === '/') { i = src.indexOf('\n', i); if (i < 0 || i > to) i = to; continue; }
    if (c === '/' && n === '*') { i = src.indexOf('*/', i + 2); i = i < 0 ? to : i + 2; continue; }
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < to && src[j] !== c && src[j] !== '\n') j += src[j] === '\\' ? 2 : 1;
      out.push({ at: i, text: src.slice(i + 1, j) }); i = j + 1; continue;
    }
    if (c === '`') {
      let j = i + 1, text = '';
      while (j < to && src[j] !== '`') {
        if (src[j] === '\\') { text += src.slice(j, j + 2); j += 2; }
        else if (src[j] === '$' && src[j + 1] === '{') { text += '${…}'; j = scanJs(src, j + 2, to, out, true); }
        else text += src[j++];
      }
      out.push({ at: i, text }); i = j + 1; continue;
    }
    if (c === '/' && regexAllowed(src, i, from)) {
      let j = i + 1, cls = false;
      while (j < to && src[j] !== '\n' && (cls || src[j] !== '/')) { if (src[j] === '\\') j++; else if (src[j] === '[') cls = true; else if (src[j] === ']') cls = false; j++; }
      i = j + 1; continue;
    }
    if (c === '{') depth++;
    else if (c === '}') { if (!depth && inBrace) return i + 1; depth--; }
    i++;
  }
  return i;
}
function regexAllowed(src, i, from) {
  let k = i - 1;
  while (k >= from && /\s/.test(src[k])) k--;
  if (k < from || '(,=:[!&|?{};+-*%<>~^'.includes(src[k])) return true;
  const word = /[\w$]+$/.exec(src.slice(Math.max(from, k - 10), k + 1));
  return !!word && /^(return|typeof|case|else|in|of|void|delete|new|throw|yield|await)$/.test(word[0]);
}

let found = 0;
for (const file of process.argv.slice(2)) {
  const src = readFileSync(file, 'utf8'), lines = src.split('\n'), starts = [0];
  for (let i = src.indexOf('\n'); i >= 0; i = src.indexOf('\n', i + 1)) starts.push(i + 1);
  const lineOf = at => { let lo = 0, hi = starts.length - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (starts[m] <= at) lo = m; else hi = m - 1; } return lo + 1; };
  const report = (at, text) => {
    const line = lineOf(at);
    if (lines[line - 1].includes('i18n-ok')) return;
    const t = norm(text);
    found++; console.log(`${file}:${line}  ${t.length > 120 ? t.slice(0, 117) + '…' : t}`);
  };

  const js = [];
  if (/\.html?$/.test(file)) {
    const open = /<x-dc(?:\s[^>]*)?>/.exec(src), close = src.lastIndexOf('</x-dc>');
    if (open && close > open.index) {
      const base = open.index + open[0].length, tpl = src.slice(base, close).replace(/<!--[\s\S]*?-->/g, blank);
      for (const m of tpl.matchAll(/>([^<>]+)</g))
        if (CJK.test(m[1]) && !keys.has(norm(m[1]))) report(base + m.index + 1 + m[1].search(/\S/), m[1]);
      for (const m of tpl.matchAll(/\s(?:title|placeholder|aria-label|alt)\s*=\s*(?:"([^"]*)"|'([^']*)')/gi)) {
        const v = m[1] ?? m[2];
        if (CJK.test(v) && !keys.has(norm(v))) report(base + m.index, v);
      }
    }
    for (const m of src.matchAll(/(<script\b[^>]*>)([\s\S]*?)<\/script>/gi)) scanJs(src, m.index + m[1].length, m.index + m[1].length + m[2].length, js);
  } else scanJs(src, 0, src.length, js);

  for (const { at, text } of js) {
    if (!CJK.test(text) || keys.has(text) || keys.has(text.replace(/\\(.)/g, '$1'))) continue;
    if (/\$t\(\s*$/.test(src.slice(Math.max(0, at - 12), at))) continue;
    report(at, text);
  }
}
if (found) { console.error(`${found} left`); process.exitCode = 1; }
