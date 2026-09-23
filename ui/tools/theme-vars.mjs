#!/usr/bin/env node
// theme-vars.mjs — re-runnable theming: literal colours → var(--kN, <literal>), learned from the designer's export.
//
//   node ui/tools/theme-vars.mjs extract --from "<designer export dir>" [--report]
//       Scans every var(--kN, <literal>) in the export and writes ui/tools/theme-map.json:
//         vars   literal → the variable the designer used most often for it (--report prints ties/ambiguities)
//         lines  every designer line carrying a colour, keyed by a hash of the line with the wrappers stripped,
//                with the variable of each colour occurrence in order ('' = the designer left it bare)
//         segs   the same per segment: a style="…" attribute value or a quoted JS string
//       and writes ui/theme-page.css (light palette for [data-light] document pages in dark mode, glass transparency).
//   node ui/tools/theme-vars.mjs apply [--dry] [--verbose] [files…]        default: ui/*.dc.html
//       Rewrites literal colours (#hex, rgb(), rgba()) inside style="…", style-hover/active/focus="…",
//       <style> blocks and the component script's string constants (e.g. color: on ? '#1D1D1F' : '#6E6E73').
//       Order of precedence per line: the skip list; a designer line with the same colours → its exact variables;
//       a designer segment with the same colours → its variables; otherwise the ROLE table (the CSS property or
//       JS key before the literal decides between text / background / line / shadow / inset variants of an
//       ambiguous literal) and finally the plain frequency mapping. Idempotent: wrapped values are never wrapped
//       again, so hand fixes on lines the designer never had survive re-runs.
//
// Skip list — colours that are document content, not chrome, stay literal:
//   * files that are document I/O are not in the default set: office-io.js (slide THEMES, DOC_CSS),
//     sheet-engine.js, md.js, pdf-kit.js (annotation colours), mindmap.js (branch palette), engine.js;
//   * the template's data-props attribute (default documents) and anything outside style attributes / <style>;
//   * lines matching SKIP_LINE (see the comments there), among them every palette — named or an inline [['#hex', 'name'], …] list —
//     since a palette colour is written into the document: a var(--kN, …) there would reach the file and follow the theme;
//   * literals in SKIP_LITERAL: the PDF highlight, sheet fills and black as a document colour;
//   * a line ending in a `// theme: keep` comment or an element with data-theme-keep, for hand-kept literals.
// apply also links ./theme-dark.css from every page's <helmet> (the dark values of every --kN), as the designer did.
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const UI_DIR = path.resolve(HERE, '..');
export const MAP_FILE = path.join(HERE, 'theme-map.json');

const LIT = String.raw`#[0-9a-fA-F]{3,8}(?![0-9a-zA-Z_-])|rgba?\([^()]*\)`;
/** One token: an existing wrapper (groups 1+2, consumed whole so it is never re-wrapped) or a bare literal (group 3). */
const TOKEN = new RegExp(String.raw`var\(\s*(--[\w-]+)\s*,\s*(${LIT})\s*\)|(${LIT})`, 'g');
const ATTR = /(\sstyle(?:-[a-z]+)?=")([^"]*)(")/g;
const SCRIPT_START = /<script type="text\/x-dc"/;

export const SKIP_LINE = [
  /\b(FILLS|TEXTS|THEMES|PAL|PALETTE|SWATCHES|CHART_COLORS|COLORS)\s*=/, // palettes the user picks document colours from
  /\[\s*\[\s*'#[0-9a-fA-F]{3,8}'\s*,/,                            // …and inline ones, [['#hex', 'name'], …] (the sheet's border colours)
  /\bs\.(fill|color|bdc)\b|bsh\.push/,                            // sheet cell rendering: cell styles are document content
  /\bstate = \{/,                                                 // component state defaults (e.g. the PDF annotation colour)
  /isColors?: true/,                                              // colour inputs and swatches: values must be real colours
  /\bconst DEF = \{/,                                             // settings defaults are stored values (e.g. accent: '#3F7D5C')
  /\.style\.[a-zA-Z]+ = '|<td style=|data-toc=|<!doctype html>/i, // HTML/CSS written into the document or its print copy
  /\b(fillStyle|strokeStyle)\b/,                                  // canvas exports render the document, always in light colours
  /data-props=/,                                                  // default documents in the component props
  /inset:0;[^"]*background:rgba\(0,0,0,/,                         // scrims (full-cover backdrops) stay black in both themes
  /theme:\s*keep|data-theme-keep/                                 // explicit opt-out: a hand-kept literal (comment or attribute)
];
export const SKIP_LITERAL = new Set(['#FFE066', '#F3E6C4', '#FFE8A3', '#EFEFF4', '#000000']);

/** Ambiguous literals: the designer used a different variable per role. Matched on the property / JS key before the literal. */
const TEXT = /^(color|fg|fill|caret-color|inCaret|accent-color)$/, BG = /^(background|background-color|bg|ringFill)$/;
const LINE = /^(border|border-[a-z-]+|outline|stroke|ringColor|line|tline|bd|hb|sep)$/;
export const ROLE = {
  '#1D1D1F': { text: '--k1', bg: '--k7', line: '--k7' },
  '#3A3A3C': { text: '--k9', bg: '--k13', line: '--k13', inset: '--k13' },
  '#C7C7CC': { text: '--k63', bg: '--k47', line: '--k47' },
  '#3F7D5C': { text: '--k35', bg: '--k59', line: '--k59' },
  '#fff': { text: 'inv', inset: '--k25', bg: '--k51', line: '--k51' },
  '#FFFFFF': { text: 'inv', inset: '--k25', bg: '--k0', line: '--k0' },
  'rgba(255,255,255,0.95)': { inset: '--k18', bg: '--k31' },
  'rgba(255,255,255,0.3)': { inset: '--k21', bg: '--k49' },
  'rgba(255,255,255,0.9)': { inset: '--k34', bg: '--k50' },
  'rgba(255,255,255,0.85)': { inset: '--k64', bg: '--k41' },
  'rgba(0,0,0,0.12)': { shadow: '--k22', bg: '--k36', line: '--k36', inset: '--k36' },
  'rgba(0,0,0,0.1)': { shadow: '--k19', bg: '--k29', line: '--k29', inset: '--k29' },
  'rgba(0,0,0,0.08)': { shadow: '--k53', bg: '--k28', line: '--k28', inset: '--k28' },
  'rgba(0,0,0,0.07)': { shadow: '--k65', bg: '--k17', line: '--k17', inset: '--k17' },
  'rgba(0,0,0,0.06)': { shadow: '--k66', bg: '--k55', line: '--k55', inset: '--k55' },
  'rgba(0,0,0,0.05)': { shadow: '--k38', bg: '--k37', line: '--k37', inset: '--k37' },
  'rgba(0,0,0,0.04)': { shadow: '--k67', bg: '--k54', line: '--k54', inset: '--k54' },
  'rgba(0,0,0,0.14)': { shadow: '--k77', bg: '--k60', line: '--k60', inset: '--k60' },
  'rgba(0,0,0,0.2)': { shadow: '--k45', bg: '--k78', line: '--k78', inset: '--k78' }
};

export const strip = s => s.replace(TOKEN, (m, v, lit, bare) => bare || lit);
export const hash = s => createHash('sha1').update(s).digest('hex').slice(0, 12);
const canon = lit => lit[0] === '#' ? (lit.length === 4 ? '#' + [...lit.slice(1)].map(c => c + c).join('') : lit).toLowerCase() : lit.replace(/\s+/g, '').toLowerCase();

/** The colour occurrences of a string as [var|''] in order; null when it has none. */
function seqOf(s) { const seq = []; s.replace(TOKEN, (m, v, lit, bare) => { seq.push(bare ? '' : v); return m; }); return seq.length ? seq : null; }
/** Segments of one line: style attribute values in the template, quoted strings in the script, the whole line in <style>. */
function segments(line, ctx) {
  if (ctx.inScript) { const out = []; let q = null, start = 0; for (let i = 0; i < line.length; i++) { const c = line[i]; if (q) { if (c === '\\') i++; else if (c === q) { out.push([start, i]); q = null; } } else if (c === "'" || c === '"' || c === '`') { q = c; start = i + 1; } } if (q) out.push([start, line.length]); return out; }
  if (ctx.inStyle) return [[0, line.length]];
  const out = []; let m; ATTR.lastIndex = 0; while ((m = ATTR.exec(line))) out.push([m.index + m[1].length, m.index + m[1].length + m[2].length]); return out;
}
function track(line, ctx) { if (SCRIPT_START.test(line)) ctx.inScript = true; if (!ctx.inScript && /<style[\s>]/.test(line)) ctx.inStyle = true; }
function untrack(line, ctx) { if (ctx.inStyle && /<\/style>/.test(line)) ctx.inStyle = false; }

/** Learns the mapping from every file of the designer's export. */
export function extract(dir) {
  const counts = new Map(), lines = {}, segCounts = {}, conflicts = [], glass = new Set(), lightCounts = {};
  const tally = (table, key, seq) => { const t = table[key] = table[key] || {}; const k = seq.join(); t[k] = (t[k] || 0) + 1; };
  for (const f of fs.readdirSync(dir).filter(f => /\.(dc\.html|js)$/.test(f)).sort()) {
    const ctx = { inScript: false, inStyle: false };
    fs.readFileSync(path.join(dir, f), 'utf8').split('\n').forEach(line => {
      track(line, ctx);
      const seq = seqOf(line);
      if (seq) {
        line.replace(TOKEN, (m, v, lit) => { if (v) { const c = counts.get(lit) || new Map(); c.set(v, (c.get(v) || 0) + 1); counts.set(lit, c); const l = lightCounts[v] = lightCounts[v] || {}; l[lit] = (l[lit] || 0) + 1; } return m; });
        // frosted surfaces: the variables in the background of any style that also has a backdrop-filter
        for (const [, css] of line.matchAll(/style="([^"]*backdrop-filter[^"]*)"/g)) for (const [, bg] of css.matchAll(/background:([^;]*)/g)) for (const [, v] of bg.matchAll(/var\((--k\d+),/g)) glass.add(v);
        const k = hash(strip(line));
        if (lines[k] && lines[k].join() !== seq.join()) conflicts.push({ file: f, line: line.trim().slice(0, 120) }); else lines[k] = seq;
        for (const [a, b] of segments(line, ctx)) { const s = line.slice(a, b), ss = seqOf(s); if (ss) tally(segCounts, hash(strip(s)), ss); }
      }
      untrack(line, ctx);
    });
  }
  const vars = {}, ambiguous = {}, ties = [];
  for (const [lit, c] of counts) {
    // the --k palette is the app's; --g*/--m* belong to the Mac windows and win only where no --k variable exists
    const sorted = [...c].sort((a, b) => (/^--k/.test(b[0]) - /^--k/.test(a[0])) || b[1] - a[1]);
    vars[lit] = sorted[0][0];
    if (sorted.length > 1) { ambiguous[lit] = Object.fromEntries(sorted); if (sorted[0][1] === sorted[1][1]) ties.push(lit); }
  }
  const segs = {}; for (const [k, t] of Object.entries(segCounts)) { const best = Object.entries(t).sort((a, b) => b[1] - a[1])[0][0]; segs[k] = best.split(','); }
  const light = {}; for (const [v, t] of Object.entries(lightCounts)) light[v] = Object.entries(t).sort((a, b) => b[1] - a[1])[0][0];
  const dark = {}; try { for (const [, v, val] of fs.readFileSync(path.join(dir, 'theme-dark.css'), 'utf8').matchAll(/(--[\w-]+):([^;}]+)/g)) dark[v] = val.trim(); } catch (e) { }
  return { from: dir, vars, ambiguous, ties, conflicts, lines, segs, light, dark, glass: [...glass].sort((a, b) => a.slice(3) - b.slice(3)) };
}

/** ui/theme-page.css: (1) document pages marked [data-light] keep the light value of every --kN in dark mode, so pages stay
 *  white with readable text unless html[data-dark-pages] is set; (2) html[data-glass] rescales the alpha of every frosted
 *  surface by --gt (glass setting / 55: 0 = solid, 1 = as designed, 1.8 = most transparent). */
export function pageCss(map) {
  const k = Object.keys(map.light).filter(v => /^--k/.test(v)).sort((a, b) => (a.slice(3) - b.slice(3)) || a.localeCompare(b));
  const decl = vars => vars.map(v => `${v}:${map.light[v]}`).join(';');
  const light = decl(k) + ';--kpage:#FFFFFF;--kpageb:rgba(0,0,0,0.06);--kinv:#FFFFFF';
  const scale = val => { const m = /^rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)$/.exec(val || ''); return m && `rgba(${m[1]},${m[2]},${m[3]},clamp(0.03,calc(1 - ${+(1 - m[4]).toFixed(3)} * var(--gt, 1)),1))`; };
  const glass = (pick) => map.glass.map(v => { const x = scale(pick(v)); return x && `${v}:${x}`; }).filter(Boolean).join(';');
  return `/* Generated by ui/tools/theme-vars.mjs extract from the designer's export — do not edit; re-run extract instead. */
/* Document pages stay white in dark mode: [data-light] restores the light palette inside them (html[data-dark-pages] opts out). */
html[data-theme="dark"]:not([data-dark-pages]) [data-light]{color-scheme:light;${light}}
/* Glass transparency (settings › 外观 › 玻璃通透度): html[data-glass] with --gt = glass / 55. */
html[data-glass]{${glass(v => map.light[v])}}
html[data-glass][data-theme="dark"]{${glass(v => map.dark[v])}}
`;
}

/** The role of the literal at `at` inside `s`: which CSS property / JS key precedes it, and whether it sits in an inset shadow layer. */
function roleAt(s, at) {
  const before = s.slice(0, at);
  const decl = before.slice(before.lastIndexOf(';') + 1);
  const km = /([\w-]+)\s*:\s*[^:;]*$/.exec(decl); // last "prop:" of the declaration
  let key = km ? km[1] : (/([\w-]+)\s*[:=]\s*[^:=]*$/.exec(before) || [])[1] || '';
  const val = km ? decl.slice(km.index + km[1].length + 1) : before;
  if (/^box-shadow$/.test(key) || /(^|[\s,])(inset\s+)?-?\d[\w.]*\s+-?\d[\w.]*\s+-?\d/.test(val) && !/\b(solid|dashed|dotted)\b/.test(val)) { // a shadow layer
    let depth = 0, cut = 0; for (let i = 0; i < val.length; i++) { const c = val[i]; if (c === '(') depth++; else if (c === ')') depth--; else if (c === ',' && !depth) cut = i + 1; }
    return /\binset\b/.test(val.slice(cut)) ? 'inset' : 'shadow';
  }
  if (TEXT.test(key)) return 'text'; if (BG.test(key)) return 'bg'; if (LINE.test(key)) return 'line';
  return '';
}

/** Applies the mapping to one file's text; returns the new text and a summary. */
export function apply(text, map, opts = {}) {
  const canonIndex = new Map(Object.entries(map.vars).map(([lit, v]) => [canon(lit), v]));
  const lookup = lit => map.vars[lit] || canonIndex.get(canon(lit)) || null;
  const s = { transplanted: 0, wrapped: {}, unmapped: new Set(), skipped: 0, changed: 0, notes: [] };
  const wrap = (v, lit) => { s.wrapped[v] = (s.wrapped[v] || 0) + 1; return `var(${v}, ${lit})`; };
  const rewrap = (str, seq) => { let k = 0; return str.replace(TOKEN, (m, v, lit, bare) => { const want = seq[k++], l = bare || lit; return want && !SKIP_LITERAL.has(l.toUpperCase()) ? `var(${want}, ${l})` : l; }); };
  const fill = (str, seq) => { let k = 0; return str.replace(TOKEN, (m, v, lit, bare) => { const want = seq[k++]; return bare && want && !SKIP_LITERAL.has(bare.toUpperCase()) ? wrap(want, bare) : m; }); };
  const mapLits = (str, prefix = '') => {
    let out = str.replace(TOKEN, (m, v, lit, bare, at) => {
      if (!bare) return m;
      if (SKIP_LITERAL.has(bare.toUpperCase())) { s.skipped++; return m; }
      const role = ROLE[bare] || ROLE[Object.keys(ROLE).find(k => canon(k) === canon(bare))];
      if (role) { const nv = role[roleAt(prefix + str, prefix.length + at)]; if (nv === 'inv') return m; if (nv) return wrap(nv, bare); }
      const nv = lookup(bare); if (!nv) { s.unmapped.add(bare); return m; }
      return wrap(nv, bare);
    });
    // white text is --kinv only on an inverted (--k7) surface; elsewhere (toasts, badges) it stays white
    if (/var\(--k7\b/.test(prefix + out)) out = out.replace(TOKEN, (m, v, lit, bare, at) => bare && /^#(fff|ffffff)$/i.test(bare) && roleAt(prefix + out, prefix.length + at) === 'text' ? wrap('--kinv', bare) : m);
    return out;
  };
  if (/<helmet>/.test(text) && !/theme-dark\.css/.test(text)) { text = text.replace('<helmet>', '<helmet>\n<link rel="stylesheet" href="./theme-dark.css">'); s.changed++; }
  const ctx = { inScript: false, inStyle: false };
  const out = text.split('\n').map((line, n) => {
    track(line, ctx);
    let next = line;
    if (SKIP_LINE.some(r => r.test(line))) s.skipped++;
    else {
      const seq = map.lines[hash(strip(line))];
      if (seq) { next = rewrap(line, seq); if (next !== line) s.transplanted++; }
      else {
        let off = 0;
        for (const [a, b] of segments(line, ctx)) {
          const seg = next.slice(a + off, b + off), ss = seqOf(seg); if (!ss) continue;
          const dseq = ss.length > 1 || !ROLE[strip(seg).match(TOKEN)?.[0]] ? map.segs[hash(strip(seg))] : null;
          // a designer segment only fills bare literals: a segment is short and context-free, so existing wrappers stay
          const rep = dseq ? fill(seg, dseq) : mapLits(seg, next.slice(0, a + off));
          next = next.slice(0, a + off) + rep + next.slice(b + off); off += rep.length - seg.length;
        }
      }
      if (next !== line) { s.changed++; if (opts.verbose && !seq) s.notes.push(`${n + 1}: ${next.trim().slice(0, 220)}`); }
    }
    untrack(line, ctx);
    return next;
  }).join('\n');
  return { text: out, summary: s };
}

function main(argv) {
  const [cmd, ...rest] = argv; const flags = {}, files = [];
  for (let i = 0; i < rest.length; i++) { if (rest[i].startsWith('--')) { const k = rest[i].slice(2); if (k === 'from' || k === 'map') flags[k] = rest[++i]; else flags[k] = true; } else files.push(rest[i]); }
  const mapFile = flags.map || MAP_FILE;
  if (cmd === 'extract') {
    if (!flags.from) throw new Error('extract needs --from <designer export dir>');
    const map = extract(flags.from);
    fs.writeFileSync(mapFile, JSON.stringify(map) + '\n');
    fs.writeFileSync(path.join(UI_DIR, 'theme-page.css'), pageCss(map));
    console.log(`${Object.keys(map.vars).length} literals, ${Object.keys(map.lines).length} designer lines, ${Object.keys(map.segs).length} segments → ${path.relative(process.cwd(), mapFile)}; ${map.glass.length} glass variables → ui/theme-page.css`);
    if (flags.report) {
      console.log(`\n${Object.keys(map.ambiguous).length} literals map to more than one variable (most frequent wins unless ROLE decides):`);
      for (const [lit, c] of Object.entries(map.ambiguous)) console.log(`  ${lit.padEnd(24)} ${Object.entries(c).map(([v, n]) => `${v}×${n}`).join('  ')}${map.ties.includes(lit) ? '   TIE' : ''}${ROLE[lit] ? '   (ROLE)' : ''}`);
      if (map.conflicts.length) { console.log(`\n${map.conflicts.length} identical designer lines carry different variables (first wins):`); map.conflicts.forEach(c => console.log(`  ${c.file}: ${c.line}`)); }
    }
    return;
  }
  if (cmd === 'apply') {
    const map = JSON.parse(fs.readFileSync(mapFile, 'utf8'));
    const list = files.length ? files : fs.readdirSync(UI_DIR).filter(f => f.endsWith('.dc.html')).map(f => path.join(UI_DIR, f));
    for (const f of list) {
      const src = fs.readFileSync(f, 'utf8'), { text, summary: s } = apply(src, map, { verbose: flags.verbose });
      const wrapped = Object.entries(s.wrapped).sort((a, b) => b[1] - a[1]).map(([v, n]) => `${v}×${n}`).join(' ');
      console.log(`${path.basename(f)}: ${s.changed} lines changed (${s.transplanted} whole lines from the designer), ${s.skipped} skipped${s.unmapped.size ? `, left bare (no variable): ${[...s.unmapped].join(' ')}` : ''}${wrapped ? `\n  wrapped by segment/role/frequency: ${wrapped}` : ''}`);
      s.notes.forEach(n => console.log('   ' + n));
      if (!flags.dry && text !== src) fs.writeFileSync(f, text);
    }
    return;
  }
  console.log('usage: theme-vars.mjs extract --from <dir> [--report] | apply [--dry] [--verbose] [files…]');
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main(process.argv.slice(2));
