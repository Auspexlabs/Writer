#!/usr/bin/env node
// Fidelity audit. Every fixture goes through the engine twice: opened and saved unchanged (`create --from`), and
// through the editors' write path (every paragraph, shape or cell rewritten and then set back to its own html /
// value, which is what a save from the app does to the blocks the user retyped). Both results are unzipped next to
// the original and compared part by part after normalising the XML; what the result no longer has is a fidelity
// bug. `--render` adds a picture comparison of page 1: the engine's `view html` in headless Chrome against Quick Look.
//
//   node tools/fidelity/roundtrip.mjs [--fixtures dir]... [--only name] [--formats docx,xlsx,pptx] [--cap 40]
//                                     [--no-touch] [--render] [--out dir] [--update-baseline]
//
// Prints Markdown, writes <out>/report.json with every detail, and exits 1 when a fixture loses more than
// tools/fidelity/baseline.json allows (the regression gate; --update-baseline accepts the current numbers).
import { spawn, spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { basename, dirname, extname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const args = process.argv.slice(2);
const flag = name => args.includes('--' + name);
const opts = name => args.flatMap((a, i) => a === '--' + name ? [args[i + 1]] : []);
if (flag('help')) { console.log(readFileSync(fileURLToPath(import.meta.url), 'utf8').split('\n').slice(1, 12).map(l => l.replace(/^\/\/ ?/, '')).join('\n')); process.exit(0); }
const fixtureDirs = opts('fixtures').length ? opts('fixtures').map(d => resolve(d)) : [join(root, 'tests', 'Writer.Tests', 'Fixtures')];
const out = resolve(opts('out')[0] || join(root, 'tools', 'fidelity', 'out'));
const formats = (opts('formats')[0] || 'docx,xlsx,pptx').split(',');
const cap = +(opts('cap')[0] || 40);
const only = opts('only')[0];
const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';

// ---------- engine ----------
const build = spawnSync('dotnet', ['build', 'src/Writer.Cli', '-c', 'Release', '-nologo', '-v', 'q'], { cwd: root, stdio: 'inherit' });
if (build.status !== 0) process.exit(build.status ?? 1);
const binDir = join(root, 'src', 'Writer.Cli', 'bin', 'Release');
const dll = join(binDir, readdirSync(binDir).find(d => d.startsWith('net')), 'writer.dll');
const port = await new Promise(r => { const s = createServer(); s.listen(0, '127.0.0.1', () => { const p = s.address().port; s.close(() => r(p)); }); });
const server = spawn('dotnet', [dll, 'serve', '--no-token', '--port', String(port)], { stdio: ['ignore', 'ignore', 'pipe'] });
await new Promise((res, rej) => {
  let buf = '';
  server.stderr.on('data', d => { buf += d; if (buf.includes('"url"')) res(); });
  server.on('exit', code => rej(new Error(`writer serve exited with ${code}: ${buf}`)));
});
process.on('exit', () => server.kill());
async function run(argv, attempt = 0) {
  let r;
  try { r = await fetch(`http://127.0.0.1:${port}/run`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) }); }
  catch (e) { if (attempt < 3) { await new Promise(res => setTimeout(res, 300)); return run(argv, attempt + 1); } throw e; } // a kept-alive socket the server dropped
  const j = await r.json();
  if (j.code !== 0) throw new Error(`${argv[0]} ${argv[2] || ''}: ${j.error?.message || 'failed'}`);
  return j.output || '';
}

// ---------- xml: parse, resolve namespaces, digest ----------
const NS = {
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w', 'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main': 'x',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic', 'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr', 'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14', 'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape': 'wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup': 'wpg',
  'urn:schemas-microsoft-com:vml': 'v', 'urn:schemas-microsoft-com:office:office': 'o', 'urn:schemas-microsoft-com:office:word': 'w10',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm', 'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel', 'http://schemas.openxmlformats.org/package/2006/content-types': 'ct',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'cx', 'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
  'http://schemas.microsoft.com/office/powerpoint/2012/main': 'p15', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main': 'x14',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac': 'x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision': 'xr',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14', 'http://schemas.microsoft.com/office/drawing/2014/main': 'a16',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm', 'http://www.w3.org/XML/1998/namespace': 'xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties': 'ep', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
};
const IGNORE_ATTR = /^(w:rsid\w*|w14:paraId|w14:textId|xr:uid|xr:revisionPtr|xr:revIDLastSave)$/;
const IGNORE_PART = /^docProps\/(core|app)\.xml$/;
const TEXT_ELEMENTS = new Set(['t', 'v', 'delText', 'instrText', 'f']);

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const decode = s => s.includes('&') ? s.replace(/&(#x[0-9a-fA-F]+|#\d+|\w+);/g, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e] ?? m) : s;

function parseXml(text) {
  const doc = { name: '#doc', attrs: [], children: [] };
  const stack = [doc];
  const tag = /<(\/)?([^\s\/>]+)((?:\s+[^\s=\/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/)?>/y;
  const attr = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let i = 0;
  const addText = t => { if (t) stack[stack.length - 1].children.push(decode(t)); };
  while (i < text.length) {
    const lt = text.indexOf('<', i);
    if (lt < 0) { addText(text.slice(i)); break; }
    addText(text.slice(i, lt));
    if (text.startsWith('<?', lt)) { i = text.indexOf('?>', lt) + 2; continue; }
    if (text.startsWith('<!--', lt)) { i = text.indexOf('-->', lt) + 3; continue; }
    if (text.startsWith('<![CDATA[', lt)) { const e = text.indexOf(']]>', lt); stack[stack.length - 1].children.push(text.slice(lt + 9, e)); i = e + 3; continue; }
    if (text.startsWith('<!', lt)) { i = text.indexOf('>', lt) + 1; continue; }
    tag.lastIndex = lt;
    const m = tag.exec(text);
    if (!m) throw new Error(`malformed xml near offset ${lt}: ${text.slice(lt, lt + 60)}`);
    i = tag.lastIndex;
    if (m[1]) { stack.pop(); continue; }
    const attrs = [];
    for (const a of m[3].matchAll(attr)) attrs.push([a[1], decode(a[2] ?? a[3])]);
    const el = { name: m[2], attrs, children: [] };
    stack[stack.length - 1].children.push(el);
    if (!m[4]) stack.push(el);
  }
  return doc.children.find(c => typeof c === 'object');
}

/** Element and attribute names as canonical `prefix:local` (by namespace URI, so prefix renames do not count), xmlns declarations dropped. */
function resolveNs(el, scope) {
  const local = { ...scope };
  for (const [k, v] of el.attrs) { if (k === 'xmlns') local[''] = v; else if (k.startsWith('xmlns:')) local[k.slice(6)] = v; }
  const qname = (name, isAttr) => {
    const colon = name.indexOf(':');
    const prefix = colon < 0 ? '' : name.slice(0, colon), localName = colon < 0 ? name : name.slice(colon + 1);
    if (isAttr && !prefix) return name;
    const uri = local[prefix];
    return uri ? `${NS[uri] ?? prefix}:${localName}` : name;
  };
  const attrs = el.attrs.filter(([k]) => k !== 'xmlns' && !k.startsWith('xmlns:')).map(([k, v]) => [qname(k, true), v]);
  return { name: qname(el.name, false), attrs, children: el.children.map(c => typeof c === 'string' ? c : resolveNs(c, local)) };
}

const bump = (map, key, n = 1) => map.set(key, (map.get(key) || 0) + n);
/** One part's canonical string plus its multisets of elements, attributes and attribute values, and its text length. */
function digest(el, acc) {
  const attrs = el.attrs.filter(([k]) => !IGNORE_ATTR.test(k)).sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  bump(acc.elems, el.name);
  for (const [k, v] of attrs) { bump(acc.attrs, `${el.name}@${k}`); bump(acc.vals, `${el.name}@${k}=${v}`); }
  let children = el.children.map(c => {
    if (typeof c !== 'string') return digest(c, acc);
    if (TEXT_ELEMENTS.has(el.name.replace(/^\w+:/, ''))) acc.text += c.length;
    return c.trim() ? c.replace(/&/g, '&amp;').replace(/</g, '&lt;') : '';
  }).filter(Boolean);
  if (/^(ct:Types|rel:Relationships)$/.test(el.name)) children = children.sort();
  return `<${el.name}${attrs.map(([k, v]) => ` ${k}="${v.replace(/"/g, '&quot;')}"`).join('')}>${children.join('')}</>`;
}
function digestPart(bytes) {
  const acc = { elems: new Map(), attrs: new Map(), vals: new Map(), text: 0 };
  acc.canon = digest(resolveNs(parseXml(bytes.toString('utf8').replace(/^﻿/, '')), {}), acc);
  return acc;
}

// ---------- packages ----------
function unzip(file, dir) {
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(dir, { recursive: true });
  const r = spawnSync('unzip', ['-o', '-q', file, '-d', dir], { encoding: 'utf8' });
  if (r.status !== 0 && r.status !== 1) throw new Error(`unzip ${file}: ${r.stderr}`);
  const parts = new Map();
  const walk = d => { for (const e of readdirSync(d, { withFileTypes: true })) e.isDirectory() ? walk(join(d, e.name)) : parts.set(relative(dir, join(d, e.name)), readFileSync(join(d, e.name))); };
  walk(dir);
  return parts;
}
const isXml = name => /\.(xml|rels)$/.test(name);
const lostBetween = (a, b) => { const lost = new Map(); for (const [k, n] of a) { const d = n - (b.get(k) || 0); if (d > 0) lost.set(k, d); } return lost; };
const sum = map => [...map.values()].reduce((s, n) => s + n, 0);

/** Compares the original package with one result. */
function compare(orig, saved) {
  const res = { partsChanged: [], partsMissing: [], partsAdded: [], mediaChanged: [], elementsLost: 0, elementsAdded: 0, attributesLost: 0, valuesChanged: 0, textLost: 0, lost: { elems: new Map(), attrs: new Map(), vals: new Map(), added: new Map() }, examples: new Map() };
  let textBefore = 0, textAfter = 0; // over the whole package: shared strings that become inline strings only move
  for (const [name, bytes] of saved) if (!orig.has(name)) { res.partsAdded.push(name); if (isXml(name)) try { textAfter += digestPart(bytes).text; } catch { } }
  for (const [name, bytes] of orig) {
    if (IGNORE_PART.test(name)) continue;
    const other = saved.get(name);
    if (!other) { res.partsMissing.push(name); continue; }
    if (!isXml(name)) { if (!bytes.equals(other)) res.mediaChanged.push(name); continue; }
    let a, b;
    try { a = digestPart(bytes); b = digestPart(other); } catch (e) { res.partsChanged.push(`${name} (${e.message})`); continue; }
    textBefore += a.text; textAfter += b.text;
    if (a.canon === b.canon) continue;
    res.partsChanged.push(name);
    const le = lostBetween(a.elems, b.elems), la = lostBetween(a.attrs, b.attrs), lv = lostBetween(a.vals, b.vals), ad = lostBetween(b.elems, a.elems);
    for (const [k, n] of le) { bump(res.lost.elems, k, n); res.examples.set(k, name); }
    for (const [k, n] of la) { bump(res.lost.attrs, k, n); res.examples.set(k, name); }
    for (const [k, n] of ad) { bump(res.lost.added, k, n); res.examples.set(k, name); }
    for (const [k, n] of lv) { const key = k.slice(0, k.indexOf('=')); if (!la.has(key)) { bump(res.lost.vals, key, n); res.examples.set(key, name); } }
    res.elementsLost += sum(le); res.attributesLost += sum(la); res.elementsAdded += sum(ad);
    res.valuesChanged += sum(lv) - sum(la);
  }
  res.textLost = Math.max(0, textBefore - textAfter);
  res.total = res.elementsLost + res.attributesLost + res.valuesChanged + res.partsMissing.length + res.mediaChanged.length;
  return res;
}

// ---------- the editors' write path ----------
/** What a save from the app sets on a block the user touched: paragraphs and headings get their html, shapes their html, cells their
 * value or formula. The writers rewrite only the words that differ, so each block is set twice: once with every word altered, then
 * back to its own value. Every run has then been written by the engine, and the file should still equal the original. */
function altered(props) {
  const out = {};
  for (const [k, v] of Object.entries(props)) out[k] = k === 'type' ? v : v.split(/(<[^>]*>)/).map(s => s.startsWith('<') ? s : s.replace(/(\S+)/g, '$1~')).join('');
  return out;
}
function targets(node, format, list = []) {
  const p = node.props || {};
  if (format === 'docx' && (node.kind === 'paragraph' || node.kind === 'heading') && p.html) list.push({ path: node.path, props: { html: p.html } });
  else if (format === 'pptx' && node.kind === 'shape' && p.html) list.push({ path: node.path, props: { html: p.html } });
  else if (format === 'xlsx' && node.kind === 'cell') {
    if (p.formula) list.push({ path: node.path, props: { formula: p.formula } });
    else if (p.value !== undefined && p.value !== '') list.push({ path: node.path, props: p.type === 'string' ? { value: p.value, type: 'string' } : { value: p.value } });
  }
  for (const c of node.children || []) targets(c, format, list);
  return list;
}

// ---------- rendering ----------
function screenshot(html, png, w, h, profile) {
  return new Promise(res => {
    const p = spawn(CHROME, ['--headless=new', '--disable-gpu', '--no-first-run', '--hide-scrollbars', '--force-device-scale-factor=1', `--user-data-dir=${profile}`, `--window-size=${w},${h}`, `--screenshot=${png}`, 'file://' + html], { stdio: ['ignore', 'ignore', 'pipe'] });
    let done = false;
    const finish = () => { if (done) return; done = true; clearTimeout(timer); p.kill('SIGKILL'); res(existsSync(png)); };
    const timer = setTimeout(finish, 30000);
    p.stderr.on('data', d => { if (String(d).includes('bytes written to file')) setTimeout(finish, 300); }); // Chrome writes the file, then hangs on exit here
    p.on('exit', finish);
  });
}
function readBmp(file) {
  const b = readFileSync(file);
  const off = b.readUInt32LE(10), w = b.readInt32LE(18), hRaw = b.readInt32LE(22), bpp = b.readUInt16LE(28);
  const h = Math.abs(hRaw), stride = Math.ceil(w * bpp / 32) * 4, bytes = bpp / 8, g = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    const row = off + (hRaw < 0 ? y : h - 1 - y) * stride;
    for (let x = 0; x < w; x++) { const p = row + x * bytes; g[y * w + x] = (b[p] + b[p + 1] + b[p + 2]) / 3; }
  }
  return { w, h, g };
}
/** Mean SSIM over 8x8 windows (stride 4) of two same-sized grey images, and the overlap of their dark pixels. */
function similarity(a, b) {
  const C1 = 6.5025, C2 = 58.5225, { w, h } = a;
  let total = 0, n = 0, inter = 0, union = 0;
  for (let i = 0; i < w * h; i++) { const da = a.g[i] < 160, db = b.g[i] < 160; if (da && db) inter++; if (da || db) union++; }
  for (let y = 0; y + 8 <= h; y += 4) for (let x = 0; x + 8 <= w; x += 4) {
    let ma = 0, mb = 0, va = 0, vb = 0, cov = 0;
    for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) { ma += a.g[(y + j) * w + x + i]; mb += b.g[(y + j) * w + x + i]; }
    ma /= 64; mb /= 64;
    for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) { const pa = a.g[(y + j) * w + x + i] - ma, pb = b.g[(y + j) * w + x + i] - mb; va += pa * pa; vb += pb * pb; cov += pa * pb; }
    va /= 63; vb /= 63; cov /= 63;
    total += ((2 * ma * mb + C1) * (2 * cov + C2)) / ((ma * ma + mb * mb + C1) * (va + vb + C2)); n++;
  }
  return { ssim: total / n, ink: union ? inter / union : 1 };
}
function toGrey(png, bmp, w, h) {
  const r = spawnSync('sips', ['-s', 'format', 'bmp', '-z', String(h), String(w), png, '--out', bmp], { encoding: 'utf8' });
  if (r.status !== 0) throw new Error(`sips: ${r.stderr}`);
  return readBmp(bmp);
}
async function render(fixture, format, src, dir, profile) {
  const ql = spawnSync('qlmanage', ['-t', '-s', '1200', '-o', dir, src], { encoding: 'utf8', timeout: 30000, killSignal: 'SIGKILL' });
  const ref = join(dir, basename(src) + '.png');
  if (ql.error?.code === 'ETIMEDOUT') { spawnSync('qlmanage', ['-r']); return { error: 'Quick Look hung on this file (30 s)' }; } // its Office previewer wedges on some files; a reset frees the next one
  if (ql.status !== 0 || !existsSync(ref)) return { error: 'Quick Look produced no picture' };
  let html = await run(['view', src, 'html']);
  let w = 794, h = 1123, cw = 300, ch = 424; // A4 at 96 dpi
  if (format === 'pptx') {
    const m = /class="slide" style="width:([\d.]+)px;height:([\d.]+)px/.exec(html);
    if (m) { w = Math.round(+m[1]); h = Math.round(+m[2]); }
    cw = 320; ch = Math.round(320 * h / w);
    html = html.replace('</style>', 'body{margin:0;padding:0;background:#fff}.slide{margin:0;box-shadow:none}.slide~.slide{display:none}</style>');
  } else if (format === 'docx') html = html.replace('</style>', 'body.docx{margin:0;padding:0;background:#fff}.docx .page{box-shadow:none;margin:0}.docx .page~*{display:none}</style>'); // the first sheet at its own size and margins
  else html = html.replace('</style>', 'body{max-width:none;margin:0;padding:72px 60px}.page{box-shadow:none;padding:0;margin:0}</style>');
  const page = join(dir, fixture + '.html'), ours = join(dir, fixture + '.writer.png');
  writeFileSync(page, html);
  if (!await screenshot(page, ours, w, h, profile)) return { error: 'Chrome produced no screenshot' };
  const s = similarity(toGrey(ref, join(dir, fixture + '.ref.bmp'), cw, ch), toGrey(ours, join(dir, fixture + '.writer.bmp'), cw, ch));
  return { ssim: s.ssim, ink: s.ink, ref, ours };
}

// ---------- ranking ----------
/** How visible a loss is: 3 = formatting or drawing the reader sees, 2 = content structure, 1 = metadata and machinery. */
const VISIBLE = /^(w:(b|bCs|i|iCs|u|strike|dstrike|color|sz|szCs|shd|highlight|jc|spacing|ind|rFonts|caps|smallCaps|vertAlign|tabs|tab|pStyle|rStyle|numPr|ilvl|numId|tbl\w*|tc\w*|tr\w*|gridCol|drawing|br|pBdr|bdr|top|bottom|left|right|insideH|insideV|keepNext|keepLines|pageBreakBefore|outlineLvl|kern|position|vanish|emboss|imprint|outline|shadow|effect|em|w|framePr|textAlignment|contextualSpacing|mirrorIndents|snapToGrid|widowControl|sectPr|pgSz|pgMar|cols|headerReference|footerReference|pict|object|sym|noProof|fitText|pgNumType|docGrid|titlePg|lnNumType|pgBorders)|a:(solidFill|noFill|gradFill|pattFill|blipFill|blip|srcRect|stretch|fillRect|srgbClr|schemeClr|sysClr|prstClr|lumMod|lumOff|alpha|tint|shade|latin|ea|cs|sym|ln|prstGeom|custGeom|xfrm|off|ext|pPr|rPr|defRPr|endParaRPr|bodyPr|normAutofit|spAutoFit|spcBef|spcAft|spcPts|spcPct|buChar|buAutoNum|buNone|buFont|buSzPct|buClr|lnSpc|effectLst|outerShdw|innerShdw|glow|reflection|softEdge|tab|tabLst|highlight|uFill|hlinkClick|scene3d|sp3d|avLst|gd|pathLst|path|moveTo|lnTo|close|pt|headEnd|tailEnd|prstDash|round|miter|bevel|gsLst|gs|lin|tileRect|tbl\w*|tc\w*|tr|gridCol|lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr|cell3D)|p:(sp|pic|grpSp|graphicFrame|cxnSp|txBody|spPr|style|nvSpPr|cNvPr|cNvSpPr|nvPr|ph|bg|bgPr|bgRef|transition|timing|ph|xfrm|clrMapOvr|clrMap|txStyles|titleStyle|bodyStyle|otherStyle|hf)|c:|cx:|x:(c|f|v|row|col|cols|mergeCell|mergeCells|xf|font|fill|border|numFmt|cellXfs|cellStyleXfs|cellStyles|fonts|fills|borders|numFmts|dxf|dxfs|conditionalFormatting|cfRule|dataValidation|dataValidations|sheetView|sheetViews|pane|selection|pageSetup|pageMargins|printOptions|drawing|hyperlink|hyperlinks|autoFilter|tableParts|tablePart|colBreaks|rowBreaks|headerFooter|oddHeader|oddFooter|sheetFormatPr|sheetPr|tabColor|outlinePr|pageSetUpPr|legacyDrawing|picture|color|b|i|u|sz|name|family|scheme|patternFill|fgColor|bgColor|alignment|protection|left|right|top|bottom|diagonal|rgbColor|indexedColors|mruColors|colors|extLst|sortState|customSheetView)|wp:|pic:|dgm:|dsp:|v:|o:|w10:|wps:|wpg:|xdr:|a14:|w15:|m:)/;
const STRUCTURE = /^(w:(t|p|r|pPr|rPr|hyperlink|fldSimple|fldChar|instrText|sdt|sdtPr|sdtContent|sdtEndPr|footnoteReference|endnoteReference|commentRangeStart|commentRangeEnd|commentReference|ins|del|delText|moveFrom|moveTo|bookmarkStart|bookmarkEnd|smartTag|customXml|cr|lastRenderedPageBreak|softHyphen|noBreakHyphen|ptab|dayShort|annotationRef|footnoteRef|endnoteRef|separator|continuationSeparator|proofErr)|a:(p|r|t|fld|br|endParaRPr)|x:(si|t|is|sheetData|dimension|definedName|definedNames|sheet|sheets|workbookView|bookViews|calcPr|workbookPr|fileVersion|calcChain)|p:(sld|cSld|spTree|nvGrpSpPr|grpSpPr|extLst|notes|notesMasterIdLst|sldIdLst|sldId|sldLayoutIdLst|sldMasterIdLst)|rel:|ct:|mc:)/;
const visibility = key => VISIBLE.test(key) ? 3 : STRUCTURE.test(key) ? 2 : 1;
function rank(results, mode, which) {
  const agg = new Map();
  for (const r of results) {
    const c = r[mode];
    if (!c) continue;
    for (const [k, n] of c.lost[which]) {
      const a = agg.get(k) || { key: k, count: 0, fixtures: new Set(), example: `${r.name} › ${c.examples.get(k)}` };
      a.count += n; a.fixtures.add(r.name); agg.set(k, a);
    }
  }
  return [...agg.values()].map(a => ({ ...a, fixtures: a.fixtures.size, visibility: visibility(a.key), score: a.count * a.fixtures.size * visibility(a.key) })).sort((a, b) => b.score - a.score);
}

// ---------- main ----------
const fixtures = [];
const walk = (d, base) => { for (const e of readdirSync(d, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) { const f = join(d, e.name); if (e.isDirectory()) walk(f, base); else if (formats.includes(extname(e.name).slice(1)) && !e.name.startsWith('~$')) fixtures.push({ file: f, name: relative(base, f) }); } };
for (const d of fixtureDirs) walk(d, d);
const picked = fixtures.filter(f => !only || f.name.includes(only));
if (!picked.length) { console.error('no fixtures found under ' + fixtureDirs.join(', ')); process.exit(2); }
const work = join(out, 'work'), renderDir = join(out, 'render'), profile = join(out, 'chrome-profile');
mkdirSync(work, { recursive: true }); mkdirSync(renderDir, { recursive: true });
const results = [];
for (const { file, name } of picked) {
  const format = extname(file).slice(1), slug = name.replace(/[\\/]/g, '__'), src = join(work, slug);
  const r = { name, format, size: statSync(file).size, touched: 0, touchErrors: [] };
  results.push(r);
  copyFileSync(file, src);
  const orig = unzip(src, join(work, 'x', slug + '.orig'));
  try {
    const saved = src.replace(/(\.\w+)$/, '.saved$1');
    rmSync(saved, { force: true });
    await run(['create', saved, '--from', src]);
    r.save = compare(orig, unzip(saved, join(work, 'x', slug + '.saved')));
    r.save.sizeDelta = statSync(saved).size - r.size;
  } catch (e) { r.error = e.message; }
  if (!flag('no-touch') && !r.error) {
    const touched = src.replace(/(\.\w+)$/, '.touched$1');
    copyFileSync(src, touched);
    try {
      const list = targets(JSON.parse(await run(['view', src, 'json'])), format).slice(0, cap);
      const propArgs = props => Object.entries(props).flatMap(([k, v]) => ['--prop', `${k}=${v}`]);
      for (const t of list) {
        try { await run(['set', touched, t.path, ...propArgs(altered(t.props))]); await run(['set', touched, t.path, ...propArgs(t.props)]); r.touched++; }
        catch (e) { r.touchErrors.push(e.message); }
      }
      r.touch = compare(orig, unzip(touched, join(work, 'x', slug + '.touched')));
      r.touch.sizeDelta = statSync(touched).size - r.size;
    } catch (e) { r.error = e.message; }
  }
  if (flag('render')) {
    if (!existsSync(CHROME)) r.render = { error: 'Chrome not found' };
    else try { r.render = await render(slug, format, src, renderDir, profile); } catch (e) { r.render = { error: e.message }; }
  }
  process.stderr.write(`${name}: save ${r.save ? r.save.total : '-'}  touch ${r.touch ? r.touch.total : '-'} (${r.touched} set)${r.render ? `  render ${r.render.ssim?.toFixed(2) ?? r.render.error}` : ''}${r.error ? '  ERROR ' + r.error : ''}\n`);
}
if (flag('render')) spawnSync('pkill', ['-f', profile]);

// ---------- output ----------
const md = [];
const n = (x, dash = '0') => x ? String(x) : dash;
const list = (a, max = 3) => a.length ? a.slice(0, max).map(p => p.replace(/^(word|xl|ppt)\//, '')).join(', ') + (a.length > max ? ` +${a.length - max}` : '') : '–';
md.push(`# Round-trip audit (${picked.length} fixtures)\n`);
md.push('Save = opened and saved unchanged. Touch = every paragraph / shape / cell rewritten and set back to its own value through the editors\' write path (capped at ' + cap + ' per file). Counts are elements, attributes and attribute values the original has and the result no longer has, after normalisation (rsid, paraId and the core properties excluded).\n');
for (const format of formats) {
  const rows = results.filter(r => r.format === format);
  if (!rows.length) continue;
  md.push(`## ${format}\n`);
  md.push('| fixture | save: parts changed | save: lost | touch: set | touch: parts changed | elements lost | attrs lost | values changed | text lost | media / parts lost | elements added | size Δ |');
  md.push('|---|---|---|---|---|---|---|---|---|---|---|---|');
  for (const r of rows) {
    if (r.error && !r.save) { md.push(`| ${r.name} | error: ${r.error} | | | | | | | | | | |`); continue; }
    const s = r.save, t = r.touch || {};
    md.push(`| ${r.name} | ${list(s.partsChanged)} | ${n(s.total)} | ${r.touch ? r.touched + (r.touchErrors.length ? ` (${r.touchErrors.length} failed)` : '') : '–'} | ${r.touch ? list(t.partsChanged) : '–'} | ${n(t.elementsLost, '–')} | ${n(t.attributesLost, '–')} | ${n(t.valuesChanged, '–')} | ${n(t.textLost, '–')} | ${r.touch ? n(t.mediaChanged.length + t.partsMissing.length) : '–'} | ${n(t.elementsAdded, '–')} | ${(t.sizeDelta ?? s.sizeDelta) > 0 ? '+' : ''}${t.sizeDelta ?? s.sizeDelta} |`);
  }
  md.push('');
}
for (const [mode, title] of [['save', 'Open → save unchanged'], ['touch', 'The editors\' write path']]) {
  const elems = rank(results, mode, 'elems'), attrs = rank(results, mode, 'attrs'), vals = rank(results, mode, 'vals'), added = rank(results, mode, 'added');
  if (!elems.length && !attrs.length && !vals.length && !added.length) { md.push(`## ${title}: nothing lost\n`); continue; }
  md.push(`## ${title}: what gets lost, ranked by count × fixtures × visibility\n`);
  for (const [label, ranked] of [['Elements lost', elems], ['Attributes lost', attrs], ['Attribute values changed', vals], ['Elements added (split runs, defaults written out)', added]]) {
    if (!ranked.length) continue;
    md.push(`### ${label}\n`);
    md.push('| # | element / attribute | count | fixtures | visibility | first seen in |');
    md.push('|---|---|---|---|---|---|');
    ranked.slice(0, 30).forEach((a, i) => md.push(`| ${i + 1} | \`${a.key}\` | ${a.count} | ${a.fixtures} | ${a.visibility} | ${a.example} |`));
    md.push('');
  }
}
if (flag('render')) {
  md.push('## Page 1 against Quick Look (SSIM 0…1, ink overlap 0…1)\n');
  md.push('| fixture | SSIM | ink overlap | note |');
  md.push('|---|---|---|---|');
  for (const r of results) md.push(`| ${r.name} | ${r.render?.ssim?.toFixed(3) ?? '–'} | ${r.render?.ink?.toFixed(3) ?? '–'} | ${r.render?.error ?? ''} |`);
  md.push('');
}
const errors = results.filter(r => r.touchErrors.length);
if (errors.length) {
  md.push('## Touch commands the engine refused\n');
  for (const r of errors) md.push(`- ${r.name}: ${[...new Set(r.touchErrors)].slice(0, 3).join('; ')}`);
  md.push('');
}
console.log(md.join('\n'));
const plain = v => v instanceof Map ? Object.fromEntries(v) : v instanceof Set ? [...v] : v;
writeFileSync(join(out, 'report.json'), JSON.stringify({ date: new Date().toISOString(), cap, results, ranking: Object.fromEntries(['save', 'touch'].flatMap(m => ['elems', 'attrs', 'vals', 'added'].map(w => [`${m}.${w}`, rank(results, m, w)]))) }, (k, v) => plain(v), 2));
writeFileSync(join(out, 'report.md'), md.join('\n'));

// ---------- gate ----------
const baselineFile = join(root, 'tools', 'fidelity', 'baseline.json');
// added: elements the write path leaves beyond the original (split runs, empty runs): no loss, but the file grows with every save
const current = Object.fromEntries(results.filter(r => r.save).map(r => [r.name, { save: r.save.total, touch: r.touch?.total ?? null, added: r.touch?.elementsAdded ?? null }]));
if (flag('update-baseline')) { writeFileSync(baselineFile, JSON.stringify(current, null, 2) + '\n'); console.error(`baseline written: ${baselineFile}`); }
else if (existsSync(baselineFile) && !only && !opts('fixtures').length) {
  const baseline = JSON.parse(readFileSync(baselineFile, 'utf8'));
  const worse = Object.entries(current).filter(([k, v]) => baseline[k] && (v.save > baseline[k].save || (v.touch ?? 0) > (baseline[k].touch ?? Infinity) || (v.added ?? 0) > (baseline[k].added ?? Infinity)));
  const broken = results.filter(r => r.error);
  if (worse.length || broken.length) {
    for (const [k, v] of worse) console.error(`REGRESSION ${k}: save ${baseline[k].save} → ${v.save}, touch ${baseline[k].touch} → ${v.touch}, added ${baseline[k].added ?? '-'} → ${v.added}`);
    for (const r of broken) console.error(`ERROR ${r.name}: ${r.error}`);
    server.kill();
    process.exit(1);
  }
  console.error('no regression against the baseline');
} else console.error('no baseline: run with --update-baseline to make this run the gate');
server.kill(); // the engine's server would otherwise keep the event loop alive
process.exit(0);
