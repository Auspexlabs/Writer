// node ui/tools/templates/build.mjs [docx|xlsx|pptx|md|mm|<id> …] [--no-thumbs]
// The new-file gallery's templates. Builds every template of docx.mjs … mm.mjs (or the types / ids named) in Chinese and in
// English from its one source, through the writer engine itself (serve --no-token, POST /run), into
// ui/templates/<zh|en>/<type>/<id>.<ext>; renders each file's thumbnail with the app's own editors in headless Chrome
// (<id>.webp beside it); and writes the gallery's list (ui/templates/index.json) and the English names and categories
// (a generated block in ui/i18n/en-shell.js). Needs dotnet (the engine is built when src/Writer.Cli has no build yet; run
// dotnet build src/Writer.Cli after changing it) and Chrome (CHROME=<path> for another build). Nothing opens on screen.
import { spawn, execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, readFileSync, readdirSync, writeFileSync, copyFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as L from './lib.mjs';

const SRC = dirname(fileURLToPath(import.meta.url)), ROOT = join(SRC, '../../..'), UI = join(ROOT, 'ui'), OUT = join(UI, 'templates');
const CLI = join(ROOT, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
export const TYPES = ['docx', 'xlsx', 'pptx', 'md', 'mm'], LANGS = ['zh', 'en'];
const BUILDERS = { docx: L.docx, xlsx: L.xlsx, pptx: L.pptx, mm: L.mm };
const sleep = ms => new Promise(r => setTimeout(r, ms));

const SOURCES = {};
for (const type of TYPES) SOURCES[type] = await import(`./${type}.mjs`);

/** writer serve on dir (with the app's ui/ when ui is given); run(argv) runs one command and returns its JSON. */
async function engine(dir, ui) {
  const proc = spawn('dotnet', [CLI, 'serve', '--dir', dir, '--port', '0', '--no-token', ...(ui ? ['--ui', ui] : [])], { stdio: ['ignore', 'ignore', 'pipe'] });
  let err = '';
  const url = await new Promise((resolve, reject) => {
    proc.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1]); });
    proc.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
  });
  const run = async argv => {
    const r = await (await fetch(url + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) })).json();
    if (r.code !== 0) throw new Error(argv.join(' ') + '\n→ ' + JSON.stringify(r.error));
    const out = (r.output || '').trim();
    return /^[{[]/.test(out) ? JSON.parse(out) : out;
  };
  return { url, run, stop: () => proc.kill() };
}

const fileOf = (lang, type, id) => join(OUT, lang, type, `${id}.${type}`);
const thumbOf = (lang, type, id) => join(OUT, lang, type, `${id}.webp`);

async function buildOne(run, type, tpl, lang) {
  const file = fileOf(lang, type, tpl.id), t = (zh, en) => lang === 'zh' ? zh : en;
  mkdirSync(dirname(file), { recursive: true }); rmSync(file, { force: true });
  if (type === 'md') { writeFileSync(file, (await tpl.build(null, t)).replace(/^\n/, '')); return; }
  await run(['create', file]);
  await tpl.build(BUILDERS[type](run, file), t);
}

// ---- thumbnails: the app opens each file (a copy) and the editor's page, grid, slide or map is captured ----
const WIDTH = { docx: 300, md: 300, xlsx: 320, pptx: 320, mm: 320 }; // px; each clip below has its type's shape
/** In the page: the rectangle to capture, once the editor of that type shows the file. */
const CLIP = {
  docx: `(() => { const ed = document.querySelector('.wd-ed[data-pg]'); if (!ed || !ed.textContent.trim()) return null; const b = ed.parentElement.children[0].getBoundingClientRect(); return { x: b.x, y: b.y, width: b.width, height: b.height }; })()`,
  md: `(() => { const ed = document.querySelector('.pt-ed'); if (!ed || !ed.textContent.trim()) return null; const b = ed.getBoundingClientRect(); return { x: b.x, y: b.y + 40, width: b.width, height: b.width * 4 / 3 }; })()`,
  pptx: `(() => { const o = document.querySelector('[data-soid]'); if (!o) return null; const b = o.parentElement.parentElement.getBoundingClientRect(); return { x: b.x, y: b.y, width: b.width, height: b.height }; })()`,
  xlsx: `(() => { const ds = [...document.querySelectorAll('[data-edroot] div')], ch = ds.find(d => d.style.gridRow === '1' && d.textContent.trim() === 'A'), rh = ds.find(d => d.style.gridColumn === '1' && d.textContent.trim() === '1');
    if (!ch || !rh) return null; const x = rh.getBoundingClientRect().x, y = ch.getBoundingClientRect().y; return { x, y, width: 880, height: 660 }; })()`,
  mm: `(() => { const gs = [...document.querySelectorAll('[data-edroot] g[data-nid]')].map(g => g.getBoundingClientRect()); if (!gs.length) return null;
    const x0 = Math.min(...gs.map(b => b.left)), x1 = Math.max(...gs.map(b => b.right)), y0 = Math.min(...gs.map(b => b.top)), y1 = Math.max(...gs.map(b => b.bottom));
    let w = x1 - x0 + 80, h = y1 - y0 + 80; if (w / h > 1.6) h = w / 1.6; else w = h * 1.6; return { x: (x0 + x1 - w) / 2, y: (y0 + y1 - h) / 2, width: w, height: h }; })()`,
};

async function thumbs(list) {
  const { launch } = await import('../../../desktop/store/screenshots/src/cdp.mjs');
  const tmp = mkdtempSync(join(process.env.TEMPLATES_TMP || tmpdir(), 'writer-templates-'));
  const eng = await engine(tmp, UI), page = await launch({ width: 1400, height: 1700, scale: 1 });
  try {
    await page.goto(eng.url + '/files', 300);
    await page.eval(`localStorage.clear(); localStorage.setItem('writer-settings', JSON.stringify({ startup: 'home', theme: 'light' })); true`);
    await page.goto(eng.url + '/app/mac.dc.html?native=1', 1500);
    await page.until('typeof window.__writerOpen === "function"');
    for (const { type, id, lang } of list) {
      const copy = join(tmp, lang, type, `${id}.${type}`);
      mkdirSync(dirname(copy), { recursive: true }); copyFileSync(fileOf(lang, type, id), copy);
      if (!(await page.eval(`window.__writerOpen(${JSON.stringify(copy)})`))) throw new Error('the app could not open ' + copy);
      if (type === 'xlsx') { await page.until(CLIP.xlsx, 20000, id); await page.click(1300, 1450); } // the selection away from the captured corner
      const clip = await page.until(`document.fonts.status === 'loaded' && !document.querySelector('[data-pop="hud"]') && ${CLIP[type]}`, 20000, `${lang}/${type}/${id}`);
      await sleep(700);
      const box = (await page.eval(CLIP[type])) || clip; // after any late layout
      await page.shot(thumbOf(lang, type, id), { format: 'webp', quality: 82, clip: Object.assign(box, { scale: WIDTH[type] / box.width }) });
      await page.eval(`window.__writerMenu('closeTab'); true`);
      await page.until(`!document.querySelector('[data-edroot]')`, 10000, 'closing ' + id);
    }
    const errs = page.log.filter(l => /EXCEPTION/.test(l)); if (errs.length) console.log('page errors:\n  ' + errs.slice(0, 8).join('\n  '));
  } finally { page.close(); eng.stop(); rmSync(tmp, { recursive: true, force: true }); }
}

// ---- the gallery's list and its English names ----
function writeIndex() {
  const list = [], en = new Map();
  for (const type of TYPES) {
    const cats = new Map(SOURCES[type].cats);
    for (const tpl of SOURCES[type].default) {
      if (!cats.has(tpl.cat)) throw new Error(`${type}/${tpl.id}: category ${tpl.cat} is not in cats`);
      list.push({ id: tpl.id, type, cat: tpl.cat, name: tpl.name[0] });
      for (const [zh, e] of [[tpl.cat, cats.get(tpl.cat)], tpl.name]) {
        if (en.has(zh) && en.get(zh) !== e) throw new Error(`${zh}: both ${en.get(zh)} and ${e}`);
        en.set(zh, e);
      }
    }
  }
  const ids = list.map(x => x.type + '/' + x.id); if (new Set(ids).size !== ids.length) throw new Error('duplicate template id');
  mkdirSync(OUT, { recursive: true });
  writeFileSync(join(OUT, 'index.json'), '[\n' + list.map(x => '  ' + JSON.stringify(x)).join(',\n') + '\n]\n');
  // the English, as a block of en-shell.js; a name the other dictionaries already have must mean the same there ($t keys are shared)
  const dir = join(UI, 'i18n'), file = join(dir, 'en-shell.js'), src = readFileSync(file, 'utf8');
  const BEGIN = '  // ---- template gallery: names and categories, written by ui/tools/templates/build.mjs from its sources ----', END = '  // ---- end of the template gallery ----';
  const a = src.indexOf(BEGIN), b = src.indexOf(END), rest = a >= 0 && b > a ? src.slice(0, a) + src.slice(b + END.length) : src, win = {};
  for (const f of readdirSync(dir).filter(f => /^en-.*\.js$/.test(f))) new Function('window', f === 'en-shell.js' ? rest : readFileSync(join(dir, f), 'utf8'))(win);
  for (const [zh, e] of en) if (Object.hasOwn(win.I18N_EN, zh) && win.I18N_EN[zh] !== e) throw new Error(`${zh}: already ${win.I18N_EN[zh]} in ui/i18n, not ${e}`);
  const q = s => "'" + s.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
  const block = [BEGIN, ...[...en].map(([zh, e]) => `  ${q(zh)}: ${q(e)},`), END].join('\n');
  const next = a >= 0 && b > a ? src.slice(0, a) + block + src.slice(b + END.length) : src.replace(/\n\}\);\s*$/, '\n\n' + block + '\n});\n');
  if (next !== src) writeFileSync(file, next);
  return list;
}

// ---------------------------------------------------------------------------------------------------------------------
const args = process.argv.slice(2), only = args.filter(a => !a.startsWith('--'));
const list = writeIndex().filter(x => !only.length || only.includes(x.type) || only.includes(x.id));
if (!existsSync(CLI)) execFileSync('dotnet', ['build', join(ROOT, 'src/Writer.Cli'), '-v', 'q', '-nologo'], { stdio: 'inherit' });
const eng = await engine(OUT);
try {
  for (const x of list) for (const lang of LANGS) {
    const t0 = Date.now();
    await buildOne(eng.run, x.type, SOURCES[x.type].default.find(t => t.id === x.id), lang);
    console.log(`built ${lang}/${x.type}/${x.id} (${Date.now() - t0} ms)`);
  }
} finally { eng.stop(); }
if (!args.includes('--no-thumbs')) {
  await thumbs(list.flatMap(x => LANGS.map(lang => Object.assign({ lang }, x))));
  console.log(`${list.length * LANGS.length} thumbnails`);
}
