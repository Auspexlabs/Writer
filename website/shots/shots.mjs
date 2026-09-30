// Renders the website and README screenshots: node website/shots/shots.mjs [hero sheet …] [--light|--dark]
// Builds the sample documents in a temporary workspace (make-samples.sh), runs the engine on it (serve --no-token with the
// repo's ui/), opens the Mac window page (mac.dc.html?native=1) in headless Chrome, stages each scene through the app
// itself and captures the window as WebP, light and dark, into website/dist/assets/shots/ (the README's docs/images/ get
// the light hero and the four format shots). The AI scenes talk to mock-model.mjs; the engine really edits the file.
// Needs macOS, Google Chrome and an engine: `dotnet` (the repo is built), WRITER=<writer binary>, or the installed
// /Applications/Writer.app. Nothing opens on screen.
import { spawn, execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, rmSync, existsSync, copyFileSync, readdirSync } from 'node:fs';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { launch, sleep } from '../../desktop/store/screenshots/src/cdp.mjs';
import { startMock, SCENES } from './mock-model.mjs';

const SRC = dirname(fileURLToPath(import.meta.url));
const ROOT = join(SRC, '../..');
const OUT = join(ROOT, 'website/dist/assets/shots'), DOCS = join(ROOT, 'docs/images');
const TMP = mkdtempSync(join(process.env.SHOTS_TMP || tmpdir(), 'writer-site-shots-'));
const WS = join(TMP, 'workspace'), PRISTINE = join(TMP, 'pristine');
const QUALITY = 88;

/** The shots: the window in CSS px (captured at 2 device px per px), the files opened, the staging before the capture.
 *  readme: also the README's copy (light). clip: a part of the window instead (theme, the card of 浅色与深色), or the whole
 *  window at another scale (show, 1 device px per px). */
const SHOTS = [
  { id: 'hero', w: 1200, h: 750, open: ['咖啡节活动方案.docx'], stage: p => stageAi(p, 'thumbs'), scene: 'dates', readme: true },
  { id: 'sheet', w: 1024, h: 680, open: ['咖啡节预算.xlsx'], stage: p => selectCell(p, 'D', 8), readme: true },
  { id: 'slides', w: 1024, h: 680, open: ['咖啡节方案.pptx'], stage: p => menu(p, 'toggleSidebar'), readme: true },
  { id: 'mindmap', w: 1024, h: 680, open: ['咖啡节筹备.mm'], stage: stageMindmap, readme: true },
  { id: 'markdown', w: 1024, h: 680, open: ['筹备会议纪要.md'], stage: p => menu(p, 'toggleSidebar'), readme: true },
  { id: 'ai', w: 920, h: 680, open: ['咖啡节预算.xlsx'], stage: p => stageAi(p), scene: 'budget' },
  { id: 'theme', w: 1200, h: 750, open: ['咖啡节活动方案.docx'], clip: { x: 0, y: 0, width: 380, height: 248, scale: 1 } },
  { id: 'show', w: 760, h: 427, open: ['咖啡节方案.pptx'], stage: stageShow, clip: { x: 0, y: 0, width: 760, height: 427, scale: 0.5 } },
];

// ---------------------------------------------------------------------------------------------------------------------

const freePort = () => new Promise(resolve => { const s = createServer(); s.listen(0, '127.0.0.1', () => { const p = s.address().port; s.close(() => resolve(p)); }); });
const has = cmd => { try { execFileSync('/bin/sh', ['-c', `command -v ${cmd}`], { stdio: 'ignore' }); return true; } catch (e) { return false; } };

/** The engine command: WRITER, else the repo's build (built here when dotnet is around), else the installed app's. */
function engineCommand() {
  if (process.env.WRITER) return [process.env.WRITER];
  const dll = join(ROOT, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
  if (has('dotnet')) { execFileSync('dotnet', ['build', join(ROOT, 'src/Writer.Cli'), '-v', 'q', '-nologo'], { stdio: 'inherit' }); return ['dotnet', dll]; }
  const app = '/Applications/Writer.app/Contents/MacOS/writer';
  if (existsSync(app)) { console.log('no dotnet: using the installed app\'s engine', app); return [app]; }
  throw new Error('no engine: install dotnet, set WRITER to a writer binary, or install Writer.app');
}

async function startEngine(cmd, mockUrl) {
  const port = await freePort();
  const env = Object.assign({}, process.env, { WRITER_AI_PROVIDER: 'custom', WRITER_AI_BASE_URL: mockUrl, WRITER_AI_MODEL: 'deepseek-chat' /* the name the AI panel shows; the replies are mock-model.mjs's */, WRITER_AI_KEY: '' });
  const proc = spawn(cmd[0], [...cmd.slice(1), 'serve', '--dir', WS, '--port', String(port), '--no-token', '--ui', join(ROOT, 'ui')], { env, stdio: ['ignore', 'ignore', 'pipe'] });
  let err = ''; proc.stderr.on('data', d => { err += d; });
  const base = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 100; i++) { try { if ((await fetch(base + '/files')).ok) return { base, stop: () => proc.kill() }; } catch (e) { } await sleep(100); }
  throw new Error('engine did not start: ' + err);
}

/** The sample files as make-samples.sh wrote them (an AI scene edits its file; the next capture starts afresh). */
function restoreSamples() { for (const f of readdirSync(PRISTINE)) copyFileSync(join(PRISTINE, f), join(WS, f)); }

/** Center of the first visible element matching a CSS selector (and, when given, whose text is `text`), in CSS px. */
async function center(p, sel, text) {
  return p.until(`(() => { const el = [...document.querySelectorAll(${JSON.stringify(sel)})].find(e => ${text == null ? 'true' : `e.textContent.trim() === ${JSON.stringify(text)}`} && e.getClientRects().length);
    if (!el) return null; const b = el.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; })()`, 15000, sel + (text ? ' ' + text : ''));
}
async function clickOn(p, sel, text) { const c = await center(p, sel, text); await p.click(c.x, c.y); return c; }
async function menu(p, id) { await p.eval(`window.__writerMenu(${JSON.stringify(id)}); true`); await sleep(700); }

/** Waits until the window is at rest: fonts in, no toast, animations done. */
async function settle(p, ms = 900) {
  await p.until(`document.fonts.status === 'loaded' && !document.querySelector('[data-pop="hud"]')`, 20000, 'toasts to clear');
  await sleep(ms);
}

// ---- staging: everything goes through the app's own controls ----

/** The AI panel with the scene's request answered: the assistant's reply, its commands, the change list and 保留 / 撤销. */
async function stageAi(p, sidebar) {
  if (sidebar) await menu(p, 'toggleSidebar');
  await menu(p, 'toggleAI');
  await clickOn(p, 'textarea[placeholder="让助手修改…"]');
  await p.type(p.scene.prompt); await sleep(300);
  await clickOn(p, 'button[data-accent-btn]');
  await p.until(`[...document.querySelectorAll('button')].some(b => b.textContent.trim() === '保留')`, 30000, 'the assistant turn');
  await p.eval(`document.activeElement && document.activeElement.blur(); true`);
  await sleep(400);
}

/** Clicks a cell of the sheet by its column letter and row number. */
async function selectCell(p, col, row) {
  const c = await p.until(`(() => { const ds = [...document.querySelectorAll('[data-edroot] div')];
    const ch = ds.find(d => d.style.gridRow === '1' && d.textContent.trim() === ${JSON.stringify(col)}), rh = ds.find(d => d.style.gridColumn === '1' && d.textContent.trim() === ${JSON.stringify(String(row))});
    if (!ch || !rh) return null; const a = ch.getBoundingClientRect(), b = rh.getBoundingClientRect(); return { x: a.left + a.width / 2, y: b.top + b.height / 2 }; })()`, 10000, 'cell ' + col + row);
  await p.click(c.x, c.y);
}

/** The map fitted to the window, the pointer back on the empty canvas. */
async function stageMindmap(p) {
  await clickOn(p, '[data-edroot] button', '适应画布'); await sleep(400);
  await p.hover(60, 120);
}

/** 放映 from the first slide (F5), as the 放映 tab's 从头开始 does. */
async function stageShow(p) {
  await p.eval(`(() => { const t = document.querySelector('[data-edroot]') || document.body; t.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', code: 'F5', bubbles: true, cancelable: true })); return true; })()`);
  await sleep(1500);
}

// ---- capture ----

async function capture(p, base, shot, dark) {
  restoreSamples();
  p.scene = shot.scene ? SCENES[shot.scene] : null;
  await p.size(shot.w, shot.h, 2);
  await p.scheme(dark);
  await p.goto(base + '/files', 300);
  // dark: 设置 › 外观 › 深色页面 too, so the page and the assistant's change marks are drawn in the dark palette as one
  await p.eval(`localStorage.clear(); localStorage.setItem('writer-settings', ${JSON.stringify(JSON.stringify({ theme: dark ? 'dark' : 'light', startup: 'home', darkPages: dark }))}); true`);
  await p.goto(base + '/app/mac.dc.html?native=1', 1200);
  await p.until('typeof window.__writerOpen === "function"');
  for (const f of shot.open) {
    if (!(await p.eval(`window.__writerOpen(${JSON.stringify(join(WS, f))})`))) throw new Error('could not open ' + f);
    await p.until(`!!document.querySelector('[data-edroot]')`, 20000, f); await settle(p, 500);
  }
  if (shot.stage) await shot.stage(p);
  await settle(p);
  const name = `${shot.id}-${dark ? 'dark' : 'light'}`;
  await p.shot(join(TMP, name + '.png'));
  const out = join(OUT, name + '.webp');
  await p.shot(out, { format: 'webp', quality: QUALITY, clip: shot.clip });
  const info = execFileSync('sips', ['-g', 'pixelWidth', '-g', 'pixelHeight', out]).toString();
  const want = shot.clip ? [shot.clip.width * shot.clip.scale * 2, shot.clip.height * shot.clip.scale * 2] : [shot.w * 2, shot.h * 2];
  if (!new RegExp(`pixelWidth: ${want[0]}\\b`).test(info) || !new RegExp(`pixelHeight: ${want[1]}\\b`).test(info)) throw new Error('bad output ' + out + '\n' + info);
  if (shot.readme && !dark) copyFileSync(out, join(DOCS, shot.id + '.webp'));
  return out;
}

// ---------------------------------------------------------------------------------------------------------------------

const args = process.argv.slice(2), only = args.filter(a => !a.startsWith('--'));
const themes = args.includes('--light') ? [false] : args.includes('--dark') ? [true] : [false, true];
const todo = SHOTS.filter(s => !only.length || only.includes(s.id));
if (!todo.length) throw new Error('no such shot; one of ' + SHOTS.map(s => s.id).join(' '));
mkdirSync(OUT, { recursive: true }); mkdirSync(DOCS, { recursive: true });
const engineCmd = engineCommand();
execFileSync('bash', [join(SRC, 'make-samples.sh'), WS], { env: Object.assign({}, process.env, { WRITER: engineCmd.join(' ') }), stdio: 'inherit' });
mkdirSync(PRISTINE); for (const f of readdirSync(WS)) copyFileSync(join(WS, f), join(PRISTINE, f));
const page = await launch({ width: 1200, height: 750, scale: 2 });
let mock, engine;
try {
  mock = await startMock();
  engine = await startEngine(engineCmd, mock.url);
  for (const shot of todo) for (const dark of themes) {
    console.log('wrote', await capture(page, engine.base, shot, dark));
    const errs = page.log.filter(l => /EXCEPTION|error/i.test(l)); if (errs.length) console.log('  page errors:', errs.slice(0, 5).join(' | ')); page.log.length = 0;
  }
} finally {
  page.close(); if (engine) engine.stop(); if (mock) mock.close();
  if (!process.env.KEEP) rmSync(TMP, { recursive: true, force: true }); else console.log('kept', TMP);
}
