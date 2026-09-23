// Renders the Mac App Store screenshots: node desktop/store/screenshots/src/shots.mjs [01 02 …]
// Draws the sample photo (photo.html), builds the sample documents in a temporary workspace (make-samples.sh), runs the
// engine on it (serve --no-token with the repo's ui/), opens the Mac window page (mac.dc.html?native=1) in headless
// Chrome, stages each scene through the app itself, captures the window, and frames it with a headline (frame.html).
// Needs macOS (sips; Apple Vision for the cut-out), dotnet and Google Chrome. Nothing opens on screen.
import { spawn, execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, rmSync, existsSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { launch, sleep } from './cdp.mjs';
import { startMock, PROMPT } from './mock-model.mjs';
import { buildVision } from '../../../scripts/build-vision.mjs';

const SRC = dirname(fileURLToPath(import.meta.url));
const ROOT = join(SRC, '../../../..');
const OUT = join(SRC, '../zh-Hans');
const TMP = mkdtempSync(join(process.env.SHOTS_TMP || tmpdir(), 'writer-shots-'));
const WS = join(TMP, 'workspace'), PHOTO = join(TMP, 'latte.jpg');

// The window is the app's default size; the frame shows it at SCALE, captured at 2 × SCALE device pixels so it lands 1:1.
const WIN = { w: 1280, h: 800 }, SCALE = 0.825, TOP = 200;
const GLOW = { word: '70,110,190', excel: '63,125,92', ppt: '214,134,60', mind: '140,100,200', md: '120,120,130' };

/** The shots in display order. open: files opened in turn (the last one in front); stage: what happens before the capture. */
const SHOTS = [
  { id: '01', name: 'all-in-one', title: '一个 App，文档·表格·演示全搞定', sub: 'Word、Excel、PowerPoint、Markdown 和思维导图，都在一个窗口里编辑', glow: GLOW.word, logo: true,
    open: ['门店销售统计.xlsx', '秋季新品发布会.pptx', '新品上市计划.mm', '京都四日行程.md', '秋季新品上市方案.docx'] },
  { id: '02', name: 'ai', title: 'AI 帮你写、帮你改', sub: '改动直接写进文档并标出来，一键保留或撤销；可接入 DeepSeek、通义千问、Kimi 等模型', glow: GLOW.excel,
    open: ['市场部周报.docx'], stage: stageAi },
  { id: '03', name: 'sheet', title: '表格：公式、图表，一样不少', sub: '常用函数即时计算，图表跟着数据变，存下来还是标准的 Excel 文件', glow: GLOW.excel,
    open: ['门店销售统计.xlsx'], stage: stageSheet },
  { id: '04', name: 'slides', title: '演示文稿，原样打开', sub: '版式、配色、表格照原样显示，改完直接放映', glow: GLOW.ppt,
    open: ['秋季新品发布会.pptx'] },
  { id: '05', name: 'cutout', title: '一键抠图', sub: '选中图片点「抠图」，背景自动去掉，只留下主体；在本机完成，原图随时可恢复', glow: GLOW.ppt,
    open: ['秋季新品发布会.pptx'], stage: stageCutout, inset: '原图' },
  { id: '06', name: 'mindmap', title: '思维导图，理清思路', sub: 'Tab 加子主题，Enter 加同级，拖动就能调整结构；还能一键转成文档或演示', glow: GLOW.mind,
    open: ['新品上市计划.mm'], stage: stageMindmap },
  { id: '07', name: 'dark', title: '深色模式，夜里也舒服', sub: '可以跟随系统自动切换，晚上写东西也不刺眼', glow: GLOW.md, theme: 'dark',
    open: ['京都四日行程.md'] },
  { id: '08', name: 'review', title: '修订与批注，和 Word 互通', sub: '对方改过的合同，修订和批注原样显示；可以接着改，一键接受或拒绝全部修订', glow: GLOW.word,
    open: ['联名合作协议.docx'], stage: stageReview },
];

// ---------------------------------------------------------------------------------------------------------------------

const freePort = () => new Promise(resolve => { const s = createServer(); s.listen(0, '127.0.0.1', () => { const p = s.address().port; s.close(() => resolve(p)); }); });

async function startEngine(mockUrl) {
  execFileSync('dotnet', ['build', join(ROOT, 'src/Writer.Cli'), '-v', 'q', '-nologo'], { stdio: 'inherit' });
  const vision = join(ROOT, 'desktop/src-tauri/binaries', `writer-vision-${process.arch === 'x64' ? 'x86_64' : 'aarch64'}-apple-darwin`);
  if (!existsSync(vision)) buildVision();
  const port = await freePort();
  const env = Object.assign({}, process.env, { WRITER_VISION: vision, WRITER_AI_PROVIDER: 'custom', WRITER_AI_BASE_URL: mockUrl, WRITER_AI_MODEL: 'screenshot-mock', WRITER_AI_KEY: '' });
  const proc = spawn('dotnet', [join(ROOT, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll'), 'serve', '--dir', WS, '--port', String(port), '--no-token', '--ui', join(ROOT, 'ui')], { env, stdio: ['ignore', 'ignore', 'pipe'] });
  let err = ''; proc.stderr.on('data', d => { err += d; });
  const base = `http://127.0.0.1:${port}`;
  for (let i = 0; i < 100; i++) { try { if ((await fetch(base + '/files')).ok) return { base, stop: () => proc.kill() }; } catch (e) { } await sleep(100); }
  throw new Error('engine did not start: ' + err);
}

/** The photo the cut-out shot works on, drawn by photo.html. */
async function drawPhoto(p) {
  await p.size(1200, 1200, 1);
  await p.goto(pathToFileURL(join(SRC, 'photo.html')).href, 300);
  await p.until(`document.body.dataset.ready === '1'`);
  const url = await p.eval(`document.getElementById('c').toDataURL('image/jpeg', 0.92)`);
  writeFileSync(PHOTO, Buffer.from(url.split(',')[1], 'base64'));
}

/** Center of the first visible element matching a CSS selector (and, when given, whose text is `text`), in CSS px. */
async function center(p, sel, text) {
  return p.until(`(() => { const el = [...document.querySelectorAll(${JSON.stringify(sel)})].find(e => ${text == null ? 'true' : `e.textContent.trim() === ${JSON.stringify(text)}`} && e.getClientRects().length);
    if (!el) return null; const b = el.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; })()`, 15000, sel + (text ? ' ' + text : ''));
}
async function clickOn(p, sel, text) { const c = await center(p, sel, text); await p.click(c.x, c.y); return c; }

/** Waits until the window is at rest: fonts in, no toast, animations done. */
async function settle(p, ms = 900) {
  await p.until(`document.fonts.status === 'loaded' && !document.querySelector('[data-pop="hud"]')`, 20000, 'toasts to clear');
  await sleep(ms);
}

// ---- staging: everything goes through the app's own controls ----

async function stageAi(p) {
  await p.eval(`window.__writerMenu('toggleSidebar'); true`); await sleep(600);
  await p.eval(`window.__writerMenu('toggleAI'); true`);
  await clickOn(p, 'textarea[placeholder="让助手修改…"]');
  await p.type(PROMPT); await sleep(300);
  await clickOn(p, 'button[data-accent-btn]');
  await p.until(`[...document.querySelectorAll('button')].some(b => b.textContent.trim() === '保留')`, 30000, 'the assistant turn');
  await p.eval(`document.activeElement && document.activeElement.blur(); true`);
}

/** Clicks a cell of the sheet by its column letter and row number. */
async function selectCell(p, col, row) {
  const c = await p.until(`(() => { const ds = [...document.querySelectorAll('[data-edroot] div')];
    const ch = ds.find(d => d.style.gridRow === '1' && d.textContent.trim() === ${JSON.stringify(col)}), rh = ds.find(d => d.style.gridColumn === '1' && d.textContent.trim() === ${JSON.stringify(String(row))});
    if (!ch || !rh) return null; const a = ch.getBoundingClientRect(), b = rh.getBoundingClientRect(); return { x: a.left + a.width / 2, y: b.top + b.height / 2 }; })()`, 10000, 'cell ' + col + row);
  await p.click(c.x, c.y);
}

/** 120% (the status bar's ＋, twice; it zooms about the middle, so scroll back to A1), then the grand total, so the formula bar shows its SUM. */
async function stageSheet(p) {
  for (let i = 0; i < 2; i++) { await clickOn(p, '[data-edroot] button[title^="放大"]'); await sleep(300); }
  await p.wheel(640, 500, -3000, -3000); await sleep(500);
  await selectCell(p, 'F', 11);
}

/** No page sidebar, the map fitted to the window, the pointer back on the empty canvas. */
async function stageMindmap(p) {
  await p.eval(`window.__writerMenu('toggleSidebar'); true`); await sleep(600);
  await clickOn(p, '[data-edroot] button', '适应画布'); await sleep(400);
  await p.hover(80, 120);
}

/** The contract without the page sidebar (the comments take the right side), and the 审阅 tab open. */
async function stageReview(p) {
  await p.eval(`window.__writerMenu('toggleSidebar'); true`); await sleep(600);
  await clickOn(p, '[data-glass-bar] button', '审阅');
}

async function stageCutout(p) {
  const t = await p.until(`(() => { const t = document.querySelectorAll('[data-thumbpanel] [draggable="true"]')[1]; if (!t) return null; const b = t.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + b.height / 2 }; })()`, 10000, 'slide 2');
  await p.click(t.x, t.y); await sleep(600);
  const pic = await p.until(`(() => { const o = [...document.querySelectorAll('[data-soid]')].filter(e => e.querySelector('img') && e.getClientRects().length).sort((a, b) => b.offsetWidth - a.offsetWidth)[0];
    if (!o) return null; const b = o.getBoundingClientRect(); return { x: b.left + b.width * 0.3, y: b.top + b.height * 0.5 }; })()`, 10000, 'the picture');
  await p.click(pic.x, pic.y); await sleep(500);
  await clickOn(p, '[data-glass-bar] button', '图片'); await sleep(700);
  const src = `[...document.querySelectorAll('[data-soid] img')].map(i => i.src).join()`, before = await p.eval(src);
  await clickOn(p, '[data-bubble] button', '抠图');
  // the picture reloads once the engine has cut it out; a toast instead means the cut-out failed
  const r = await p.until(`(() => { const t = document.querySelector('[data-pop="hud"]'); return t ? 'failed: ' + t.textContent : ${src} !== ${JSON.stringify(before)} && 'ok'; })()`, 60000, 'the cut-out');
  if (r !== 'ok') throw new Error('cut-out ' + r);
  await sleep(1500);
}

// ---- capture + frame ----

async function capture(p, base, shot) {
  const dark = shot.theme === 'dark';
  await p.size(WIN.w, WIN.h, 2 * SCALE);
  await p.scheme(dark);
  await p.goto(base + '/files', 300);
  await p.eval(`localStorage.clear(); localStorage.setItem('writer-settings', ${JSON.stringify(JSON.stringify({ theme: dark ? 'dark' : 'light', startup: 'home' }))}); true`);
  await p.goto(base + '/app/mac.dc.html?native=1', 1200);
  await p.until('typeof window.__writerOpen === "function"');
  for (const f of shot.open) {
    if (!(await p.eval(`window.__writerOpen(${JSON.stringify(join(WS, f))})`))) throw new Error('could not open ' + f);
    await p.until(`!!document.querySelector('[data-edroot]')`, 20000, f); await settle(p, 500);
  }
  if (shot.stage) await shot.stage(p);
  await settle(p);
  const file = join(TMP, `window-${shot.id}.png`);
  await p.shot(file);
  return file;
}

async function frame(p, shot, windowPng) {
  await p.size(1440, 900, 2);
  await p.scheme(shot.theme === 'dark');
  const q = new URLSearchParams({ t: shot.title, s: shot.sub || '', img: pathToFileURL(windowPng).href, w: WIN.w * SCALE, h: WIN.h * SCALE, top: TOP,
    theme: shot.theme || 'light', glow: shot.glow || '', logo: shot.logo ? '1' : '', inset: shot.inset ? pathToFileURL(PHOTO).href : '', insetLabel: shot.inset || '', insetTop: 470, logoSrc: pathToFileURL(join(ROOT, 'ui/assets', shot.theme === 'dark' ? 'writer-logo-dark.svg' : 'writer-logo-light.svg')).href });
  await p.goto(pathToFileURL(join(SRC, 'frame.html')).href + '?' + q, 300);
  await p.until(`document.body.dataset.ready === '1' && document.fonts.status === 'loaded'`);
  await sleep(200);
  const out = join(OUT, `${shot.id}-${shot.name}.png`);
  await p.shot(out);
  execFileSync('sips', ['--embedProfile', '/System/Library/ColorSync/Profiles/sRGB Profile.icc', out], { stdio: 'ignore' });
  const info = execFileSync('sips', ['-g', 'pixelWidth', '-g', 'pixelHeight', '-g', 'hasAlpha', out]).toString();
  if (!/pixelWidth: 2880/.test(info) || !/pixelHeight: 1800/.test(info) || !/hasAlpha: no/.test(info)) throw new Error('bad output ' + out + '\n' + info);
  return out;
}

// ---------------------------------------------------------------------------------------------------------------------

const only = process.argv.slice(2);
const todo = SHOTS.filter(s => !only.length || only.includes(s.id));
mkdirSync(OUT, { recursive: true });
const page = await launch({ width: 1440, height: 900, scale: 2 });
let mock, engine;
try {
  await drawPhoto(page);
  execFileSync('bash', [join(SRC, 'make-samples.sh'), WS, PHOTO], { stdio: 'inherit' });
  mock = await startMock();
  engine = await startEngine(mock.url);
  for (const shot of todo) {
    const win = await capture(page, engine.base, shot);
    console.log('wrote', await frame(page, shot, win));
    const errs = page.log.filter(l => /EXCEPTION|error/i.test(l)); if (errs.length) console.log('  page errors:', errs.slice(0, 5).join(' | ')); page.log.length = 0;
  }
} finally {
  page.close(); if (engine) engine.stop(); if (mock) mock.close();
  if (!process.env.KEEP) rmSync(TMP, { recursive: true, force: true }); else console.log('kept', TMP);
}
