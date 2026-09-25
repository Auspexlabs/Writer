// Headless Chrome over --remote-debugging-pipe (CDP on fds 3/4, NUL-separated JSON): no port, no window, its own profile.
import { spawn } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const CHROME = process.env.CHROME || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
export const sleep = ms => new Promise(r => setTimeout(r, ms));

/** A headless Chrome with one page of width × height CSS px at `scale` device pixels per px. */
export async function launch({ width = 1440, height = 900, scale = 2 } = {}) {
  const profile = mkdtempSync(join(process.env.SHOTS_TMP || tmpdir(), 'chrome-'));
  const proc = spawn(CHROME, ['--headless=new', '--remote-debugging-pipe', '--user-data-dir=' + profile, '--no-first-run',
    '--no-default-browser-check', '--hide-scrollbars', '--force-color-profile=srgb', `--window-size=${width},${height}`, 'about:blank'],
    { stdio: ['ignore', 'ignore', 'ignore', 'pipe', 'pipe'] });
  const out = proc.stdio[3], inp = proc.stdio[4];
  let id = 0, buf = '';
  const waiting = new Map(), listeners = [];
  inp.on('data', d => {
    buf += d;
    let i;
    while ((i = buf.indexOf('\0')) >= 0) {
      const msg = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1);
      if (msg.id && waiting.has(msg.id)) { const w = waiting.get(msg.id); waiting.delete(msg.id); msg.error ? w.reject(new Error(JSON.stringify(msg.error))) : w.resolve(msg.result); }
      else for (const l of listeners) l(msg);
    }
  });
  const send = (method, params = {}, sessionId) => new Promise((resolve, reject) => {
    const n = ++id; waiting.set(n, { resolve, reject });
    out.write(JSON.stringify(Object.assign({ id: n, method, params }, sessionId ? { sessionId } : {})) + '\0');
  });
  const { targetId } = await send('Target.createTarget', { url: 'about:blank' });
  const { sessionId } = await send('Target.attachToTarget', { targetId, flatten: true });
  const s = (m, p) => send(m, p, sessionId);
  await s('Page.enable'); await s('Runtime.enable');
  await s('Emulation.setTimezoneOverride', { timezoneId: process.env.SHOTS_TZ || 'Asia/Shanghai' }); // times shown in the app read as Beijing time
  const log = [];
  listeners.push(m => {
    if (m.sessionId !== sessionId) return;
    if (m.method === 'Runtime.exceptionThrown') log.push('EXCEPTION ' + (m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text));
    else if (m.method === 'Runtime.consoleAPICalled') log.push(m.params.type + ' ' + m.params.args.map(a => a.value ?? a.description).join(' '));
  });
  const page = {
    log,
    async size(w, h, sc = scale) { await s('Emulation.setDeviceMetricsOverride', { width: w, height: h, deviceScaleFactor: sc, mobile: false }); },
    /** prefers-color-scheme for the page (the app itself follows its own theme setting). */
    async scheme(dark) { await s('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: dark ? 'dark' : 'light' }] }); },
    async goto(url, wait = 1500) { await s('Page.navigate', { url }); await sleep(wait); },
    async eval(expr) {
      const r = await s('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true });
      if (r.exceptionDetails) throw new Error('eval: ' + (r.exceptionDetails.exception?.description || r.exceptionDetails.text));
      return r.result.value;
    },
    /** Polls `expr` until it is truthy; returns its value. */
    async until(expr, ms = 15000, what = expr) {
      const end = Date.now() + ms;
      for (;;) { const v = await page.eval(expr).catch(() => null); if (v) return v; if (Date.now() > end) throw new Error('timed out waiting for ' + what); await sleep(150); }
    },
    async mouse(type, x, y, extra = {}) { await s('Input.dispatchMouseEvent', Object.assign({ type, x, y, button: 'left', clickCount: 1 }, extra)); },
    async click(x, y) { await page.mouse('mouseMoved', x, y); await page.mouse('mousePressed', x, y); await page.mouse('mouseReleased', x, y); },
    async hover(x, y) { await page.mouse('mouseMoved', x, y, { button: 'none' }); },
    async wheel(x, y, deltaX, deltaY) { await page.mouse('mouseWheel', x, y, { button: 'none', deltaX, deltaY }); },
    async type(text) { await s('Input.insertText', { text }); },
    /** PNG of the viewport; opts go to Page.captureScreenshot as they are (format, quality, clip). */
    async shot(file, opts) {
      const r = await s('Page.captureScreenshot', Object.assign({ format: 'png', fromSurface: true }, opts));
      writeFileSync(file, Buffer.from(r.data, 'base64'));
      return file;
    },
    close() { proc.kill(); setTimeout(() => rmSync(profile, { recursive: true, force: true }), 500); }
  };
  await page.size(width, height, scale);
  return page;
}
