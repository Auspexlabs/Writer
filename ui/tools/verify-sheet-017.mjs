// Real DOM + engine round trips. Run after dotnet build src/Writer.Cli.
// Set WRITER_NODE_MODULES when Playwright is supplied outside this repository.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../..', import.meta.url));
const { chromium } = await import(process.env.WRITER_NODE_MODULES ? pathToFileURL(join(process.env.WRITER_NODE_MODULES, 'playwright/index.mjs')).href : 'playwright');
const dir = mkdtempSync(join(tmpdir(), 'writer-sheet-017-'));
const engine = spawn('dotnet', [join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll'), 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
let browser, server;
try {
  const base = await new Promise((res, rej) => { let text = ''; engine.stderr.on('data', d => { text += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(text); if (m) res(m[1]); }); engine.once('exit', code => rej(new Error(`engine exit ${code}: ${text}`))); });
  server = createServer(async (req, res) => {
    try {
      if (req.url === '/favicon.ico') { res.writeHead(204).end(); return; }
      if (req.url === '/ui/test.dc.html') { res.setHeader('Content-Type', 'text/html'); res.end(`<!doctype html><meta charset="utf-8"><style>html,body,#dc-root{height:100%;margin:0;overflow:hidden}</style><script>window.__testBook = {active:0,sheets:[{name:'Data',cells:{A1:{v:'Label'},A2:{v:'one'},AAA10001:{v:'far cell'}}}]};window.__WRITER_LANG='zh';</script><script src="/ui/memo.js"></script><script src="/ui/i18n.js"></script><script src="/ui/vendor/offline.js"></script><script src="/ui/support.js"></script><body><x-dc><dc-import name="SheetEditor" doc="{{ doc }}" on-change="{{ onChange }}" toast="{{ toast }}" style="position:absolute;inset:0"></dc-import></x-dc><script type="text/x-dc" data-dc-script>class Component extends DCLogic { state = {doc:window.__testBook}; constructor(p){super(p);window.__sheetParent=this;} renderVals(){return {doc:this.state.doc,onChange:doc=>this.setState({doc}),toast:message=>{window.__lastToast=message;}};} }</script></body>`); return; }
      if (req.url.startsWith('/ui/')) {
        const path = resolve(root, '.' + new URL(req.url, 'http://local').pathname);
        if (!path.startsWith(join(root, 'ui') + '/')) { res.writeHead(403).end(); return; }
        res.setHeader('Content-Type', path.endsWith('.js') ? 'text/javascript' : 'text/html'); res.end(readFileSync(path)); return;
      }
      const chunks = []; for await (const chunk of req) chunks.push(chunk);
      const response = await fetch(base + req.url, { method: req.method, headers: { 'Content-Type': req.headers['content-type'] || 'application/json' }, ...(chunks.length ? { body: Buffer.concat(chunks) } : {}) });
      res.writeHead(response.status, { 'Content-Type': response.headers.get('Content-Type') || 'application/json' }); res.end(Buffer.from(await response.arrayBuffer()));
    } catch (error) { res.writeHead(500).end(String(error)); }
  });
  await new Promise(res => server.listen(0, '127.0.0.1', res));
  browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage();
  page.on('console', m => { if (m.type() === 'error') console.error('BROWSER', m.text()); });
  page.on('pageerror', e => console.error('PAGE', e.message));
  await page.goto(`http://127.0.0.1:${server.address().port}/ui/test.dc.html`);
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await page.locator('.sh-cell').first().waitFor();
  assert.ok(await page.locator('.sh-cell').count() < 1000, 'grid mounts a viewport, not all rows');
  const name = page.locator('input').first(); await name.fill('AAA10001'); await name.press('Enter');

  await page.getByText('far cell', {exact:true}).waitFor({timeout:3000});
  assert.ok(await page.locator('.sh-cell').count() < 1000, 'far selection stays virtualized');
  const results = await page.evaluate(async () => {
    const E = await import('/ui/sheet-engine.js'), EN = await import('/ui/engine.js');
    const check = (yes, msg) => { if (!yes) throw new Error(msg); };
    const clip = E.readTableHTML('<style>.xl1{font-weight:700;color:#ff0000}</style><table><tr><td class="xl1" colspan="2">hello<br>world</td></tr><tr><td>001</td><td style="background:#00ff00">12</td></tr></table>');
    check(clip.cells[0][0].v === 'hello\nworld' && clip.cells[0][0].s.b, 'Office HTML text/format');
    check(clip.merges[0].cs === 2, 'Office HTML merged cell');
    await EN.run(['create', 'grid.xlsx']);
    let doc = await EN.open({id:'grid',path:'grid.xlsx',type:'xlsx'});
    doc = E.editWorkbook(doc, (sh,d) => { sh.cells.AAA10001 = {v:'far saved'}; sh.cells.A1={v:'7'}; sh.cells.B1={v:'=SEQUENCE(2,2)'}; d.names={Total:'Sheet1!$A$1'}; sh.cells.E1={v:'=Total*2'}; sh.print={orientation:'landscape',fitWidth:1,fitHeight:0,titles:'Sheet1!$1:$1',footer:'&P / &N'}; });
    await EN.save(doc); doc = await EN.open(doc);
    check(doc.sheets[0].cells.AAA10001.v === 'far saved', 'far cell round trip');
    check(new E.Calc(doc).value(0,0,4) === 14, 'defined name after reopen');
    check(doc.sheets[0].cells.C2.spill === 'B1', 'cached dynamic array children');
    check(await EN.save(doc) === 0, 'unchanged workbook rewrites nothing');
    const src = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jFq8AAAAASUVORK5CYII=';
    doc = E.editWorkbook(doc, sh => { sh.images = [{ id:'test-image',src,x:96,y:48,w:120,h:80,look:{} }]; }); await EN.save(doc); doc = await EN.open(doc);
    check(doc.sheets[0].images.length === 1, 'inserted image missing');
    doc = E.editWorkbook(doc, sh => { sh.images[0].x=192; sh.images[0].w=240; }); await EN.save(doc); doc = await EN.open(doc);
    check(Math.abs(doc.sheets[0].images[0].x - 192) < 1 && Math.abs(doc.sheets[0].images[0].w - 240) < 1, 'image move/resize lost');
    const image = {...doc.sheets[0].images[0],src:await EN.dataUrlOf(doc.sheets[0].images[0].src)};
    doc = E.editWorkbook(doc, sh => { sh.images=[]; }); await EN.save(doc); doc = await EN.open(doc); check(doc.sheets[0].images.length === 0, 'image removal lost');
    doc = E.editWorkbook(doc, sh => { sh.images=[image]; }); await EN.save(doc); doc = await EN.open(doc); check(doc.sheets[0].images.length === 1, 'deleted image undo lost bytes');
    return ['virtual grid navigation to AAA10001', 'Office HTML multiline, style and merge import', 'far cell, names, dynamic arrays and print settings save/reopen', 'image insert, move, resize, delete and undo round trips'];
  });
  assert.deepEqual(errors, []); for (const result of results) console.log('PASS', result);
} finally {
  await browser?.close(); server?.close(); engine.kill(); rmSync(dir, { recursive: true, force: true });
}
