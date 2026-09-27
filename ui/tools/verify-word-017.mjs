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
const dir = mkdtempSync(join(tmpdir(), 'writer-word-017-'));
const engine = spawn('dotnet', [join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll'), 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
let browser, server;
try {
  const base = await new Promise((res, rej) => { let text = ''; engine.stderr.on('data', d => { text += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(text); if (m) res(m[1]); }); engine.once('exit', code => rej(new Error(`engine exit ${code}: ${text}`))); });
  server = createServer(async (req, res) => {
    try {
      if (req.url === '/') { res.setHeader('Content-Type', 'text/html'); res.end('<!doctype html><meta charset="utf-8"><div id="editor" contenteditable="true"></div>'); return; }
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
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const results = await page.evaluate(async () => {
    const EN = await import('/ui/engine.js'), ed = document.querySelector('#editor');
    const check = (condition, message) => { if (!condition) throw new Error(message); };
    const results = [];
    await EN.run(['create', 'comments.docx']);
    for (const html of ['first <b>bold</b> ending', 'second <i>italic</i> ending', 'third paragraph']) await EN.run(['add', 'comments.docx', '/body', '--type', 'paragraph', '--prop', 'html=' + html]);
    let doc = await EN.open({ id: 'comments', path: 'comments.docx', type: 'docx' });
    ed.innerHTML = doc.html;
    const before = ed.textContent, range = document.createRange();
    range.setStart(ed.children[0].firstChild, 3); range.setEnd(ed.children[2].firstChild, 5);
    EN.wrapCommentSelection(ed, range, 'c-test');
    check(ed.children.length === 3 && ed.textContent === before, 'cross-paragraph comment changed text or paragraph count');
    check(!ed.querySelector('span p'), 'paragraph nested inside a comment span');
    doc.comments = [{ id: 'c-test', text: 'Across three paragraphs' }]; doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.children.length === 3 && ed.textContent === before, 'comment round trip changed text');
    check(ed.querySelectorAll('[data-cid]').length >= 3, 'comment anchors did not reopen across paragraphs');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged cross-paragraph comment was rewritten');
    results.push('cross-paragraph formatted comment: text, anchors, save and reopen');

    await EN.run(['create', 'objects.docx']);
    const ole = '<w:object><v:shape id="_x0000_i1025" style="width:72pt;height:36pt"/><o:OLEObject ProgID="Equation.DSMT4"/></w:object>';
    await EN.run(['set', 'objects.docx', '/body', '--raw', '<w:body><w:p><w:r>' + ole + '</w:r></w:p><w:p><w:r><w:t>left</w:t></w:r><w:r>' + ole + '</w:r><w:r><w:t>right</w:t></w:r></w:p></w:body>']);
    doc = await EN.open({ id: 'objects', path: 'objects.docx', type: 'docx' }); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 2, 'object display missing');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged object save mutated file');
    ed.querySelector('p').append(' edited'); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 2, 'text edit lost objects');
    ed.querySelector('p [data-office-object]').remove(); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 1, 'deleted inline object came back');
    ed.querySelector('[data-office-object]').remove(); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 0, 'deleted standalone object came back');
    check(ed.textContent === 'leftright edited', 'object deletion changed text');
    results.push('Office objects: display, no-op save, text edit, inline and block deletion');
    return results;
  });
  assert.equal(results.length, 2);
  for (const result of results) console.log('PASS', result);
} finally {
  await browser?.close(); server?.close(); engine.kill(); rmSync(dir, { recursive: true, force: true });
}
