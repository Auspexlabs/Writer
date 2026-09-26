// node --test ui/tests/ — the engine inside the page (ui/embed/server.js): the editors' server requests answered by the engine
// (here a stand-in with the WebAssembly engine's calls), changes told to the editor's /events and to the site, errors with the
// server's status and body, blob: URLs for what the page loads by itself.
import { test } from 'node:test';
import assert from 'node:assert/strict';

// a page: fetch and EventSource to replace, URL.createObjectURL for the blob: URLs
const made = [];
globalThis.window = globalThis;
globalThis.location = new URL('https://embed.example/writer/embed.dc.html');
globalThis.fetch = async (u) => new Response('static ' + u);
globalThis.EventSource = class { constructor(u) { this.url = u; this.native = true; } };
URL.createObjectURL = b => { made.push(b); return 'blob:' + made.length; };
URL.revokeObjectURL = () => { };

/** The Writer.Browser.Engine calls, over a Map of files: versions go up by one on each write, as the engine's do. */
function fakeEngine() {
  const files = new Map(), changed = [], versions = new Map(); let v = 1000;
  const bump = p => { versions.set(p, ++v); if (!changed.includes(p)) changed.push(p); };
  const need = p => { if (!files.has(p)) throw new Error(JSON.stringify({ status: 404, error: { code: 'FILE_NOT_FOUND', message: p + ' not found', hint: 'Check the path.' } })); return p; };
  return {
    files, calls: [],
    Run(body) { const { command, argv } = JSON.parse(body); const a = argv || command.split(' '); this.calls.push(a); if (a[0] === 'create') { files.set(a[1], new Uint8Array([80, 75])); bump(a[1]); } if (a[0] === 'add') { files.set(a[1], new Uint8Array([...files.get(a[1]), 33])); bump(a[1]); } return JSON.stringify({ code: 0, output: 'ok' }); },
    Files() { return JSON.stringify({ workspace: '/work', chat: false, files: [...files.keys()].map(p => ({ path: p })) }); },
    Stat(p) { need(p); return JSON.stringify({ path: p, mtime: versions.get(p) || 1, size: files.get(p).length }); },
    Read(p) { return files.get(need(p)); },
    Write(p, bytes, ifMtime) { if (ifMtime && files.has(p) && String(versions.get(p)) !== ifMtime) throw new Error(JSON.stringify({ status: 409, error: { code: 'CONFLICT', message: 'changed', hint: '' } })); files.set(p, bytes); bump(p); return JSON.stringify({ path: p, size: bytes.length, mtime: versions.get(p) }); },
    Move(from, to, keep) { need(from); files.set(to, files.get(from)); if (!keep) files.delete(from); bump(to); return JSON.stringify({ path: to }); },
    Delete(p) { need(p); files.delete(p); bump(p); return JSON.stringify({ deleted: p }); },
    Json(p, skip) { need(p); return JSON.stringify({ type: 'document', skip }); },
    View(p, mode) { need(p); return mode + ' of ' + p; },
    Binary(p, node) { need(p); this.bin = node; return new Uint8Array([1, 2, 3]); },
    BinaryType() { return 'image/png'; },
    Drain() { const out = JSON.stringify(changed); changed.length = 0; return out; },
    ChatSystem(p, outline) { return 'SYSTEM ' + p + ' ' + outline; },
    ChatTools() { return JSON.stringify([{ name: 'writer' }, { name: 'batch' }, { name: 'plan' }]); },
    ChatTool(name, input) { this.calls.push([name, JSON.parse(input)]); return JSON.stringify({ display: 'x', code: 0, output: 'done', wrote: false }); }
  };
}

const S = await import('../embed/server.js');
const E = S.useEngine(fakeEngine());
S.install();
const ask = (path, init) => fetch('https://embed.example' + path, init);

test('the editors\' requests reach the engine: /files, /stat, /file, /json, /html, /binary, /run — and nothing else is taken', async () => {
  E.files.set('a.docx', new Uint8Array([80, 75]));
  const list = await (await ask('/files')).json();
  assert.deepEqual([list.files.map(f => f.path), list.chat], [['a.docx'], false], 'the workspace, and no assistant yet');
  assert.equal((await (await ask('/stat?file=a.docx')).json()).size, 2);
  const got = await ask('/file?file=a.docx');
  assert.deepEqual([got.status, got.headers.get('content-type'), [...new Uint8Array(await got.arrayBuffer())]], [200, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', [80, 75]]);
  assert.deepEqual(await (await ask('/json?file=a.docx&skip=run')).json(), { type: 'document', skip: 'run' });
  assert.equal(await (await ask('/html?file=a.docx')).text(), 'html of a.docx');
  const bin = await ask('/binary?file=a.docx&path=/body/picture[1]');
  assert.deepEqual([bin.headers.get('content-type'), E.bin], ['image/png', '/body/picture[1]']);
  assert.deepEqual(await (await ask('/run', { method: 'POST', body: JSON.stringify({ argv: ['add', 'a.docx', '/body'] }) })).json(), { code: 0, output: 'ok' });
  assert.equal(await (await fetch('https://embed.example/writer/WordEditor.dc.html')).text(), 'static https://embed.example/writer/WordEditor.dc.html', 'the page\'s own files load as before');
  assert.equal(await (await fetch('https://other.example/run')).text(), 'static https://other.example/run', 'another origin\'s /run is not the engine');
});

test('saves: a PUT writes, from= moves or copies (keep=1), a stale ifMtime is a 409 with the server\'s error body, DELETE removes', async () => {
  const w = await (await ask('/file?file=b.md', { method: 'PUT', body: '# Hi' })).json();
  assert.equal(new TextDecoder().decode(E.files.get('b.md')), '# Hi');
  const stale = await ask('/file?file=b.md&ifMtime=1', { method: 'PUT', body: 'x' });
  assert.deepEqual([stale.status, (await stale.json()).error.code], [409, 'CONFLICT']);
  assert.equal((await ask('/file?file=b.md&ifMtime=' + w.mtime, { method: 'PUT', body: 'y' })).status, 200);
  await ask('/file?file=c.md&from=b.md&keep=1', { method: 'PUT', body: new Uint8Array(0) });
  assert.deepEqual([E.files.has('b.md'), E.files.has('c.md')], [true, true]);
  await ask('/file?file=d.md&from=c.md', { method: 'PUT', body: new Uint8Array(0) });
  assert.deepEqual([E.files.has('c.md'), E.files.has('d.md')], [false, true]);
  assert.equal((await ask('/file?file=d.md', { method: 'DELETE' })).status, 200);
  const missing = await ask('/stat?file=nope.docx');
  assert.deepEqual([missing.status, (await missing.json()).error.message], [404, 'nope.docx not found']);
});

test('a change reaches the editor\'s /events for that file and the site, which hears who made it', async () => {
  const seen = [], heard = [];
  const es = new EventSource('/events?file=a.docx');
  assert.equal(es.native, undefined, 'the engine\'s own stream, not the network\'s');
  es.addEventListener('change', e => seen.push(JSON.parse(e.data).file));
  const stop = S.engine.onChange((files, from) => heard.push([files, from]));
  await ask('/run', { method: 'POST', body: JSON.stringify({ argv: ['add', 'a.docx', '/body'] }) });
  await S.engine.run(['add', 'a.docx', '/body']);
  assert.deepEqual(seen, ['a.docx', 'a.docx']);
  assert.deepEqual(heard, [[['a.docx'], 'editor'], [['a.docx'], 'api']]);
  es.close(); await S.engine.run(['add', 'a.docx', '/body']);
  assert.equal(seen.length, 2, 'closed: no more');
  stop();
  assert.equal(new EventSource('https://elsewhere.example/events?file=a.docx').native, true);
});

test('pictures and downloads get blob: URLs, made again only when the file changed', async () => {
  const url = globalThis.__writerUrl;
  const a = url('binary', 'a.docx', '/body/picture[1]'), b = url('binary', 'a.docx', '/body/picture[1]');
  assert.equal(a, b, 'the same file version: the same URL');
  const f1 = url('file', 'a.docx');
  await S.engine.run(['add', 'a.docx', '/body']);
  const f2 = url('file', 'a.docx');
  assert.notEqual(f1, f2, 'a new version, a new URL');
  assert.equal(url('file', 'missing.md'), '', 'nothing to show');
});

test('the assistant: /files says whether the site set one up, /chat goes to it, /ai says the site manages it', async () => {
  let body = null;
  S.setChat(async b => { body = b; return new Response('event: done\ndata: {}\n\n', { headers: { 'Content-Type': 'text/event-stream' } }); }, 'site-model');
  const list = await (await ask('/files')).json();
  assert.deepEqual([list.chat, list.model], [true, 'site-model']);
  const r = await ask('/chat', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ file: 'a.docx', messages: [{ role: 'user', content: 'hi' }] }) });
  assert.equal(await r.text(), 'event: done\ndata: {}\n\n');
  assert.equal(body.file, 'a.docx');
  assert.deepEqual(await (await ask('/ai')).json(), { provider: '', baseUrl: '', model: 'site-model', hasKey: false, source: 'host' });
  assert.equal((await ask('/ai', { method: 'PUT', body: '{}' })).status, 403);
  S.setChat(null);
  assert.equal((await ask('/chat', { method: 'POST', body: '{}' })).status, 503);
  assert.deepEqual([await S.engine.system('a.docx'), (await S.engine.tools()).map(t => t.name)], ['SYSTEM a.docx true', ['writer', 'batch', 'plan']]);
});

test('while the site\'s first documents go in, the editors wait for them; the commands that put them there do not', async () => {
  let open; S.hold(new Promise(r => { open = r; }));
  let listed = false;
  const list = ask('/files').then(r => { listed = true; return r.json(); });
  assert.equal((await S.engine.run(['create', 'blank.docx'], { seeding: true })).code, 0, 'no deadlock: the gate waits for this very command');
  await new Promise(r => setTimeout(r, 10));
  assert.equal(listed, false, 'the editors see the workspace only once it is ready');
  open();
  assert.ok((await list).files.some(f => f.path === 'blank.docx'));
});

test('AI 自动补全 in the page: /complete goes to the site\'s completer with where it is typed; none set up is a 503', async () => {
  let got = null;
  S.setChat(async () => new Response(''), 'm', async (b, signal) => { got = [b, !!signal]; return ' rest'; });
  const ok = await ask('/complete', { method: 'POST', body: JSON.stringify({ before: 'Dear', after: '', hint: 'a spreadsheet cell' }), signal: new AbortController().signal });
  assert.deepEqual([ok.status, await ok.json(), got], [200, { text: ' rest' }, [{ before: 'Dear', after: '', hint: 'a spreadsheet cell' }, true]]);
  assert.equal((await ask('/ai/models', { method: 'POST', body: '{}' })).status, 403, 'the site manages the models');
  S.setChat(async () => new Response(''), 'm');
  assert.equal((await ask('/complete', { method: 'POST', body: '{}' })).status, 503, 'an assistant without a completer');
  S.setChat(null, '', async () => 'x');
  assert.equal((await ask('/complete', { method: 'POST', body: '{}' })).status, 503, 'no assistant: no completer either');
});
