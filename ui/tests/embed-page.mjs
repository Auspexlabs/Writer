// The embed page (ui/embed.dc.html) in an iframe, for the tests of ui/embed/host.js: its window, the site around it (parent),
// the options the page reads from its URL, an engine stand-in and the shell played by the test. host.js starts on import, so
// each test file loads one page: page(options) before anything else.

/** A stand-in for the WebAssembly engine's calls (Writer.Browser.Engine), over a Map of files. */
export function fakeEngine() {
  const files = new Map(), changed = [];
  const bump = p => { if (!changed.includes(p)) changed.push(p); };
  return {
    files, calls: [],
    Run(body) { const { argv, command } = JSON.parse(body); const a = argv || command.split(' '); this.calls.push(a); if (a[0] === 'create') { files.set(a[1], new Uint8Array([1])); bump(a[1]); } if (a[0] === 'add') bump(a[1]); return JSON.stringify({ code: 0, output: 'ran ' + a.join(' ') }); },
    Files() { return JSON.stringify({ files: [...files.keys()].map(p => ({ path: p })) }); },
    Stat(p) { return JSON.stringify({ path: p, mtime: 1, size: files.get(p).length }); },
    Read(p) { if (!files.has(p)) throw new Error(JSON.stringify({ status: 404, error: { code: 'FILE_NOT_FOUND', message: p + ' not found' } })); return files.get(p); },
    Write(p, bytes) { files.set(p, bytes); bump(p); return JSON.stringify({ path: p, size: bytes.length }); },
    Json(p) { return JSON.stringify({ type: 'document', path: '/', file: p }); },
    Drain() { const out = JSON.stringify(changed); changed.length = 0; return out; },
    ChatSystem(p) { return 'SYSTEM for ' + p; },
    ChatTools() { return JSON.stringify([{ name: 'writer', description: 'd', input_schema: {} }]); },
    ChatTool(name, input) { this.calls.push([name, JSON.parse(input)]); return JSON.stringify({ display: 'writer ' + JSON.parse(input).command, code: 0, output: 'done', wrote: true }); }
  };
}

export const SITE = 'https://site.example';

/** Loads the page with these widget options (window.__WRITER_EMBED__): the page's hello is the first thing posted. */
export async function page(options) {
  const bus = new EventTarget(), posted = [];
  globalThis.window = globalThis;
  globalThis.addEventListener = bus.addEventListener.bind(bus);
  globalThis.removeEventListener = bus.removeEventListener.bind(bus);
  globalThis.dispatchEvent = bus.dispatchEvent.bind(bus);
  const parent = { postMessage(msg, origin, transfer) { posted.push({ msg: structuredClone(msg), origin, transfer }); } };
  globalThis.parent = parent;
  globalThis.location = new URL('https://cdn.example/writer/embed.dc.html?id=' + options.id);
  globalThis.fetch = async u => new Response('static ' + u);
  globalThis.EventSource = class { };
  URL.createObjectURL = () => 'blob:x';
  URL.revokeObjectURL = () => { };
  window.__WRITER_EMBED__ = options;
  const S = await import('../embed/server.js');
  const E = S.useEngine(fakeEngine());
  await import('../embed/host.js');

  /** A message to the page: from the site by default. */
  const say = (data, from = {}) => dispatchEvent(Object.assign(new Event('message'), { data: Object.assign({ writer: 1, id: options.id }, data), origin: from.origin || SITE, source: 'source' in from ? from.source : parent }));
  const until = async (what, ms = 3000) => { for (const end = Date.now() + ms; Date.now() < end; await new Promise(r => setTimeout(r, 5))) { const x = what(); if (x) return x; } throw new Error('timed out'); };
  const find = (type, more = () => true) => posted.find(p => p.msg.type === type && more(p.msg));
  /** The shell (index.dc.html) as the page reports it: open a document, and it becomes the one shown. */
  const embed = window.__WRITER_EMBED__, shown = [];
  const api = { async openPath(name) { shown.push(name); setTimeout(() => embed.shell({ path: name, isDoc: true }, api), 5); } };
  let seq = 0;
  const call = async (method, ...args) => { const n = ++seq; say({ type: 'call', seq: n, method, args }); return until(() => find('result', m => m.seq === n)); };
  return { S, E, posted, parent, say, until, find, embed, api, shown, call };
}
