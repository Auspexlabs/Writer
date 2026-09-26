// ui/embed/host.js — inside the embed page (embed.dc.html), loaded before the editors: starts the engine in the page
// (server.js) and talks to the site that embedded it (writer-embed.js) over postMessage. The protocol, every message
// carrying { writer: 1, id } (id: the widget's, from the URL):
//   page → site  hello                                   the page is up; the site answers with init
//   site → page  init { files: [{ name, data }], open }  the first documents (ArrayBuffers) and the one to show
//   page → site  ready { file }                          the editor is up, showing file ('' on the 新建 page)
//   site → page  call { seq, method, args }              run / put / get / list / open / tree / theme (below)
//   page → site  result { seq, ok, value | error, code, hint }
//   page → site  event { name, data }                    change { files, from }: files written by the editor (from 'editor')
//                                                        or by a call (from 'api'); open { file }: the document shown
// Opened on its own (no site around it) the page starts with an empty workspace, or with ?src= fetched into it.
import { install, boot, hold, engine, setChat } from './server.js';

const q = new URLSearchParams(location.search);
const opts = window.__WRITER_EMBED__ || {};
const id = opts.id || '';
const site = window.parent !== window ? window.parent : null;
let siteOrigin = null; // the embedding site's origin, taken from its init: every later message must come from it, and go to it

install();
boot().catch(e => console.error('Writer: the engine did not start', e));

let seed; const seeded = new Promise(r => { seed = r; });
hold(seeded);
opts.starting = true; // the page shows 正在打开… until the site's documents are open (embed.dc.html)
let shellApi = null, shellInfo = null, firstOpen = null;
const shellWaiters = [];

function post(msg, transfer) { if (site) site.postMessage(Object.assign({ writer: 1, id }, msg), siteOrigin || '*', transfer || []); }
const emit = (name, data) => post({ type: 'event', name, data });

let seeding = true; // the site's own first documents going in are no change to tell it about
engine.onChange((files, from) => { if (!seeding) emit('change', { files, from }); });

/** The embed page reports the shell (the editors) as it changes: onShell(info, api) from index.dc.html. */
opts.shell = (info, api) => {
  const was = shellInfo && shellInfo.path;
  shellInfo = info; shellApi = api;
  while (shellWaiters.length) shellWaiters.shift()();
  if (info.path !== was && info.isDoc) emit('open', { file: info.path });
};
const shell = () => shellApi ? Promise.resolve(shellApi) : new Promise(r => shellWaiters.push(() => r(shellApi)));

/** Shows a document of the workspace in the editor; resolves once it is the current one. */
async function openDoc(name) {
  const api = await shell();
  await api.openPath(name);
  for (let i = 0; i < 200 && !(shellInfo && shellInfo.isDoc && shellInfo.path === name); i++) await new Promise(r => setTimeout(r, 25));
  return !!(shellInfo && shellInfo.path === name);
}

async function start(files, open) {
  try { await begin(files, open); } finally { opts.starting = false; window.dispatchEvent(new Event('writer-settings')); }
}

async function begin(files, open) {
  const given = (files || []).filter(f => f && f.name && f.data).map(f => f.name);
  for (const f of files || []) if (f && f.name && f.data) await engine.put(f.name, new Uint8Array(f.data));
  if (!files || !files.length) {
    // a single editor with nothing given: a blank document of its kind (a PDF is only ever opened)
    const blank = opts.mode && opts.mode !== 'app' && opts.mode !== 'pdf' ? opts.mode : null;
    if (blank) {
      const name = (opts.blankName || (window.$t ? window.$t('未命名') : 'Untitled')) + '.' + blank;
      if ((await engine.run(['create', name], { seeding: true })).code === 0) open = open || name;
    }
  }
  seeding = false; seed();
  firstOpen = open || given[0] || '';
  // the whole app: every document the site gave is a tab, in its order, and the one to show comes to the front
  if ((!opts.mode || opts.mode === 'app') && given.length > 1) for (const name of given) await openDoc(name);
  const shown = !!firstOpen && await openDoc(firstOpen);
  if (!shown && opts.mode && opts.mode !== 'app') opts.empty = true; // one editor with nothing to show says so
  post({ type: 'ready', file: shown ? firstOpen : '' });
}

// ---- calls from the site ----
const methods = {
  /** One writer command (the CLI's, as a string or an argv array): { code, output } or { code, error }. */
  run: command => engine.run(command),
  /** Adds or replaces a document: put(name, ArrayBuffer | Uint8Array | string). */
  put: (name, data) => engine.put(name, typeof data === 'string' ? new TextEncoder().encode(data) : new Uint8Array(data)),
  /** A document's bytes, as an ArrayBuffer (transferred). */
  get: async name => { const b = await engine.get(name); return b.buffer.byteLength === b.byteLength ? b.buffer : b.slice().buffer; },
  list: () => engine.list(),
  tree: name => engine.tree(name),
  open: name => openDoc(name),
  /** The document shown now ('' on the 新建 page). */
  current: () => (shellInfo && shellInfo.isDoc ? shellInfo.path : ''),
  /** For the site's own agent: the system prompt the assistant works with on a file, its tools, and one tool call. */
  system: (file, outline) => engine.system(file, outline !== false),
  tools: () => engine.tools(),
  callTool: (name, input) => engine.callTool(name, input, 'api'),
  theme: theme => { window.__WRITER_PREFS = Object.assign({}, window.__WRITER_PREFS, { theme }); window.dispatchEvent(new Event('writer-settings')); return true; }
};

/** A failed call, for the site: the engine's error ({ status, error: { code, message, hint } }) as its message, code and hint. */
function reason(err) {
  const text = String(err && err.message || err);
  try { const r = JSON.parse(text); if (r && r.error) return { error: r.error.message || text, code: r.error.code || '', hint: r.error.hint || '' }; } catch (e) { }
  return { error: text, code: '', hint: '' };
}

window.addEventListener('message', async e => {
  const m = e.data;
  if (!site || e.source !== site || !m || m.writer !== 1 || (id && m.id !== id)) return;
  if (m.type === 'init') {
    if (siteOrigin) return; // once
    siteOrigin = e.origin === 'null' ? '*' : e.origin;
    try { await start(m.files, m.open); } catch (err) { console.error(err); seeding = false; seed(); post({ type: 'ready', file: '', error: String(err && err.message || err) }); }
    return;
  }
  if (e.origin !== siteOrigin && siteOrigin !== '*') return;
  if (m.type === 'call') {
    const f = methods[m.method];
    try {
      if (!f) throw new Error('Unknown method ' + m.method);
      const value = await f(...(m.args || []));
      post({ type: 'result', seq: m.seq, ok: true, value }, value instanceof ArrayBuffer ? [value] : []);
    } catch (err) { post(Object.assign({ type: 'result', seq: m.seq, ok: false }, reason(err))); }
  }
  if (m.type === 'chat-event' && chatStreams.has(m.seq)) chatStreams.get(m.seq)(m);
});

// ---- the assistant: the site answers chat turns (writer-embed.js: options.ai); ai.js runs the tool loop here ----
const chatStreams = new Map(); let chatSeq = 0;
// registered at once (the editors ask /files whether there is an assistant before anything else); the loop loads on first use
if (opts.ai) setChat(async (body, signal) => (await import('./ai.js')).chatTurn(body, signal, { ask: askSite, engine }), opts.aiModel || '');
/** One model call through the site: request { system, messages, tools } → { text, toolCalls }; onDelta gets the text as it streams. */
function askSite(request, onDelta, signal) {
  return new Promise((resolve, reject) => {
    const seq = ++chatSeq;
    chatStreams.set(seq, m => {
      if (m.name === 'delta') onDelta && onDelta(m.data);
      else if (m.name === 'done') { chatStreams.delete(seq); resolve(m.data); }
      else if (m.name === 'error') { chatStreams.delete(seq); reject(new Error(m.data)); }
    });
    if (signal) signal.addEventListener('abort', () => { chatStreams.delete(seq); post({ type: 'chat-abort', seq }); reject(new DOMException('Aborted', 'AbortError')); });
    post({ type: 'chat', seq, request });
  });
}

// ---- no site around it: the page is Writer in the browser ----
if (!site) (async () => {
  const src = q.get('src');
  if (!src) return start([], '');
  try {
    const name = decodeURIComponent(new URL(src, location.href).pathname.split('/').pop() || 'document');
    const data = await (await fetch(src)).arrayBuffer();
    await start([{ name, data }], name);
  } catch (e) { console.error(e); start([], ''); }
})();
else post({ type: 'hello' });
