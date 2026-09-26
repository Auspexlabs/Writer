// ui/embed/server.js — the writer engine inside the page. An editor embedded in another site (docs/embed.md) has no
// `writer serve` behind it: the engine runs here, compiled to WebAssembly (src/Writer.Browser), with the documents in the
// page's memory. install() answers what the editors ask the server (engine.js: /run, /files, /file, /stat, /json, /html,
// /binary, /events; the shells' own /files) from that engine, so the editors do not change. Images and downloads, which the
// page loads without fetch, get blob: URLs instead (engine.js asks globalThis.__writerUrl for them).

const ROUTE = /^\/(run|files|file|stat|binary|json|html|outline|text|chat|ai|ai\/test)$/;
const TYPES = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf: 'application/pdf', md: 'text/markdown; charset=utf-8', txt: 'text/plain; charset=utf-8',
  csv: 'text/csv; charset=utf-8', mm: 'application/xml', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' };
const typeOf = path => TYPES[String(path).split('.').pop().toLowerCase()] || 'application/octet-stream';

let E = null; // the engine's exports (Writer.Browser.Engine)
let booted = null, gate = Promise.resolve();
const watchers = new Map(); // path → the /events streams the editors opened for it
const hooks = new Set(); // (paths, from) → void, for the embedding site: from is 'editor' (the user) or 'api' (the site, its AI)

/** Starts the engine (once): base is the folder of the published engine's _framework/. */
export function boot(base = new URL('../_framework/', import.meta.url).href) {
  return booted || (booted = (async () => {
    const { dotnet } = await import(base + 'dotnet.js');
    const rt = await dotnet.create();
    E = (await rt.getAssemblyExports('Writer.Browser')).Writer.Browser.Engine;
    await rt.runMain();
    return E;
  })());
}

/** Uses an engine that is already loaded (a page that loads it itself; the tests' stand-in) instead of the published one. */
export function useEngine(exports) { E = exports; booted = Promise.resolve(E); return E; }

/** Requests wait for this too: the site's first documents are in place before the editors list the workspace. */
export function hold(p) { gate = Promise.all([gate, p]); }

/** The engine's error, from the exception it threw: { status, error: { code, message, hint } }. */
function failure(e) {
  try { const r = JSON.parse(String(e && e.message || e)); if (r && r.error) return r; } catch (x) { }
  return { status: 500, error: { code: 'INTERNAL', message: String(e && e.message || e), hint: '' } };
}
const json = (status, text) => new Response(text, { status, headers: { 'Content-Type': 'application/json; charset=utf-8' } });
const fail = (status, code, message, hint = '') => json(status, JSON.stringify({ error: { code, message, hint } }));

/** Tells the editors and the site which files changed (the engine lists what it wrote since the last time). */
function announce(from) {
  const paths = JSON.parse(E.Drain());
  if (!paths.length) return paths;
  for (const p of paths) for (const es of watchers.get(p) || []) es.dispatchEvent(Object.assign(new Event('change'), { data: JSON.stringify({ file: p }) }));
  for (const h of hooks) { try { h(paths, from); } catch (e) { console.error(e); } }
  return paths;
}

async function bodyBytes(init) {
  const b = init && init.body;
  if (b == null) return new Uint8Array(0);
  if (b instanceof Uint8Array) return b;
  return new Uint8Array(await new Response(b).arrayBuffer());
}

/** One request the editors made to the server, answered by the engine. */
export async function answer(url, init = {}) {
  await boot(); await gate;
  const q = url.searchParams, file = q.get('file'), method = String(init.method || 'GET').toUpperCase();
  try {
    switch (url.pathname) {
      case '/run': { const out = E.Run(typeof init.body === 'string' ? init.body : new TextDecoder().decode(await bodyBytes(init))); announce('editor'); return json(200, out); }
      case '/files': {
        if (q.get('drafts') === '1') return json(200, JSON.stringify({ files: [] }));
        const list = JSON.parse(E.Files()); list.chat = !!chat; if (chat) list.model = chatModel; // the assistant is the site's
        return json(200, JSON.stringify(list));
      }
      case '/stat': return json(200, E.Stat(file));
      case '/file':
        if (method === 'GET') return new Response(E.Read(file), { status: 200, headers: { 'Content-Type': typeOf(file) } });
        if (method === 'DELETE') { const out = E.Delete(file); announce('editor'); return json(200, out); }
        if (method === 'PUT') {
          const out = q.get('from') ? E.Move(q.get('from'), file, q.get('keep') === '1') : E.Write(file, await bodyBytes(init), q.get('ifMtime'));
          announce('editor'); return json(200, out);
        }
        break;
      case '/binary': { const bytes = E.Binary(file, q.get('path') || ''); return new Response(bytes, { status: 200, headers: { 'Content-Type': E.BinaryType() } }); }
      case '/json': return json(200, E.Json(file, q.get('skip') || ''));
      case '/html': return new Response(E.View(file, 'html'), { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' } });
      case '/outline': case '/text': return new Response(E.View(file, url.pathname.slice(1)), { status: 200, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
      case '/chat': return chat ? chat(JSON.parse(typeof init.body === 'string' ? init.body : new TextDecoder().decode(await bodyBytes(init))), init.signal)
        : fail(503, 'NO_MODEL', 'No model is set up', 'The site embedding Writer connects its assistant (docs/embed.md).');
      case '/ai': return method === 'GET' ? json(200, JSON.stringify({ provider: '', baseUrl: '', model: chatModel || '', hasKey: false, source: chat ? 'host' : 'none' }))
        : fail(403, 'HOST_MANAGED', 'The site embedding Writer sets up its assistant');
      case '/ai/test': return json(200, JSON.stringify({ ok: !!chat, error: chat ? '' : 'No model is set up' }));
    }
    return fail(404, 'NOT_FOUND', 'No such endpoint');
  } catch (e) {
    const r = failure(e); if (E) announce('editor');
    return json(r.status || 500, JSON.stringify({ error: r.error }));
  }
}

// ---- the assistant: the embedding site supplies it (setChat), see ai.js ----
let chat = null, chatModel = '';
/** handler(body, signal) → Response (server-sent events, as POST /chat answers); model: the name the settings show. */
export function setChat(handler, model) { chat = handler || null; chatModel = model || ''; }

// ---- for the embedding site (host.js): the same engine, called directly ----
export const engine = {
  /** One writer command, as the CLI takes it: { code, output } or { code, error }. The editor showing the file reloads it.
   *  seeding: a command that puts the first documents in place (host.js), which the gate is waiting for. */
  async run(command, { seeding = false } = {}) { await boot(); if (!seeding) await gate; const body = JSON.stringify(Array.isArray(command) ? { argv: command } : { command: String(command) }); const r = JSON.parse(E.Run(body)); announce('api'); return r; },
  async put(path, bytes) { await boot(); const r = JSON.parse(E.Write(path, bytes, null)); announce('api'); return r; },
  async get(path) { await boot(); return E.Read(path); },
  async list() { await boot(); return JSON.parse(E.Files()).files; },
  async stat(path) { await boot(); return JSON.parse(E.Stat(path)); },
  /** A document's tree ({type, path, props, children}), as `writer get <file> /` would give it in full. */
  async tree(path) { await boot(); return JSON.parse(E.Json(path, '')); },
  /** Called with (paths, from) whenever files change; returns a function that stops it. */
  onChange(fn) { hooks.add(fn); return () => hooks.delete(fn); },
  // the assistant's side in the engine (src/Writer.Cli/Assistant.cs): the desktop assistant's system prompt for a file, its
  // editing tools, and running one tool call — for the assistant in the editor (ai.js) and for the site's own agent
  async system(file, outline = true) { await boot(); await gate; return E.ChatSystem(file || '', !!outline); },
  async tools() { await boot(); return JSON.parse(E.ChatTools()); },
  /** { display, code, output, wrote }; from: who asked ('assistant' in the editor, 'api' for the site). */
  async callTool(name, input, from = 'api') { await boot(); await gate; const r = JSON.parse(E.ChatTool(String(name), JSON.stringify(input || {}))); announce(from); return r; }
};

// ---- blob: URLs for what the page loads by itself (an <img>, a download link) ----
const urls = new Map(); // kind|path|node → { v, url }
function blobUrl(kind, path, node) {
  if (!E) return kind === 'file' ? '/file?file=' + encodeURIComponent(path) : '/binary?file=' + encodeURIComponent(path) + '&path=' + encodeURIComponent(node);
  const key = kind + '|' + path + '|' + (node || '');
  let v; try { v = JSON.parse(E.Stat(path)).mtime; } catch (e) { return ''; }
  const hit = urls.get(key); if (hit && hit.v === v) return hit.url;
  let bytes, type;
  try { if (kind === 'file') { bytes = E.Read(path); type = typeOf(path); } else { bytes = E.Binary(path, node); type = E.BinaryType(); } } catch (e) { return ''; }
  if (hit) URL.revokeObjectURL(hit.url);
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  urls.set(key, { v, url });
  return url;
}

/** Answers the editors' server requests from the engine in this page, from now on. */
export function install() {
  const native = window.fetch.bind(window);
  window.fetch = (input, init) => {
    const href = typeof input === 'string' ? input : input instanceof URL ? input.href : input && input.url;
    const url = new URL(href, location.href);
    if (url.origin === location.origin && ROUTE.test(url.pathname)) {
      if (input && typeof input === 'object' && !(input instanceof URL) && !init) init = { method: input.method, body: input.body, signal: input.signal };
      return answer(url, init || {});
    }
    return native(input, init);
  };
  const NativeES = window.EventSource;
  class LocalEvents extends EventTarget { // GET /events?file=: 'change' whenever the file is written, by the editor or the site
    constructor(href) {
      super(); this.url = href; this.readyState = 1; this.onmessage = null; this.onerror = null; this.onopen = null;
      this.file = new URL(href, location.href).searchParams.get('file') || '';
      const set = watchers.get(this.file) || new Set(); set.add(this); watchers.set(this.file, set);
    }
    close() { this.readyState = 2; const set = watchers.get(this.file); if (set) set.delete(this); }
  }
  window.EventSource = function (href, opts) {
    const url = new URL(href, location.href);
    return url.origin === location.origin && url.pathname === '/events' ? new LocalEvents(href) : new NativeES(href, opts);
  };
  globalThis.__writerUrl = blobUrl;
}
