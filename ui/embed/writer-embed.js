/*! Writer embed — Writer's editors (Word, Excel, PowerPoint, Markdown, mind maps, PDF) or the whole app, in a page.
 *
 *   <div id="doc" style="height:640px"></div>
 *   <script src="https://YOUR-HOST/writer/embed/writer-embed.js"></script>
 *   <script type="module">
 *     const writer = Writer.embed('#doc', { mode: 'docx', file: '/files/report.docx' });
 *     await writer.ready;
 *     await writer.run('add report.docx /body --type paragraph --prop text="Hello"');
 *     const blob = await writer.get('report.docx');
 *   </script>
 *
 * Everything runs in the visitor's browser: the engine is WebAssembly, the documents stay in the page. docs/embed.md has the
 * options, the API and the assistant. A classic script: it defines window.Writer. */
(function (global) {
  'use strict';
  var me = document.currentScript;
  var BASE = me && me.src ? new URL('../', me.src).href : new URL('./', location.href).href; // the Writer folder: …/writer/
  var MODES = { app: 1, docx: 1, xlsx: 1, pptx: 1, md: 1, mm: 1, pdf: 1 };
  var TYPES = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf: 'application/pdf', md: 'text/markdown', mm: 'application/xml', csv: 'text/csv', txt: 'text/plain' };
  var count = 0;

  /** A document given to the widget, as { name, data: ArrayBuffer }: a URL (a string or URL, fetched by this page), a File or
   *  Blob, bytes, { name, url } (a URL that does not end in the name), or { name, data } with text, bytes or a Blob as data. */
  function load(f, fallbackName) {
    if (f == null) return Promise.resolve(null);
    if (typeof f === 'string' || f instanceof URL) {
      var url = new URL(String(f), location.href);
      return fetch(url.href, { credentials: 'same-origin' }).then(function (r) {
        if (!r.ok) throw new Error('Writer: ' + r.status + ' fetching ' + url.href);
        return r.arrayBuffer();
      }).then(function (data) { return { name: decodeURIComponent(url.pathname.split('/').pop()) || fallbackName, data: data }; });
    }
    if (typeof Blob !== 'undefined' && f instanceof Blob) return f.arrayBuffer().then(function (data) { return { name: f.name || fallbackName, data: data }; });
    if (f instanceof ArrayBuffer) return Promise.resolve({ name: fallbackName, data: f.slice(0) });
    if (ArrayBuffer.isView(f)) return Promise.resolve({ name: fallbackName, data: f.buffer.slice(f.byteOffset, f.byteOffset + f.byteLength) });
    if (typeof f === 'object' && (f.url != null || 'data' in f)) {
      var named = function (x) { return x && { name: f.name || x.name, data: x.data }; };
      if (f.url != null) return load(f.url instanceof URL ? f.url : String(f.url), fallbackName).then(named);
      return (typeof f.data === 'string' ? Promise.resolve({ name: fallbackName, data: new TextEncoder().encode(f.data).buffer }) : load(f.data, fallbackName)).then(named);
    }
    return Promise.reject(new Error('Writer: a file is a URL, a File or Blob, bytes, { name, url } or { name, data }'));
  }
  var bytesOf = function (data) {
    if (typeof data === 'string') return Promise.resolve(new TextEncoder().encode(data).buffer);
    return load(data, 'x').then(function (x) { return x.data; });
  };

  /** Writer.embed(target, options): the widget. target: an element or a selector; the iframe fills it (or options.height). */
  function Widget(target, options) {
    var self = this, o = options || {};
    var host = typeof target === 'string' ? document.querySelector(target) : target;
    if (!host) throw new Error('Writer: no element ' + target);
    var mode = MODES[o.mode] ? o.mode : 'app';
    this.id = 'w' + (++count) + Math.random().toString(36).slice(2, 8);
    this.options = o;
    var base = o.base ? new URL(o.base, location.href).href : BASE;
    var q = new URLSearchParams({ id: this.id, mode: mode, theme: o.theme || 'light' });
    if (o.lang) q.set('lang', o.lang);
    if (o.ai) { q.set('ai', '1'); if (o.aiModel) q.set('model', o.aiModel); }
    if (o.blankName) q.set('blank', o.blankName);
    var src = new URL('embed.dc.html?' + q.toString(), base);
    this.origin = src.origin;
    var frame = this.frame = document.createElement('iframe');
    frame.src = src.href;
    frame.title = o.title || 'Writer';
    frame.setAttribute('allow', 'clipboard-read; clipboard-write; fullscreen');
    frame.style.cssText = 'display:block;width:' + css(o.width || '100%') + ';height:' + css(o.height || '100%') + ';border:0;border-radius:' + css(o.radius == null ? 12 : o.radius) + ';background:transparent;color-scheme:normal';
    host.appendChild(frame);

    this._calls = {}; this._seq = 0; this._on = {}; this._chats = {};
    this.ready = new Promise(function (resolve, reject) { self._resolveReady = resolve; self._rejectReady = reject; });
    this._listener = function (e) { self._message(e); };
    global.addEventListener('message', this._listener);
    // the documents are read now, while the engine loads in the frame
    var list = [].concat(o.files || [], o.file != null ? [o.file] : []);
    var fallback = 'document.' + (mode === 'app' ? 'docx' : mode);
    this._files = Promise.all(list.map(function (f, i) { return load(f, i ? i + '-' + fallback : fallback); }));
    this._files.catch(function () { }); // a file that cannot be read fails ready, on hello
  }
  function css(v) { return typeof v === 'number' ? v + 'px' : String(v); }

  Widget.prototype._post = function (msg, transfer) {
    if (!this.frame.contentWindow) return;
    msg.writer = 1; msg.id = this.id;
    this.frame.contentWindow.postMessage(msg, this.origin, transfer || []);
  };

  Widget.prototype._message = function (e) {
    var m = e.data, self = this;
    if (!m || m.writer !== 1 || m.id !== this.id || e.source !== this.frame.contentWindow || e.origin !== this.origin) return;
    if (m.type === 'hello') {
      this._files.then(function (files) {
        files = files.filter(Boolean);
        var open = self.options.open || (files[0] && files[0].name) || '';
        self._post({ type: 'init', files: files, open: open }, files.map(function (f) { return f.data; }));
      }, function (err) { self._rejectReady(err); });
    } else if (m.type === 'ready') {
      if (m.error) this._rejectReady(new Error(m.error)); else this._resolveReady({ file: m.file });
      this._emit('ready', { file: m.file });
    } else if (m.type === 'result') {
      var c = this._calls[m.seq]; if (!c) return; delete this._calls[m.seq];
      if (m.ok) c.resolve(m.value); else { var err = new Error(m.error); err.code = m.code || ''; err.hint = m.hint || ''; c.reject(err); }
    } else if (m.type === 'event') {
      this._emit(m.name, m.data);
    } else if (m.type === 'chat') {
      this._chat(m.seq, m.request);
    } else if (m.type === 'chat-abort') {
      var ac = this._chats[m.seq]; if (ac) { ac.abort(); delete this._chats[m.seq]; }
    }
  };

  /** One model call for the assistant in the editor, answered by options.ai. */
  Widget.prototype._chat = function (seq, request) {
    var self = this, ac = new AbortController();
    this._chats[seq] = ac;
    var send = function (name, data) { if (self._chats[seq]) self._post({ type: 'chat-event', seq: seq, name: name, data: data }); };
    Promise.resolve().then(function () {
      if (typeof self.options.ai !== 'function') throw new Error('No assistant: pass options.ai');
      return self.options.ai(request, { onDelta: function (text) { send('delta', String(text)); }, signal: ac.signal });
    }).then(function (reply) {
      reply = reply || {};
      send('done', { text: String(reply.text || ''), toolCalls: (reply.toolCalls || []).map(function (c) { return { id: String(c.id || ''), name: String(c.name), input: c.input || {} }; }) });
      delete self._chats[seq];
    }, function (err) { send('error', String(err && err.message || err)); delete self._chats[seq]; });
  };

  Widget.prototype._call = function (method, args, transfer) {
    var self = this;
    if (this._gone) return Promise.reject(this._gone);
    return this.ready.then(function () {
      if (self._gone) throw self._gone;
      return new Promise(function (resolve, reject) {
        var seq = ++self._seq;
        self._calls[seq] = { resolve: resolve, reject: reject };
        self._post({ type: 'call', seq: seq, method: method, args: args || [] }, transfer);
      });
    });
  };
  Widget.prototype._emit = function (name, data) { (this._on[name] || []).slice().forEach(function (fn) { try { fn(data); } catch (e) { console.error(e); } }); };

  /** Runs one writer command (the CLI's syntax, as a string or an argv array): { code, output } or { code, error }. The editor
   *  showing the file reloads it. */
  Widget.prototype.run = function (command) { return this._call('run', [command]); };
  /** Adds or replaces a document: data is a URL, a File or Blob, bytes or text. */
  Widget.prototype.put = function (name, data) { var self = this; return bytesOf(data).then(function (buf) { return self._call('put', [name, buf], [buf]); }); };
  /** A document as it is now (the editor saves as the user types), as a Blob. */
  Widget.prototype.get = function (name) { return this._call('get', [name]).then(function (buf) { return new Blob([buf], { type: TYPES[String(name).split('.').pop().toLowerCase()] || 'application/octet-stream' }); }); };
  /** The documents: [{ path, name, format, size, modified }]. */
  Widget.prototype.list = function () { return this._call('list'); };
  /** Shows a document of the widget (a name from list()). */
  Widget.prototype.open = function (name) { return this._call('open', [name]); };
  /** The document shown now ('' on the 新建 page). */
  Widget.prototype.current = function () { return this._call('current'); };
  /** A document's tree: { type, path, props, children }. */
  Widget.prototype.tree = function (name) { return this._call('tree', [name]); };
  /** For an agent of the site's own: the system prompt Writer's assistant works with on a file (rules, the command reference,
   *  the outline), its tools ([{ name, description, input_schema }]: writer, batch, plan), and one tool call
   *  ({ display, code, output, wrote }). */
  Widget.prototype.system = function (file) { return this._call('system', [file]); };
  Widget.prototype.tools = function () { return this._call('tools'); };
  Widget.prototype.callTool = function (name, input) { return this._call('callTool', [name, input]); };
  /** light, dark or auto. */
  Widget.prototype.setTheme = function (theme) { return this._call('theme', [theme]); };
  /** Events: ready { file }, open { file }, change { files, from } (from: 'editor' for the user, 'assistant', 'api' for calls). */
  Widget.prototype.on = function (name, fn) { (this._on[name] = this._on[name] || []).push(fn); return this; };
  Widget.prototype.off = function (name, fn) { this._on[name] = (this._on[name] || []).filter(function (f) { return f !== fn; }); return this; };
  /** Removes the widget; calls still waiting fail. */
  Widget.prototype.destroy = function () {
    var gone = this._gone = new Error('Writer: the widget was removed');
    global.removeEventListener('message', this._listener);
    Object.keys(this._chats).forEach(function (k) { this._chats[k].abort(); }, this);
    Object.keys(this._calls).forEach(function (k) { this._calls[k].reject(gone); }, this);
    this._chats = {}; this._calls = {};
    this.ready.catch(function () { }); this._rejectReady(gone); // no-op once ready
    if (this.frame.parentNode) this.frame.parentNode.removeChild(this.frame);
  };

  // ---------- the assistant: options.ai(request, { onDelta, signal }) → { text, toolCalls: [{ id, name, input }] } ----------
  // request: { system, messages, tools }. messages is provider-neutral (ui/embed/ai.js): { role: 'user', content } |
  // { role: 'assistant', content, toolCalls } | { role: 'tool', toolCallId, name, content, isError }. The adapters below speak
  // the Anthropic Messages API and OpenAI-compatible /chat/completions; point url at the site's own endpoint, which adds the key.

  /** The lines of a server-sent event stream, as they arrive. */
  function readEvents(res, onData) {
    var reader = res.body.getReader(), dec = new TextDecoder(), buf = '';
    function pump() {
      return reader.read().then(function (r) {
        if (r.done) { if (buf.trim()) buf.split('\n').forEach(line); return; }
        buf += dec.decode(r.value, { stream: true });
        var i; while ((i = buf.indexOf('\n')) >= 0) { line(buf.slice(0, i)); buf = buf.slice(i + 1); }
        return pump();
      });
    }
    function line(l) { // a line that is not JSON is skipped; an error the stream reports (onData throws) ends the call
      l = l.replace(/\r$/, ''); if (l.indexOf('data:') !== 0) return;
      var d = l.slice(5).trim(), j; if (!d || d === '[DONE]') return;
      try { j = JSON.parse(d); } catch (e) { return; }
      onData(j);
    }
    return pump();
  }
  function failed(res) { return res.text().then(function (t) { var msg = t; try { var j = JSON.parse(t); msg = (j.error && (j.error.message || j.error)) || j.message || t; } catch (e) { } throw new Error(res.status + ' ' + String(msg).slice(0, 300)); }); }

  function anthropic(o) {
    o = o || {};
    return function (request, ctx) {
      var messages = [];
      request.messages.forEach(function (m) {
        if (m.role === 'user') messages.push({ role: 'user', content: m.content });
        else if (m.role === 'assistant') {
          var content = [];
          if (m.content) content.push({ type: 'text', text: m.content });
          (m.toolCalls || []).forEach(function (c) { content.push({ type: 'tool_use', id: c.id, name: c.name, input: c.input || {} }); });
          if (content.length) messages.push({ role: 'assistant', content: content });
        } else if (m.role === 'tool') {
          var block = { type: 'tool_result', tool_use_id: m.toolCallId, content: m.content, is_error: !!m.isError }, last = messages[messages.length - 1];
          if (last && last.role === 'user' && Array.isArray(last.content) && last.content[0] && last.content[0].type === 'tool_result') last.content.push(block);
          else messages.push({ role: 'user', content: [block] });
        }
      });
      var url = o.url || 'https://api.anthropic.com/v1/messages', headers = { 'content-type': 'application/json', 'anthropic-version': '2023-06-01' };
      // straight from the page to Anthropic (a key the visitors can see: for trying it out; a site puts its own endpoint in url)
      if (new URL(url, location.href).hostname === 'api.anthropic.com') headers['anthropic-dangerous-direct-browser-access'] = 'true';
      headers = Object.assign(headers, o.headers || {});
      return fetch(url, {
        method: 'POST', headers: headers, signal: ctx.signal, credentials: o.credentials || 'same-origin',
        body: JSON.stringify({ model: o.model || 'claude-sonnet-5', max_tokens: o.maxTokens || 8192, system: request.system, tools: request.tools, messages: messages, stream: true })
      }).then(function (res) {
        if (!res.ok) return failed(res);
        var blocks = {}, text = '';
        return readEvents(res, function (e) {
          if (e.type === 'content_block_start') { var b = e.content_block || {}; blocks[e.index] = { type: b.type, id: b.id, name: b.name, text: b.text || '', json: '' }; if (b.text) { text += b.text; ctx.onDelta(b.text); } }
          else if (e.type === 'content_block_delta' && blocks[e.index]) { var d = e.delta || {}; if (d.text) { blocks[e.index].text += d.text; text += d.text; ctx.onDelta(d.text); } if (d.partial_json) blocks[e.index].json += d.partial_json; }
          else if (e.type === 'error') throw new Error((e.error && e.error.message) || 'model error');
        }).then(function () {
          var calls = Object.keys(blocks).sort(function (a, b) { return a - b; }).map(function (k) { return blocks[k]; }).filter(function (b) { return b.type === 'tool_use'; })
            .map(function (b) { var input = {}; try { input = b.json ? JSON.parse(b.json) : {}; } catch (e) { } return { id: b.id, name: b.name, input: input }; });
          return { text: text, toolCalls: calls };
        });
      });
    };
  }

  function openai(o) {
    o = o || {};
    return function (request, ctx) {
      var messages = [{ role: 'system', content: request.system }];
      request.messages.forEach(function (m) {
        if (m.role === 'user') messages.push({ role: 'user', content: m.content });
        else if (m.role === 'assistant') {
          var x = { role: 'assistant', content: m.content || null };
          if (m.toolCalls && m.toolCalls.length) x.tool_calls = m.toolCalls.map(function (c) { return { id: c.id, type: 'function', function: { name: c.name, arguments: JSON.stringify(c.input || {}) } }; });
          messages.push(x);
        } else if (m.role === 'tool') messages.push({ role: 'tool', tool_call_id: m.toolCallId, content: m.content });
      });
      var tools = request.tools.map(function (t) { return { type: 'function', function: { name: t.name, description: t.description, parameters: t.input_schema } }; });
      return fetch(o.url || 'https://api.openai.com/v1/chat/completions', {
        method: 'POST', headers: Object.assign({ 'content-type': 'application/json' }, o.headers || {}), signal: ctx.signal, credentials: o.credentials || 'same-origin',
        body: JSON.stringify({ model: o.model || 'gpt-4o', messages: messages, tools: tools, stream: true })
      }).then(function (res) {
        if (!res.ok) return failed(res);
        var text = '', calls = [];
        return readEvents(res, function (e) {
          var d = e.choices && e.choices[0] && e.choices[0].delta; if (!d) return;
          if (d.content) { text += d.content; ctx.onDelta(d.content); }
          (d.tool_calls || []).forEach(function (t) {
            var last = calls[calls.length - 1], i = t.index != null ? t.index // some compatible servers leave out index: a new id is a new call
              : !last ? 0 : t.id && last.id && t.id !== last.id ? calls.length : calls.length - 1;
            var c = calls[i] = calls[i] || { id: '', name: '', args: '' };
            if (t.id) c.id = t.id; if (t.function && t.function.name) c.name += t.function.name; if (t.function && t.function.arguments) c.args += t.function.arguments;
          });
        }).then(function () {
          return { text: text, toolCalls: calls.filter(Boolean).map(function (c) { var input = {}; try { input = c.args ? JSON.parse(c.args) : {}; } catch (e) { } return { id: c.id, name: c.name, input: input }; }) };
        });
      });
    };
  }

  global.Writer = {
    version: '0.1.5',
    embed: function (target, options) { return new Widget(target, options); },
    Widget: Widget,
    ai: { anthropic: anthropic, openai: openai }
  };
})(window);
