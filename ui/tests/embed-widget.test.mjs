// node --test ui/tests/ — the script a site includes (ui/embed/writer-embed.js): the iframe it makes, the documents it hands over,
// calls and events over postMessage (only with its own frame), the assistant's model calls it answers, and the adapters for the
// Anthropic and OpenAI-compatible APIs.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

// a page: its window (message listeners), document (the element to fill, the iframe), location and fetch
const listeners = new Set(), hosts = {}, fetched = [];
let respond = () => new Response('', { status: 404 });
globalThis.window = { addEventListener: (t, f) => t === 'message' && listeners.add(f), removeEventListener: (t, f) => listeners.delete(f) };
globalThis.location = new URL('https://site.example/shop/page.html');
globalThis.fetch = async (url, init) => { fetched.push([String(url), init]); return respond(String(url), init); };
globalThis.document = {
  currentScript: { src: 'https://cdn.example/writer/embed/writer-embed.js' },
  querySelector: sel => hosts[sel] || null,
  createElement: tag => ({
    tag, style: {}, attrs: {}, parentNode: null,
    setAttribute(k, v) { this.attrs[k] = v; },
    contentWindow: { sent: [], postMessage(msg, origin, transfer) { this.sent.push({ msg: structuredClone(msg), origin, transfer }); } }
  })
};
const element = () => ({ children: [], appendChild(c) { this.children.push(c); c.parentNode = this; }, removeChild(c) { this.children = this.children.filter(x => x !== c); c.parentNode = null; c.contentWindow = null; } });
vm.runInThisContext(readFileSync(new URL('../embed/writer-embed.js', import.meta.url), 'utf8'));
const Writer = globalThis.window.Writer;

/** A message from the widget's frame (or, with from, from somewhere else). */
function say(w, data, from = {}) {
  const e = { data: Object.assign({ writer: 1, id: w.id }, data), origin: from.origin || 'https://cdn.example', source: 'source' in from ? from.source : w.frame.contentWindow };
  for (const f of [...listeners]) f(e);
}
const sent = w => w.frame.contentWindow.sent;
const tick = () => new Promise(r => setTimeout(r, 0));
const text = buf => new TextDecoder().decode(buf);

test('the iframe: the embed page next to the script, with the options in its URL, filling the element', () => {
  hosts['#doc'] = element();
  const w = Writer.embed('#doc', { mode: 'docx', theme: 'dark', lang: 'en', ai: () => ({}), aiModel: 'site-model', blankName: 'Draft', height: 500, radius: 0, title: 'Report' });
  const f = hosts['#doc'].children[0], u = new URL(f.src);
  assert.equal(f, w.frame);
  assert.equal(u.origin + u.pathname, 'https://cdn.example/writer/embed.dc.html');
  assert.deepEqual(Object.fromEntries(u.searchParams), { id: w.id, mode: 'docx', theme: 'dark', lang: 'en', ai: '1', model: 'site-model', blank: 'Draft' });
  assert.deepEqual([f.title, f.attrs.allow, w.origin], ['Report', 'clipboard-read; clipboard-write; fullscreen', 'https://cdn.example']);
  assert.match(f.style.cssText, /width:100%;height:500px;border:0;border-radius:0px/);
  const app = Writer.embed(element(), { mode: 'nonsense', base: 'https://other.example/w/' });
  assert.equal(new URL(app.frame.src).searchParams.get('mode'), 'app', 'an unknown mode is the whole app');
  assert.equal(new URL(app.frame.src).origin, 'https://other.example', 'base: Writer from another folder');
  assert.throws(() => Writer.embed('#nowhere', {}), /no element #nowhere/);
  w.destroy(); app.destroy();
});

test('the documents: URLs fetched by the site, Files, bytes, { name, url }, { name, data }; handed over on hello, then ready', async () => {
  respond = url => /\.docx$|\/download\?/.test(url) ? new Response(new Uint8Array([80, 75, 3, 4])) : new Response('', { status: 404 });
  const file = new File(['# 标题'], '笔记.md'), fetchedBefore = fetched.length;
  const w = Writer.embed(element(), { mode: 'app', files: [file, new Uint8Array([1, 2]), { name: 'data.csv', data: 'a,b' }, { name: '报告.docx', url: '/download?id=7' }], file: '/files/%E4%BC%9A%E8%AE%AE.docx', open: 'data.csv' });
  const readyEvents = []; w.on('ready', e => readyEvents.push(e));
  await tick();
  assert.deepEqual(fetched.slice(fetchedBefore), [['https://site.example/download?id=7', { credentials: 'same-origin' }], ['https://site.example/files/%E4%BC%9A%E8%AE%AE.docx', { credentials: 'same-origin' }]], 'URLs on the site\'s page, with its cookies');
  say(w, { type: 'hello' }); await tick(); await tick();
  const init = sent(w).at(-1);
  assert.equal(init.origin, 'https://cdn.example', 'only to the frame\'s origin');
  assert.equal(init.msg.type, 'init');
  assert.deepEqual(init.msg.files.map(f => f.name), ['笔记.md', '1-document.docx', 'data.csv', '报告.docx', '会议.docx']);
  assert.deepEqual(init.msg.files.map((f, i) => i === 0 || i === 2 ? text(f.data) : [...new Uint8Array(f.data)]), ['# 标题', [1, 2], 'a,b', [80, 75, 3, 4], [80, 75, 3, 4]]);
  assert.equal(init.msg.open, 'data.csv');
  assert.equal(init.transfer.length, 5, 'the bytes are moved, not copied');
  say(w, { type: 'ready', file: 'data.csv' });
  assert.deepEqual([await w.ready, readyEvents], [{ file: 'data.csv' }, [{ file: 'data.csv' }]]);
  w.destroy();
});

test('a document that cannot be read fails ready: a URL that is not there, something that is not a file', async () => {
  respond = () => new Response('', { status: 404 });
  const missing = Writer.embed(element(), { mode: 'docx', file: 'missing.docx' });
  say(missing, { type: 'hello' });
  await assert.rejects(missing.ready, /404 fetching https:\/\/site\.example\/shop\/missing\.docx/);
  const odd = Writer.embed(element(), { mode: 'docx', file: 42 });
  say(odd, { type: 'hello' });
  await assert.rejects(odd.ready, /a file is a URL, a File or Blob, bytes, \{ name, url \} or \{ name, data \}/);
  const broken = Writer.embed(element(), { mode: 'docx' });
  say(broken, { type: 'ready', file: '', error: 'the engine did not start' });
  await assert.rejects(broken.ready, /the engine did not start/);
  missing.destroy(); odd.destroy(); broken.destroy();
});

async function readyWidget(options) {
  const w = Writer.embed(element(), Object.assign({ mode: 'docx' }, options));
  say(w, { type: 'hello' }); await tick();
  say(w, { type: 'ready', file: 'a.docx' }); await w.ready;
  return w;
}

test('calls wait for ready, go to the frame with a number, and settle with its answer; others\' messages change nothing', async () => {
  const w = await readyWidget();
  const run = w.run('add a.docx /body --type paragraph'), put = w.put('b.md', '# Hi'), get = w.get('a.docx'), list = w.list();
  await tick(); await tick();
  const calls = sent(w).filter(m => m.msg.type === 'call').map(m => m.msg);
  assert.deepEqual([calls.map(c => c.method).sort(), calls.map(c => c.seq).sort()], [['get', 'list', 'put', 'run'], [1, 2, 3, 4]]);
  const by = name => calls.find(c => c.method === name);
  assert.deepEqual(by('run').args, ['add a.docx /body --type paragraph']);
  assert.deepEqual([by('put').args[0], text(by('put').args[1])], ['b.md', '# Hi'], 'text to put is the document\'s content');
  assert.equal(sent(w).find(m => m.msg.method === 'put').transfer.length, 1);
  // a stranger's answers: another origin, another window, another widget's id
  say(w, { type: 'result', seq: by('run').seq, ok: true, value: 'forged' }, { origin: 'https://evil.example' });
  say(w, { type: 'result', seq: by('run').seq, ok: true, value: 'forged' }, { source: {} });
  say(w, { type: 'result', seq: by('run').seq, ok: true, value: 'forged', id: 'someone-else' });
  say(w, { type: 'result', seq: by('run').seq, ok: true, value: { code: 0, output: 'Added' } });
  assert.deepEqual(await run, { code: 0, output: 'Added' });
  say(w, { type: 'result', seq: by('list').seq, ok: false, error: 'x.docx not found', code: 'FILE_NOT_FOUND', hint: 'Check the path.' });
  await assert.rejects(list, { message: 'x.docx not found', code: 'FILE_NOT_FOUND', hint: 'Check the path.' });
  say(w, { type: 'result', seq: by('get').seq, ok: true, value: new Uint8Array([80, 75]).buffer });
  const blob = await get;
  assert.deepEqual([blob.type, blob.size], ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 2]);
  say(w, { type: 'result', seq: by('put').seq, ok: true, value: { path: 'b.md' } });
  assert.deepEqual(await put, { path: 'b.md' });
  w.destroy();
});

test('events reach on() until off(); destroy removes the frame and fails what still waits', async () => {
  const w = await readyWidget(), seen = [];
  const fn = e => seen.push(e);
  w.on('change', fn).on('open', fn);
  say(w, { type: 'event', name: 'change', data: { files: ['a.docx'], from: 'editor' } });
  say(w, { type: 'event', name: 'open', data: { file: 'a.docx' } });
  w.off('change', fn);
  say(w, { type: 'event', name: 'change', data: { files: ['a.docx'], from: 'api' } });
  assert.deepEqual(seen, [{ files: ['a.docx'], from: 'editor' }, { file: 'a.docx' }]);
  const waiting = w.current();
  await tick();
  const host = w.frame.parentNode;
  w.destroy();
  await assert.rejects(waiting, /the widget was removed/);
  await assert.rejects(w.run('view a.docx outline'), /the widget was removed/);
  assert.deepEqual(host.children, []);
  const before = listeners.size;
  w.destroy(); // twice is fine
  assert.equal(listeners.size, before);
});

test('the assistant\'s model calls: the site\'s ai() answers, its text streams back, errors and stops reach it', async () => {
  let seen = null, stopped = false;
  const ai = async (request, ctx) => {
    seen = request;
    if (request.system === 'fail') throw new Error('429 rate limited');
    if (request.system === 'slow') return new Promise((_, reject) => ctx.signal.addEventListener('abort', () => { stopped = true; reject(new Error('aborted')); }));
    ctx.onDelta('好'); ctx.onDelta('的');
    return { text: '好的', toolCalls: [{ id: 7, name: 'writer', input: { command: 'view a.docx outline' } }, { name: 'plan' }] };
  };
  const w = await readyWidget({ ai });
  const events = seq => sent(w).filter(m => m.msg.type === 'chat-event' && m.msg.seq === seq).map(m => [m.msg.name, m.msg.data]);
  say(w, { type: 'chat', seq: 1, request: { system: 'S', messages: [{ role: 'user', content: 'hi' }], tools: [] } });
  await tick(); await tick();
  assert.deepEqual(seen.messages, [{ role: 'user', content: 'hi' }]);
  assert.deepEqual(events(1), [['delta', '好'], ['delta', '的'], ['done', { text: '好的', toolCalls: [{ id: '7', name: 'writer', input: { command: 'view a.docx outline' } }, { id: '', name: 'plan', input: {} }] }]]);
  say(w, { type: 'chat', seq: 2, request: { system: 'fail', messages: [], tools: [] } });
  await tick(); await tick();
  assert.deepEqual(events(2), [['error', '429 rate limited']]);
  say(w, { type: 'chat', seq: 3, request: { system: 'slow', messages: [], tools: [] } });
  await tick();
  say(w, { type: 'chat-abort', seq: 3 });
  await tick(); await tick();
  assert.deepEqual([stopped, events(3)], [true, []], 'stopped in the editor: the site\'s call is cancelled, nothing more is sent');
  w.destroy();
  const none = await readyWidget();
  say(none, { type: 'chat', seq: 1, request: { system: 'S', messages: [], tools: [] } });
  await tick(); await tick();
  assert.deepEqual(sent(none).filter(m => m.msg.type === 'chat-event').map(m => m.msg.data), ['No assistant: pass options.ai']);
  none.destroy();
});

// ---------- the adapters ----------

/** A streamed answer, cut into small pieces of bytes (a character may be split between two). */
function stream(lines, cut = 5) {
  const bytes = new TextEncoder().encode(lines.join(''));
  return new Response(new ReadableStream({ start(c) { for (let i = 0; i < bytes.length; i += cut) c.enqueue(bytes.slice(i, i + cut)); c.close(); } }), { headers: { 'content-type': 'text/event-stream' } });
}
const anthropicEvents = events => events.map(e => 'event: ' + e.type + '\r\ndata: ' + JSON.stringify(e) + '\r\n\r\n');
const openaiChunks = chunks => chunks.map(c => 'data: ' + JSON.stringify(c) + '\n\n').concat(['data: [DONE]\n\n']);
const request = {
  system: 'SYSTEM', tools: [{ name: 'writer', description: 'Runs a command', input_schema: { type: 'object', properties: { command: { type: 'string' } } } }],
  messages: [
    { role: 'user', content: '加一段' },
    { role: 'assistant', content: '好的', toolCalls: [{ id: 'a', name: 'writer', input: { command: 'add a.docx /body' } }, { id: 'b', name: 'writer', input: { command: 'view a.docx outline' } }] },
    { role: 'tool', toolCallId: 'a', name: 'writer', content: 'Added', isError: false },
    { role: 'tool', toolCallId: 'b', name: 'writer', content: 'No such file', isError: true },
    { role: 'assistant', content: '', toolCalls: [{ id: 'c', name: 'writer', input: {} }] },
    { role: 'tool', toolCallId: 'c', name: 'writer', content: '(no output)', isError: false }
  ]
};

test('Anthropic: the conversation as the Messages API has it, the stream read back as text and tool calls', async () => {
  let call = null;
  respond = (url, init) => { call = [url, init]; return stream(anthropicEvents([
    { type: 'message_start', message: { id: 'm' } },
    { type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } },
    { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: '好的，' } },
    { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: '再加一段。' } },
    { type: 'content_block_stop', index: 0 },
    { type: 'ping' },
    { type: 'content_block_start', index: 1, content_block: { type: 'tool_use', id: 'tu1', name: 'writer', input: {} } },
    { type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: '{"command": "add a.docx /body ' } },
    { type: 'content_block_delta', index: 1, delta: { type: 'input_json_delta', partial_json: '--prop text=段落"}' } },
    { type: 'content_block_stop', index: 1 },
    { type: 'message_delta', delta: { stop_reason: 'tool_use' } },
    { type: 'message_stop' }
  ])); };
  const deltas = [], ac = new AbortController();
  const reply = await Writer.ai.anthropic({ url: '/api/anthropic', model: 'claude-x', headers: { 'x-site': '1' } })(request, { onDelta: d => deltas.push(d), signal: ac.signal });
  assert.deepEqual(reply, { text: '好的，再加一段。', toolCalls: [{ id: 'tu1', name: 'writer', input: { command: 'add a.docx /body --prop text=段落' } }] });
  assert.deepEqual(deltas, ['好的，', '再加一段。']);
  const [url, init] = call, body = JSON.parse(init.body);
  assert.equal(url, '/api/anthropic');
  assert.deepEqual([init.method, init.signal, init.credentials], ['POST', ac.signal, 'same-origin']);
  assert.deepEqual(init.headers, { 'content-type': 'application/json', 'anthropic-version': '2023-06-01', 'x-site': '1' }, 'the site\'s own endpoint: no browser-access header');
  assert.deepEqual([body.model, body.max_tokens, body.system, body.tools, body.stream], ['claude-x', 8192, 'SYSTEM', request.tools, true]);
  assert.deepEqual(body.messages, [
    { role: 'user', content: '加一段' },
    { role: 'assistant', content: [{ type: 'text', text: '好的' }, { type: 'tool_use', id: 'a', name: 'writer', input: { command: 'add a.docx /body' } }, { type: 'tool_use', id: 'b', name: 'writer', input: { command: 'view a.docx outline' } }] },
    { role: 'user', content: [{ type: 'tool_result', tool_use_id: 'a', content: 'Added', is_error: false }, { type: 'tool_result', tool_use_id: 'b', content: 'No such file', is_error: true }] },
    { role: 'assistant', content: [{ type: 'tool_use', id: 'c', name: 'writer', input: {} }] },
    { role: 'user', content: [{ type: 'tool_result', tool_use_id: 'c', content: '(no output)', is_error: false }] }
  ]);
});

test('Anthropic: straight to the API only with the browser-access header; errors in the answer or the stream fail the call', async () => {
  let call = null;
  respond = (url, init) => { call = [url, init]; return stream(anthropicEvents([{ type: 'message_start', message: {} }, { type: 'error', error: { type: 'overloaded_error', message: 'Overloaded' } }])); };
  const direct = Writer.ai.anthropic({ headers: { 'x-api-key': 'k' } });
  await assert.rejects(direct(request, { onDelta() { }, signal: null }), /Overloaded/);
  assert.equal(call[0], 'https://api.anthropic.com/v1/messages');
  assert.equal(call[1].headers['anthropic-dangerous-direct-browser-access'], 'true');
  assert.equal(JSON.parse(call[1].body).model, 'claude-sonnet-5', 'the desktop assistant\'s default model');
  respond = () => new Response(JSON.stringify({ type: 'error', error: { type: 'authentication_error', message: 'invalid x-api-key' } }), { status: 401 });
  await assert.rejects(direct(request, { onDelta() { }, signal: null }), /^Error: 401 invalid x-api-key$/);
  respond = () => new Response('Bad Gateway', { status: 502 });
  await assert.rejects(direct(request, { onDelta() { }, signal: null }), /^Error: 502 Bad Gateway$/);
});

test('OpenAI-compatible: the system message first, tool calls as functions; the streamed deltas put back together', async () => {
  let call = null;
  respond = (url, init) => { call = [url, init]; return stream(openaiChunks([
    { choices: [{ index: 0, delta: { role: 'assistant', content: '' } }] },
    { choices: [{ index: 0, delta: { content: '先看' } }] },
    { choices: [{ index: 0, delta: { content: '大纲。' } }] },
    { choices: [{ index: 0, delta: { tool_calls: [{ index: 0, id: 'call_1', type: 'function', function: { name: 'writer', arguments: '' } }] } }] },
    { choices: [{ index: 0, delta: { tool_calls: [{ index: 0, function: { arguments: '{"command":' } }] } }] },
    { choices: [{ index: 0, delta: { tool_calls: [{ index: 0, function: { arguments: '"view a.docx outline"}' } }] } }] },
    { choices: [{ index: 0, delta: { tool_calls: [{ index: 1, id: 'call_2', type: 'function', function: { name: 'plan', arguments: '{}' } }] } }] },
    { choices: [{ index: 0, delta: {}, finish_reason: 'tool_calls' }] },
    { choices: [] }
  ]), 3); };
  const deltas = [];
  const reply = await Writer.ai.openai({ url: 'https://llm.example/v1/chat/completions', model: 'qwen', headers: { Authorization: 'Bearer t' } })(request, { onDelta: d => deltas.push(d), signal: null });
  assert.deepEqual(reply, { text: '先看大纲。', toolCalls: [{ id: 'call_1', name: 'writer', input: { command: 'view a.docx outline' } }, { id: 'call_2', name: 'plan', input: {} }] });
  assert.deepEqual(deltas, ['先看', '大纲。']);
  const [url, init] = call, body = JSON.parse(init.body);
  assert.equal(url, 'https://llm.example/v1/chat/completions');
  assert.deepEqual(init.headers, { 'content-type': 'application/json', Authorization: 'Bearer t' });
  assert.deepEqual([body.model, body.stream], ['qwen', true]);
  assert.deepEqual(body.tools, [{ type: 'function', function: { name: 'writer', description: 'Runs a command', parameters: request.tools[0].input_schema } }]);
  assert.deepEqual(body.messages, [
    { role: 'system', content: 'SYSTEM' },
    { role: 'user', content: '加一段' },
    { role: 'assistant', content: '好的', tool_calls: [
      { id: 'a', type: 'function', function: { name: 'writer', arguments: '{"command":"add a.docx /body"}' } },
      { id: 'b', type: 'function', function: { name: 'writer', arguments: '{"command":"view a.docx outline"}' } }] },
    { role: 'tool', tool_call_id: 'a', content: 'Added' },
    { role: 'tool', tool_call_id: 'b', content: 'No such file' },
    { role: 'assistant', content: null, tool_calls: [{ id: 'c', type: 'function', function: { name: 'writer', arguments: '{}' } }] },
    { role: 'tool', tool_call_id: 'c', content: '(no output)' }
  ]);
});

test('OpenAI-compatible: servers that leave out the tool call index; an error answer fails the call', async () => {
  respond = () => stream(openaiChunks([
    { choices: [{ delta: { tool_calls: [{ id: 'x1', function: { name: 'writer', arguments: '{"command":"view a.docx outline"}' } }] } }] },
    { choices: [{ delta: { tool_calls: [{ id: 'x2', function: { name: 'writer', arguments: '{"command":' } }] } }] },
    { choices: [{ delta: { tool_calls: [{ function: { arguments: '"view b.docx outline"}' } }] } }] }
  ]));
  const reply = await Writer.ai.openai({ url: '/v1/chat/completions' })(request, { onDelta() { }, signal: null });
  assert.deepEqual(reply.toolCalls, [{ id: 'x1', name: 'writer', input: { command: 'view a.docx outline' } }, { id: 'x2', name: 'writer', input: { command: 'view b.docx outline' } }]);
  respond = () => new Response(JSON.stringify({ error: { message: 'model not found' } }), { status: 404 });
  await assert.rejects(Writer.ai.openai({ url: '/v1/chat/completions' })(request, { onDelta() { }, signal: null }), /^Error: 404 model not found$/);
});
