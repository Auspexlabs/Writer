// node --test ui/tests/ — the embedded editor's assistant (ui/embed/ai.js): the desktop assistant's loop in the page, the engine
// giving the prompt and the tools and running them, the site answering each model call; POST /chat's events as the desktop's.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chatTurn, MAX_STEPS } from '../embed/ai.js';

/** The server-sent events of a /chat answer, as [event, data]. */
async function events(res) {
  assert.equal(res.headers.get('content-type'), 'text/event-stream');
  return (await res.text()).split('\n\n').filter(Boolean).map(block => {
    const [e, d] = block.split('\n');
    return [e.replace(/^event: /, ''), JSON.parse(d.replace(/^data: /, ''))];
  });
}

/** server.js's engine as ai.js uses it: the prompt for a file, the tools, one tool call (results given in order). */
function engine(results = []) {
  const calls = [];
  return {
    calls,
    system: async (file, outline) => { calls.push(['system', file, outline]); return 'SYSTEM for ' + file; },
    tools: async () => [{ name: 'writer', description: 'Runs a writer command', input_schema: { type: 'object' } }],
    callTool: async (name, input, from) => { calls.push(['tool', name, input, from]); return results.shift() || { display: 'writer ' + input.command, code: 0, output: 'ok', wrote: true }; }
  };
}

test('a turn: the reply streams, a tool runs, the model sees its result and answers; the editor gets the events in order', async () => {
  const E = engine([{ display: 'writer add a.docx /body', code: 0, output: 'Added /body/p[3]', wrote: true }]);
  const asked = [], replies = [
    { text: '好的，加一段。', toolCalls: [{ id: 't1', name: 'writer', input: { command: 'add a.docx /body' } }] },
    { text: '加好了。' }
  ];
  const ask = async (request, onDelta) => {
    asked.push(structuredClone(request));
    const r = replies.shift();
    for (const piece of r.text.match(/.{1,3}/gu)) onDelta(piece);
    return r;
  };
  const body = { file: 'a.docx', instructions: '  回答简短。 ', messages: [{ role: 'assistant', content: '你好' }, { role: 'user', content: '  ' }, { role: 'system', content: 'x' }, { role: 'user', content: '加一段' }] };
  const got = await events(chatTurn(body, null, { ask, engine: E }));
  assert.deepEqual(got, [
    ['delta', { text: '好的，' }], ['delta', { text: '加一段' }], ['delta', { text: '。' }], ['text', { text: '好的，加一段。' }],
    ['tool', { command: 'writer add a.docx /body', code: 0, output: 'Added /body/p[3]', wrote: true }],
    ['delta', { text: '加好了' }], ['delta', { text: '。' }], ['text', { text: '加好了。' }],
    ['done', { steps: 2 }]
  ]);
  assert.equal(asked[0].system, "SYSTEM for a.docx\n\n## The user's preferences\n\n回答简短。\n", 'the desktop prompt, the user\'s preferences after it');
  assert.deepEqual(asked[0].tools.map(t => t.name), ['writer']);
  assert.deepEqual(asked[0].messages, [{ role: 'assistant', content: '你好' }, { role: 'user', content: '加一段' }], 'blank and unknown messages left out');
  assert.deepEqual(asked[1].messages.slice(2), [
    { role: 'assistant', content: '好的，加一段。', toolCalls: [{ id: 't1', name: 'writer', input: { command: 'add a.docx /body' } }] },
    { role: 'tool', toolCallId: 't1', name: 'writer', content: 'Added /body/p[3]', isError: false }
  ]);
  assert.deepEqual(E.calls, [['system', 'a.docx', true], ['tool', 'writer', { command: 'add a.docx /body' }, 'assistant']]);
});

test('a selection goes before the question, and the prompt leaves the outline out; a failed tool is an error result', async () => {
  const E = engine([{ display: 'writer set a.docx /x', code: 1, output: '', wrote: false }]);
  const asked = [], replies = [{ text: '', toolCalls: [{ id: 'x', name: 'writer', input: { command: 'set a.docx /x' } }] }, { text: '改不了。' }];
  const ask = async request => { asked.push(structuredClone(request)); return replies.shift(); };
  const got = await events(chatTurn({ file: 'a.docx', selection: ' 第一段 ', messages: [{ role: 'user', content: '改短' }] }, null, { ask, engine: E }));
  assert.equal(asked[0].messages[0].content, 'The user selected this part of the document:\n<selection>\n第一段\n</selection>\n\n改短');
  assert.deepEqual(E.calls[0], ['system', 'a.docx', false]);
  assert.deepEqual(asked[1].messages[2], { role: 'tool', toolCallId: 'x', name: 'writer', content: '(no output)', isError: true });
  assert.deepEqual(got.map(e => e[0]), ['tool', 'text', 'done'], 'no text event for an empty reply');
});

test('the last message must be the user\'s; a model call that fails is an error event', async () => {
  let asked = 0;
  const none = await events(chatTurn({ messages: [{ role: 'user', content: 'a' }, { role: 'assistant', content: 'b' }] }, null, { ask: async () => { asked++; }, engine: engine() }));
  assert.deepEqual([none, asked], [[['error', { message: 'The last message must be from the user' }]], 0]);
  const failed = await events(chatTurn({ messages: [{ role: 'user', content: 'a' }] }, null, { ask: async () => { throw new Error('401 bad key'); }, engine: engine() }));
  assert.deepEqual(failed, [['error', { message: '401 bad key', hint: '' }]]);
});

test(`at most ${MAX_STEPS} model calls a turn, then the desktop's note and done`, async () => {
  let asked = 0;
  const ask = async () => { asked++; return { text: '', toolCalls: [{ id: 'i' + asked, name: 'writer', input: { command: 'view a.docx outline' } }] }; };
  const got = await events(chatTurn({ file: 'a.docx', messages: [{ role: 'user', content: 'go' }] }, null, { ask, engine: engine() }));
  assert.equal(asked, MAX_STEPS);
  assert.deepEqual(got.slice(-2), [['text', { text: '已达到本轮的步数上限，先停在这里；回复「继续」可以接着做。' }], ['done', { steps: MAX_STEPS }]]); // i18n-ok: the desktop assistant's words
  assert.equal(got.filter(e => e[0] === 'tool').length, MAX_STEPS);
});

test('stopped by the user: the model call is cancelled, no tool runs after it, and no error is shown', async () => {
  const ac = new AbortController(), E = engine();
  let signalled = null;
  const ask = (request, onDelta, signal) => new Promise((resolve, reject) => {
    signalled = signal; onDelta('半');
    signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')));
  });
  const res = chatTurn({ file: 'a.docx', messages: [{ role: 'user', content: 'go' }] }, ac.signal, { ask, engine: E });
  setTimeout(() => ac.abort(), 10);
  assert.deepEqual(await events(res), [['delta', { text: '半' }]]);
  assert.equal(signalled, ac.signal, 'the site gets the editor\'s signal');
  assert.deepEqual(E.calls.filter(c => c[0] === 'tool'), []);
});
