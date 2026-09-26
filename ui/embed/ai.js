// ui/embed/ai.js — the embedded editor's assistant: the desktop assistant's loop (src/Writer.Cli/Chat.cs) run in the page. The
// engine gives the system prompt and the editing tools (writer, batch, plan) and runs the tool calls; the site that embedded
// Writer answers each model call (writer-embed.js: options.ai). POST /chat answers with the server's events: delta {text} as the
// reply streams, text {text} once whole, tool {command, code, output, wrote}, done {steps}, error {message, hint}.
//
// The site sees one call as { system, messages, tools } and answers { text, toolCalls: [{ id, name, input }] }. messages is the
// turn so far, provider-neutral: { role: 'user', content } | { role: 'assistant', content, toolCalls } | { role: 'tool',
// toolCallId, name, content, isError } (writer-embed.js has adapters for the Anthropic and OpenAI-compatible APIs).
export const MAX_STEPS = 24;

/** One user turn: body is what the editor posts to /chat ({ file, messages, instructions?, selection? }). */
export function chatTurn(body, signal, { ask, engine }) {
  const enc = new TextEncoder();
  const stream = new ReadableStream({
    async start(ctl) {
      const send = (event, data) => { try { ctl.enqueue(enc.encode('event: ' + event + '\ndata: ' + JSON.stringify(data) + '\n\n')); } catch (e) { /* the editor went away */ } };
      try { await turn(body, signal, ask, engine, send); } catch (e) { if (!(e && e.name === 'AbortError')) send('error', { message: String((e && e.message) || e), hint: '' }); }
      try { ctl.close(); } catch (e) { }
    }
  });
  return new Response(stream, { status: 200, headers: { 'Content-Type': 'text/event-stream' } });
}

async function turn(body, signal, ask, engine, send) {
  const messages = (body.messages || []).filter(m => (m.role === 'user' || m.role === 'assistant') && String(m.content || '').trim())
    .map(m => ({ role: m.role, content: String(m.content) }));
  if (!messages.length || messages[messages.length - 1].role !== 'user') { send('error', { message: 'The last message must be from the user' }); return; }
  const selection = body.selection && String(body.selection).trim();
  if (selection) messages[messages.length - 1].content = 'The user selected this part of the document:\n<selection>\n' + selection + '\n</selection>\n\n' + messages[messages.length - 1].content;
  let system = await engine.system(body.file || '', !selection);
  if (body.instructions && String(body.instructions).trim()) system += "\n\n## The user's preferences\n\n" + String(body.instructions).trim() + '\n';
  const tools = await engine.tools();
  let steps = 0;
  for (; steps < MAX_STEPS; steps++) {
    if (signal && signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const reply = await ask({ system, messages: messages.map(m => Object.assign({}, m)), tools }, text => send('delta', { text }), signal) || {};
    const text = String(reply.text || ''), calls = (reply.toolCalls || []).filter(c => c && c.name);
    if (text) send('text', { text });
    messages.push({ role: 'assistant', content: text, toolCalls: calls });
    if (!calls.length) break;
    for (const c of calls) {
      if (signal && signal.aborted) throw new DOMException('Aborted', 'AbortError');
      const r = await engine.callTool(c.name, c.input || {}, 'assistant');
      send('tool', { command: r.display, code: r.code, output: r.output, wrote: r.wrote });
      messages.push({ role: 'tool', toolCallId: c.id || '', name: c.name, content: r.output || '(no output)', isError: r.code !== 0 });
    }
    if (steps === MAX_STEPS - 1) send('text', { text: '已达到本轮的步数上限，先停在这里；回复「继续」可以接着做。' }); // i18n-ok: the model's own words to the user, as the desktop assistant says them
  }
  send('done', { steps: Math.min(steps + 1, MAX_STEPS) });
}
