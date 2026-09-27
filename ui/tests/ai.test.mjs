// node --test ui/tests/   — 设置 › AI: the provider list, the form ↔ request mapping, and the engine calls behind them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { AI_PROVIDERS, aiPreset, aiForm, pickProvider, keyStored, aiRequest, aiMissing, fastModel, modelMenu, modelName, listedModels, keepListed, accountLine } from '../prefs.js';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
if (!globalThis.dispatchEvent) { // a browser window is an EventTarget; Node's global is not
  const target = new EventTarget();
  for (const k of ['addEventListener', 'removeEventListener', 'dispatchEvent']) globalThis[k] = target[k].bind(target);
}
const EN = await import('../engine.js');

test('the providers match the engine\'s list: same ids, addresses, and which ones run without a key', () => {
  const cs = readFileSync(new URL('../../src/Writer.Cli/AiConfig.cs', import.meta.url), 'utf8');
  const engine = [...cs.matchAll(/\["(\w+)"\] = "([^"]*)",/g)].map(m => [m[1], m[2]]);
  assert.deepEqual(AI_PROVIDERS.map(p => [p.id, p.url]), engine);
  const keyless = cs.match(/NeedsKey => Provider is not \(([^)]*)\)/)[1].match(/\w+/g).filter(w => w !== 'or');
  assert.deepEqual(AI_PROVIDERS.filter(p => !p.key).map(p => p.id), keyless);
  assert.ok(AI_PROVIDERS.every(p => p.name), 'every provider has a name for the menu');
  assert.equal(aiPreset('nope').id, 'custom');
});

test('aiForm: GET /ai becomes the form; a stored key is a flag, never a value', () => {
  assert.deepEqual(aiForm({ provider: 'anthropic', baseUrl: 'https://api.anthropic.com', model: '', hasKey: false, source: 'none' }),
    { provider: 'anthropic', baseUrl: 'https://api.anthropic.com', model: 'claude-sonnet-5', completeModel: 'claude-haiku-4-5', key: '', hasKey: false, savedUrl: 'https://api.anthropic.com', source: 'none', account: null }, 'nothing saved yet: the recommended models are chosen');
  assert.deepEqual([aiForm({ provider: 'deepseek', model: 'deepseek-v4-pro', source: 'file' }).completeModel, aiForm({ provider: 'custom', model: '', source: 'file' }).model], ['', ''], 'saved settings are shown as they are');
  assert.equal(aiForm({ provider: 'openai', completeModel: 'gpt-6-luna' }).completeModel, 'gpt-6-luna');
  const saved = aiForm({ provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', hasKey: true, source: 'file' });
  assert.equal(saved.key, '');
  assert.equal(keyStored(saved), true);
  assert.deepEqual(aiForm(null), { provider: 'anthropic', baseUrl: 'https://api.anthropic.com', model: 'claude-sonnet-5', completeModel: 'claude-haiku-4-5', key: '', hasKey: false, savedUrl: '', source: 'none', account: null });
  assert.equal(aiForm({ provider: 'someday' }).provider, 'custom');
});

test('pickProvider fills in the address, the recommended model and the fast one; the stored key stays with its server', () => {
  const saved = aiForm({ provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', hasKey: true, source: 'file' });
  const typed = { ...saved, key: 'sk-typed' };
  const kimi = pickProvider(typed, 'kimi');
  assert.deepEqual([kimi.provider, kimi.baseUrl, kimi.model, kimi.completeModel, kimi.key], ['kimi', 'https://api.moonshot.cn/v1', 'kimi-k3', 'kimi-k2.6', '']);
  assert.deepEqual([pickProvider(saved, 'ollama').model, pickProvider(saved, 'ollama').completeModel], ['', ''], 'a local server: its own models, read from it');
  assert.equal(keyStored(kimi), false, 'another server: the DeepSeek key is not offered to Kimi');
  assert.equal(keyStored(pickProvider(kimi, 'deepseek')), true, 'back on the same server');
  assert.equal(pickProvider(typed, 'deepseek'), typed, 'picking the same service changes nothing');
  assert.equal(pickProvider(saved, 'custom').baseUrl, '');
  assert.equal(keyStored({ ...saved, baseUrl: 'https://api.deepseek.com/v1' }), true, 'another path on the same server keeps it');
  assert.equal(keyStored({ ...saved, baseUrl: 'not a url' }), false);
});

test('aiRequest sends a key only when one was typed; aiMissing names what is left to fill in', () => {
  const form = { provider: 'deepseek', baseUrl: ' https://api.deepseek.com/ ', model: ' deepseek-flash ', key: '', hasKey: true };
  assert.deepEqual(aiRequest(form), { provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', completeModel: '' });
  assert.deepEqual(aiRequest({ ...form, key: ' sk-1 ', completeModel: 'deepseek-flash' }), { provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', completeModel: 'deepseek-flash', apiKey: 'sk-1' });
  assert.equal(aiRequest({ provider: 'anthropic', baseUrl: 'https://api.anthropic.com', model: '' }).model, 'claude-sonnet-5', 'Anthropic has a default model');
  assert.equal(aiMissing(form), '');
  assert.equal(aiMissing({ ...form, model: '' }), '请选择模型');
  assert.match(aiMissing({ ...form, baseUrl: '' }), /接口地址/);
  assert.match(aiMissing({ ...form, baseUrl: 'ftp://x.test' }), /http/);
  assert.equal(aiMissing({ provider: 'ollama', baseUrl: 'http://localhost:11434/v1', model: 'qwen3:8b' }), '');
});

test('every service with a key has models to choose from, one marked for the autocomplete, its recommended one first', () => {
  for (const p of AI_PROVIDERS.filter(x => x.key)) {
    assert.ok(p.models.length >= 2, p.id + ' has a menu');
    assert.equal(p.models[0].id, p.model, p.id + ': the recommended model leads its menu');
    assert.equal(p.models.filter(m => m.fast).length, 1, p.id + ': one model for the autocomplete');
    assert.ok(p.models.every(m => m.id && m.name && m.note && !/[\u3400-\u9fff]/.test(m.name)), p.id + ': ids, names without Chinese, notes');
    assert.equal(new Set(p.models.map(m => m.id)).size, p.models.length);
  }
  assert.deepEqual(AI_PROVIDERS.filter(x => !x.key && !x.login).map(x => x.models.length), [0, 0, 0], 'local servers and custom addresses list their own');
  const gpt = AI_PROVIDERS.filter(x => x.login);
  assert.deepEqual(gpt.map(x => [x.id, x.key, x.model, x.models[0].id, x.models.filter(m => m.fast).map(m => m.id)]), [['chatgpt', false, 'gpt-6-astra', 'gpt-6-astra', ['gpt-5.6-luna']]], 'a sign-in, no key; the Codex default first');
  assert.ok(gpt[0].models.every(m => m.id && m.name && m.note && !/[\u3400-\u9fff]/.test(m.name)));
  assert.deepEqual([fastModel('anthropic'), fastModel('openai'), fastModel('glm'), fastModel('custom')], ['claude-haiku-4-5', 'gpt-6-luna', 'glm-4.7-flashx', '']);
  assert.deepEqual([modelName('anthropic', 'claude-opus-5-5'), modelName('anthropic', 'claude-x'), modelName('openai', '')], ['Claude Opus 5.5', 'claude-x', '']);
});

test('用 ChatGPT 登录: the form carries the account, never a key or an address of its own; saving waits for the sign-in', () => {
  const signedOut = aiForm({ provider: 'chatgpt', baseUrl: 'https://chatgpt.com/backend-api/codex', model: 'gpt-6-astra', completeModel: 'gpt-5.6-luna', hasKey: false, source: 'file' });
  assert.equal(signedOut.account, null);
  assert.equal(aiMissing(signedOut), '请先用 ChatGPT 登录');
  const account = { email: 'ada@example.com', plan: 'plus' }, form = { ...signedOut, account, key: 'sk-typed', baseUrl: 'https://elsewhere.example/v1' };
  assert.equal(aiMissing(form), '');
  assert.deepEqual(aiRequest(form), { provider: 'chatgpt', baseUrl: 'https://chatgpt.com/backend-api/codex', model: 'gpt-6-astra', completeModel: 'gpt-5.6-luna' }, 'the sign-in\'s tokens only go to the Codex backend; a typed key is not sent');
  assert.deepEqual(aiForm({ provider: 'deepseek', model: 'deepseek-flash', source: 'file', account }).account, account, 'the sign-in stays while another service is in use');
  const picked = pickProvider({ ...aiForm({ provider: 'deepseek', model: 'deepseek-flash', source: 'file', account }), key: 'sk-1' }, 'chatgpt');
  assert.deepEqual([picked.provider, picked.baseUrl, picked.model, picked.completeModel, picked.key, picked.account], ['chatgpt', 'https://chatgpt.com/backend-api/codex', 'gpt-6-astra', 'gpt-5.6-luna', '', account]);
  assert.deepEqual([accountLine(account), accountLine({ email: 'b@x.test', plan: 'team' }), accountLine({ email: 'c@x.test', plan: 'someday' }), accountLine({ email: 'd@x.test' }), accountLine(null)],
    ['ada@example.com · Plus', 'b@x.test · Business', 'c@x.test · someday', 'd@x.test', '']);
  assert.deepEqual(modelMenu('chatgpt', [{ id: 'gpt-6-astra' }, { id: 'gpt-5.7-sol', name: 'GPT-5.7 Sol' }], 'gpt-6-astra').map(m => m.id), ['gpt-6-astra', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.5', 'gpt-5.7-sol'], 'what the plan lists joins the menu');
});

test('chatgptSignIn starts the sign-in, follows it to its end and announces the new settings; cancel and sign-out go to the engine', async () => {
  const heard = [], stored = {}, states = [];
  globalThis.localStorage = { setItem: (k, v) => { stored[k] = v; } };
  const listen = () => heard.push('writer-ai');
  globalThis.addEventListener('writer-ai', listen);
  try {
    let polls = 0;
    await withEngine((url, opts) => url === '/ai/chatgpt/login' && opts.method === 'POST' ? [200, { url: 'https://auth.openai.com/oauth/authorize?state=s', opened: false }]
      : url === '/ai/chatgpt/login' ? [200, ++polls < 3 ? { state: 'waiting' } : { state: 'done', email: 'ada@example.com', plan: 'plus' }]
      : url === '/ai/chatgpt/cancel' ? [200, { state: 'cancelled' }] : [200, { provider: 'chatgpt', model: 'gpt-6-astra', source: 'file' }], async calls => {
      const end = await EN.chatgptSignIn(s => states.push(s), null, 1);
      assert.deepEqual(end, { state: 'done', email: 'ada@example.com', plan: 'plus' });
      assert.deepEqual(states, [{ state: 'waiting', url: 'https://auth.openai.com/oauth/authorize?state=s', opened: false }, end], 'the waiting polls are not announced one by one');
      assert.deepEqual(heard, ['writer-ai'], 'done: the panel and other windows pick up ChatGPT');
      assert.deepEqual(await EN.chatgptCancel(), { state: 'cancelled' });
      assert.deepEqual((await EN.chatgptLogout()).provider, 'chatgpt');
      assert.deepEqual(heard, ['writer-ai', 'writer-ai']);
      assert.deepEqual(calls.map(c => [c.method, c.url]), [['POST', '/ai/chatgpt/login'], ['GET', '/ai/chatgpt/login'], ['GET', '/ai/chatgpt/login'], ['GET', '/ai/chatgpt/login'], ['POST', '/ai/chatgpt/cancel'], ['POST', '/ai/chatgpt/logout']]);
      assert.deepEqual([calls[0].body, calls[4].body, calls[5].body], [{}, {}, {}]);
    });
    await withEngine(url => url === '/ai/chatgpt/login' ? [200, { state: 'waiting', url: 'u', opened: true }] : [404, {}], async () => {
      const ac = new AbortController(), following = EN.chatgptSignIn(null, ac.signal, 50);
      setTimeout(() => ac.abort(), 20);
      await assert.rejects(following, e => e.name === 'AbortError', 'a closed window stops following; the sign-in goes on in the engine');
    });
    assert.equal(heard.length, 2);
  } finally { globalThis.removeEventListener('writer-ai', listen); delete globalThis.localStorage; }
});

test('modelMenu: the service\'s own models with their notes, then what it listed, then a name typed in; each once', () => {
  const menu = modelMenu('openai', [{ id: 'gpt-6-sol' }, { id: 'gpt-6-nova', name: 'GPT-6 Nova' }, { id: 'o9' }, {}], 'my-fine-tune');
  assert.deepEqual(menu.map(m => [m.id, m.name, m.listed]), [['gpt-6-sol', 'GPT-6 Sol', undefined], ['gpt-6-astra', 'GPT-6 Astra', undefined], ['gpt-6-luna', 'GPT-6 Luna', undefined],
    ['gpt-6-nova', 'GPT-6 Nova', true], ['o9', 'o9', true], ['my-fine-tune', 'my-fine-tune', false]]);
  assert.equal(menu[0].note, '能力和价格均衡');
  assert.deepEqual(modelMenu('ollama', [{ id: 'qwen3:8b' }], 'qwen3:8b').map(m => m.id), ['qwen3:8b'], 'the model in use is not listed twice');
});

test('listedModels / keepListed: what a service listed is kept per service and address', () => {
  const store = new Map(), storage = { getItem: k => store.get(k) ?? null, setItem: (k, v) => store.set(k, v) };
  assert.deepEqual(listedModels('openai', 'https://api.openai.com/v1', storage), []);
  keepListed('openai', 'https://api.openai.com/v1', [{ id: 'a', name: 'A', created: 1 }, { id: 'b' }], storage);
  keepListed('ollama', 'http://localhost:11434/v1', [{ id: 'qwen3:8b' }], storage);
  assert.deepEqual(listedModels('openai', 'https://api.openai.com/v1', storage), [{ id: 'a', name: 'A' }, { id: 'b' }]);
  assert.deepEqual(listedModels('openai', 'https://proxy.example/v1', storage), [], 'another address: its own list');
  assert.deepEqual(listedModels('ollama', 'http://localhost:11434/v1', storage), [{ id: 'qwen3:8b' }]);
});

/** Swaps fetch for a fake engine: records each call, answers with reply(url, opts) → [status, body]. */
async function withEngine(reply, run) {
  const prev = globalThis.fetch, calls = [];
  globalThis.fetch = async (url, opts = {}) => {
    calls.push({ url, method: opts.method || 'GET', headers: opts.headers || {}, body: opts.body ? JSON.parse(opts.body) : undefined, credentials: opts.credentials, keepalive: !!opts.keepalive });
    const [status, body] = reply(url, opts);
    return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
  };
  try { await run(calls); } finally { globalThis.fetch = prev; }
}

test('aiConfig / saveAiConfig / testAi go to the engine, with the session cookie and JSON bodies', async () => {
  const cfg = { provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', hasKey: true, source: 'file' };
  const heard = [], stored = {};
  globalThis.localStorage = { setItem: (k, v) => { stored[k] = v; } };
  const listen = () => heard.push('writer-ai');
  globalThis.addEventListener('writer-ai', listen);
  try {
    await withEngine(url => url === '/ai/test' ? [200, { ok: false, error: 'Key 无效' }] : [200, cfg], async calls => {
      assert.deepEqual(await EN.aiConfig(), cfg);
      const body = { provider: 'deepseek', baseUrl: 'https://api.deepseek.com', model: 'deepseek-flash', apiKey: 'sk-1' };
      assert.deepEqual(await EN.saveAiConfig(body), cfg);
      assert.deepEqual(await EN.testAi(body), { ok: false, error: 'Key 无效' });
      assert.deepEqual(await EN.testAi(), { ok: false, error: 'Key 无效' });
      assert.deepEqual(calls.map(c => [c.method, c.url]), [['GET', '/ai'], ['PUT', '/ai'], ['POST', '/ai/test'], ['POST', '/ai/test']]);
      assert.ok(calls.every(c => c.credentials === 'same-origin'));
      assert.deepEqual(calls[1].body, body);
      assert.equal(calls[1].headers['Content-Type'], 'application/json');
      assert.equal(calls[1].keepalive, true, 'a save as the settings window closes still lands');
      assert.deepEqual(calls[3].body, {}, 'no settings: the engine tests the saved ones');
    });
    assert.deepEqual(heard, ['writer-ai'], 'this window hears about the save');
    assert.ok(stored['writer-ai'], 'other windows hear it through storage');
  } finally { globalThis.removeEventListener('writer-ai', listen); delete globalThis.localStorage; }
});

test('aiModels and complete go to the engine; complete carries where the text is typed and can be cancelled', async () => {
  await withEngine(url => url === '/ai/models' ? [200, { models: [{ id: 'gpt-6-sol' }] }] : [200, { text: ' and the rest' }], async calls => {
    assert.deepEqual(await EN.aiModels({ provider: 'openai', baseUrl: 'https://api.openai.com/v1' }), [{ id: 'gpt-6-sol' }]);
    assert.deepEqual(await EN.aiModels(), [{ id: 'gpt-6-sol' }]);
    assert.equal(await EN.complete('The start', 'after', 'a spreadsheet cell'), ' and the rest');
    assert.deepEqual(calls.map(c => [c.method, c.url]), [['POST', '/ai/models'], ['POST', '/ai/models'], ['POST', '/complete']]);
    assert.deepEqual([calls[1].body, calls[2].body], [{}, { before: 'The start', after: 'after', hint: 'a spreadsheet cell' }]);
  });
  const prev = globalThis.fetch;
  globalThis.fetch = (url, opts) => new Promise((_, reject) => opts.signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError'))));
  try {
    const ac = new AbortController(), pending = EN.complete('x', '', '', ac.signal);
    ac.abort();
    await assert.rejects(pending, e => e.name === 'AbortError');
  } finally { globalThis.fetch = prev; }
});

test('saveAiConfig rejects with the engine\'s message when the settings are refused, and announces nothing', async () => {
  const heard = [];
  const listen = () => heard.push(1);
  globalThis.addEventListener('writer-ai', listen);
  try {
    await withEngine(() => [400, { error: { code: 'VALIDATION', message: '请选择模型', hint: '在模型列表里选一个。' } }], async () => {
      await assert.rejects(EN.saveAiConfig({ provider: 'openai', baseUrl: 'https://api.openai.com/v1', model: '' }), e => e instanceof EN.EngineError && e.message === '请选择模型' && e.code === 'VALIDATION');
    });
    assert.deepEqual(heard, []);
  } finally { globalThis.removeEventListener('writer-ai', listen); }
});

test('the settings window and the shell use these: rows, the write-only key field, and the settings hook', () => {
  const settings = readFileSync(new URL('../MacSettings.dc.html', import.meta.url), 'utf8');
  for (const s of ['EN.aiConfig()', 'EN.saveAiConfig(', 'EN.testAi(', "type: type || 'text'", "'password'", "'清除'", "'已保存'", "'测试连接'", "'服务商'", "'接口地址'", "'模型'"])
    assert.ok(settings.includes(s), 'MacSettings: ' + s);
  const shell = readFileSync(new URL('../index.dc.html', import.meta.url), 'utf8');
  assert.match(shell, /const NO_CHAT = '在「设置 › AI」里用 ChatGPT 账号登录，或填写你自己的模型 API/);
  for (const s of ['EN.chatgptSignIn(', 'EN.chatgptCancel()', 'EN.chatgptLogout()', 'data-chatgpt-login', 'gptOn: !!P && !EMBED']) assert.ok(shell.includes(s), 'the AI panel: ' + s);
  for (const s of ['EN.chatgptSignIn(', 'EN.chatgptCancel()', 'EN.chatgptLogout()', "'ChatGPT 账号'", "'用 ChatGPT 登录'"]) assert.ok(settings.includes(s), 'MacSettings: ' + s);
  assert.ok(!shell.includes('ANTHROPIC_API_KEY'), 'no env-var hint in the panel any more');
  assert.ok(shell.includes("openAiSettings: () => this.openSettings('ai')") && shell.includes('this.props.onPrefs(tab)'));
  assert.ok(shell.includes("addEventListener('writer-ai', this.onAi)") && shell.includes("e.key === 'writer-ai'"), 'the panel picks up a saved model without a reload');
  assert.ok(shell.includes("addEventListener('focus', this.onAi)"), 'a folder window of its own (own engine) asks again when it comes to the front');
  const mac = readFileSync(new URL('../mac.dc.html', import.meta.url), 'utf8');
  assert.ok(mac.includes("openPrefs = tab => this.showWindow('settings', tab)"), 'the Mac page opens settings on the tab the shell asks for');
});
