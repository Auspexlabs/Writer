// prefs.js — the settings the Settings window (MacSettings.dc.html) keeps in localStorage 'writer-settings',
// and the pure decisions the shell makes from them. Applying them to <html> (theme, accent, density, glass,
// motion, AI, dark pages) happens in the inline script at the top of index.dc.html so it runs before first paint.
export const KEY = 'writer-settings';
export const SESSION = 'writer-session';
export const MAC = 'writer-mac';

/** The same keys and defaults as DEF in MacSettings.dc.html (tests/prefs.test.mjs keeps the two in step). */
export const DEF = { iconStyle: 'b', quickTab: '开始', theme: 'light', startup: 'last', newType: 'docx', restore: true, autoUpdate: true, lang: '跟随系统', paper: 'A4', accent: '#3F7D5C', density: 'std', glass: 55, grid: true, motion: false, darkPages: false, font: '思源宋体', size: 12, md: true, spell: true, quote: true, track: false, ai: true, complete: true, grammar: true, tone: 'bal', preview: true, ctx: true, web: true, instr: '', autosave: true, interval: '30s', history: true, fmt_docx: '.docx', fmt_xlsx: '.xlsx', fmt_pptx: '.pptx', fmt_md: '.md' };

/** Puts the appearance settings on <html> (el) as attributes the stylesheets key on; the defaults add nothing, so the
 *  designed look is untouched: data-theme, data-density (compact|loose), data-motion=reduce (also when the system asks),
 *  data-ai=off, data-dark-pages, data-accent + --accent (non-default accents), data-glass + --gt = glass / 55. */
export function applyTo(el, p, matches = q => globalThis.matchMedia(q).matches) {
  const attr = (k, v) => v ? el.setAttribute(k, v) : el.removeAttribute(k);
  const prop = (k, v) => v ? el.style.setProperty(k, v) : el.style.removeProperty(k);
  el.setAttribute('data-theme', p.theme === 'auto' ? (matches('(prefers-color-scheme: dark)') ? 'dark' : 'light') : p.theme === 'dark' ? 'dark' : 'light');
  attr('data-density', p.density === 'compact' || p.density === 'loose' ? p.density : '');
  attr('data-motion', p.motion || matches('(prefers-reduced-motion: reduce)') ? 'reduce' : '');
  attr('data-ai', p.ai === false ? 'off' : '');
  attr('data-dark-pages', p.darkPages ? '1' : '');
  const accent = /^#[0-9a-f]{6}$/i.test(p.accent || '') && p.accent.toUpperCase() !== DEF.accent ? p.accent : '';
  attr('data-accent', accent ? '1' : ''); prop('--accent', accent);
  const g = Number(p.glass), gt = Number.isFinite(g) && g !== DEF.glass ? String(+(Math.max(0, Math.min(100, g)) / DEF.glass).toFixed(3)) : '';
  attr('data-glass', gt ? '1' : ''); prop('--gt', gt);
}

const read = (storage, key) => { try { return JSON.parse(storage.getItem(key) || 'null'); } catch (e) { return null; } };
/** The settings: the defaults, what 设置 saved, and, embedded in a site, what the site set (ui/embed.dc.html: __WRITER_PREFS). */
export function load(storage = globalThis.localStorage) { return { ...DEF, ...(storage && read(storage, KEY)), ...(globalThis.__WRITER_PREFS || {}) }; }
/** Changes one setting outside 设置 (the AI panel's 自动补全 switch): stored, and every page of this window told ('writer-settings'). */
export function set(k, v, storage = globalThis.localStorage) {
  try { storage.setItem(KEY, JSON.stringify({ ...(read(storage, KEY) || {}), [k]: v })); } catch (e) { }
  try { globalThis.dispatchEvent(new Event('writer-settings')); } catch (e) { }
}
export function loadSession(storage = globalThis.localStorage) { return (storage && read(storage, SESSION)) || {}; }
export function saveSession(s, storage = globalThis.localStorage) { try { storage.setItem(SESSION, JSON.stringify(s)); } catch (e) { } }

/** The left sidebar and assistant panel, the shell's own part of 'writer-mac' (mac.dc.html/win.dc.html keep the
 *  collapsed toolbar in the same key; this merges rather than replaces). Shared by every window and kept across
 *  launches (desktop/src-tauri), so Writer reopens showing what the last window left on screen. Presentation mode
 *  is never saved here. */
export function loadLayout(storage = globalThis.localStorage) { const m = (storage && read(storage, MAC)) || {}; return { showThumbs: m.thumbs === true, showAI: !!m.ai }; }
export function saveLayout(showThumbs, showAI, storage = globalThis.localStorage) { try { const m = (storage && read(storage, MAC)) || {}; storage.setItem(MAC, JSON.stringify({ ...m, thumbs: showThumbs, ai: showAI })); } catch (e) { } }

const TONE = { brief: 'Keep every reply to one or two short sentences.', detail: 'Reply in more detail: say what you changed, where, and why.' };
/** Extra system-prompt text for POST /chat: the reply style (回复风格) plus the user's own instructions (自定义说明). */
export function instructions(p) { return [TONE[p.tone], String(p.instr || '').trim()].filter(Boolean).join('\n'); }

/** The optional fields of a /chat request: instructions always when set; the selection only when the whole document is not
 *  to be sent; web only to turn 联网搜索 off (the engine defaults it on). */
export function chatOptions(p, selection) {
  const o = {}, ins = instructions(p);
  if (ins) o.instructions = ins;
  if (p.ctx === false) o.selection = String(selection || '').slice(0, 8000);
  if (p.web === false) o.web = false;
  return o;
}

/** A model a service offers: id is what the API takes, name what the menus show, note one line on what it is for; fast marks the one
 *  the autocomplete (AI 自动补全) uses by default: quick, cheap, and able to answer without thinking first. */
const m = (id, name, note, fast) => ({ id, name, note, fast: !!fast });

/** 设置 › AI: the model services on offer, the same as AiConfig.Providers in src/Writer.Cli/AiConfig.cs (tests/ai.test.mjs checks).
 *  url fills 接口地址; model is the service's recommended model, picked with it (Anthropic's is also the default when none is given);
 *  key: the service needs an API Key (local servers and custom endpoints may run without); login: a ChatGPT sign-in instead of a key
 *  (the Codex CLI's, src/Writer.Cli/ChatGpt.cs; its models are the ones a ChatGPT plan has in Codex, from OpenClaw's docs, checked
 *  2026-09-27, and what the plan lists once signed in); models: the menu's models, checked
 *  2026-09-26 on each service's model page (developers.openai.com/api/docs/models, platform.claude.com/docs/en/about-claude/models,
 *  api-docs.deepseek.com/quick_start/pricing, help.aliyun.com/zh/model-studio/models, platform.kimi.com/docs/pricing/chat,
 *  docs.bigmodel.cn/cn/guide/start/model-overview, the Volcengine Ark model list). A key adds what the service itself lists
 *  (POST /ai/models); local servers list only what they have. */
export const AI_PROVIDERS = [
  { id: 'anthropic', name: 'Anthropic Claude', url: 'https://api.anthropic.com', model: 'claude-sonnet-5', key: true, models: [
    m('claude-sonnet-5', 'Claude Sonnet 5', '速度和能力兼顾'), m('claude-opus-5-5', 'Claude Opus 5.5', '更强，适合长文和复杂修改'),
    m('claude-fable-5-1', 'Claude Fable 5.1', '推理最强，较慢'), m('claude-haiku-4-5', 'Claude Haiku 4.5', '最快，适合自动补全', true)] },
  { id: 'openai', name: 'OpenAI', url: 'https://api.openai.com/v1', model: 'gpt-6-sol', key: true, models: [
    m('gpt-6-sol', 'GPT-6 Sol', '能力和价格均衡'), m('gpt-6-astra', 'GPT-6 Astra', '最强，适合复杂任务'), m('gpt-6-luna', 'GPT-6 Luna', '最快最省，适合自动补全', true)] },
  { id: 'chatgpt', name: 'ChatGPT（账号登录）', url: 'https://chatgpt.com/backend-api/codex', model: 'gpt-6-astra', key: false, login: true, models: [
    m('gpt-6-astra', 'GPT-6 Astra', 'Codex 默认，最强'), m('gpt-5.6-terra', 'GPT-5.6 Terra', '能力和速度均衡'),
    m('gpt-5.6-luna', 'GPT-5.6 Luna', '最快、最省额度，适合自动补全', true), m('gpt-5.5', 'GPT-5.5', '上一代')] },
  { id: 'deepseek', name: 'DeepSeek', url: 'https://api.deepseek.com', model: 'deepseek-flash', key: true, models: [
    m('deepseek-flash', 'DeepSeek V4.1 Flash', '快而便宜，也适合自动补全', true), m('deepseek-v4-pro', 'DeepSeek V4 Pro', '更强')] },
  { id: 'qwen', name: '通义千问', url: 'https://dashscope.aliyuncs.com/compatible-mode/v1', model: 'qwen3.7-plus', key: true, models: [
    m('qwen3.7-plus', 'Qwen3.7 Plus', '能力和价格均衡'), m('qwen3.8-max', 'Qwen3.8 Max', '最强'), m('qwen3.8-flash', 'Qwen3.8 Flash', '最快最省，适合自动补全', true)] },
  { id: 'kimi', name: 'Kimi', url: 'https://api.moonshot.cn/v1', model: 'kimi-k3', key: true, models: [
    m('kimi-k3', 'Kimi K3', '最新旗舰，回答前总会先思考'), m('kimi-k2.6', 'Kimi K2.6', '上一代，可以不思考直接回答，适合自动补全', true),
    m('kimi-k2.7-code', 'Kimi K2.7 Code', '擅长代码'), m('kimi-k2.7-code-highspeed', 'Kimi K2.7 Code Highspeed', '擅长代码，更快')] },
  { id: 'glm', name: '智谱 GLM', url: 'https://open.bigmodel.cn/api/paas/v4', model: 'glm-5.3', key: true, models: [
    m('glm-5.3', 'GLM-5.3', '旗舰，回答前总会先思考'), m('glm-5-turbo', 'GLM-5 Turbo', '长任务稳定'), m('glm-4.7', 'GLM-4.7', '通用对话'),
    m('glm-4.7-flashx', 'GLM-4.7 FlashX', '轻量高速，适合自动补全', true), m('glm-4.7-flash', 'GLM-4.7 Flash', '免费')] },
  { id: 'doubao', name: '豆包', url: 'https://ark.cn-beijing.volces.com/api/v3', model: 'doubao-seed-2-1-pro-260628', key: true, models: [
    m('doubao-seed-2-1-pro-260628', 'Doubao Seed 2.1 Pro', '旗舰'), m('doubao-seed-2-1-turbo-260628', 'Doubao Seed 2.1 Turbo', '效果和成本均衡'),
    m('doubao-seed-2-0-lite-260428', 'Doubao Seed 2.0 Lite', '轻量'), m('doubao-seed-2-0-mini-260428', 'Doubao Seed 2.0 Mini', '延迟低、最省，适合自动补全', true)] },
  { id: 'ollama', name: 'Ollama（本机）', url: 'http://localhost:11434/v1', model: '', key: false, models: [] },
  { id: 'lmstudio', name: 'LM Studio（本机）', url: 'http://localhost:1234/v1', model: '', key: false, models: [] },
  { id: 'custom', name: '自定义（OpenAI 兼容）', url: '', model: '', key: false, models: [] }
];
export const aiPreset = id => AI_PROVIDERS.find(p => p.id === id) || AI_PROVIDERS[AI_PROVIDERS.length - 1];
const origin = u => { try { return new URL(u).origin; } catch (e) { return ''; } };

/** The 设置 › AI form from GET /ai. The key field starts empty: a stored key is only ever the flag hasKey, for the address savedUrl.
 *  Nothing saved yet: the service's recommended model and its fast one are chosen already. */
export function aiForm(cfg) {
  const c = cfg || {}, p = aiPreset(c.provider || 'anthropic'), fresh = !c.source || c.source === 'none';
  return { provider: p.id, baseUrl: c.baseUrl || p.url, model: c.model || (fresh ? p.model : ''), completeModel: c.completeModel || (fresh && !c.model ? fastModel(p.id) : ''), key: '', hasKey: !!c.hasKey, savedUrl: c.baseUrl || '', source: c.source || 'none', account: c.account || null };
}

const PLANS = { free: 'Free', go: 'Go', plus: 'Plus', pro: 'Pro', team: 'Business', business: 'Business', enterprise: 'Enterprise', edu: 'Edu' };
/** The ChatGPT sign-in as the settings show it: its email and plan (ada@example.com · Plus), '' when signed out. */
export const accountLine = a => !a ? '' : [a.email, PLANS[a.plan] || a.plan].filter(Boolean).join(' · ');

/** The model the autocomplete uses by default with a service: its fast one, else the assistant's (''). */
export const fastModel = id => (aiPreset(id).models.find(x => x.fast) || { id: '' }).id;

/** Choosing a service fills in its address, its recommended model and its fast model for the autocomplete, and starts a typed key over. */
export function pickProvider(form, id) {
  const p = aiPreset(id);
  return p.id === form.provider ? form : { ...form, provider: p.id, baseUrl: p.url, model: p.model, completeModel: fastModel(p.id), key: '' };
}

/** A model menu for a service: its own models (with their notes) first, then what the service lists (POST /ai/models) that is not
 *  among them, then the model in use when it is neither (a name typed under 其他模型). [{ id, name, note, fast, listed }] */
export function modelMenu(id, listed, current) {
  const own = aiPreset(id).models, seen = new Set(own.map(x => x.id));
  const more = (listed || []).filter(x => x && x.id && !seen.has(x.id) && seen.add(x.id)).map(x => ({ id: x.id, name: x.name || x.id, note: '', fast: false, listed: true }));
  const out = own.concat(more);
  if (current && !seen.has(current)) out.push({ id: current, name: current, note: '', fast: false, listed: false });
  return out;
}

/** What a service listed last time (POST /ai/models), kept per service and address so the AI panel's menu has it too. */
const LISTED = 'writer-ai-models';
export function listedModels(provider, baseUrl, storage = globalThis.localStorage) {
  const all = (storage && read(storage, LISTED)) || {}, x = all[provider + ' ' + baseUrl];
  return Array.isArray(x) ? x : [];
}
export function keepListed(provider, baseUrl, list, storage = globalThis.localStorage) {
  try { const all = read(storage, LISTED) || {}; all[provider + ' ' + baseUrl] = (list || []).map(x => x.name ? { id: x.id, name: x.name } : { id: x.id }); storage.setItem(LISTED, JSON.stringify(all)); } catch (e) { }
}

/** The name the AI panel shows for a model id: the menu's name when the service has it, else the id. */
export const modelName = (id, model) => { const x = aiPreset(id).models.find(y => y.id === model); return x ? x.name : model || ''; };

/** A stored key counts only while the address stays on its server: the engine never sends it anywhere else. */
export const keyStored = form => !!form.hasKey && !!origin(form.baseUrl) && origin(form.baseUrl) === origin(form.savedUrl);

/** The PUT /ai (and POST /ai/test) body: apiKey only when one was typed; left out, the engine keeps the stored key. */
export function aiRequest(form) {
  const p = aiPreset(form.provider), key = String(form.key || '').trim();
  const body = { provider: p.id, baseUrl: p.login ? p.url : String(form.baseUrl || '').trim().replace(/\/+$/, ''), model: String(form.model || '').trim() || (p.id === 'anthropic' ? p.model : ''), completeModel: String(form.completeModel || '').trim() };
  if (key && !p.login) body.apiKey = key;
  return body;
}

const httpUrl = u => { try { const x = new URL(u); return /^https?:$/.test(x.protocol) && !!x.host; } catch (e) { return false; } };

/** What the form still lacks before it can be saved, or ''. */
export function aiMissing(form) {
  const b = aiRequest(form);
  return !httpUrl(b.baseUrl) ? '请填写接口地址（http:// 或 https:// 开头）' : !b.model ? '请选择模型' : aiPreset(b.provider).login && !form.account ? '请先用 ChatGPT 登录' : '';
}

export const PAIRS = { '「': '」', '『': '』', '《': '》', '“': '”', '‘': '’', '（': '）', '【': '】' }; // i18n-ok
const CLOSERS = new Set(Object.values(PAIRS));
/** Smart CJK punctuation for text about to be typed (`data`) with `next` the character after the caret:
 *  { skip: true } — a closing mark typed over the same closing mark just moves the caret;
 *  { close } — an opening mark gets its partner inserted after the caret, unless the caret sits before a word;
 *  null — nothing to do. */
export function pairQuote(data, next) {
  const ch = String(data || '').slice(-1);
  if (!ch) return null;
  if (data.length === 1 && CLOSERS.has(ch) && ch === next) return { skip: true };
  if (PAIRS[ch] && (!next || /[\s\p{P}]/u.test(next))) return { close: PAIRS[ch] };
  return null;
}

/** What the shell shows at launch. files: the workspace, newest first ({ path }); session: { order, closed, last } saved last time.
 *  restore (退出时保留窗口) keeps the documents that were open, in their order (files new since then come first);
 *  startup (打开 Writer 时) picks what opens: 'last' the last active document, 'home' the create view, 'blank' a new document. */
export function launch(files, p, session) {
  const s = session || {};
  let docs = files;
  if (p.restore !== false) {
    const closed = new Set(s.closed || []), at = new Map((s.order || []).map((x, i) => [x, i]));
    docs = files.filter(f => !closed.has(f.path));
    const fresh = docs.filter(f => !at.has(f.path)), known = docs.filter(f => at.has(f.path)).sort((a, b) => at.get(a.path) - at.get(b.path));
    docs = [...fresh, ...known];
  }
  if (p.startup === 'home') return { docs, open: null, blank: null };
  if (p.startup === 'blank') return { docs, open: null, blank: p.newType || 'docx' };
  const last = p.restore !== false && s.last ? docs.find(f => f.path === s.last) : null;
  return { docs, open: last || docs[0] || null, blank: null };
}

/** The writer commands that apply the new-document settings to a file just created: paper size (纸张大小) and,
 *  when the engine has the property (caps.track), revision tracking (新文档默认开启修订). Word documents only. */
export function newFileArgs(path, type, p, caps) {
  if (type !== 'docx') return [];
  const props = [];
  if (p.paper && p.paper !== 'A4') props.push('page=' + p.paper);
  if (p.track && caps && caps.track) props.push('track=true');
  return props.length ? [['set', path, '/', ...props.flatMap(x => ['--prop', x])]] : [];
}
