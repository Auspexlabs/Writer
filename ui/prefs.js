// prefs.js — the settings the Settings window (MacSettings.dc.html) keeps in localStorage 'writer-settings',
// and the pure decisions the shell makes from them. Applying them to <html> (theme, accent, density, glass,
// motion, AI, dark pages) happens in the inline script at the top of index.dc.html so it runs before first paint.
export const KEY = 'writer-settings';
export const SESSION = 'writer-session';
export const MAC = 'writer-mac';

/** The same keys and defaults as DEF in MacSettings.dc.html (tests/prefs.test.mjs keeps the two in step). */
export const DEF = { iconStyle: 'b', quickTab: '开始', theme: 'light', startup: 'last', newType: 'docx', restore: true, autoUpdate: true, lang: '跟随系统', paper: 'A4', accent: '#3F7D5C', density: 'std', glass: 55, grid: true, motion: false, darkPages: false, font: '思源宋体', size: 12, md: true, spell: true, quote: true, track: false, ai: true, tone: 'bal', preview: true, ctx: true, web: true, instr: '', autosave: true, interval: '30s', history: true, fmt_docx: '.docx', fmt_xlsx: '.xlsx', fmt_pptx: '.pptx', fmt_md: '.md' };

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
export function load(storage = globalThis.localStorage) { return { ...DEF, ...(storage && read(storage, KEY)) }; }
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

/** 设置 › AI: the model services on offer, the same as AiConfig.Providers in src/Writer.Cli/AiConfig.cs (tests/ai.test.mjs checks).
 *  url fills 接口地址; model is the example the empty 模型 field shows (a model id from the provider's docs; Anthropic's is also the
 *  default when the field stays empty); key: the service needs an API Key (local servers and custom endpoints may run without). */
export const AI_PROVIDERS = [
  { id: 'anthropic', name: 'Anthropic Claude', url: 'https://api.anthropic.com', model: 'claude-sonnet-5', key: true },
  { id: 'openai', name: 'OpenAI', url: 'https://api.openai.com/v1', model: 'gpt-6-luna', key: true },
  { id: 'deepseek', name: 'DeepSeek', url: 'https://api.deepseek.com', model: 'deepseek-flash', key: true },
  { id: 'qwen', name: '通义千问', url: 'https://dashscope.aliyuncs.com/compatible-mode/v1', model: 'qwen3.7-plus', key: true },
  { id: 'kimi', name: 'Kimi', url: 'https://api.moonshot.cn/v1', model: 'kimi-k3', key: true },
  { id: 'glm', name: '智谱 GLM', url: 'https://open.bigmodel.cn/api/paas/v4', model: 'glm-5.3', key: true },
  { id: 'doubao', name: '豆包', url: 'https://ark.cn-beijing.volces.com/api/v3', model: 'doubao-seed-2-1-pro-260628', key: true },
  { id: 'ollama', name: 'Ollama（本机）', url: 'http://localhost:11434/v1', model: 'qwen3:8b', key: false },
  { id: 'lmstudio', name: 'LM Studio（本机）', url: 'http://localhost:1234/v1', model: '', key: false },
  { id: 'custom', name: '自定义（OpenAI 兼容）', url: '', model: '', key: false }
];
export const aiPreset = id => AI_PROVIDERS.find(p => p.id === id) || AI_PROVIDERS[AI_PROVIDERS.length - 1];
const origin = u => { try { return new URL(u).origin; } catch (e) { return ''; } };

/** The 设置 › AI form from GET /ai. The key field starts empty: a stored key is only ever the flag hasKey, for the address savedUrl. */
export function aiForm(cfg) {
  const c = cfg || {}, p = aiPreset(c.provider || 'anthropic');
  return { provider: p.id, baseUrl: c.baseUrl || p.url, model: c.model || '', key: '', hasKey: !!c.hasKey, savedUrl: c.baseUrl || '', source: c.source || 'none' };
}

/** Choosing a service fills in its address and starts the model and a typed key over. */
export function pickProvider(form, id) {
  const p = aiPreset(id);
  return p.id === form.provider ? form : { ...form, provider: p.id, baseUrl: p.url, model: '', key: '' };
}

/** A stored key counts only while the address stays on its server: the engine never sends it anywhere else. */
export const keyStored = form => !!form.hasKey && !!origin(form.baseUrl) && origin(form.baseUrl) === origin(form.savedUrl);

/** The PUT /ai (and POST /ai/test) body: apiKey only when one was typed; left out, the engine keeps the stored key. */
export function aiRequest(form) {
  const p = aiPreset(form.provider), key = String(form.key || '').trim();
  const body = { provider: p.id, baseUrl: String(form.baseUrl || '').trim().replace(/\/+$/, ''), model: String(form.model || '').trim() || (p.id === 'anthropic' ? p.model : '') };
  if (key) body.apiKey = key;
  return body;
}

const httpUrl = u => { try { const x = new URL(u); return /^https?:$/.test(x.protocol) && !!x.host; } catch (e) { return false; } };

/** What the form still lacks before it can be saved, or ''. */
export function aiMissing(form) {
  const b = aiRequest(form);
  return !httpUrl(b.baseUrl) ? '请填写接口地址（http:// 或 https:// 开头）' : !b.model ? '请填写模型' : '';
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
