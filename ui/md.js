// Minimal Markdown → HTML (GFM-ish): headings, emphasis, code, links, images, lists, tasks, tables, quotes, hr, math (KaTeX).
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const attr = s => esc(s).replace(/"/g, '&quot;');

// KaTeX, vendored under ./vendor/katex (no network at run time). It's a UMD build: a dynamic import() still runs it as
// plain script, so with no "exports"/"module" globals (real ESM, browser or Node) it falls through to the browser-global
// branch and assigns self.katex — and under Node's CJS interop it lands on the default export instead. Either way one
// of the two is the library.
let KX = null;
try { const m = await import('./vendor/katex/katex.min.js'); KX = (m && (m.default || globalThis.katex)) || null; } catch (e) { KX = null; }

/** LaTeX -> HTML for one formula (inline, or display when display is true). `trust` stays false: without it
 *  \href/\includegraphics/\url could smuggle an arbitrary (e.g. javascript:) URL or a foreign image into the page.
 *  `throwOnError` stays false: KaTeX then renders bad LaTeX as "$src$"-ish text in red with the parse error as its
 *  title, instead of throwing — one malformed formula must never take the rest of the document down with it. */
let MATHML = false; // mdToHtml(src, {mathml:true}): MathML only, for a standalone HTML export that carries no KaTeX CSS or fonts
export function renderMath(src, display) {
  const raw = () => esc(display ? '$$' + src + '$$' : '$' + src + '$');
  if (!KX) return raw();
  try { return KX.renderToString(src, { throwOnError: false, trust: false, displayMode: !!display, output: MATHML ? 'mathml' : 'htmlAndMathml' }); }
  catch (e) { return `<span class="katex-error" title="${attr(String((e && e.message) || e))}">${raw()}</span>`; }
}
/** KaTeX's parse error for a formula being typed, or '' when it renders — the math box shows it live under the preview. */
export function mathError(src) { if (!KX || !String(src).trim()) return ''; try { KX.renderToString(src, { throwOnError: true, trust: false, displayMode: true }); return ''; } catch (e) { return String((e && e.message) || e).replace(/^KaTeX parse error:\s*/, ''); } }
/** The exact markup mdToHtml would produce for one formula — the rich editor calls these again once the user closes
 *  the source box, so a re-render always matches what a fresh parse of the saved Markdown would show. */
export const mathBlockHtml = src => `<div class="md-math-block" contenteditable="false" data-src="${attr(src)}">${renderMath(src, true)}</div>`;
export const mathInlineHtml = src => `<span class="md-math" contenteditable="false" data-src="${attr(src)}">${renderMath(src, false)}</span>`;

// highlight.js (the ES build of its "common" bundle, ~36 languages plus their aliases: js, py, sh, yml…), vendored under
// ./vendor/highlight. Loaded on demand: mdToHtml leaves code blocks plain, the editor calls highlight() on each once loaded.
let HL = null, HLP = null;
export function loadHljs() { return HLP || (HLP = import('./vendor/highlight/highlight.min.js').then(m => (HL = (m && m.default) || null)).catch(() => (HL = null))); }
/** Highlighted HTML for one code block, or null when the library isn't loaded (yet), lang is empty or unknown to it. */
export function highlight(src, lang) {
  if (!HL || !lang || !HL.getLanguage(lang)) return null;
  try { return HL.highlight(src, { language: lang, ignoreIllegals: true }).value; } catch (e) { return null; }
}

// Mermaid diagrams (```mermaid fences). The library is 2.7 MB, so like PDF.js it comes from jsdelivr the first time a document
// needs a diagram and is never bundled; a document without one never loads it. MIT, Knut Sveidqvist and contributors.
const MERMAID = 'https://cdn.jsdelivr.net/npm/mermaid@11.12.0/dist/mermaid.esm.min.mjs';
let MM = null, mmSeq = 0;
/** The diagram block: its source shown as code until the editor draws it; data-src is what saves back inside the fence. */
export const mermaidBlockHtml = (src, extra = '') => `<div class="md-mermaid"${extra} contenteditable="false" data-src="${attr(src)}"><pre><code data-lang="mermaid">${esc(src)}</code></pre></div>`;
/** The SVG for one diagram's source; rejects with mermaid's own message for bad source. securityLevel 'strict' (its default) keeps any HTML in labels as text. */
export async function renderMermaid(src) {
  if (!MM) MM = import(MERMAID).then(m => { const mm = m.default || m; mm.initialize({ startOnLoad: false, securityLevel: 'strict', theme: typeof document !== 'undefined' && document.documentElement.dataset.theme === 'dark' ? 'dark' : 'neutral', fontFamily: "'IBM Plex Sans','Noto Sans SC',sans-serif" }); return mm; }).catch(e => { MM = null; throw e; });
  const mm = await MM, { svg } = await mm.render('mmd' + (++mmSeq), src);
  return svg;
}

let SAFE = false; // set for the run by mdToHtml(src, {safe:true}) — the AI chat panel, where the text is untrusted
const schemeOf = u => { const m = /^([a-z][a-z0-9+.-]*):/i.exec(String(u).trim()); return m ? m[1].toLowerCase() : ''; };
const _t = zh => (globalThis.$t ? globalThis.$t(zh) : String(zh).replace(/@@.*$/, '')); // ui/i18n.js in the page; the Chinese itself under node
let FN = {}, FN_ORDER = [], IDS = [], TOC = ''; // per mdToHtml run: the footnote definitions and the order they are first cited, the headings' anchor ids, the [TOC] markup
const ALERT = { note: '注意@@alert', tip: '提示@@alert', important: '重要@@alert', warning: '警告@@alert', caution: '小心@@alert' };
/** A heading's anchor id, as GitHub makes them: lower case, punctuation dropped, spaces to hyphens; CJK stays as it is. */
export const slug = text => String(text || '').replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/[*_`~]/g, '').trim().toLowerCase().replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s+/g, '-');
/** The [TOC] block's markup for the document's headings — mdToHtml at render time, the editor again whenever the text changes. */
export const tocHtml = heads => '<ul>' + heads.map(h => `<li style="padding-left:${(h.lvl - 1) * 14}px"><a href="#${attr(h.id)}">${esc(h.text)}</a></li>`).join('') + '</ul>';
/** The YAML front matter card; its source goes back out between --- lines. */
export const frontHtml = src => `<div class="md-front" contenteditable="false" data-src="${attr(src)}"><pre>${esc(src)}</pre></div>`;
/** GitHub's emoji shortcodes, the common ones: :smile: renders as 😄 and saves back as :smile:. */
const EMOJI = { smile: '😄', smiley: '😃', grinning: '😀', laughing: '😆', joy: '😂', rofl: '🤣', wink: '😉', blush: '😊', heart_eyes: '😍', kissing_heart: '😘', thinking: '🤔', neutral_face: '😐', unamused: '😒', sweat_smile: '😅', sweat: '😓', cry: '😢', sob: '😭', angry: '😠', rage: '😡', scream: '😱', fearful: '😨', flushed: '😳', sleeping: '😴', mask: '😷', sunglasses: '😎', nerd_face: '🤓', innocent: '😇', smirk: '😏', relieved: '😌', yum: '😋', stuck_out_tongue: '😛', confused: '😕', disappointed: '😞', worried: '😟', open_mouth: '😮', zipper_mouth_face: '🤐', skull: '💀', ghost: '👻', alien: '👽', robot: '🤖', poop: '💩', clown_face: '🤡',
  heart: '❤️', broken_heart: '💔', two_hearts: '💕', sparkling_heart: '💖', orange_heart: '🧡', yellow_heart: '💛', green_heart: '💚', blue_heart: '💙', purple_heart: '💜', black_heart: '🖤', white_heart: '🤍', '+1': '👍', thumbsup: '👍', '-1': '👎', thumbsdown: '👎', ok_hand: '👌', wave: '👋', clap: '👏', pray: '🙏', muscle: '💪', point_up: '☝️', point_down: '👇', point_left: '👈', point_right: '👉', raised_hands: '🙌', handshake: '🤝', v: '✌️', fist: '✊', eyes: '👀', brain: '🧠', see_no_evil: '🙈', hear_no_evil: '🙉', speak_no_evil: '🙊',
  dog: '🐶', cat: '🐱', mouse: '🐭', rabbit: '🐰', fox_face: '🦊', bear: '🐻', panda_face: '🐼', koala: '🐨', tiger: '🐯', lion: '🦁', cow: '🐮', pig: '🐷', frog: '🐸', monkey: '🐒', chicken: '🐔', penguin: '🐧', bird: '🐦', bee: '🐝', bug: '🐛', butterfly: '🦋', snail: '🐌', turtle: '🐢', snake: '🐍', dragon: '🐉', whale: '🐳', dolphin: '🐬', fish: '🐟', octopus: '🐙', unicorn: '🦄',
  sunny: '☀️', cloud: '☁️', umbrella: '☔', zap: '⚡', snowflake: '❄️', fire: '🔥', droplet: '💧', ocean: '🌊', rainbow: '🌈', star: '⭐', star2: '🌟', sparkles: '✨', boom: '💥', dizzy: '💫', moon: '🌙', earth_asia: '🌏', earth_americas: '🌎', earth_africa: '🌍', seedling: '🌱', herb: '🌿', four_leaf_clover: '🍀', cactus: '🌵', palm_tree: '🌴', evergreen_tree: '🌲', deciduous_tree: '🌳', maple_leaf: '🍁', fallen_leaf: '🍂', cherry_blossom: '🌸', rose: '🌹', sunflower: '🌻', tulip: '🌷', bouquet: '💐',
  apple: '🍎', green_apple: '🍏', banana: '🍌', grapes: '🍇', strawberry: '🍓', watermelon: '🍉', lemon: '🍋', peach: '🍑', cherries: '🍒', pizza: '🍕', hamburger: '🍔', fries: '🍟', taco: '🌮', sushi: '🍣', ramen: '🍜', rice: '🍚', bento: '🍱', dumpling: '🥟', cake: '🍰', birthday: '🎂', cookie: '🍪', doughnut: '🍩', ice_cream: '🍨', candy: '🍬', chocolate_bar: '🍫', coffee: '☕', tea: '🍵', beer: '🍺', beers: '🍻', wine_glass: '🍷', cocktail: '🍸', tropical_drink: '🍹', champagne: '🍾',
  soccer: '⚽', basketball: '🏀', football: '🏈', baseball: '⚾', tennis: '🎾', trophy: '🏆', medal_sports: '🏅', first_place_medal: '🥇', second_place_medal: '🥈', third_place_medal: '🥉', dart: '🎯', video_game: '🎮', game_die: '🎲', guitar: '🎸', musical_note: '🎵', notes: '🎶', microphone: '🎤', headphones: '🎧', art: '🎨', clapper: '🎬', tada: '🎉', confetti_ball: '🎊', balloon: '🎈', gift: '🎁', ribbon: '🎀', christmas_tree: '🎄', jack_o_lantern: '🎃', red_envelope: '🧧', lantern: '🏮',
  car: '🚗', taxi: '🚕', bus: '🚌', truck: '🚚', bike: '🚲', rocket: '🚀', airplane: '✈️', helicopter: '🚁', ship: '🚢', train: '🚋', bullettrain_side: '🚄', house: '🏠', office: '🏢', school: '🏫', hospital: '🏥', bank: '🏦', hotel: '🏨', mountain: '⛰️', volcano: '🌋', camping: '🏕️', beach_umbrella: '🏖️', city_sunset: '🌇', world_map: '🗺️',
  watch: '⌚', iphone: '📱', computer: '💻', desktop_computer: '🖥️', keyboard: '⌨️', printer: '🖨️', camera: '📷', video_camera: '📹', tv: '📺', radio: '📻', telephone: '☎️', battery: '🔋', electric_plug: '🔌', bulb: '💡', flashlight: '🔦', candle: '🕯️', moneybag: '💰', dollar: '💵', yen: '💴', credit_card: '💳', gem: '💎', wrench: '🔧', hammer: '🔨', gear: '⚙️', link: '🔗', paperclip: '📎', scissors: '✂️', pencil2: '✏️', pen: '🖊️', memo: '📝', book: '📖', books: '📚', bookmark: '🔖', newspaper: '📰', page_facing_up: '📄', clipboard: '📋', calendar: '📅', date: '📅', chart_with_upwards_trend: '📈', chart_with_downwards_trend: '📉', bar_chart: '📊', pushpin: '📌', round_pushpin: '📍', triangular_flag_on_post: '🚩', lock: '🔒', unlock: '🔓', key: '🔑', bell: '🔔', mag: '🔍', hourglass: '⌛', alarm_clock: '⏰', email: '📧', envelope: '✉️', inbox_tray: '📥', outbox_tray: '📤', package: '📦', label: '🏷️',
  white_check_mark: '✅', heavy_check_mark: '✔️', x: '❌', question: '❓', exclamation: '❗', warning: '⚠️', no_entry: '⛔', no_entry_sign: '🚫', recycle: '♻️', 100: '💯', arrow_right: '➡️', arrow_left: '⬅️', arrow_up: '⬆️', arrow_down: '⬇️', arrows_counterclockwise: '🔄', heavy_plus_sign: '➕', heavy_minus_sign: '➖', copyright: '©️', registered: '®️', tm: '™️', information_source: 'ℹ️', bangbang: '‼️', red_circle: '🔴', large_blue_circle: '🔵', white_circle: '⚪', black_circle: '⚫', small_red_triangle: '🔺', checkered_flag: '🏁', cn: '🇨🇳', us: '🇺🇸', jp: '🇯🇵', gb: '🇬🇧', de: '🇩🇪', fr: '🇫🇷', kr: '🇰🇷' };
export const emojiHtml = name => `<span class="md-emoji" contenteditable="false" data-src=":${attr(name)}:">${EMOJI[name]}</span>`;
/** Smart punctuation while typing (the 视图 › 智能标点 switch, off by default). `before` is the paragraph's text up to and including
 *  the character just typed; the answer is what to put in its place — { del: characters to take back, text } — or null. " → “ or ”
 *  (opening after nothing, a space, a bracket or a dash), ' → ‘ or ’, -- → –, –- → —, ... → …, -> → →, <- → ←. A paragraph that is
 *  only dashes so far is left alone: that is a rule (---) being typed. */
export function smartPunct(before) {
  const c = before.slice(-1), prev = before.slice(0, -1), open = !prev || /[\s([{“‘—–-]$/.test(prev);
  if (c === '"') return { del: 1, text: open ? '“' : '”' };
  if (c === "'") return { del: 1, text: open ? '‘' : '’' };
  if (/^[-–—]*$/.test(prev)) return null;
  if (c === '-') return prev.endsWith('<') ? { del: 2, text: '←' } : prev.endsWith('–') ? { del: 2, text: '—' } : prev.endsWith('-') ? { del: 2, text: '–' } : null;
  if (c === '>' && prev.endsWith('-')) return { del: 2, text: '→' };
  if (c === '.' && prev.endsWith('..')) return { del: 3, text: '…' };
  return null;
}
/** "---" on the first line up to the next "---" line is YAML front matter: { front, body: the lines after it, base: the lines it took }. */
function splitFront(lines) { const end = lines[0] === '---' ? lines.indexOf('---', 1) : -1; return end < 1 ? { front: null, body: lines, base: 0 } : { front: lines.slice(1, end).join('\n'), body: lines.slice(end + 1), base: end + 1 }; }

export function inline(s) {
  const codes = [], maths = [];
  s = s.replace(/`([^`\n]+)`/g, (m, c) => { codes.push(c); return '\u0000' + (codes.length - 1) + '\u0000'; });
  // math is pulled out (and rendered to trusted HTML) before escaping, like code spans above — its LaTeX may contain
  // < > & that must reach KaTeX untouched, and the HTML KaTeX hands back must not be escaped a second time
  // (with the data-src wrapper, so saving the file writes $…$ again rather than the rendered text)
  s = s.replace(/\$(?=\S)([^$\n]*?\S)\$/g, (m, expr) => { maths.push(mathInlineHtml(expr)); return '\u0001' + (maths.length - 1) + '\u0001'; });
  // esc leaves quotes alone, so every value that goes inside an attribute gets q: a " in a URL or alt text would
  // otherwise close the attribute and open a new one (onmouseover=…), in the chat panel and in any .md file opened
  const q = v => v.replace(/"/g, '&quot;');
  // a picture with a size, as Typora writes one it resized: <img src="…" alt="…" width="240"> — the one html tag read here, only from
  // a file (never a chat reply), rebuilt from its src, alt, width and height; anything else stays text
  const imgs = [];
  if (!SAFE) s = s.replace(/<img\s+([^<>]*?)\/?>/gi, (m, a) => {
    const at = k => { const x = new RegExp('(?:^|\\s)' + k + '\\s*=\\s*"([^"]*)"', 'i').exec(a); return x ? x[1] : ''; };
    const u = at('src'), sc = schemeOf(u), w = at('width'), h = at('height');
    if (!(sc === 'http' || sc === 'https' || sc === 'data')) return m;
    imgs.push(`<img alt="${q(at('alt'))}" src="${q(u)}"${/^\d+$/.test(w) ? ` width="${w}"` : ''}${/^\d+$/.test(h) ? ` height="${h}"` : ''}>`);
    return '\u0002' + (imgs.length - 1) + '\u0002';
  });
  s = esc(s);
  // a footnote mark [^id] for a definition the document has: numbered in citation order, the note's text on hover, a link down to it
  s = s.replace(/\[\^([^\]\s]+)\]/g, (m, id) => { if (!(id in FN)) return m; let n = FN_ORDER.indexOf(id); if (n < 0) n = FN_ORDER.push(id) - 1; return `<sup class="md-fn" contenteditable="false" data-fn="${q(id)}" title="${attr(FN[id].replace(/[*_`~]/g, ''))}"><a href="#fn-${q(id)}">${n + 1}</a></sup>`; });
  if (!SAFE) s = s.replace(/\[\[([^\]|\n]+)(?:\|([^\]\n]+))?\]\]/g, (m, t, l) => `<a href="${q(t.trim())}" data-wiki="1">${(l || t).trim()}</a>`); // [[Page]] / [[Page|label]]: a wiki link, the user's own files only
  s = s.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+"([^"]*)")?\)/g, (m, a, u) => {
    const sc = schemeOf(u); // http(s)/data cover every real use (remote images, pasted/local images); anything else (e.g. javascript:) is dropped
    return (!sc || sc === 'http' || sc === 'https' || sc === 'data') ? `<img alt="${q(a)}" src="${q(u)}">` : a;
  });
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (m, label, u) => {
    const sc = schemeOf(u); // web and mail links anywhere; a relative link only in the editor; never javascript: and the like
    return (sc ? ['http', 'https', 'mailto'].includes(sc) : !SAFE) ? `<a href="${q(u)}" target="_blank" rel="noopener">${label}</a>` : label;
  });
  s = s.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, (m, u) => `<a href="${q(u)}" data-auto="angle" target="_blank" rel="noopener">${u}</a>`);
  // a bare URL becomes a link (data-auto, so it saves back bare); links, pictures and attribute values already there are skipped
  s = s.replace(/(<a\b[^>]*>[\s\S]*?<\/a>|<img\b[^>]*>|="[^"]*")|(https?:\/\/[^\s<]+)/g, (m, tag, url) => {
    if (tag) return tag;
    let u = url, tail = (/[.,;:!?'"\])]+$/.exec(u) || [''])[0]; u = u.slice(0, u.length - tail.length);
    if (tail.startsWith(')') && (u.match(/\(/g) || []).length > (u.match(/\)/g) || []).length) { u += ')'; tail = tail.slice(1); } // a Wikipedia-style (…) inside the URL keeps its bracket
    return `<a href="${q(u)}" data-auto="bare" target="_blank" rel="noopener">${u}</a>${tail}`;
  });
  // the markup made so far steps out of the emphasis passes' way: a _ inside a URL or target="_blank" is not emphasis, nor is ~ or ^ in a href
  const tags = []; s = s.replace(/<\/?a\b[^>]*>|<img\b[^>]*>|<sup class="md-fn"[\s\S]*?<\/sup>/g, m => { tags.push(m); return '\u0003' + (tags.length - 1) + '\u0003'; });
  s = s.replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>').replace(/__(?=\S)([\s\S]*?\S)__/g, '<strong>$1</strong>');
  s = s.replace(/(^|[^*])\*(?=\S)([^*\n]*?\S)\*(?!\*)/g, '$1<em>$2</em>').replace(/(^|[^\w])_(?=\S)([^_\n]*?\S)_(?!\w)/g, '$1<em>$2</em>');
  s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<del>$1</del>').replace(/==(?=\S)([\s\S]*?\S)==/g, '<mark>$1</mark>');
  s = s.replace(/(^|[^~])~(?=\S)([^~\s]+?)~(?!~)/g, '$1<sub>$2</sub>').replace(/\^(?=\S)([^^\s]+?)\^/g, '<sup>$1</sup>'); // H~2~O, x^2^
  s = s.replace(/(<[^>]*>)|:([a-z0-9_+-]+):/g, (m, tag, name) => tag ? tag : (name in EMOJI ? emojiHtml(name) : m));
  s = s.replace(/ {2,}\n/g, '<br>').replace(/\n/g, ' ').replace(/&lt;br\s*\/?&gt;/gi, '<br>'); // a literal <br> is how GFM breaks a line inside a table cell
  s = s.replace(/\u0003(\d+)\u0003/g, (m, i) => tags[+i]);
  s = s.replace(/\u0000(\d+)\u0000/g, (m, i) => '<code>' + esc(codes[+i]) + '</code>').replace(/\u0002(\d+)\u0002/g, (m, i) => imgs[+i]);
  return s.replace(/\u0001(\d+)\u0001/g, (m, i) => maths[+i]);
}
const RE = {
  fence: /^\s*(```|~~~)\s*([\w+-]*)\s*$/, mathfence: /^\s*\$\$\s*$/, head: /^(#{1,6})\s+(.*?)\s*#*\s*$/, hr: /^\s*([-*_])(\s*\1){2,}\s*$/,
  quote: /^\s*>\s?(.*)$/, list: /^(\s*)([-*+]|\d+[.)])\s+(.*)$/, tsep: /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/,
  fndef: /^\[\^([^\]\s]+)\]:\s?(.*)$/, toc: /^\s*\[\[?toc\]\]?\s*$/i
};
const splitRow = l => l.trim().replace(/^\||\|$/g, '').split(/(?<!\\)\|/).map(c => c.trim().replace(/\\\|/g, '|'));
let hid = 0;
function blocks(lines, base, top) {
  let out = '', i = 0;
  const L = n => top ? ` data-line="${base + n}"` : '';
  while (i < lines.length) {
    const ln = lines[i];
    if (!ln.trim()) { i++; continue; }
    let m;
    if ((m = RE.fence.exec(ln))) { const start = i, fence = m[1], lang = m[2]; const body = []; i++; while (i < lines.length && !lines[i].trim().startsWith(fence)) body.push(lines[i++]); i++; out += lang === 'mermaid' && !SAFE ? mermaidBlockHtml(body.join('\n'), L(start)) : `<pre${L(start)}><code${lang ? ` data-lang="${attr(lang)}"` : ''}>${esc(body.join('\n'))}</code></pre>`; continue; }
    if (RE.mathfence.test(ln)) { const start = i; const body = []; i++; while (i < lines.length && !RE.mathfence.test(lines[i])) body.push(lines[i++]); i++; const src = body.join('\n'); out += `<div class="md-math-block"${L(start)} contenteditable="false" data-src="${attr(src)}">${renderMath(src, true)}</div>`; continue; }
    if ((m = RE.head.exec(ln))) { const n = m[1].length, id = top && IDS[hid] ? IDS[hid] : 'h-' + hid; hid++; out += `<h${n}${L(i)} id="${attr(id)}">${inline(m[2])}</h${n}>`; i++; continue; } // the anchor ids headings(src) gives (top level, in order; a heading inside a quote is not in that list)
    if (RE.hr.test(ln)) { out += `<hr${L(i)}>`; i++; continue; }
    if (RE.toc.test(ln)) { out += `<nav class="md-toc"${L(i)} contenteditable="false">${TOC}</nav>`; i++; continue; }
    if (RE.fndef.test(ln)) { // footnote definitions, consecutive ones as one block where they stand; an indented line continues the note above
      const start = i, items = []; let d;
      while (i < lines.length && (d = RE.fndef.exec(lines[i]))) { let text = d[2]; i++; while (i < lines.length && /^\s+\S/.test(lines[i])) text += '\n' + lines[i++].trim(); items.push([d[1], text]); }
      out += `<section class="md-footnotes"${L(start)}><ol>${items.map(([id, t]) => `<li id="fn-${attr(id)}" data-fn="${attr(id)}">${inline(t)}</li>`).join('')}</ol></section>`; continue;
    }
    if (RE.quote.test(ln)) {
      const start = i, body = []; while (i < lines.length && (RE.quote.test(lines[i]) || (lines[i].trim() && body.length && !RE.list.test(lines[i]) && !RE.head.test(lines[i])))) { const q = RE.quote.exec(lines[i]); body.push(q ? q[1] : lines[i]); i++; }
      const am = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$/i.exec(body[0] || ''); // GitHub's alerts: > [!NOTE] on the quote's first line
      if (am) { const kind = am[1].toLowerCase(); out += `<blockquote${L(start)} class="md-alert md-alert-${kind}" data-alert="${am[1].toUpperCase()}"><p class="md-alert-title" contenteditable="false">${esc(_t(ALERT[kind]))}</p>${blocks(body.slice(1), 0, false)}</blockquote>`; }
      else out += `<blockquote${L(start)}>${blocks(body, 0, false)}</blockquote>`;
      continue;
    }
    if (ln.includes('|') && i + 1 < lines.length && RE.tsep.test(lines[i + 1])) {
      const start = i, head = splitRow(ln), al = splitRow(lines[i + 1]).map(c => c.startsWith(':') && c.endsWith(':') ? 'center' : c.endsWith(':') ? 'right' : c.startsWith(':') ? 'left' : ''); i += 2;
      const rows = []; while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(splitRow(lines[i++]));
      const td = (t, c, j) => `<${t}${al[j] ? ` style="text-align:${al[j]}"` : ''}>${inline(c || '')}</${t}>`;
      out += `<table${L(start)}><thead><tr>${head.map((c, j) => td('th', c, j)).join('')}</tr></thead><tbody>${rows.map(r => '<tr>' + head.map((_, j) => td('td', r[j], j)).join('') + '</tr>').join('')}</tbody></table>`;
      continue;
    }
    if (RE.list.test(ln)) {
      const items = []; const start = i;
      while (i < lines.length) {
        const l = lines[i]; const lm = RE.list.exec(l);
        if (lm) { const it = { ind: lm[1].replace(/\t/g, '    ').length, ord: /\d/.test(lm[2]), num: parseInt(lm[2]) || 1, text: lm[3], line: base + i }; if (items.length && it.ind <= items[0].ind && it.ord !== items[0].ord) break; items.push(it); i++; continue; } // a numbered list right after a bulleted one (or the reverse) is its own list
        if (l.trim() && /^\s+/.test(l) && items.length) { items[items.length - 1].text += '\n' + l.trim(); i++; continue; }
        if (!l.trim() && i + 1 < lines.length && RE.list.test(lines[i + 1])) { i++; continue; }
        break;
      }
      const build = (arr, lvl) => {
        if (!arr.length) return '';
        const ord = arr[0].ord; let h = `<${ord ? 'ol' : 'ul'}${lvl === 0 ? L(start) : ''}${ord && arr[0].num !== 1 ? ` start="${arr[0].num}"` : ''}>`;
        let k = 0;
        while (k < arr.length) {
          const it = arr[k], kids = []; k++;
          while (k < arr.length && arr[k].ind > it.ind) kids.push(arr[k++]);
          const tm = /^\[([ xX])\]\s+(.*)$/s.exec(it.text);
          const body = tm ? `<input type="checkbox" data-task="${it.line}"${tm[1] !== ' ' ? ' checked' : ''}> <span>${inline(tm[2])}</span>` : inline(it.text);
          h += `<li${tm ? ' class="task"' : ''} data-li="${it.line}">${body}${build(kids, lvl + 1)}</li>`;
        }
        return h + `</${ord ? 'ol' : 'ul'}>`;
      };
      out += build(items, 0); continue;
    }
    const start = i, para = [];
    while (i < lines.length && lines[i].trim() && !RE.fence.test(lines[i]) && !RE.head.test(lines[i]) && !RE.hr.test(lines[i]) && !RE.quote.test(lines[i]) && !RE.list.test(lines[i]) && !RE.mathfence.test(lines[i]) && !RE.fndef.test(lines[i]) && !RE.toc.test(lines[i]) && !(lines[i].includes('|') && i + 1 < lines.length && RE.tsep.test(lines[i + 1]))) para.push(lines[i++]);
    out += `<p${L(start)}>${inline(para.join('\n'))}</p>`;
  }
  return out;
}
export function mdToHtml(src, opts) {
  hid = 0; SAFE = !!(opts && opts.safe); MATHML = !!(opts && opts.mathml); FN = {}; FN_ORDER = [];
  const { front, body, base } = splitFront(String(src || '').replace(/\r\n?/g, '\n').split('\n'));
  for (const l of body) { const m = RE.fndef.exec(l); if (m && !(m[1] in FN)) FN[m[1]] = m[2]; } // known before any mark renders: its number and hover text
  const heads = headings(src); IDS = heads.map(h => h.id); TOC = tocHtml(heads);
  return (front != null ? frontHtml(front) : '') + blocks(body, base, true);
}
/** The document's headings (outside fences and front matter) with their anchor ids: slug(text), a repeat getting -1, -2… as on GitHub. */
export function headings(src) {
  const out = [], seen = {}; let inFence = false;
  const { body, base } = splitFront(String(src || '').replace(/\r\n?/g, '\n').split('\n'));
  body.forEach((l, i) => {
    if (RE.fence.test(l)) inFence = !inFence; if (inFence) return; const m = RE.head.exec(l); if (!m) return;
    let id = slug(m[2]) || 'h'; if (seen[id] == null) seen[id] = 0; else id += '-' + (++seen[id]);
    out.push({ lvl: m[1].length, text: m[2].replace(/[*_`~]/g, ''), line: base + i, id });
  });
  return out;
}
export function wordStats(src) {
  const t = String(src || '').replace(/```[\s\S]*?```/g, '').replace(/[#>*_`~\-|[\]()!]/g, ' ');
  const cjk = (t.match(/[一-龥]/g) || []).length, en = (t.replace(/[一-龥]/g, ' ').match(/[A-Za-z0-9]+/g) || []).length;
  return { words: cjk + en, chars: t.replace(/\s/g, '').length, minutes: Math.max(1, Math.round((cjk + en) / 400)) };
}

// ---- live "Typora-style" shortcuts (MarkdownEditor's rich view): pure text matching, so the DOM glue that turns a
// match into real formatting can stay a thin wrapper, and the matching itself is testable without a browser. ----

/** The inline shortcut that was just completed at the end of `before` (the current paragraph's text from its start up
 *  to the caret), if any. Longer/more specific markers are tried first so e.g. "**x**" doesn't read as "*" + literal.
 *  For the two markers that need a non-word/non-* character before them (so `*`/`_` don't fire inside `snake_case` or
 *  `a*b`), that guard character rides along in the match; `len` strips it back off before the caller replaces text. */
export function inlineTrigger(before) {
  const P = [
    ['**', /\*\*(?=\S)([^*\n]*?\S)\*\*$/], ['__', /__(?=\S)([^_\n]*?\S)__$/],
    ['~~', /~~(?=\S)([^~\n]*?\S)~~$/], ['==', /==(?=\S)([^=\n]*?\S)==$/],
    ['`', /`(?=\S)([^`\n]*?\S)`$/], ['$', /\$(?=\S)([^$\n]*?\S)\$$/],
    ['*', /(?:^|[^*])\*(?=\S)([^*\n]*?\S)\*$/], ['_', /(?:^|[^\w])_(?=\S)([^_\n]*?\S)_$/],
    [':', /:([a-z0-9_+-]+):$/]
  ];
  const TAG = { '**': 'strong', '__': 'strong', '~~': 'del', '==': 'mark', '`': 'code', $: 'math', '*': 'em', _: 'em', ':': 'emoji' };
  for (const [marker, re] of P) {
    const m = re.exec(before); if (!m) continue;
    if (marker === ':' && !(m[1] in EMOJI)) continue; // :something: that is not a shortcode is just text
    const full = m[0].startsWith(marker) ? m[0] : m[0].slice(1); // drop the borrowed guard character, if any
    return { tag: TAG[marker], text: m[1], len: full.length };
  }
  return null;
}
/** The block shortcut for a paragraph whose entire text so far (start to caret) is `before`, just after its trailing
 *  trigger character (a space, normally). `len` is how much of `before` is the marker, to strip before applying it. */
export function blockTrigger(before) {
  before = before.replace(/\u00a0/g, ' '); // a space typed at the end of a text node arrives as a no-break space
  let m;
  if ((m = /^(#{1,6}) $/.exec(before))) return { type: 'h' + m[1].length, len: m[0].length };
  if (/^[-*+] $/.test(before)) return { type: 'ul', len: 2 };
  if (/^1[.)] $/.test(before)) return { type: 'ol', len: 3 };
  if (/^> $/.test(before)) return { type: 'quote', len: 2 };
  if (/^\[ ?\] $/.test(before)) return { type: 'task', len: before.length };
  return null;
}
export const MATH_BLOCK_MARK = '$$'; // a paragraph whose entire text is exactly this opens the math block editor

// ---- pasting from other apps ----
/** Plain clipboard text as Markdown when it reads like some: the bullets other apps use (• ◦ ▪ ‣ ·) become "- ", so a list from
 *  Pages, Notion or a chat pastes as a list. Null for plain prose, where the browser's own paste is right. */
export function pasteText(t) {
  const s = String(t || '').replace(/^([ \t]*)[•◦▪‣·]\s+/gm, '$1- ');
  return /^[ \t]*(#{1,6} |[-*+] |\d+[.)] |>|```|\|)/m.test(s) ? s : null;
}
/** Word's HTML lists — <p class="MsoListParagraph…" style="…mso-list:l0 level2 lfo1"> with the bullet or number in an
 *  [if !supportLists] span — as real <ul>/<ol>, nested by level ("1." / "a." / "i." marks make an <ol>). Other HTML passes through. */
export function wordLists(h) {
  const P = /<p\b([^>]*class="[^"]*MsoList[^"]*"[^>]*)>([\s\S]*?)<\/p>/gi, MARK = /<!--\[if !supportLists\]-->([\s\S]*?)<!--\[endif\]-->|<span\b[^>]*mso-list:\s*Ignore[^>]*>([\s\S]*?)<\/span>/i;
  let out = '', last = 0, m; const stack = [];
  const close = to => { while (stack.length > to) out += `</li></${stack.pop()}>`; };
  while ((m = P.exec(h))) {
    const seg = h.slice(last, m.index); if (seg.trim()) close(0); out += seg; last = P.lastIndex;
    const lvl = +((/level(\d+)/i.exec(m[1]) || [])[1] || 1), mk = MARK.exec(m[2]), mark = mk ? (mk[1] || mk[2] || '').replace(/<[^>]*>|&nbsp;/g, ' ').trim() : '';
    const tag = /^\w{1,3}[.)]$/.test(mark) ? 'ol' : 'ul', body = m[2].replace(MARK, '');
    if (lvl > stack.length) { while (stack.length < lvl) { if (stack.length && out.endsWith('</li>')) out = out.slice(0, -5); out += `<${tag}>`; stack.push(tag); if (stack.length < lvl) out += '<li>'; } out += '<li>' + body; }
    else { close(lvl); out += '</li><li>' + body; }
  }
  if (!stack.length && !out) return h;
  const tail = h.slice(last); if (tail.trim()) close(0); out += tail; close(0);
  return out;
}

// Rich DOM → Markdown
/** A picture: ![alt](src), or once it was resized (a width or height attribute) the <img> tag Typora writes, which keeps the size. */
const imgMd = c => {
  const src = c.getAttribute('data-rel') || c.getAttribute('src') || '', alt = c.getAttribute('alt') || '', w = c.getAttribute('width'), h = c.getAttribute('height'); // data-rel: the link as written (assets/x.png) while src shows it through the engine
  return w || h ? `<img src="${src}" alt="${alt.replace(/"/g, '&quot;')}"${w ? ` width="${w}"` : ''}${h ? ` height="${h}"` : ''}>` : `![${alt}](${src})`;
};
export function htmlToMd(root) {
  const inl = n => {
    let s = '';
    n.childNodes.forEach(c => {
      if (c.nodeType === 3) { s += c.nodeValue.replace(/​/g, '').replace(/ /g, ' '); return; }
      if (c.nodeType !== 1) return;
      const t = c.tagName.toLowerCase(), st = c.style || {};
      if (t === 'br') { s += '  \n'; return; }
      if (t === 'input') return;
      if (t === 'img') { s += imgMd(c); return; }
      if (t === 'span' && c.classList && c.classList.contains('md-math')) { s += '$' + (c.getAttribute('data-src') || '') + '$'; return; }
      if (t === 'sup' && c.classList && c.classList.contains('md-fn')) { s += '[^' + (c.getAttribute('data-fn') || '') + ']'; return; }
      if (t === 'span' && c.classList && c.classList.contains('md-emoji')) { s += c.getAttribute('data-src') || c.textContent; return; }
      const x = inl(c); if (!x.trim() && t !== 'code') { s += x; return; }
      const wrap = m => { const lead = x.match(/^\s*/)[0], trail = x.match(/\s*$/)[0]; return lead + m + x.trim() + m + trail; };
      if (t === 'strong' || t === 'b' || +st.fontWeight >= 600 || st.fontWeight === 'bold') s += wrap('**');
      else if (t === 'em' || t === 'i' || st.fontStyle === 'italic') s += wrap('*');
      else if (t === 'del' || t === 's' || t === 'strike' || (st.textDecoration || '').includes('line-through')) s += wrap('~~');
      else if (t === 'mark') s += wrap('==');
      else if (t === 'code') s += '`' + c.textContent + '`';
      else if (t === 'sub') s += '~' + x.trim() + '~';
      else if (t === 'sup') s += '^' + x.trim() + '^';
      else if (t === 'a') { const href = c.getAttribute('href') || '', auto = c.getAttribute('data-auto'); s += c.getAttribute('data-wiki') ? (x === href ? `[[${href}]]` : `[[${href}|${x}]]`) : auto === 'angle' ? `<${x}>` : auto ? x : `[${x}](${href})`; }
      else s += x;
    });
    return s;
  };
  const list = (el, depth) => {
    const ord = el.tagName.toLowerCase() === 'ol'; let i = parseInt(el.getAttribute('start') || '1');
    const out = [];
    Array.from(el.children).forEach(li => {
      if (/^(ul|ol)$/i.test(li.tagName)) { out.push(list(li, depth + 1)); return; } // a sublist beside its item rather than inside it (WebKit's indent, some pasted HTML): nested under the item above
      if (li.tagName.toLowerCase() !== 'li') return;
      const cb = li.querySelector(':scope > input[type=checkbox]');
      const clone = li.cloneNode(true); clone.querySelectorAll(':scope > ul, :scope > ol').forEach(x => x.remove());
      let text = inl(clone).replace(/^\s+/, '').replace(/\s+$/, '');
      const mark = ord ? (i++) + '. ' : '- ';
      out.push('  '.repeat(depth) + mark + (cb ? (cb.checked ? '[x] ' : '[ ] ') : '') + text);
      li.querySelectorAll(':scope > ul, :scope > ol').forEach(sub => out.push(list(sub, depth + 1)));
    });
    return out.join('\n');
  };
  const block = el => {
    if (el.nodeType === 3) { const t = el.nodeValue.replace(/​/g, ''); return t.trim() ? t.trim() : null; }
    if (el.nodeType !== 1) return null;
    const t = el.tagName.toLowerCase();
    const cls = c => !!(el.classList && el.classList.contains(c)), live = () => { const ta = el.querySelector('textarea'); return ta ? ta.value : (el.getAttribute('data-src') || ''); }; // a block still open for editing saves its live value, not the last-rendered data-src
    if (cls('md-math-block')) return '$$\n' + live() + '\n$$';
    if (cls('md-mermaid')) return '```mermaid\n' + live() + '\n```';
    if (cls('md-front')) return '---\n' + live() + '\n---';
    if (cls('md-toc')) return '[TOC]';
    if (cls('md-alert-title')) return null; // the alert's label is drawn, not written
    if (cls('md-footnotes')) return Array.from(el.querySelectorAll('li')).map(li => '[^' + (li.getAttribute('data-fn') || '') + ']: ' + inl(li).trim().replace(/\n/g, '\n    ')).join('\n');
    const m = /^h([1-6])$/.exec(t);
    if (m) return '#'.repeat(+m[1]) + ' ' + inl(el).trim();
    if (t === 'ul' || t === 'ol') return list(el, 0);
    if (t === 'blockquote') { const inner = Array.from(el.childNodes).map(block).filter(x => x != null), kind = el.getAttribute('data-alert'); const body = (kind ? '[!' + kind + ']\n' : '') + (inner.length ? inner.join('\n\n') : inl(el)); return body.split('\n').map(l => '> ' + l).join('\n'); }
    if (t === 'pre') { const code = el.querySelector('code'); const lang = code && code.getAttribute('data-lang') || ''; return '```' + lang + '\n' + (el.innerText || el.textContent).replace(/\n$/, '') + '\n```'; }
    if (t === 'hr') return '---';
    if (t === 'table') {
      const trs = Array.from(el.querySelectorAll('tr'));
      const rows = trs.map(tr => Array.from(tr.children).map(c => inl(c).trim().replace(/ {2}\n/g, '<br>').replace(/\s*\n\s*/g, ' ').replace(/\|/g, '\\|')));
      if (!rows.length) return null; const w = Math.max(...rows.map(r => r.length));
      const al = Array.from(trs[0].children).map(c => (c.style && c.style.textAlign) || c.getAttribute('align') || ''); // the header row's alignment is the column's (the toolbar sets every cell, a file only the header)
      const line = r => '| ' + Array.from({ length: w }, (_, i) => r[i] || '').join(' | ') + ' |';
      const sep = Array.from({ length: w }, (_, i) => ({ center: ':---:', right: '---:', left: ':---' })[al[i]] || '---');
      return [line(rows[0]), '| ' + sep.join(' | ') + ' |'].concat(rows.slice(1).map(line)).join('\n');
    }
    if (t === 'img') return imgMd(el);
    if (t === 'div' && el.querySelector('p,h1,h2,h3,ul,ol,pre,blockquote,table')) return Array.from(el.childNodes).map(block).filter(x => x != null).join('\n\n');
    const s = inl(el).replace(/\s+$/, '');
    return s.trim() ? s : '';
  };
  const parts = []; let buf = '';
  root.childNodes.forEach(n => {
    const inline = n.nodeType === 3 || (n.nodeType === 1 && /^(b|strong|i|em|a|code|span|mark|del|s|img|br|sup|sub)$/i.test(n.tagName));
    if (inline) { const d = document.createElement('div'); d.appendChild(n.cloneNode(true)); buf += inl(d); return; }
    if (buf.trim()) parts.push(buf.trim()); buf = '';
    const b = block(n); if (b != null) parts.push(b);
  });
  if (buf.trim()) parts.push(buf.trim());
  return parts.join('\n\n').replace(/\n{3,}/g, '\n\n').replace(/^\n+/, '') + '\n';
}
export function wordHtml(html) {
  return html.replace(/<table/g, '<table style="border-collapse:collapse;width:100%;margin:8px 0"').replace(/<(t[hd])( style="[^"]*")?>/g, (m, t, st) => `<${t} style="border:1px solid #C7C7CC;padding:6px 8px;${t === 'th' ? 'font-weight:600;background:#F5F5F7;' : ''}${st ? st.slice(8, -1) : ''}">`).replace(/<input type="checkbox"[^>]*checked>/g, '☑').replace(/<input type="checkbox"[^>]*>/g, '☐').replace(/ data-(line|li|task)="\d+"/g, '');
}
