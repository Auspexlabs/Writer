// grammar.js — 语法检查. The paragraph the writer just left goes to the fast model (engine.js ask → POST /ask), which answers with the
// corrections a copy editor would make, as JSON: {find, replace, why, kind}. Each is anchored to the paragraph's own text here — what
// the model cannot quote exactly is dropped, so nothing is marked that is not there — and the Word editor draws them over the text
// (a deletion struck through, a change underlined) without touching it until the writer takes one with Tab (WordEditor's gram*).

/** Characters of one paragraph sent; a longer one is checked up to there. */
export const MAX = 6000;
export const KINDS = ['spelling', 'grammar', 'punctuation', 'style'];

export const SYSTEM = [
  'You are the grammar checker of a document editor. The user hands you one paragraph of a document; you reply with the corrections a careful copy editor would make, as JSON, and nothing else.',
  '',
  'Rules:',
  '- Only real errors: spelling, grammar (agreement, tense, articles, prepositions), punctuation, a wrong, missing or doubled word, 错别字, 语病, 搭配不当, 成分残缺, 标点误用. Do not rewrite for style or taste, do not change the meaning, and do not flag what is merely unusual.',
  '- Leave alone: quotations, titles, names, numbers, URLs, code, and citations such as (Smith 12), (Smith, 2020, p. 4), [3] or a footnote number.',
  '- When <conventions> are given, also correct what breaks them (a citation after the period, a contraction in formal prose, a missing serial comma); mark those with kind "style".',
  '- Each correction is {"find": the exact shortest stretch of the paragraph to change, copied character for character, "replace": what it becomes, "" to delete it, "why": one short reason in the language of the paragraph (at most ten words), "kind": one of "spelling", "grammar", "punctuation", "style"}.',
  '- "find" must occur in the paragraph exactly once: when the same words occur more than once, extend it by a word or two until it is unique. Corrections never overlap.',
  '- At most twelve corrections. A paragraph with nothing to correct gets [].',
  '- Reply with the JSON array only: no prose, no code fence.',
].join('\n');

/** The language a paragraph is written in: zh when it has at least as many Chinese characters as Latin words. */
export function langOf(text) {
  const s = String(text || ''), cjk = (s.match(/[一-鿿]/g) || []).length, words = (s.match(/[A-Za-z]{2,}/g) || []).length;
  return cjk > 0 && cjk >= words ? 'zh' : 'en';
}

/** The request for a paragraph: the system prompt, the message (the paragraph, its language, the conventions of the paper's style when
 *  writing assistance names one — assist.js grammarRules), and the language found. */
export function prompt(text, { lang, rules = '' } = {}) {
  const t = String(text || '').slice(0, MAX), language = lang || langOf(t);
  const user = `<language>${language === 'zh' ? 'Chinese' : 'English'}</language>\n` + (rules ? `<conventions>\n${rules}\n</conventions>\n` : '')
    + `<paragraph>\n${t}\n</paragraph>\nList the corrections for the paragraph as a JSON array.`;
  return { system: SYSTEM, user, lang: language };
}

/** The corrections in a reply: the JSON array in it (a fence or a word around it tolerated), each with its find, replace, why and kind;
 *  [] when there is none to read. */
export function parse(reply) {
  let s = String(reply || '').trim();
  s = s.replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/, '');
  const a = s.indexOf('['), b = s.lastIndexOf(']'); if (a < 0 || b <= a) return [];
  let list; try { list = JSON.parse(s.slice(a, b + 1)); } catch (e) { return []; }
  if (!Array.isArray(list)) return [];
  return list.filter(x => x && typeof x === 'object' && typeof x.find === 'string' && x.find.length > 0)
    .map(x => ({ find: x.find, replace: typeof x.replace === 'string' ? x.replace : '', why: typeof x.why === 'string' ? x.why.trim() : '', kind: KINDS.includes(x.kind) ? x.kind : 'grammar' }));
}

/** Quotes, dashes and spaces as one character each, so a curly quote in the text meets a straight one in the reply at the same offset. */
const fold = s => s.replace(/[‘’‚′]/g, "'").replace(/[“”„″]/g, '"').replace(/ /g, ' ').replace(/[–—]/g, '-');

/** The corrections placed in the paragraph's text as [start, end) offsets: found as written, else with quotes and dashes folded, else
 *  without the spaces around them; one the text does not hold, one that changes nothing, and one overlapping an earlier one are left
 *  out. In text order. */
export function anchor(issues, text) {
  const t = String(text || ''), ft = fold(t), out = [];
  for (const x of issues || []) {
    let find = x.find, replace = x.replace, at = t.indexOf(find);
    if (at < 0 && (at = ft.indexOf(fold(find))) >= 0) find = t.slice(at, at + find.length);
    if (at < 0) { const trimmed = x.find.trim(); if (trimmed && trimmed !== x.find && (at = t.indexOf(trimmed)) >= 0) { find = trimmed; replace = replace.trim(); } }
    if (at < 0 || replace === find) continue;
    out.push({ start: at, end: at + find.length, find, replace, why: x.why || '', kind: x.kind || 'grammar' });
  }
  out.sort((a, b) => a.start - b.start || a.end - b.end);
  const kept = []; let last = -1;
  for (const x of out) { if (x.start < last) continue; kept.push(x); last = x.end; }
  return kept;
}

/** How a correction shows: a deletion (its words struck through) or a fix (its words underlined, the new ones offered). */
export const kindOf = x => x.replace === '' ? 'delete' : 'fix';

/** The text with one correction taken. */
export const apply = (text, x) => String(text).slice(0, x.start) + x.replace + String(text).slice(x.end);

/** The other corrections of a paragraph once `taken` is applied: those after it move by the length it added or removed. */
export function shift(issues, taken) {
  const delta = taken.replace.length - (taken.end - taken.start);
  return issues.filter(x => x !== taken).map(x => x.start >= taken.end ? Object.assign({}, x, { start: x.start + delta, end: x.end + delta }) : x);
}
