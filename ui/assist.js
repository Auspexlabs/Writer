// assist.js — 写作辅助 for academic papers. A paper is written in a citation style (MLA, APA, Chicago…). When Writer knows which, the
// Word editor's grammar check follows that style's conventions (grammarRules), the AI assistant is told what the paper is and how its
// citations are kept (chatInstructions), and the works-cited list takes what is pasted into it. The style comes from the template the
// paper started from, from the document itself — its works-cited title, its citations, its header (detect) — or, failing those, from
// the fast model's reading of its opening and its end (classifyPrompt / parseStyle). The writer is asked before assistance is turned
// on (index.dc.html's assist*), and the answer is kept per file (assistOf / setAssist).

/** The styles, as the engine names them (its citationStyle), with the list each ends in. The names are the engine's own (engine.js
 *  CITE_STYLES), so the dictionaries translate them. */
export const STYLES = {
  mla: { short: 'MLA', name: 'MLA 第 9 版', list: 'Works Cited', notes: false },
  apa: { short: 'APA', name: 'APA 第 7 版', list: 'References', notes: false },
  chicago: { short: 'Chicago', name: 'Chicago 第 18 版（脚注）', list: 'Bibliography', notes: true },
  'chicago-date': { short: 'Chicago', name: 'Chicago 第 18 版（作者-日期）', list: 'References', notes: false },
  gb7714: { short: 'GB/T 7714', name: 'GB/T 7714—2015（顺序编码）', list: '参考文献', notes: false },
  ieee: { short: 'IEEE', name: 'IEEE', list: 'References', notes: false },
};
export const KEYS = Object.keys(STYLES);

// ---- what the writer decided, per file ----
const KEY = 'writerAssist:';
/** The style assistance is on for a file, 'off' when the writer said no, '' when not asked yet. */
export function assistOf(path, storage = globalThis.localStorage) { try { return (storage && storage.getItem(KEY + path)) || ''; } catch (e) { return ''; } }
export function setAssist(path, value, storage = globalThis.localStorage) { try { if (value) storage.setItem(KEY + path, value); else storage.removeItem(KEY + path); } catch (e) { } }
/** A file moved or saved as: its decision goes with it. */
export function moveAssist(from, to, storage = globalThis.localStorage) { const v = assistOf(from, storage); if (v && from !== to) { setAssist(to, v, storage); setAssist(from, '', storage); } }

// ---- telling the style from the document ----
const TITLES = { 'works cited': 'mla', 'work cited': 'mla', references: 'apa', 'reference list': 'apa', bibliography: 'chicago', '参考文献': 'gb7714' };
const name = "[A-Z][\\p{L}'’-]+(?: (?:and|&) [A-Z][\\p{L}'’-]+)?(?: et al\\.)?";
const INTEXT = [
  ['apa', new RegExp(`\\(${name}, (?:n\\.d\\.|\\d{4}[a-z]?)(?:, pp?\\. \\d+(?:[-–]\\d+)?)?\\)`, 'gu')],
  ['chicago-date', new RegExp(`\\(${name} \\d{4}[a-z]?(?:, \\d+(?:[-–]\\d+)?)?\\)`, 'gu')],
  ['mla', new RegExp(`\\(${name} \\d{1,3}(?:[-–]\\d{1,3})?\\)`, 'gu')],
  ['numbered', /\[\d{1,3}(?:\s?[,–-]\s?\d{1,3})*\]/g],
];
const clean = s => String(s || '').replace(/\s+/g, ' ').trim();
const count = (re, s) => { re.lastIndex = 0; return (s.match(re) || []).length; };
/** Whether the text is written in English (Latin letters outweigh Chinese characters two to one). */
export const english = text => { const s = String(text || ''); const cjk = (s.match(/[一-鿿]/g) || []).length, latin = (s.match(/[A-Za-z]/g) || []).length; return latin > cjk * 2; };
const words = text => (String(text || '').match(/[A-Za-z][A-Za-z'’-]*|[一-鿿]/g) || []).length;

/** The style a document is written in, from what it shows: the title of its works-cited list (the last Works Cited, References,
 *  Bibliography or 参考文献 standing alone), its in-text citations ((Author page), (Author, Year), (Author Year, page), [1]), an
 *  MLA header (Surname page), footnotes in an English text with no citations in it, or the style its sources are kept in. '' when
 *  none of these says. blocks: [{ kind: 'heading' | 'paragraph', text }], notes: how many footnotes, header: the header's text. */
export function detect({ blocks = [], header = '', notes = 0, citeStyle = '', sources = 0 } = {}) {
  const texts = blocks.map(b => clean(b.text)), all = texts.join('\n'), en = english(all);
  const hits = Object.fromEntries(INTEXT.map(([k, re]) => [k, count(re, all)]));
  const top = Object.entries(hits).sort((a, b) => b[1] - a[1])[0];
  const numbered = () => en ? 'ieee' : 'gb7714';
  for (let i = texts.length - 1; i >= 0; i--) {
    const t = texts[i].replace(/[:：.]$/, '').toLowerCase();
    if (!(t in TITLES) || (blocks[i].kind !== 'heading' && texts[i].split(' ').length > 3)) continue;
    const style = TITLES[t];
    if (style === 'apa') return hits['chicago-date'] > hits.apa && hits['chicago-date'] > 0 ? 'chicago-date' : hits.numbered > hits.apa && hits.numbered > 0 ? numbered() : 'apa';
    if (style === 'gb7714') return en && hits.numbered > 0 ? 'ieee' : 'gb7714';
    return style;
  }
  if (top && top[1] >= 2) return top[0] === 'numbered' ? numbered() : top[0];
  const h = clean(header);
  if (/^[A-Z][\p{L}'’-]+(?: [A-Z][\p{L}'’-]+)? (?:\{page\}|\d+)$/u.test(h)) return 'mla';
  if (/^[A-Z][A-Z0-9 :,'’-]{6,}\s+(?:\{page\}|\d+)$/.test(h) && en) return 'apa';
  if (top && top[1] === 1 && top[0] !== 'numbered' && en) return top[0];
  if (notes >= 2 && en && !top?.[1]) return 'chicago';
  if (citeStyle && sources > 0 && KEYS.includes(citeStyle)) return citeStyle;
  return '';
}

/** Whether a document reads like an academic paper (worth asking the model which style it follows): 250 words or more, in English,
 *  with a year, a quotation or a paper's own words (abstract, introduction, argue, thesis…) somewhere. */
export function looksAcademic(blocks = []) {
  const all = blocks.map(b => clean(b.text)).join('\n');
  return words(all) >= 250 && english(all) && (/\b(19|20)\d{2}\b/.test(all) || /["“][^"”]{20,}["”]/.test(all) || /\b(abstract|introduction|conclusion|argues?|argued|thesis|this (essay|paper|study)|literature|hypothes[ie]s|methodolog)/i.test(all));
}

/** The question for the model when the document itself does not say: its header, its opening and its end, one word back. */
export function classifyPrompt(blocks = [], header = '') {
  const all = blocks.map(b => (b.kind === 'heading' ? '# ' : '') + clean(b.text)).filter(Boolean).join('\n');
  const head = all.slice(0, 2500), tail = all.length > 4000 ? all.slice(-1500) : all.slice(2500);
  return {
    system: 'You read the header, the opening and the end of a document and say which academic citation style it follows. Reply with exactly one word: mla, apa, chicago, chicago-date, gb7714, ieee, or none.\n'
      + 'mla: a four-line heading (name, instructor, course, date), a "Surname page" header, (Author page) citations, a Works Cited list. apa: a title page with an affiliation, the page number alone in the header, (Author, Year, p. n) citations, a References list. chicago: footnotes carrying full citations and a Bibliography. chicago-date: (Author Year, page) citations and a References list. gb7714: a Chinese paper with 参考文献 numbered [1]. ieee: an English paper with references numbered [1].\n'
      + 'none: not an academic paper, or no sign of a style. One word, lower case, nothing else.',
    user: `<header>${clean(header)}</header>\n<opening>\n${head}\n</opening>\n` + (tail ? `<ending>\n${tail}\n</ending>\n` : '') + 'Which style?',
  };
}
/** The model's one word as a style key, '' for none or anything else. */
export function parseStyle(reply) { const w = String(reply || '').trim().toLowerCase().replace(/[^a-z0-9-]+/g, ' ').trim().split(' ')[0] || ''; return KEYS.includes(w) ? w : ''; }

// ---- what each style asks of the text and of the assistant ----
const COMMON = 'Formal academic prose: no contractions (do not → "do not", not "don\'t"), no second person, numbers under ten in words, consistent spelling (one of US or UK), one space after a period.';
const RULES = {
  mla: COMMON + ' MLA 9: an in-text citation sits before the period that ends the sentence — "… the regulars set the tone (Oldenburg 16)." — never after it; when the sentence names the author only the page is given, (16); a short quotation keeps its closing quotation mark before the citation and the period after it: "… the street" (Jacobs 35). Titles of books in italics, of articles in quotation marks (leave formatting to the writer; flag only the words).',
  apa: COMMON + ' APA 7: citations are (Author, Year) or (Author, Year, p. 12) before the period, with "&" between authors inside the parentheses and "and" in the sentence; a narrative citation reads "Putnam (1995) found"; use the serial comma; "et al." for three or more authors; percentages as numerals with %.',
  chicago: COMMON + ' Chicago 18 (notes and bibliography): a note number comes after the period or comma, never before it, and never two numbers at one place; the text carries no parenthetical citations; "ibid." is not used.',
  'chicago-date': COMMON + ' Chicago 18 (author-date): citations are (Author Year, page) with no comma between author and year, before the period.',
  gb7714: '正式的学术汉语：全角标点，句末句号在引文序号之前（……研究表明[1]。），文献序号用上标方括号；不用口语和第二人称；数字用阿拉伯数字，统计量和单位之间留空格；中英文之间留一个空格。',
  ieee: COMMON + ' IEEE: citations are numbered in square brackets, [1], placed before the punctuation, in the order first cited; "in [3]" rather than "in reference [3]".',
};
/** What the grammar check holds a paragraph of this style to, beside the grammar itself; '' for no style. */
export const grammarRules = style => RULES[style] || '';

const SOURCE_JSON = '{"type":"article|book|chapter|webpage|newspaper|report|thesis","authors":[{"last":"Putnam","first":"Robert D."}],"title":"…","container":"the journal, book or site","year":"1995","volume":"6","issue":"1","pages":"65-78","publisher":"…","url":"…","doi":"…"}';
const TOOLS = `Sources, citations and the works-cited list are Word's own objects in the file, never typed text:
- Add or change a source: \`set <file> / --prop source='${SOURCE_JSON}'\` (a tag is given on add; \`get <file> / --prop sources\` lists them).
- Cite in the text: \`add <file> <paragraph path> --type citation --prop sources=<tag> --prop pages=<n> --prop at=<character offset>\` — the offset is where the citation goes, right before the period that ends the sentence (count the paragraph's text; the engine draws the citation in the paper's style, so write only the sentence's words, with a space before the period).
- Cite in a note (Chicago): \`add <file> <paragraph path> --type footnote --prop cite=<tag> --prop pages=<n>\` — the mark lands at the end of the paragraph; give at=<offset> for an earlier sentence, after its period.
- The list: \`add <file> /body --type bibliography\` once, at the end; the engine keeps it sorted and formatted. Never edit its text, and never write references as paragraphs.
- The style: \`set <file> / --prop citationStyle=<style>\` redraws every citation and the list.`;
const PAPER = {
  mla: 'an academic paper in MLA 9 style: Times New Roman 12 double-spaced, half-inch first-line indents, a header of "Surname {page}" top right, a four-line heading (student, instructor, course, date) and the title centered on the first page, no title page; the Works Cited list on its own page.',
  apa: 'an academic paper in APA 7 style (student paper): Times New Roman 12 double-spaced, half-inch first-line indents, the page number alone top right, a title page (title in bold, author, affiliation, course, instructor, date), the title again in bold above the text, Level 1 headings centered bold and Level 2 flush-left bold (heading and heading2 nodes); the References list on its own page.',
  chicago: 'an academic paper in Chicago 18 notes-bibliography style: Times New Roman 12 double-spaced, half-inch first-line indents, a title page without a page number, page numbers from the first text page, citations as footnotes (full the first time, short after), the Bibliography on its own page.',
  'chicago-date': 'an academic paper in Chicago 18 author-date style: Times New Roman 12 double-spaced, (Author Year, page) citations in the text, the References list on its own page.',
  gb7714: '一篇按 GB/T 7714—2015（顺序编码制）写的中文论文：文中引用用上标序号 [1]，参考文献按引用顺序排在文末。',
  ieee: 'a paper in IEEE style: citations numbered [1] in order of first citation, the References list numbered the same way.',
};
/** What the AI assistant is told when assistance is on for a paper: what the paper is, and how its citations are kept. */
export function chatInstructions(style) {
  if (!PAPER[style]) return '';
  return `## The open document\n\nThe open document is ${PAPER[style]} The writer has turned on writing assistance for it: keep the paper in this style, and when asked to check it, go through the header, the title or title page, spacing and indents, the headings, every citation against the list and the list against the citations, and the style's punctuation around citations; report what is off and fix what the writer asks.\n\n${TOOLS}`;
}

// ---- a source read by the model when the lookup finds nothing (the works-cited list's paste box, WordEditor's bibGuess) ----
/** A web page as text for the model: its title, its description, author and date tags when it has them, and the words of its body. */
export function pageText(html, max = 5000) {
  const h = String(html || ''), meta = name => { const m = new RegExp(`<meta[^>]+(?:name|property)=["'](?:${name})["'][^>]*content=["']([^"']*)["']`, 'i').exec(h) || new RegExp(`<meta[^>]+content=["']([^"']*)["'][^>]*(?:name|property)=["'](?:${name})["']`, 'i').exec(h); return m ? m[1].trim() : ''; };
  const title = (/<title[^>]*>([\s\S]*?)<\/title>/i.exec(h) || [])[1] || '';
  const body = h.replace(/<(script|style|noscript|svg|nav|footer|header|head|title)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, ' ').replace(/<br\s*\/?>|<\/(p|div|li|h\d|tr|section|article)>/gi, '\n').replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/[ \t]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
  const head = [['title', clean(title)], ['description', meta('description|og:description')], ['author', meta('author|article:author|dc.creator')], ['date', meta('article:published_time|date|dc.date|pubdate')], ['site', meta('og:site_name')]].filter(([, v]) => v).map(([k, v]) => `${k}: ${v}`).join('\n');
  return (head + '\n\n' + body).trim().slice(0, max);
}
/** The question: what was pasted, and the page it names when it is an address, read into the engine's source JSON. */
export function sourcePrompt(pasted, page = '') {
  return {
    system: 'You turn what a writer pasted into a source record for a citation manager. Reply with one JSON object and nothing else: {"type": "article" | "book" | "chapter" | "webpage" | "newspaper" | "magazine" | "report" | "conference" | "thesis" | "other", "authors": [{"last": "…", "first": "…"}] (an organisation as [{"name": "…"}]), "title": "…", "container": the journal, the book, the newspaper or the website, "publisher": "…", "place": "…", "year": "…", "month": "…", "day": "…", "volume": "…", "issue": "…", "pages": "…", "url": "…", "doi": "…"}.\n'
      + 'Only what the text states: leave out every field it does not give, and never invent an author, a year or a page. The pasted text may be a reference in any style (MLA, APA, Chicago, GB/T 7714, IEEE, or by hand), a web address with the page\'s text, or a few facts in words. A page\'s title is the title; its site is the container; a byline is the author; a dateline is the date. Reply with the JSON object only, no code fence.',
    user: `<pasted>\n${String(pasted || '').trim()}\n</pasted>\n` + (page ? `<page>\n${page}\n</page>\n` : '') + 'The source as JSON.',
  };
}
const TYPES = ['article', 'book', 'chapter', 'webpage', 'newspaper', 'magazine', 'report', 'conference', 'thesis', 'other'];
const FIELDS = ['title', 'container', 'publisher', 'place', 'year', 'month', 'day', 'volume', 'issue', 'pages', 'edition', 'url', 'doi', 'isbn', 'accessed'];
/** The model's object as a source: its known fields as text, its people as {last, first} or {name}; null without a title. */
export function parseSource(reply) {
  let s = String(reply || '').trim().replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/, '');
  const a = s.indexOf('{'), b = s.lastIndexOf('}'); if (a < 0 || b <= a) return null;
  let o; try { o = JSON.parse(s.slice(a, b + 1)); } catch (e) { return null; }
  if (!o || typeof o !== 'object' || typeof o.title !== 'string' || !o.title.trim()) return null;
  const src = { type: TYPES.includes(o.type) ? o.type : o.url || o.doi ? (o.doi ? 'article' : 'webpage') : 'other' };
  for (const k of FIELDS) if (o[k] != null && String(o[k]).trim()) src[k] = String(o[k]).trim();
  for (const k of ['authors', 'editors']) {
    const people = (Array.isArray(o[k]) ? o[k] : []).map(p => typeof p === 'string' ? (p.includes(',') ? { last: p.split(',')[0].trim(), first: p.split(',').slice(1).join(',').trim() } : { name: p.trim() })
      : p && typeof p === 'object' ? (p.last ? { last: String(p.last).trim(), first: String(p.first || '').trim() } : p.name ? { name: String(p.name).trim() } : null) : null).filter(p => p && (p.last || p.name));
    if (people.length) src[k] = people;
  }
  return src;
}
