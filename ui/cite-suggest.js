// Sentences that draw on a source of the document and have no citation yet, found from the sources' own words — an author's
// name, the words of a title that the rest of the paper does not use everywhere, the year — without asking a model. A quotation,
// a figure or a reporting verb ("found", "according to", 研究表明) makes a claim that wants a citation even when no source matches.
// The Word editor marks such a sentence in grey and offers the citation where it goes (WordEditor's citeSuggest).

const STOP = new Set(('a an and are as at be been being but by can could did do does for from had has have how in into is it its of on or our ' +
  'that the their them then there these they this those to was were what when where which while who why will with within without would ' +
  'about above after again against all also among any because before below between both during each few further here more most other ' +
  'over same should some such than through under until very upon using used use new study studies effect effects role case based toward towards').split(' '));
/** Words that end in a period without ending the sentence: et al., e.g., i.e., Fig., vol., pp., Dr., and single letters (initials). */
const ABBR = /(?:^|[\s(])(?:al|e\.g|i\.e|etc|vs|cf|fig|figs|no|nos|vol|vols|pp|p|ch|chap|ed|eds|dr|mr|mrs|ms|st|jr|sr|approx|ca|op|cit|ibid|[a-z])$/i;
const REPORT = /\b(found|finds|showed|shows|shown|reported|reports|according to|argues?|argued|suggests?|suggested|demonstrat(?:e|es|ed)|estimat(?:e|es|ed)|measured|observed|concluded|conclude[sd]?|indicate[sd]?|revealed|noted|claims?|claimed|researchers?|evidence|surveys?|data show)\b|研究(表明|发现|显示)|数据显示|指出|认为|据.{0,8}(统计|报道|报告)|根据/i;
const QUOTE = /[“"「『][^”"」』]{12,}[”"」』]/;
const FIGURE = /\d+(?:\.\d+)?\s?(?:%|percent|per cent|times|fold)|\b\d[\d,]{2,}\b|\d+(?:\.\d+)?\s?倍/;

/** Lower case, without accents, as words: Latin words of two letters or more, runs of Chinese characters. */
export const wordsOf = s => String(s || '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '').match(/[a-z][a-z'-]*[a-z]|[一-鿿]{2,}/g) || [];
/** A word's plain form, so "processes" meets "process" and "studies" "study". */
export const stem = w => w.replace(/'s$/, '').replace(/ies$/, 'y').replace(/sses$/, 'ss').replace(/([^s])s$/, (m, c) => w.length > 4 ? c : m);

/** A text's sentences as [start, end) offsets: ended by . ! ? (not after an abbreviation or an initial) or 。！？, closing quotes
 *  and brackets with them; the space between them belongs to neither. */
export function sentencesOf(text) {
  const out = [], re = /[.!?]+["”’')\]]*(?=\s|$)|[。！？]+[”’」』）]*/g; let start = 0, m;
  const push = end => { let s = start; while (s < end && /\s/.test(text[s])) s++; if (s < end) out.push({ start: s, end, text: text.slice(s, end) }); start = end; };
  while ((m = re.exec(text))) {
    if (m[0] === '.' && ABBR.test(text.slice(start, m.index))) continue;
    push(m.index + m[0].length);
  }
  if (start < text.length) push(text.length);
  return out;
}

/** What names a source in a sentence: its authors' (or editors') surnames, the words of its title (stemmed), its year. */
export function termsOf(src) {
  const people = [].concat(src.authors || [], (src.authors || []).length ? [] : src.editors || []);
  const names = people.map(p => typeof p === 'string' ? (p.includes(',') ? p.split(',')[0] : p.trim().split(/\s+/).pop()) : p.last || '').map(n => n.trim()).filter(n => n.length > 1);
  const title = [...new Set(wordsOf(src.title).filter(w => w.length > 3 && !STOP.has(w)).map(stem))];
  return { tag: src.tag, names, title, year: /^\d{4}$/.test(String(src.year || '')) ? String(src.year) : '' };
}
/** A surname as a whole word: not inside another (Kim in Kimball). */
const nameRe = names => names.length ? new RegExp(`(^|[^\\p{L}])(${names.map(n => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})($|[^\\p{L}])`, 'u') : null;

/** How strongly a sentence draws on a source: its author named (3), each of its title's words the paper does not use everywhere
 *  (up to 1 each, less when other sources' titles have it too, at most 4), its year (1). `rare` weighs a word by how few of the
 *  paper's sentences have it; `shared` by how many of the sources' titles do. */
export function scoreOf(sentence, t, rare, shared, has = new Set(wordsOf(sentence).map(stem))) {
  const re = t.re !== undefined ? t.re : (t.re = nameRe(t.names)), name = !!re && re.test(sentence);
  const words = t.title.filter(w => has.has(w)), topic = Math.min(4, words.reduce((s, w) => s + rare(w) / (shared.get(w) || 1), 0));
  const year = !!t.year && sentence.includes(t.year);
  return { score: (name ? 3 : 0) + topic + (year ? 1 : 0), name, topic, words };
}

/** A claim that wants a citation whatever it cites: a quotation, a figure, a reporting verb. */
export const claimOf = s => (QUOTE.test(s) ? 2 : 0) + (FIGURE.test(s) ? 0.5 : 0) + (REPORT.test(s) ? 1 : 0);

/** The sentences to mark: paras is [{ text, cited: [offsets of the citations and note marks in it] }], sources the document's.
 *  Returns [{ para, start, end, tag, score }] — tag '' when the sentence makes a claim but no source is its. A sentence cites already
 *  when a citation or a note's mark sits in it or right after its end; short sentences (under six words) and `skip`ped ones (the
 *  writer said no: their text) are left alone. */
export function suggest(paras, sources, skip) {
  const terms = (sources || []).map(termsOf), all = [];
  paras.forEach((p, i) => { for (const s of sentencesOf(p.text || '')) { const ws = wordsOf(s.text); all.push(Object.assign(s, { para: i, n: ws.length, has: new Set(ws.map(stem)), cited: (p.cited || []).some(o => o > s.start && o <= s.end) })); } });
  const df = new Map(); for (const s of all) for (const w of s.has) df.set(w, (df.get(w) || 0) + 1);
  const n = all.length, cut = Math.max(0.15, 1.5 / Math.max(1, n));
  const rare = w => (df.get(w) || 0) / Math.max(1, n) <= cut ? 1 : 0.3; // the paper's own subject is in every other sentence: no sign of a source
  const shared = new Map(); for (const t of terms) for (const w of t.title) shared.set(w, (shared.get(w) || 0) + 1);
  const out = [];
  for (const s of all) {
    if (s.cited || s.n < 6 || (skip && skip.has(s.text.trim()))) continue;
    let best = null;
    for (const t of terms) { const r = scoreOf(s.text, t, rare, shared, s.has); if ((r.name || r.topic >= 1.8) && r.score >= 2.5 && (!best || r.score > best.score)) best = Object.assign(r, { tag: t.tag }); }
    const claim = claimOf(s.text);
    if (best) out.push({ para: s.para, start: s.start, end: s.end, tag: best.tag, score: best.score + claim / 2 });
    else if (claim >= 2) out.push({ para: s.para, start: s.start, end: s.end, tag: '', score: claim });
  }
  return out;
}

/** Where a citation goes in a sentence ending at `end` of `text`: in the parentheses styles before its closing punctuation, or after a
 *  closing quotation mark with a period after the citation (“…retention” (Pegg 288). — the period inside the quotation marks moves
 *  out); in Chicago's notes (`notes`) the mark goes after it all. { at, drop: the offset of a period to take out, or -1, after: what
 *  follows the citation }. */
export function placeOf(text, end, notes) {
  const tail = /([.!?。！？]+)(["”’')\]」』）]*)$/.exec(text.slice(0, end));
  if (!tail || notes) return { at: end, drop: -1, after: '' };
  const p = tail.index, close = tail[2];
  if (/^["”’」』]/.test(close)) return tail[1] === '.' ? { at: end, drop: p, after: '.' } : { at: end, drop: -1, after: '.' };
  return { at: close ? end : p, drop: -1, after: '' };
}
