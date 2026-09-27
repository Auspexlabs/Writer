// cite.js — sources for citations. What people paste (a DOI, an ISBN, a web address, or a whole reference copied from somewhere)
// becomes a source the engine keeps (its JSON: type, authors, title, container, year, volume, issue, pages, doi…), looked up where
// the facts are — Crossref for anything with a DOI and for a reference given in words, Open Library for an ISBN, a page's own
// citation tags (the ones Google Scholar reads) for a web address — so no model is asked and nothing is guessed.

const DOI = /\b(10\.\d{4,9}\/[^\s"<>]+[^\s"<>.,;)\]])/i;
const ISBN = /^(?:isbn[:\s-]*)?((?:97[89][\s-]?)?(?:\d[\s-]?){9}[\dx])$/i;

/** What a pasted string is: a DOI (or a link to one), an ISBN, a web address, or a reference in words. */
export function kindOf(q) {
  const s = String(q || '').trim();
  if (!s) return '';
  if (/^(doi:\s*|https?:\/\/(dx\.)?doi\.org\/)?10\.\d{4,9}\//i.test(s)) return 'doi';
  if (ISBN.test(s.replace(/\s+/g, ' '))) return 'isbn';
  if (/^https?:\/\/\S+$/i.test(s) || /^www\.\S+$/i.test(s)) return 'url';
  return 'text';
}

/** The DOI in a string (a link, a whole reference), without what trails it. */
export const doiIn = q => { const m = DOI.exec(String(q || '')); return m ? decodeURIComponent(m[1]).replace(/[.,;]+$/, '') : ''; };

/** Looks a source up. `http(url)` fetches JSON or text straight from the web (Crossref and Open Library allow it); `page(url)` gets
 *  a web page's html when the page is on another site (the app's engine fetches it; the browser version tries directly). Resolves
 *  with { source, from } — from says where it came from (crossref, openlibrary, page) — or null when nothing matched. */
export async function lookup(q, { http = fetchJson, text = fetchText, page = fetchText, today = new Date() } = {}) {
  const s = String(q || '').trim(), kind = kindOf(s);
  if (!kind) return null;
  const doi = doiIn(s);
  if (doi) { const m = await crossref(doi, http); if (m) return { source: fromCrossref(m), from: 'crossref' }; if (kind === 'doi') return null; }
  if (kind === 'isbn') {
    // Open Library's edition, and its authors' names (an edition names them by key, or only its work does)
    const isbn = s.replace(/[^\dx]/gi, '').toUpperCase(), ol = 'https://openlibrary.org', ed = await http(`${ol}/isbn/${isbn}.json`).catch(() => null);
    if (!ed || !ed.title) return null;
    let keys = (ed.authors || []).map(a => a.key).filter(Boolean);
    if (!keys.length && ed.works && ed.works[0]) { const work = await http(ol + ed.works[0].key + '.json').catch(() => null); keys = ((work && work.authors) || []).map(a => a.author && a.author.key).filter(Boolean); }
    const names = await Promise.all(keys.slice(0, 12).map(k => http(ol + k + '.json').then(a => a && a.name || '').catch(() => '')));
    return { source: fromOpenLibrary(Object.assign({}, ed, { authors: names.filter(Boolean).map(name => ({ name })) }), isbn), from: 'openlibrary' };
  }
  if (kind === 'url') {
    const url = /^www\./i.test(s) ? 'https://' + s : s;
    const html = await page(url).catch(() => null);
    if (!html) return null;
    const src = fromPage(html, url, today), inPage = src.doi && await crossref(src.doi, http);
    return inPage ? { source: fromCrossref(inPage), from: 'crossref' } : { source: src, from: 'page' }; // the DOI's record is the work's own
  }
  // a reference in words: Crossref finds the work it names; kept only when its title is really in the words given
  const r = await http('https://api.crossref.org/works?rows=3&query.bibliographic=' + encodeURIComponent(s.slice(0, 600))).catch(() => null);
  const items = (r && r.message && r.message.items) || [], words = fold(s);
  const hit = items.find(m => { const t = fold(titleOf(m)); return t.length > 8 && (words.includes(t) || overlap(t, words) >= 0.85); });
  return hit ? { source: fromCrossref(hit), from: 'crossref' } : null;
}

async function crossref(doi, http) {
  const r = await http('https://api.crossref.org/works/' + encodeURIComponent(doi)).catch(() => null);
  return r && r.message || null;
}
async function fetchJson(url) { const r = await fetch(url, { headers: { Accept: 'application/json' } }); if (!r.ok) throw new Error(r.status); return r.json(); }
async function fetchText(url) { const r = await fetch(url); if (!r.ok) throw new Error(r.status); return r.text(); }

const fold = s => String(s || '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '').replace(/<[^>]+>/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
/** The share of a title's words found in the words given. */
function overlap(title, words) { const t = title.split(' ').filter(w => w.length > 2), have = new Set(words.split(' ')); return t.length ? t.filter(w => have.has(w)).length / t.length : 0; }
const strip = s => String(s || '').replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
const titleOf = m => strip((m.title || [])[0]) + ((m.subtitle || [])[0] ? ': ' + strip(m.subtitle[0]) : '');

const CROSSREF_TYPES = { 'journal-article': 'article', 'book': 'book', 'monograph': 'book', 'edited-book': 'book', 'reference-book': 'book', 'book-chapter': 'chapter',
  'book-section': 'chapter', 'book-part': 'chapter', 'reference-entry': 'chapter', 'proceedings-article': 'conference', 'report': 'report', 'report-component': 'report',
  'dissertation': 'thesis', 'posted-content': 'article', 'magazine-article': 'magazine', 'newspaper-article': 'newspaper' };

/** A Crossref work as a source. */
export function fromCrossref(m) {
  const people = list => (list || []).map(a => a.family ? { last: a.family, first: a.given || '' } : { name: a.name || a.literal || '' }).filter(a => a.last || a.name);
  const parts = ((m.issued || m['published-print'] || m['published-online'] || m.published || {})['date-parts'] || [[]])[0] || [];
  const type = CROSSREF_TYPES[m.type] || 'other';
  const s = { type, authors: people(m.author), title: titleOf(m), container: strip((m['container-title'] || [])[0]), year: parts[0] ? String(parts[0]) : '',
    volume: m.volume || '', issue: m.issue || '', pages: (m.page || '').replace(/-/g, '–'), doi: m.DOI || '' };
  if (m.editor && m.editor.length) s.editors = people(m.editor);
  if (type !== 'article') { if (m.publisher) s.publisher = m.publisher; if (m['publisher-location']) s.place = m['publisher-location']; }
  if (parts[1] && type !== 'article' && type !== 'book' && type !== 'chapter') { s.month = String(parts[1]); if (parts[2]) s.day = String(parts[2]); }
  if (m['edition-number']) s.edition = String(m['edition-number']);
  if (m.ISBN && m.ISBN.length) s.isbn = m.ISBN[0];
  return clean(s);
}

/** An Open Library edition as a source: a book. Its publishers and places are names or { name }; its authors { name }. */
export function fromOpenLibrary(b, isbn) {
  const year = (/\b(1[5-9]\d\d|20\d\d)\b/.exec(b.publish_date || '') || [])[1] || '', name = x => typeof x === 'string' ? x : x && x.name || '';
  const edition = (/^(\d+)/.exec(b.edition_name || '') || [])[1] || '';
  return clean({ type: 'book', authors: (b.authors || []).map(name).filter(Boolean), title: [b.title, b.subtitle].filter(Boolean).join(': '),
    publisher: name((b.publishers || [])[0]), place: name((b.publish_places || [])[0]), year, edition, isbn: isbn || '' });
}

/** A web page as a source, from the tags it carries for citation managers (citation_title, citation_author…), else its Open Graph
 *  and JSON-LD tags and its title: a journal's article page is an article, anything else a web page read today. */
export function fromPage(html, url, today = new Date()) {
  const tags = metaTags(html), metas = name => tags.filter(t => t.name === name.toLowerCase()).map(t => t.content).filter(Boolean);
  const one = (...names) => { for (const n of names) { const v = metas(n)[0]; if (v) return v; } return ''; };
  let ld = {};
  for (const m of html.matchAll(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi)) {
    try { const v = JSON.parse(m[1]); const items = [].concat(v['@graph'] || v); const a = items.find(x => /Article|WebPage|Report|Book/i.test([].concat(x['@type'] || '').join(' '))); if (a) { ld = a; break; } } catch (e) { /* not ours */ }
  }
  const ldAuthors = [].concat(ld.author || []).map(a => typeof a === 'string' ? a : a && a.name).filter(Boolean);
  const authors = metas('citation_author').length ? metas('citation_author') : metas('author').length ? metas('author') : ldAuthors.length ? ldAuthors : metas('article:author').filter(a => !/^https?:/i.test(a));
  const date = one('citation_publication_date', 'citation_date', 'citation_online_date', 'article:published_time', 'dc.date', 'date') || ld.datePublished || '';
  const d = /(\d{4})(?:[-/](\d{1,2})(?:[-/](\d{1,2}))?)?/.exec(date) || [];
  const journal = one('citation_journal_title');
  const site = one('og:site_name') || (ld.publisher && ld.publisher.name) || new URL(url).hostname.replace(/^www\./, '');
  const first = one('citation_firstpage'), last = one('citation_lastpage');
  const pageTitle = entities(((/<title[^>]*>([\s\S]*?)<\/title>/i.exec(html) || [])[1] || ''));
  const s = { type: journal ? 'article' : /NewsArticle/i.test([].concat(ld['@type'] || '').join(' ')) ? 'newspaper' : 'webpage',
    authors: authors.map(a => a.replace(/\s+/g, ' ')), title: strip(one('citation_title', 'og:title', 'twitter:title') || ld.headline || ld.name || pageTitle),
    container: journal || site, year: d[1] || '', month: journal ? '' : d[2] ? String(+d[2]) : '', day: journal ? '' : d[3] ? String(+d[3]) : '',
    volume: one('citation_volume'), issue: one('citation_issue'), pages: first ? (last && last !== first ? first + '–' + last : first) : '',
    doi: doiIn(one('citation_doi', 'dc.identifier', 'prism.doi')), publisher: journal ? '' : one('citation_publisher', 'dc.publisher'), url,
    accessed: journal ? '' : today.toISOString().slice(0, 10) };
  // a site's name in its page title ("Title | Site") is not the page's
  if (s.title && site && s.title.endsWith(site)) s.title = s.title.slice(0, -site.length).replace(/\s*[|\-–—:]\s*$/, '');
  return clean(s);
}

/** A page's meta tags as { name, content }: name from its name, property or itemprop, lower case, the content's entities decoded. */
function metaTags(html) {
  const out = [];
  for (const m of String(html).matchAll(/<meta\b([^>]*)>/gi)) {
    const attrs = {}; for (const a of m[1].matchAll(/([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)'|([^\s"'>]+))/g)) attrs[a[1].toLowerCase()] = a[3] ?? a[4] ?? a[5] ?? '';
    const name = (attrs.name || attrs.property || attrs.itemprop || '').toLowerCase();
    if (name) out.push({ name, content: entities(attrs.content || '').trim() });
  }
  return out;
}
const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ndash: '–', mdash: '—', rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“' };
const entities = s => String(s).replace(/&(#x[\da-f]+|#\d+|\w+);/gi, (m, e) => e[0] === '#' ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENTITIES[e.toLowerCase()] ?? m);

/** The fields worth keeping: empty ones and empty lists go. */
function clean(s) { for (const k of Object.keys(s)) if (s[k] === '' || (Array.isArray(s[k]) && !s[k].length)) delete s[k]; return s; }

/** A short line for a source in a list: its first author and year, and its title. */
export function label(s) {
  const a = (s.authors || [])[0], who = a ? (a.last || a.name || a) : '', etal = (s.authors || []).length > 2 ? ' et al.' : (s.authors || []).length === 2 ? ' & ' + (s.authors[1].last || s.authors[1].name || s.authors[1]) : '';
  return [who && who + etal, s.year && '(' + s.year + ')', s.title].filter(Boolean).join(' ');
}
