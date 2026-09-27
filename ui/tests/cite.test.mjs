// node --test ui/tests/ — looking sources up (cite.js): a DOI, an ISBN, a web address or a reference in words becomes a source the
// engine keeps, from where the facts are (Crossref, Open Library, the page's own citation tags). The web is stood in for by the
// records Crossref and Open Library really answer.
import { test } from 'node:test';
import assert from 'node:assert/strict';

const C = await import('../cite.js');

// Crossref's records of two sources of the annotated bibliography the citation work started from (fields as Crossref sends them)
const PEGG = { type: 'journal-article', title: ['Behavior of technetium in nuclear waste vitrification processes'], subtitle: [], 'container-title': ['Journal of Radioanalytical and Nuclear Chemistry'],
  author: [{ given: 'Ian L.', family: 'Pegg', sequence: 'first' }], volume: '305', issue: '1', page: '287-292', DOI: '10.1007/s10967-014-3900-9', publisher: 'Springer Science and Business Media LLC', issued: { 'date-parts': [[2015, 1, 14]] } };
const LUKSIC = { type: 'journal-article', title: ['Effect of Technetium-99 sources on its retention in low activity waste glass'], 'container-title': ['Journal of Nuclear Materials'],
  author: [{ given: 'Steven A.', family: 'Luksic' }, { given: 'Dong-Sang', family: 'Kim' }], volume: '503', page: '235-244', DOI: '10.1016/j.jnucmat.2018.02.019', issued: { 'date-parts': [[2018, 5]] } };
const REDOX = { type: 'journal-article', title: ['Redox-dependent solubility of technetium in low activity waste glass'], 'container-title': ['Journal of Nuclear Materials'], author: [{ given: 'Chuck Z.', family: 'Soderquist' }],
  volume: '449', issue: '1-3', page: '173-180', DOI: '10.1016/j.jnucmat.2014.03.008', issued: { 'date-parts': [[2014, 6]] } };

/** A stand-in for the web: answers the addresses it knows, fails the others; records what was asked. */
function web(pages) {
  const asked = [];
  const answer = async url => { asked.push(url); const k = Object.keys(pages).find(p => url.startsWith(p)); if (!k) throw new Error('404 ' + url); return pages[k]; };
  return { http: answer, page: answer, asked };
}

test('what is pasted: a DOI or a link to one, an ISBN, a web address, a reference in words', () => {
  assert.deepEqual(['10.1007/s10967-014-3900-9', 'https://doi.org/10.1007/s10967-014-3900-9', 'doi: 10.1007/x', '978-0-226-45808-3', 'ISBN 0226458083', 'https://www.hanford.gov/page.cfm/AboutHanford', 'Pegg, I. L. (2015). Behavior of technetium.', '']
    .map(C.kindOf), ['doi', 'doi', 'doi', 'isbn', 'isbn', 'url', 'text', '']);
  assert.equal(C.doiIn('Soderquist, C. Z. (2014). Redox. Journal, 449(1–3), 173–180. https://doi.org/10.1016/j.jnucmat.2014.03.008'), '10.1016/j.jnucmat.2014.03.008', 'the period ending the reference is not the DOI\'s');
});

test('a Crossref record becomes a source: its type, people, title, journal, year, volume, issue and pages', () => {
  assert.deepEqual(C.fromCrossref(PEGG), { type: 'article', authors: [{ last: 'Pegg', first: 'Ian L.' }], title: 'Behavior of technetium in nuclear waste vitrification processes',
    container: 'Journal of Radioanalytical and Nuclear Chemistry', year: '2015', volume: '305', issue: '1', pages: '287–292', doi: '10.1007/s10967-014-3900-9' }, 'a journal\'s publisher and month are not the article\'s');
  const chapter = C.fromCrossref({ type: 'book-chapter', title: ['Glass <i>chemistry</i>'], 'container-title': ['Waste forms'], author: [{ name: 'Pacific Northwest National Laboratory' }], editor: [{ given: 'A.', family: 'Kruger' }],
    publisher: 'Springer', 'publisher-location': 'Cham', issued: { 'date-parts': [[2020]] }, page: '1-20', ISBN: ['9783030000000'] });
  assert.deepEqual(chapter, { type: 'chapter', authors: [{ name: 'Pacific Northwest National Laboratory' }], title: 'Glass chemistry', container: 'Waste forms', year: '2020', pages: '1–20',
    editors: [{ last: 'Kruger', first: 'A.' }], publisher: 'Springer', place: 'Cham', isbn: '9783030000000' });
});

test('lookup: a DOI from Crossref, a reference in words only when Crossref finds that very work', async () => {
  const w = web({ 'https://api.crossref.org/works/10.1007': { message: PEGG }, 'https://api.crossref.org/works?': { message: { items: [LUKSIC, REDOX] } } });
  assert.deepEqual(await C.lookup('https://doi.org/10.1007/s10967-014-3900-9', w), { source: C.fromCrossref(PEGG), from: 'crossref' });
  const words = 'Luksic, S. A., Kim, D.-S., Um, W., Wang, G., Schweiger, M. J., Soderquist, C. Z., Lukens, W., & Kruger, A. A. (2018). Effect of technetium-99 sources on its retention in low activity waste glass. Journal of Nuclear Materials, 503, 235–244.';
  assert.equal((await C.lookup(words, w)).source.doi, '10.1016/j.jnucmat.2018.02.019');
  assert.equal(await C.lookup('Smith, J. (2020). A book nobody wrote about glass kilns in the valley.', w), null, 'Crossref\'s nearest works do not carry that title: nothing, rather than the wrong source');
  assert.ok(w.asked.some(u => u.includes('query.bibliographic=Smith')));
});

test('lookup: an ISBN from Open Library, its authors named', async () => {
  const w = web({ 'https://openlibrary.org/isbn/9780226458083.json': { title: 'The structure of scientific revolutions', authors: [{ key: '/authors/OL531166A' }], publish_date: '1996', publishers: ['University of Chicago Press'], edition_name: '3rd ed.' },
    'https://openlibrary.org/authors/OL531166A.json': { name: 'Thomas S. Kuhn' } });
  assert.deepEqual(await C.lookup('978-0-226-45808-3', w), { source: { type: 'book', authors: ['Thomas S. Kuhn'], title: 'The structure of scientific revolutions', publisher: 'University of Chicago Press', year: '1996', edition: '3', isbn: '9780226458083' }, from: 'openlibrary' });
});

test('lookup: a web page from its citation tags, or Crossref when the page names its DOI; else its Open Graph tags, read today', async () => {
  const article = '<html><head><meta name="citation_title" content="Behavior of technetium in nuclear waste vitrification processes"><meta name="citation_doi" content="10.1007/s10967-014-3900-9"></head></html>';
  const news = `<html><head><title>Tank leaks at Hanford &amp; the river | The Seattle Times</title><meta property="og:site_name" content="The Seattle Times">
    <meta name="author" content="Hal Bernton"><meta property="article:published_time" content="2023-03-05T10:00:00Z">
    <script type="application/ld+json">{"@type":"NewsArticle","headline":"Tank leaks at Hanford & the river"}</script></head></html>`;
  const w = web({ 'https://link.springer.com/article/1': article, 'https://api.crossref.org/works/10.1007': { message: PEGG }, 'https://www.seattletimes.com/x': news });
  assert.deepEqual(await C.lookup('https://link.springer.com/article/1', w), { source: C.fromCrossref(PEGG), from: 'crossref' });
  const r = await C.lookup('https://www.seattletimes.com/x', Object.assign({ today: new Date('2026-09-27T12:00:00Z') }, w));
  assert.deepEqual(r, { source: { type: 'newspaper', authors: ['Hal Bernton'], title: 'Tank leaks at Hanford & the river', container: 'The Seattle Times', year: '2023', month: '3', day: '5',
    url: 'https://www.seattletimes.com/x', accessed: '2026-09-27' }, from: 'page' });
  assert.equal(await C.lookup('https://gone.example/page', w), null, 'a page that cannot be read: nothing');
});

test('a source\'s line in a list', () => {
  assert.equal(C.label(C.fromCrossref(PEGG)), 'Pegg (2015) Behavior of technetium in nuclear waste vitrification processes');
  assert.equal(C.label(C.fromCrossref(LUKSIC)), 'Luksic & Kim (2018) Effect of Technetium-99 sources on its retention in low activity waste glass');
});
