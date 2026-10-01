// node --test ui/tests/ — assist.js: a paper's citation style told from the document (its list's title, its citations, its header,
// its notes, its sources), the question for the model and its answer, the decision kept per file, and what each style asks for.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { detect, looksAcademic, classifyPrompt, parseStyle, assistOf, setAssist, moveAssist, grammarRules, chatInstructions, pageText, sourcePrompt, parseSource, STYLES, KEYS } from '../assist.js';

const P = text => ({ kind: 'paragraph', text }), H = text => ({ kind: 'heading', text });
const ESSAY = Array.from({ length: 24 }, (_, i) => P(`Paragraph ${i + 1} of an English essay argues that public libraries, as Oldenburg wrote in 1999, remain the quiet commons of the city, and that this matters for how a civic life is made and kept over time.`));

test('the works-cited list names the style, whatever the text does', () => {
  assert.equal(detect({ blocks: [P('Body.'), H('Works Cited'), P('Jacobs, Jane. The Death and Life of Great American Cities. Random House, 1961.')] }), 'mla');
  assert.equal(detect({ blocks: [P('Body.'), P('References'), P('Jacobs, J. (1961). The death and life.')] }), 'apa');
  assert.equal(detect({ blocks: [P('Body.'), P('Bibliography'), P('Jacobs, Jane. The Death and Life. New York: Random House, 1961.')] }), 'chicago');
  assert.equal(detect({ blocks: [P('正文。'), H('参考文献'), P('[1] 王芳. 协作工具的用户留存研究[J]. 管理评论, 2025.')] }), 'gb7714');
  assert.equal(detect({ blocks: [P('As shown in [1] and [2], the method works.'), H('References'), P('[1] A. Author, "Title," 2020.')] }), 'ieee');
  assert.equal(detect({ blocks: [P('The point stands (Putnam 1995, 67).'), P('It holds (Jacobs 1961, 35).'), H('References')] }), 'chicago-date');
  assert.equal(detect({ blocks: [P('A long paragraph that happens to mention references in passing, like this one does.'), P('More text.')] }), '');
});

test('in-text citations tell the style when there are two or more; one needs an English text', () => {
  assert.equal(detect({ blocks: [P('The trust went with it (Putnam 67). The tone is set by regulars (Oldenburg 16).')] }), 'mla');
  assert.equal(detect({ blocks: [P('The trust went with it (Putnam, 1995, p. 67). Others agree (Jacobs & Smith, 1961).')] }), 'apa');
  assert.equal(detect({ blocks: [P('The trust went with it (Putnam 1995, 67). Others agree (Jacobs 1961).')] }), 'chicago-date');
  assert.equal(detect({ blocks: [P('As shown in [1], [2] and [3-5].')] }), 'ieee');
  assert.equal(detect({ blocks: [P('如文献 [1] 和 [2] 所述，该方法有效。')] }), 'gb7714');
  assert.equal(detect({ blocks: [P('One citation only (Putnam 67) in English prose.')] }), 'mla');
  assert.equal(detect({ blocks: [P('只有一处 (Putnam 67) 的中文段落。')] }), '');
});

test('the header, the notes and the sources say when the text does not', () => {
  assert.equal(detect({ blocks: [P('An essay.')], header: 'Chen {page}' }), 'mla');
  assert.equal(detect({ blocks: [P('An essay.')], header: 'Chen 3' }), 'mla');
  assert.equal(detect({ blocks: [P('An essay.')], header: '{page}' }), '');
  assert.equal(detect({ blocks: [P('An essay.')], header: 'THE QUIET COMMONS {page}' }), 'apa');
  assert.equal(detect({ blocks: ESSAY, notes: 3 }), 'chicago');
  assert.equal(detect({ blocks: [P('一篇中文文章。')], notes: 3 }), '');
  assert.equal(detect({ blocks: ESSAY, notes: 1 }), '');
  assert.equal(detect({ blocks: [P('An essay.')], citeStyle: 'apa', sources: 2 }), 'apa');
  assert.equal(detect({ blocks: [P('An essay.')], citeStyle: 'apa', sources: 0 }), '');
  assert.equal(detect({ blocks: [P('The point (Putnam 67). Again (Jacobs 35).')], citeStyle: 'apa', sources: 2 }), 'mla');
  assert.equal(detect(), '');
});

test('a paper worth asking the model about, the question, and its one-word answer', () => {
  assert.equal(looksAcademic(ESSAY), true);
  assert.equal(looksAcademic(ESSAY.slice(0, 3)), false);
  assert.equal(looksAcademic(Array.from({ length: 30 }, () => P('会议纪要：讨论了下个季度的计划，决定先上线基础版，再逐步补齐功能，相关负责人下周给出排期。'))), false);
  assert.equal(looksAcademic(Array.from({ length: 20 }, () => P('Meeting notes: we agreed to ship the basic version first and fill in the rest of the features over the following weeks.'))), false);
  const q = classifyPrompt([H('The Quiet Commons'), ...ESSAY, H('Works Cited'), P('Jacobs, Jane. The Death and Life.')], 'Chen {page}');
  assert.match(q.system, /exactly one word: mla, apa, chicago, chicago-date, gb7714, ieee, or none/);
  assert.match(q.user, /<header>Chen \{page\}<\/header>/);
  assert.match(q.user, /<opening>\n# The Quiet Commons\nParagraph 1/);
  assert.match(q.user, /<ending>[\s\S]*# Works Cited\nJacobs, Jane\. The Death and Life\.\n<\/ending>/);
  assert.ok(q.user.length < 4800);
  assert.equal(parseStyle('mla'), 'mla');
  assert.equal(parseStyle(' Chicago-Date.\n'), 'chicago-date');
  assert.equal(parseStyle('The style is APA'), '');
  assert.equal(parseStyle('none'), '');
  assert.equal(parseStyle(''), '');
});

test('the decision is kept per file and follows a file that moves', () => {
  const store = new Map(), storage = { getItem: k => store.has(k) ? store.get(k) : null, setItem: (k, v) => store.set(k, v), removeItem: k => store.delete(k) };
  assert.equal(assistOf('a.docx', storage), '');
  setAssist('a.docx', 'mla', storage);
  assert.equal(assistOf('a.docx', storage), 'mla');
  moveAssist('a.docx', 'b/c.docx', storage);
  assert.equal(assistOf('a.docx', storage), '');
  assert.equal(assistOf('b/c.docx', storage), 'mla');
  setAssist('b/c.docx', 'off', storage);
  assert.equal(assistOf('b/c.docx', storage), 'off');
  setAssist('b/c.docx', '', storage);
  assert.equal(assistOf('b/c.docx', storage), '');
  assert.equal(assistOf('x', { getItem: () => { throw new Error('no storage'); } }), '');
});

test('every style has a name, a list, conventions for the grammar check and instructions for the assistant', () => {
  for (const k of KEYS) {
    assert.ok(STYLES[k].name && STYLES[k].list, k);
    assert.ok(grammarRules(k).length > 40, k + ' rules');
    const chat = chatInstructions(k);
    assert.match(chat, /writing assistance/, k);
    assert.match(chat, /--type citation|--type footnote/, k + ' tells how to cite');
    assert.match(chat, /--type bibliography/, k + ' tells about the list');
  }
  assert.match(grammarRules('mla'), /before the period/);
  assert.match(grammarRules('chicago'), /after the period/);
  assert.match(grammarRules('gb7714'), /全角标点/);
  assert.equal(grammarRules(''), '');
  assert.equal(chatInstructions('nope'), '');
});

test('a source the model reads: the page as text, the question, the object back with its people in the engine form', () => {
  const html = '<html><head><title> A Page Title </title><meta name="author" content="Eric Klinenberg"><meta property="og:site_name" content="The Times"><script>x()</script><style>p{}</style></head><body><nav>menu</nav><h1>Headline</h1><p>First&nbsp;paragraph &amp; more.</p><p>Second.</p></body></html>';
  const text = pageText(html);
  assert.match(text, /^title: A Page Title\nauthor: Eric Klinenberg\nsite: The Times\n\nHeadline\nFirst paragraph & more\.\nSecond\.$/);
  assert.equal(pageText('<p>' + 'word '.repeat(3000) + '</p>', 100).length, 100);
  const q = sourcePrompt('https://example.org/a', 'title: A Page');
  assert.match(q.system, /Only what the text states/);
  assert.match(q.user, /<pasted>\nhttps:\/\/example\.org\/a\n<\/pasted>\n<page>\ntitle: A Page\n<\/page>/);
  assert.doesNotMatch(sourcePrompt('Jacobs 1961').user, /<page>/);
  const src = parseSource('```json\n{"type":"newspaper","authors":[{"last":"Klinenberg","first":"Eric"}, "Smith, Ann", "Pew Research Center", {"name":""}],"title":"To Restore Civil Society","container":"The New York Times","year":"2018","month":"9","day":8,"pages":"","url":"https://example.org/a","extra":"dropped"}\n```');
  assert.deepEqual(src, { type: 'newspaper', title: 'To Restore Civil Society', container: 'The New York Times', year: '2018', month: '9', day: '8', url: 'https://example.org/a',
    authors: [{ last: 'Klinenberg', first: 'Eric' }, { last: 'Smith', first: 'Ann' }, { name: 'Pew Research Center' }] });
  assert.equal(parseSource('{"type":"made-up","title":"T","url":"https://x"}').type, 'webpage');
  assert.equal(parseSource('{"title":"T","doi":"10.1/x"}').type, 'article');
  assert.equal(parseSource('{"title":"T"}').type, 'other');
  assert.equal(parseSource('{"authors":[{"last":"A"}]}'), null);
  assert.equal(parseSource('not json'), null);
});
