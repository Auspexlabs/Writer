// node --test ui/tests/ — md.js: the AI chat panel's Markdown safety, and the Typora-style live shortcuts' matching
// logic (the DOM side that turns a match into real formatting lives in MarkdownEditor.dc.html and isn't unit-tested,
// same as this project's other rich-text editors — see word-editor.test.mjs, sheet-editor.test.mjs).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mdToHtml, inline, inlineTrigger, blockTrigger, renderMath, htmlToMd, mathBlockHtml } from '../md.js';

test('inline math from a file keeps its source, so saving writes $…$ back instead of the rendered text', () => {
  assert.match(mdToHtml('公式 $a^2$ 在这'), /<span class="md-math"[^>]*data-src="a\^2"/);
  const span = { nodeType: 1, tagName: 'SPAN', classList: { contains: c => c === 'md-math' }, getAttribute: k => (k === 'data-src' ? 'a^2' : null) };
  assert.match(htmlToMd({ childNodes: [{ nodeType: 1, tagName: 'P', classList: { contains: () => false }, childNodes: [span] }] }), /\$a\^2\$/);
});

test('block shortcuts also fire when the typed space arrived as a no-break space', () => {
  assert.deepEqual(blockTrigger('## '), { type: 'h2', len: 3 });
  assert.deepEqual(blockTrigger('- '), { type: 'ul', len: 2 });
});

test('a quote inside a link, image or autolink cannot open a new attribute, and script links never render', () => {
  // the attribute names a browser would see: each name="value" pair, where a value ends at the first raw quote
  const attrs = html => [...html.matchAll(/<(?:a|img)\b([^>]*)>/g)].flatMap(m => [...m[1].matchAll(/\s([a-z-]+)="[^"]*"/gi)].map(x => x[1]));
  for (const safe of [true, false]) {
    const out = mdToHtml('[x](https://a.test/"onmouseover="alert(1)) ![y" onerror="alert(2)](https://a.test/i.png) <https://a.test/"onfocus="z>', { safe });
    assert.deepEqual(attrs(out).filter(n => /^on/i.test(n)), [], `no event-handler attribute (safe=${safe})`);
    assert.ok(!mdToHtml('[x](javascript:alert(1))', { safe }).includes('href'), `no javascript: link (safe=${safe})`);
  }
  assert.match(mdToHtml('![a](https://a.test/i.png "caption")'), /<img alt="a" src="https:\/\/a.test\/i.png">/, 'an image with a title still renders');
});

test('safe mode (chat replies): raw HTML in the text never becomes a real tag or attribute', () => {
  const img = mdToHtml('<img src=x onerror="alert(1)">', { safe: true });
  assert.ok(!/<img/i.test(img), 'no real <img> with an onerror attribute');
  assert.match(img, /&lt;img/);

  const script = mdToHtml('ignore me <script>alert(1)</script> please', { safe: true });
  assert.ok(!/<script/i.test(script), 'no real <script> tag');
  assert.match(script, /&lt;script&gt;/);
});

test('safe mode: only http/https/mailto links survive as real links, others fall back to plain text', () => {
  const bad = mdToHtml("[click me](javascript:window.location='https://evil.example')", { safe: true });
  assert.ok(!/href\s*=\s*"javascript:/i.test(bad), 'javascript: never reaches an href');
  assert.match(bad, />click me</, 'the label still shows, just not as a link'); // no <a> wrapper — nothing to click into a script

  const good = mdToHtml('[docs](https://example.com/a)', { safe: true });
  assert.match(good, /<a href="https:\/\/example\.com\/a" target="_blank" rel="noopener">docs<\/a>/);
  const mail = mdToHtml('[me](mailto:a@b.com)', { safe: true });
  assert.match(mail, /<a href="mailto:a@b\.com"/);
});

test('permissive mode (the user\'s own documents) is unchanged: relative links and pasted data: images still work', () => {
  assert.match(mdToHtml('[a](notes/b.md)'), /<a href="notes\/b\.md"/, 'a relative link used to work and still does');
  assert.match(mdToHtml('![x](data:image/png;base64,AAAA)'), /<img alt="x" src="data:image\/png;base64,AAAA">/, 'a locally pasted image is not an http(s) URL and must still render');
});

test('inline math $...$ renders via KaTeX; a literal price is left alone', () => {
  const html = mdToHtml('costs $5 and $10 today');
  assert.equal(html, '<p data-line="0">costs $5 and $10 today</p>', 'no space-adjacent $ pair is not math, exactly like the * / _ emphasis rule');
  const math = mdToHtml('let $x+y=z$ hold');
  assert.match(math, /<span class="katex">/, 'a real formula renders');
  assert.ok(!math.includes('$x+y=z$'), 'the raw source is gone once it renders');
});

test('display math $$...$$ is its own block, keeps the source for a round trip, and centers via katex-display', () => {
  const html = mdToHtml('before\n\n$$\nE=mc^2\n$$\n\nafter');
  assert.match(html, /<div class="md-math-block"[^>]*data-src="E=mc\^2"[^>]*><span class="katex-display">/);
  assert.match(html, /<p data-line="0">before<\/p>/);
  assert.match(html, /<p data-line="6">after<\/p>/);
});

test('a malformed formula never throws and shows its source in red with the error on hover', () => {
  assert.doesNotThrow(() => renderMath('\\frac{1}', false));
  const bad = renderMath('\\frac{1}', false);
  assert.match(bad, /class="katex-error"/);
  assert.match(bad, /title="ParseError/);
  assert.doesNotThrow(() => mdToHtml('$$\n\\left(\n$$'));
});

test('KaTeX\'s trust:false default holds: \\href cannot smuggle a link into the output', () => {
  const html = renderMath("\\href{javascript:alert(1)}{click}", false);
  assert.ok(!/href="javascript:/i.test(html), 'no live javascript: href from inside a formula');
});

test('htmlToMd round-trips a $$ block: both once it is rendered and while it is still open for editing', () => {
  const rendered = { nodeType: 1, tagName: 'DIV', classList: { contains: c => c === 'md-math-block' }, getAttribute: k => (k === 'data-src' ? 'E=mc^2' : null), querySelector: () => null };
  assert.equal(htmlToMd({ childNodes: [rendered] }), '$$\nE=mc^2\n$$\n');

  const editing = { nodeType: 1, tagName: 'DIV', classList: { contains: c => c === 'md-math-block' }, getAttribute: k => (k === 'data-src' ? 'E=mc^2' : null), querySelector: () => ({ value: 'E=mc^2 + \\Delta' }) };
  assert.equal(htmlToMd({ childNodes: [editing] }), '$$\nE=mc^2 + \\Delta\n$$\n', 'an in-progress edit is saved, not the stale last-rendered source');

  // and mdToHtml -> (open for edit) -> mathBlockHtml is the same shape a real close-and-recommit produces
  assert.match(mathBlockHtml('a^2+b^2'), /data-src="a\^2\+b\^2"/);
});

test('live inline shortcuts: the trigger fires on the closing marker, for every documented form', () => {
  assert.deepEqual(inlineTrigger('**bold**'), { tag: 'strong', text: 'bold', len: 8 });
  assert.deepEqual(inlineTrigger('__bold__'), { tag: 'strong', text: 'bold', len: 8 });
  assert.deepEqual(inlineTrigger('*x*'), { tag: 'em', text: 'x', len: 3 });
  assert.deepEqual(inlineTrigger('*吧*'), { tag: 'em', text: '吧', len: 3 }, 'the pinyin repro: 吧 arrives as one committed character, same as any other');
  assert.deepEqual(inlineTrigger('_x_'), { tag: 'em', text: 'x', len: 3 });
  assert.equal(inlineTrigger('foo_bar_'), null, 'an underscore inside a word is not emphasis');
  assert.deepEqual(inlineTrigger('`code`'), { tag: 'code', text: 'code', len: 6 });
  assert.deepEqual(inlineTrigger('~~gone~~'), { tag: 'del', text: 'gone', len: 8 });
  assert.deepEqual(inlineTrigger('$x+y$'), { tag: 'math', text: 'x+y', len: 5 });
  assert.equal(inlineTrigger('$5 and $10'), null, 'a price is not a trigger either, live or in the static renderer');
});

test('live block shortcuts: a trigger only at the very start of the (otherwise empty) paragraph', () => {
  assert.deepEqual(blockTrigger('# '), { type: 'h1', len: 2 });
  assert.deepEqual(blockTrigger('###### '), { type: 'h6', len: 7 });
  assert.deepEqual(blockTrigger('- '), { type: 'ul', len: 2 });
  assert.deepEqual(blockTrigger('* '), { type: 'ul', len: 2 });
  assert.deepEqual(blockTrigger('1. '), { type: 'ol', len: 3 });
  assert.deepEqual(blockTrigger('> '), { type: 'quote', len: 2 });
  assert.equal(blockTrigger('hello '), null);
});
