// node --test ui/tests/ — md.js: the AI chat panel's Markdown safety, and the Typora-style live shortcuts' matching
// logic (the DOM side that turns a match into real formatting lives in MarkdownEditor.dc.html and isn't unit-tested,
// same as this project's other rich-text editors — see word-editor.test.mjs, sheet-editor.test.mjs).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mdToHtml, inline, inlineTrigger, blockTrigger, renderMath, htmlToMd, mathBlockHtml, loadHljs, highlight, pasteText, wordLists, smartPunct, wordStats, mathError } from '../md.js';

// A small DOM for round trips: parses the well-formed HTML mdToHtml emits (and hand-written fragments) into nodes with
// the handful of APIs htmlToMd uses — enough to check that what was rendered saves back as the same Markdown.
const VOID = /^(br|hr|img|input)$/i;
const dec = s => s.replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const T = v => ({ nodeType: 3, nodeValue: v, parentNode: null, cloneNode() { return T(v); }, remove() { const p = this.parentNode; if (p) { p.childNodes.splice(p.childNodes.indexOf(this), 1); this.parentNode = null; } } });
class N {
  constructor(tag, attrs = {}) {
    this.nodeType = 1; this.tagName = tag.toUpperCase(); this.attrs = attrs; this.childNodes = []; this.parentNode = null; this.style = {};
    (attrs.style || '').split(';').forEach(d => { const i = d.indexOf(':'); if (i > 0) this.style[d.slice(0, i).trim().replace(/-([a-z])/g, (m, c) => c.toUpperCase())] = d.slice(i + 1).trim(); });
    this.classList = { contains: c => (this.attrs.class || '').split(/\s+/).includes(c) };
  }
  get children() { return this.childNodes.filter(c => c.nodeType === 1); }
  get firstChild() { return this.childNodes[0] || null; }
  get textContent() { return this.childNodes.map(c => c.nodeType === 3 ? c.nodeValue : c.textContent).join(''); }
  get innerText() { return this.textContent; }
  get checked() { return 'checked' in this.attrs; }
  getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; }
  hasAttribute(k) { return k in this.attrs; }
  appendChild(c) { if (c.parentNode) c.remove(); c.parentNode = this; this.childNodes.push(c); return c; }
  remove() { const p = this.parentNode; if (p) { p.childNodes.splice(p.childNodes.indexOf(this), 1); this.parentNode = null; } }
  cloneNode() { const n = new N(this.tagName, { ...this.attrs }); this.childNodes.forEach(c => n.appendChild(c.cloneNode())); return n; }
  querySelectorAll(sel) {
    const out = [];
    for (const part of sel.split(',')) {
      const scope = /^\s*:scope\s*>/.test(part), m = /^\s*(?::scope\s*>\s*)?([a-z0-9]*)(?:\.([\w-]+))?(?:\[([\w-]+)(?:=([^\]]+))?\])?\s*$/i.exec(part);
      const ok = n => (!m[1] || n.tagName === m[1].toUpperCase()) && (!m[2] || n.classList.contains(m[2])) && (!m[3] || (m[4] === undefined ? n.hasAttribute(m[3]) : n.getAttribute(m[3]) === m[4].replace(/^["']|["']$/g, '')));
      const walk = n => n.children.forEach(c => { if (ok(c)) out.push(c); if (!scope) walk(c); });
      walk(this);
    }
    return out;
  }
  querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
}
function parse(html) {
  const root = new N('div'); let cur = root, m;
  const re = /<!--[\s\S]*?-->|<\/([a-z0-9]+)\s*>|<([a-z0-9]+)((?:\s+[\w:-]+(?:=(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*(\/?)>|([^<]+)/gi;
  while ((m = re.exec(html))) {
    if (m[1]) cur = cur.parentNode || root;
    else if (m[2]) { const attrs = {}; for (const a of m[3].matchAll(/([\w:-]+)(?:=(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)) attrs[a[1]] = dec(a[2] ?? a[3] ?? a[4] ?? ''); const n = new N(m[2], attrs); cur.appendChild(n); if (!VOID.test(m[2]) && !m[4]) cur = n; }
    else if (m[5]) cur.appendChild(T(dec(m[5])));
  }
  return root;
}
globalThis.document = { createElement: t => new N(t), createTextNode: T };
const roundTrip = md => htmlToMd(parse(mdToHtml(md)));

test('tables: per-column alignment survives a round trip, cells keep <br> line breaks, untouched rows come back as written', () => {
  const md = '| 名称 | 数量 | 备注 |\n| :--- | :---: | ---: |\n| a | 1 | x<br>y |\n| b \\| c | 2 |  |\n';
  const html = mdToHtml(md);
  assert.match(html, /<th style="text-align:left">名称<\/th><th style="text-align:center">数量<\/th><th style="text-align:right">备注<\/th>/);
  assert.match(html, /<td style="text-align:right">x<br>y<\/td>/, 'a literal <br> breaks the line inside a cell');
  assert.equal(roundTrip(md), md);
});

test('code blocks: the fence and its language survive a round trip, highlighted or not; unknown languages stay plain', async () => {
  const md = '```js\nconst a = "<b>" && 1;\n```\n';
  assert.equal(roundTrip(md), md, 'a fresh parse keeps the fence, the language and the escaped code');
  assert.equal(highlight('const a = 1;', 'js'), null, 'nothing is highlighted before the library is loaded');
  await loadHljs();
  const html = highlight('const a = "<b>" && 1;', 'js');
  assert.match(html, /<span class="hljs-keyword">const<\/span>/);
  assert.equal(highlight('x', 'no-such-language'), null);
  assert.equal(highlight('x', ''), null, 'a fence without a language stays plain text');
  assert.equal(htmlToMd(parse('<pre><code data-lang="js">' + html + '</code></pre>')), md, 'the highlighted block saves as the same fence');
});

test('lists: nesting and tasks round-trip, numbers renumber from start, a sibling sublist nests, pastes from other apps become lists', () => {
  const md = '- a\n  - b\n    - [x] c\n- d\n\n3. x\n4. y\n';
  assert.equal(roundTrip(md), md);
  assert.equal(htmlToMd(parse('<ol start="3"><li>a</li><li>b</li><li>c</li></ol>')), '3. a\n4. b\n5. c\n', 'the DOM keeps no numbers: they are written in order from start');
  assert.equal(htmlToMd(parse('<ul><li>a</li><ul><li>b</li></ul><li>c</li></ul>')), '- a\n  - b\n- c\n', 'a sublist beside its item (WebKit indent) belongs to the item above');
  assert.equal(pasteText('• one\n  ◦ two\n• three'), '- one\n  - two\n- three');
  assert.equal(pasteText('just a sentence.'), null, 'prose is left to the browser');
  const word = '<p class="MsoListParagraphCxSpFirst" style="mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span style="font-family:Symbol">·<span>&nbsp;</span></span><!--[endif]-->a</p>'
    + '<p class="MsoListParagraphCxSpMiddle" style="mso-list:l0 level2 lfo1"><!--[if !supportLists]--><span style="mso-list:Ignore">1.<span>&nbsp;</span></span><!--[endif]-->b</p>'
    + '<p class="MsoListParagraphCxSpLast" style="mso-list:l0 level1 lfo1"><!--[if !supportLists]--><span>·<span>&nbsp;</span></span><!--[endif]-->c</p><p class="MsoNormal">after</p>';
  assert.equal(htmlToMd(parse(wordLists(word))), '- a\n  1. b\n- c\n\nafter\n');
  assert.equal(wordLists('<p>plain</p>'), '<p>plain</p>', 'HTML without Word lists is untouched');
});

test('blocks: front matter card, [TOC], GitHub alerts, footnotes with hover text and heading anchors all round-trip unchanged', () => {
  const md = '---\ntitle: 说明\ntags: [a, b]\n---\n\n[TOC]\n\n# 第一章 Intro\n\n> [!WARNING]\n> 小心 *这里*\n\n正文[^1]，再来[^note]。\n\n## 第一章 Intro\n\n[^1]: 第一条脚注\n[^note]: 第二条 **注**\n';
  const html = mdToHtml(md);
  assert.match(html, /^<div class="md-front" contenteditable="false" data-src="title: 说明\ntags: \[a, b\]"><pre>title: 说明\ntags: \[a, b\]<\/pre><\/div>/);
  assert.match(html, /<nav class="md-toc" data-line="5" contenteditable="false"><ul><li style="padding-left:0px"><a href="#第一章-intro">第一章 Intro<\/a><\/li><li style="padding-left:14px"><a href="#第一章-intro-1">第一章 Intro<\/a><\/li><\/ul><\/nav>/);
  assert.match(html, /<h1 data-line="7" id="第一章-intro">/, 'a heading gets its anchor id; lines count from the top of the file, front matter included');
  assert.match(html, /<h2 data-line="14" id="第一章-intro-1">/, 'a repeated heading gets -1');
  assert.match(html, /<blockquote data-line="9" class="md-alert md-alert-warning" data-alert="WARNING"><p class="md-alert-title" contenteditable="false">警告<\/p><p>小心 <em>这里<\/em><\/p><\/blockquote>/);
  assert.match(html, /<sup class="md-fn" contenteditable="false" data-fn="1" title="第一条脚注"><a href="#fn-1">1<\/a><\/sup>，再来<sup class="md-fn" contenteditable="false" data-fn="note" title="第二条 注"><a href="#fn-note">2<\/a><\/sup>。/);
  assert.match(html, /<section class="md-footnotes" data-line="16"><ol><li id="fn-1" data-fn="1">第一条脚注<\/li><li id="fn-note" data-fn="note">第二条 <strong>注<\/strong><\/li><\/ol><\/section>/);
  assert.equal(roundTrip(md), md);
  assert.equal(mdToHtml('看 [^x] 这里'), '<p data-line="0">看 [^x] 这里</p>', 'a mark without a definition stays text');
});

test('inline: bare autolinks, wiki links, sub/superscript and emoji shortcodes render and save back as written; smart punctuation', () => {
  const md = '见 https://a.test/p?x=1&y=2. 还有 [[Notes/Plan|计划]] 与 [[Home]]，H~2~O 和 x^2^，来杯 :coffee: :nope: <https://b.test>\n';
  const html = mdToHtml(md);
  assert.match(html, /见 <a href="https:\/\/a\.test\/p\?x=1&amp;y=2" data-auto="bare" target="_blank" rel="noopener">https:\/\/a\.test\/p\?x=1&amp;y=2<\/a>\. /, 'the sentence\'s period stays outside the link');
  assert.match(html, /<a href="Notes\/Plan" data-wiki="1">计划<\/a> 与 <a href="Home" data-wiki="1">Home<\/a>/);
  assert.match(html, /H<sub>2<\/sub>O 和 x<sup>2<\/sup>/);
  assert.match(html, /<span class="md-emoji" contenteditable="false" data-src=":coffee:">☕<\/span> :nope:/, 'an unknown shortcode stays text');
  assert.equal(roundTrip(md), md);
  assert.equal(roundTrip('[https://a.test](https://a.test) 和 <https://b.test>\n'), '[https://a.test](https://a.test) 和 <https://b.test>\n', 'links already written out keep their form');
  assert.equal((mdToHtml('[x](https://a.test/) https://en.wikipedia.org/wiki/Foo_(bar).').match(/<a /g) || []).length, 2, 'a link is never linked twice, and (…) inside a URL keeps its bracket');
  assert.match(mdToHtml('https://en.wikipedia.org/wiki/Foo_(bar).'), /Foo_\(bar\)<\/a>\./);
  assert.equal(mdToHtml('[[Secret]]', { safe: true }), '<p data-line="0">[[Secret]]</p>', 'wiki links are the editor\'s, never a chat reply\'s');
  assert.deepEqual(inlineTrigger('嗯 :smile:'), { tag: 'emoji', text: 'smile', len: 7 });
  assert.equal(inlineTrigger(':nope:'), null);
  assert.deepEqual(smartPunct('he said "'), { del: 1, text: '“' });
  assert.deepEqual(smartPunct('he said “hi"'), { del: 1, text: '”' });
  assert.deepEqual(smartPunct("don'"), { del: 1, text: '’' });
  assert.deepEqual(smartPunct('a --'), { del: 2, text: '–' });
  assert.deepEqual(smartPunct('a –-'), { del: 2, text: '—' });
  assert.deepEqual(smartPunct('wait...'), { del: 3, text: '…' });
  assert.deepEqual(smartPunct('a ->'), { del: 2, text: '→' });
  assert.equal(smartPunct('---'), null, 'a rule being typed is left alone');
  assert.equal(smartPunct('a-'), null);
});

test('diagrams: a ```mermaid fence is a diagram block keeping its source (drawn later by the editor) and saves back as the fence; a chat reply keeps it as code', () => {
  const md = '```mermaid\ngraph TD\n  A-->B\n```\n';
  assert.equal(mdToHtml(md), '<div class="md-mermaid" data-line="0" contenteditable="false" data-src="graph TD\n  A--&gt;B"><pre><code data-lang="mermaid">graph TD\n  A--&gt;B</code></pre></div>');
  assert.equal(roundTrip(md), md);
  assert.match(mdToHtml(md, { safe: true }), /^<pre data-line="0"><code data-lang="mermaid">/, 'no diagram block for untrusted text');
});

test('files: a picture shown through the engine saves as the relative link it was written with; the HTML export renders math as MathML alone', () => {
  assert.equal(htmlToMd(parse('<p><img alt="图" src="/file?file=%2Fdocs%2Fassets%2Fa.png" data-rel="assets/a.png"></p>')), '![图](assets/a.png)\n');
  const out = mdToHtml('看 $a^2$', { mathml: true });
  assert.match(out, /<math/); assert.ok(!out.includes('katex-html'), 'no HTML layer that would need KaTeX\'s CSS and fonts');
  assert.match(mdToHtml('看 $a^2$'), /katex-html/, 'the editor keeps KaTeX\'s HTML rendering');
});

test('math box: the live error line has KaTeX\'s message for a broken formula and nothing for a good or empty one', () => {
  assert.match(mathError('\\frac{1}'), /expected '\}'/);
  assert.equal(mathError('\\frac{1}{2}'), '');
  assert.equal(mathError('   '), '');
});

test('the status line: words (CJK characters and Latin words), characters without whitespace, minutes at 400 a minute', () => {
  assert.deepEqual(wordStats('# 你好 world\n\n```\nskipped code\n```\n'), { words: 3, chars: 7, minutes: 1 });
});

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

test('a resized picture is Markdown as Typora writes it, <img src width>; unresized it is ![]() again', () => {
  assert.equal(mdToHtml('看 <img src="https://a.test/i.png" alt="图" width="240"> 这里'), '<p data-line="0">看 <img alt="图" src="https://a.test/i.png" width="240"> 这里</p>');
  assert.ok(!mdToHtml('<img src="javascript:alert(1)" width="9">').includes('<img'), 'only http(s) and data pictures, as for ![]()');
  const el = (tagName, attrs, childNodes = []) => ({ nodeType: 1, tagName, classList: { contains: () => false }, style: {}, childNodes, getAttribute: k => k in attrs ? attrs[k] : null, querySelector: () => null });
  assert.equal(htmlToMd({ childNodes: [el('P', {}, [el('IMG', { src: 'a.png', alt: 'x', width: '240' })])] }), '<img src="a.png" alt="x" width="240">\n');
  assert.equal(htmlToMd({ childNodes: [el('P', {}, [el('IMG', { src: 'a.png', alt: 'x' })])] }), '![x](a.png)\n');
});
