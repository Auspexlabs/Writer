// node --test ui/tests/ — grammar.js: the request for a paragraph, the corrections read from a reply, and how they are placed in
// the text (what the model cannot quote exactly is dropped; overlaps and no-ops too) and moved when one is taken.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { prompt, parse, anchor, apply, shift, kindOf, langOf, SYSTEM } from '../grammar.js';

test('the request names the language, carries the paragraph and the conventions, and asks for JSON only', () => {
  const en = prompt('The cat sat on teh mat.');
  assert.equal(en.lang, 'en');
  assert.equal(en.system, SYSTEM);
  assert.match(en.user, /<language>English<\/language>/);
  assert.match(en.user, /<paragraph>\nThe cat sat on teh mat\.\n<\/paragraph>/);
  assert.doesNotMatch(en.user, /<conventions>/);
  const zh = prompt('这篇文章的作者认为，城市的活力来自街道。', { rules: 'MLA 9: the citation before the period.' });
  assert.equal(zh.lang, 'zh');
  assert.match(zh.user, /<language>Chinese<\/language>/);
  assert.match(zh.user, /<conventions>\nMLA 9: the citation before the period\.\n<\/conventions>/);
  assert.equal(langOf('Putnam 1995 的研究表明，社会资本在下降。'), 'zh');
  assert.equal(langOf('Writer 是一个 app'), 'zh');
  assert.equal(langOf('A paragraph with one 词 in it, mostly English words though.'), 'en');
});

test('a reply is read whether bare, fenced or wrapped in words; what is not a correction is left out', () => {
  const bare = parse('[{"find":"teh","replace":"the","why":"spelling","kind":"spelling"}]');
  assert.deepEqual(bare, [{ find: 'teh', replace: 'the', why: 'spelling', kind: 'spelling' }]);
  const fenced = parse('```json\n[{"find":"has went","replace":"has gone","why":"past participle","kind":"grammar"}]\n```');
  assert.equal(fenced[0].replace, 'has gone');
  const wrapped = parse('Here are the corrections:\n[{"find":"very unique","replace":"unique"}]\nDone.');
  assert.deepEqual(wrapped, [{ find: 'very unique', replace: 'unique', why: '', kind: 'grammar' }]);
  assert.deepEqual(parse('[]'), []);
  assert.deepEqual(parse('No errors.'), []);
  assert.deepEqual(parse('[{"replace":"x"}, {"find":"", "replace":"y"}, 3, null]'), []);
  assert.equal(parse('[{"find":"a","replace":"b","kind":"made-up"}]')[0].kind, 'grammar');
});

test('corrections are anchored to the paragraph: exact, with folded quotes, or trimmed; the rest dropped; overlaps and no-ops too', () => {
  const text = 'The cat sat on teh mat. It’s “eyes upon the street” that that matter, matter very unique.';
  const placed = anchor([
    { find: 'teh', replace: 'the', why: 'spelling', kind: 'spelling' },
    { find: "It's \"eyes", replace: "It is \"eyes", why: 'no contraction', kind: 'style' }, // straight quotes for the text's curly ones
    { find: ' that that ', replace: ' that ', why: 'doubled word', kind: 'grammar' },
    { find: 'nowhere in the text', replace: 'x', why: '', kind: 'grammar' },
    { find: 'very unique', replace: 'very unique', why: 'no change', kind: 'style' },
    { find: 'that matter', replace: 'which matter', why: 'overlaps the doubled word', kind: 'grammar' },
    { find: '  mat. ', replace: '  mat! ', why: 'trimmed to what the text has', kind: 'punctuation' },
  ], text);
  assert.deepEqual(placed.map(x => [x.start, x.find, x.replace]), [
    [15, 'teh', 'the'],
    [19, 'mat.', 'mat!'],
    [24, 'It’s “eyes', 'It is "eyes'],
    [51, ' that that ', ' that '],
  ]);
  assert.equal(apply(text, placed[0]), 'The cat sat on the mat. It’s “eyes upon the street” that that matter, matter very unique.');
  assert.equal(kindOf(placed[0]), 'fix');
  assert.equal(kindOf({ replace: '' }), 'delete');
  const rest = shift(placed, placed[1]);
  assert.deepEqual(rest.map(x => [x.start, x.find]), [[15, 'teh'], [24, 'It’s “eyes'], [51, ' that that ']]);
  const after = shift(placed, placed[0]);
  assert.deepEqual(after.map(x => x.start), [19, 24, 51]);
  const two = anchor([{ find: 'cat', replace: 'small cat' }, { find: 'mat', replace: 'rug' }], text), longer = shift(two, two[0]);
  assert.deepEqual(longer.map(x => [x.start, x.find]), [[25, 'mat']]);
});
