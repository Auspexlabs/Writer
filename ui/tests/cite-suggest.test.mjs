// node --test ui/tests/ — which sentences want a citation (cite-suggest.js): found from the document's own sources — an author named,
// the words of a title the paper does not use everywhere, the year — or from the claim a sentence makes, with no model asked.
import { test } from 'node:test';
import assert from 'node:assert/strict';

const C = await import('../cite-suggest.js');

const PEGG = { tag: 'Peg15', type: 'article', authors: [{ last: 'Pegg', first: 'Ian L.' }], title: 'Behavior of technetium in nuclear waste vitrification processes', year: '2015' };
const LUKSIC = { tag: 'Luk18', type: 'article', authors: ['Steven A. Luksic', 'Kim, Dong-Sang'], title: 'Effect of technetium-99 sources on its retention in low activity waste glass', year: '2018' };
const REDOX = { tag: 'Sod14', type: 'article', authors: [{ last: 'Soderquist', first: 'Chuck Z.' }], title: 'Redox-dependent solubility of technetium in low activity waste glass', year: '2014' };

test('sentences: a period after et al., an initial or an abbreviation does not end one; Chinese punctuation does', () => {
  const text = 'I. L. Pegg et al. found that 35% of it stays in the glass. It is retained (Pegg 288). “Glass is forever,” she wrote! 研究表明玻璃能留住锝。然后呢？';
  assert.deepEqual(C.sentencesOf(text).map(s => s.text), ['I. L. Pegg et al. found that 35% of it stays in the glass.', 'It is retained (Pegg 288).', '“Glass is forever,” she wrote!', '研究表明玻璃能留住锝。', '然后呢？']);
  assert.deepEqual(C.sentencesOf('No period at the end yet').map(s => [s.start, s.end]), [[0, 24]]);
});

test('a source by its words: the authors\' surnames however they are written, its title\'s words in their plain form, its year', () => {
  assert.deepEqual(C.termsOf(LUKSIC), { tag: 'Luk18', names: ['Luksic', 'Kim'], title: ['technetium', 'source', 'retention', 'activity', 'waste', 'glass'], year: '2018' });
  assert.deepEqual(C.termsOf(PEGG).title, ['behavior', 'technetium', 'nuclear', 'waste', 'vitrification', 'process'], 'processes and process meet');
});

test('the sentences to mark: one naming an author, one on a source\'s own subject, a quotation with no source; not the paper\'s subject everywhere, not one cited', () => {
  const paras = [
    { text: 'Technetium is hard to keep in nuclear waste glass because it is volatile in the melter. Most of the technetium leaves with the off-gas when the melt runs hot and oxidizing. Behavior during vitrification depends on the processes in the cold cap above the melt.', cited: [] },
    { text: 'Luksic and his colleagues traced where the technetium went in their melter runs. The technetium in the glass rose when the feed held less nitrate, as the plant data show (Pegg 288). A reviewer called it “the hardest problem in the whole flowsheet of the plant.” Short one here.', cited: [] }];
  paras[1].cited = [paras[1].text.indexOf('(Pegg 288)')]; // the citation's place in the text (its own text is not the paragraph's)
  paras[1].text = paras[1].text.replace(' (Pegg 288)', '');
  const got = C.suggest(paras, [PEGG, LUKSIC, REDOX]);
  const said = got.map(x => [x.para, paras[x.para].text.slice(x.start, x.end).slice(0, 30), x.tag]);
  assert.deepEqual(said, [[0, 'Behavior during vitrification ', 'Peg15'], [1, 'Luksic and his colleagues trac', 'Luk18'], [1, 'A reviewer called it “the hard', '']]);
  assert.deepEqual(C.suggest(paras, [PEGG, LUKSIC, REDOX], new Set(['Luksic and his colleagues traced where the technetium went in their melter runs.'])).map(x => x.tag), ['Peg15', ''], 'the writer said no to that one');
  assert.deepEqual(C.suggest([{ text: 'We then describe the layout of the rest of this report in the next section.', cited: [] }], [PEGG]), [], 'nothing to go on');
});

test('where the citation goes: before the closing period, after a closing quotation mark with the period moved out, after it all in notes', () => {
  const at = (t, notes) => C.placeOf(t, t.length, notes);
  assert.deepEqual(at('It stays in the glass.'), { at: 21, drop: -1, after: '' });
  assert.deepEqual(at('She called it “the hardest problem.”'), { at: 36, drop: 34, after: '.' });
  assert.deepEqual(at('Why does it “leave the melt?”'), { at: 29, drop: -1, after: '.' });
  assert.deepEqual(at('It stays in the glass.', true), { at: 22, drop: -1, after: '' });
  assert.deepEqual(at('研究表明玻璃能留住锝。'), { at: 10, drop: -1, after: '' });
  assert.deepEqual(at('Still typing it'), { at: 15, drop: -1, after: '' });
});
