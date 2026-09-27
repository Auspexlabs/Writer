import { test } from 'node:test';
import assert from 'node:assert/strict';
import { install } from './dom-stub.mjs';
import { editableText, textMatches } from '../text-find.js';
const parse = install();
test('find joins formatting runs but separates paragraphs, cells and noneditable objects', () => {
  const root = parse('<p><b>Hel</b><i>lo</i> <span contenteditable="false">Hello preview</span> HELLO</p><table><tr><td>A</td><td>B</td></tr></table><p><del>Hello deleted</del>after</p>');
  const { text } = editableText(root);
  assert.equal(textMatches(text, 'hello').length, 2);
  assert.equal(textMatches(text, 'preview').length, 0);
  assert.equal(textMatches(text, 'deleted').length, 0);
  assert.equal(textMatches(text, 'AB').length, 0);
  assert.equal(textMatches(text, 'hello', { caseSensitive: true }).length, 0);
});
test('whole words handle Unicode letters and literal regex punctuation', () => {
  assert.equal(textMatches('hello shelloworld HELLO hello_ 𐐀hello hello𐐀', 'hello', { wholeWord: true }).length, 2);
  assert.equal(textMatches('a+b aab a+b', 'a+b').length, 2);
  assert.equal(textMatches('中文 中文字', '中文', { wholeWord: true }).length, 1);
});
