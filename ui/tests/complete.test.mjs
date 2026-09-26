// node --test ui/tests/ — AI 自动补全 (ui/complete.js): what a suggestion is asked with and where it may show. The grey text itself,
// Tab and typing through it are checked in a browser (the editors' own pages); here the parts that decide them.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Element, parseNodes } from './dom-stub.mjs';
import { cssString, pathSelector, context, fieldCaret, advance, surfaceOf } from '../complete.js';

const root = html => { const r = new Element('div'); r.append(...parseNodes(html)); return r; };

test('the grey text as CSS content: quotes and backslashes escaped, line breaks as spaces', () => {
  assert.equal(cssString('say "hi" \\ bye'), '"say \\"hi\\" \\\\ bye"');
  assert.equal(cssString('a\nb\r\nc'), '"a b  c"');
  assert.equal(cssString('中文，标点'), '"中文，标点"');
});

test('pathSelector: one :nth-child per level from the root; the root itself is ""; outside it, null', () => {
  const r = root('<p>one</p><ul><li>a</li><li><b>b</b></li></ul><p>two</p>');
  const li = r.children[1].children[1];
  assert.equal(pathSelector(r, li), '>:nth-child(2)>:nth-child(2)');
  assert.equal(pathSelector(r, r.children[2]), '>:nth-child(3)');
  assert.equal(pathSelector(r, r), '');
  assert.equal(pathSelector(r, new Element('p')), null);
});

test('context: the blocks before the caret\'s block (it last) and after it, one per line, capped at each end', () => {
  const r = root('<h1>标题</h1><p>第一段。</p><ul><li>要点一</li><li>要点二</li><li>要点三</li></ul><p>结尾​</p>');
  const li = r.children[2].children[1];
  assert.deepEqual(context(r, li), { before: '标题\n第一段。\n要点一\n要点二', after: '要点三\n结尾\n' });
  const long = root('<p>' + 'x'.repeat(50) + '</p><p>end</p>');
  assert.deepEqual(context(long, long.children[1], 20, 5), { before: 'x'.repeat(16) + '\nend', after: '' });
});

test('fieldCaret: a suggestion needs the caret at the end of the value, or of its line in a textarea', () => {
  assert.equal(fieldCaret('Hello', 5, 5, false), 5);
  assert.equal(fieldCaret('Hello  ', 5, 5, false), 5, 'only spaces after it');
  assert.equal(fieldCaret('Hello', 3, 3, false), -1, 'in the middle of the value');
  assert.equal(fieldCaret('Hello', 1, 4, false), -1, 'a selection');
  assert.equal(fieldCaret('one\ntwo\nthree', 3, 3, true), 3, 'the end of a line in a textarea');
  assert.equal(fieldCaret('one\ntwo', 3, 3, false), -1, 'an input has one line: the rest is text');
  assert.equal(fieldCaret('', null, null, false), -1);
});

test('advance: typing the suggestion\'s own letters walks through it; anything else drops it', () => {
  assert.equal(advance('world today', 'w'), 'orld today');
  assert.equal(advance('world today', 'world '), 'today');
  assert.equal(advance('world', 'world'), null, 'typed all of it: nothing left to show');
  assert.equal(advance('world', 'x'), null);
  assert.equal(advance('world', ''), null);
  assert.equal(advance('开会讨论', '开会'), '讨论');
});

test('surfaceOf: only what an editor marked with data-complete, and only while it can be typed in', () => {
  const el = (tag, attrs, props) => Object.assign(new Element(tag, attrs), props || {});
  const body = el('div', { 'data-complete': 'the body of a Word document' }, { isContentEditable: true });
  const p = el('p'); body.appendChild(p); const text = { nodeType: 3, parentElement: p };
  assert.equal(surfaceOf(text), body, 'a text node in the Word body');
  assert.equal(surfaceOf(el('div', { 'data-complete': 'x' }, { isContentEditable: false })), null, 'a slide text box that is not being edited');
  const cell = el('input', { 'data-complete': 'a spreadsheet cell' });
  assert.equal(surfaceOf(cell), cell);
  assert.equal(surfaceOf(el('input', { 'data-complete': '1', type: 'password' })), null, 'never a password');
  assert.equal(surfaceOf(el('input', { 'data-complete': '1', type: 'number' })), null);
  assert.equal(surfaceOf(el('textarea', { 'data-complete': '1' }, { readOnly: true })), null, 'a read-only PDF form field');
  assert.equal(surfaceOf(el('input', { placeholder: '查找' })), null, 'the find box is not marked');
  assert.equal(surfaceOf(null), null);
});

test('install: once per page; a second install (设置 inside the shell\'s page) shares the first, and its remove leaves it be', async () => {
  const { install } = await import('../complete.js');
  const listeners = new Map(), on = (t, f) => listeners.set(t + ':' + (listeners.size), f);
  const doc = { head: new Element('head'), body: new Element('body'), createElement: t => new Element(t), addEventListener: on, removeEventListener() { }, defaultView: { addEventListener() { }, removeEventListener() { }, getComputedStyle: () => ({}) } };
  const first = install(doc, { enabled: () => true, request: async () => '' });
  const count = listeners.size;
  const second = install(doc, { enabled: () => false, request: async () => 'x' });
  assert.equal(listeners.size, count, 'no second set of listeners');
  assert.equal(second.state, first.state);
  second.remove();
  assert.equal(doc.__writerComplete, first, 'the first stays installed');
  assert.equal(doc.head.children.length, 1, 'its style is still there');
  first.remove();
  assert.equal(doc.__writerComplete, undefined);
  assert.equal(doc.head.children.length, 0);
});
