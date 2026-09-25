// node --test ui/tests/ — character formatting the editor adds beyond bold and colour is a change the save sees: superscript and
// subscript, character spacing (letter-spacing), Word's outline and shadow effects, and a highlight pen colour.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { install } from './dom-stub.mjs';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const parse = install();
const EN = await import('../engine.js');

test('sup / sub, letter-spacing, text effects and a highlight each make the runs differ from plain text', () => {
  const plain = EN.runsOf(parse('<p>x2 wide shade line pen</p>'));
  const marked = EN.runsOf(parse('<p>x<sup>2</sup> <span style="letter-spacing:2pt">wide</span> <span style="text-shadow:1px 1px 0 #B3B3B3">shade</span> <span style="-webkit-text-stroke:0.5px currentColor">line</span> <span style="background-color:#FFFF00">pen</span></p>'));
  assert.notEqual(plain, marked);
  const runs = JSON.parse(marked);
  assert.equal(runs.length, 10, 'each mark is a run of its own between plain stretches');
  assert.ok(new Set(runs.map(r => r.s)).size >= 6, 'and each mark has a key of its own');
  assert.equal(EN.runsOf(parse('<p>a<sup>2</sup></p>')), EN.runsOf(parse('<p>a<sup><span>2</span></sup></p>')), 'wrapping alone changes nothing');
});

test('a slide run\'s capitals, small capitals and underline line are changes too', () => {
  const plain = EN.runsOf(parse('<p>A b c</p>'));
  assert.notEqual(plain, EN.runsOf(parse('<p><span style="text-transform:uppercase">A</span> b c</p>')));
  assert.notEqual(plain, EN.runsOf(parse('<p>A <span style="font-variant:small-caps">b</span> c</p>')));
  assert.notEqual(EN.runsOf(parse('<p>A b <u>c</u></p>')), EN.runsOf(parse('<p>A b <u style="text-decoration-style:double">c</u></p>')));
  assert.equal(plain, EN.runsOf(parse('<p><span style="text-transform:none">A</span> b c</p>')), 'no capitals is plain text');
});
