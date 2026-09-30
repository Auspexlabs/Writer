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

test('Word wildcard patterns handle Unicode boundaries, sets, repetition and capture replacement',async()=>{
 const {replacementText}=await import('../text-find.js');
 assert.equal(textMatches('cat cot cut cattle','<c[ao]t>',{wildcards:true}).length,2);
 assert.equal(textMatches('1 12 123 1234','<[0-9]{2,3}>',{wildcards:true}).length,2);
 const hits=textMatches('Zhao Hong; Wang Ming','([A-Z][a-z]@) ([A-Z][a-z]@)',{wildcards:true,caseSensitive:true});
 assert.equal(hits.length,2);assert.equal(replacementText(hits[0],'\\2, \\1'),'Hong, Zhao');
 assert.equal(textMatches('a\ufffcb a\nb','a*b',{wildcards:true}).length,0);
 assert.equal(textMatches('a?b','a\\?b',{wildcards:true}).length,1);
 assert.throws(()=>textMatches('x','(a*)@',{wildcards:true}));assert.throws(()=>textMatches('x','[a',{wildcards:true}));assert.throws(()=>textMatches('x','a{9,1}',{wildcards:true}));
});
test('format filters respect explicit CSS normal overrides and font sizes',async()=>{
 const {matchesTextFormat}=await import('../text-find.js');const root=parse('<p><b>bold<span style="font-weight:400;font-size:16px;font-family:Arial">normal</span></b><i>italic</i></p>');
 assert.equal(matchesTextFormat(root.querySelector('b').firstChild,{bold:true}),true);
 const n=root.querySelector('span').firstChild;assert.equal(matchesTextFormat(n,{bold:false,font:'Arial',size:12}),true);assert.equal(matchesTextFormat(n,{bold:true}),false);assert.equal(matchesTextFormat(n,{size:13}),false);
});

test('pathological wildcard searches stop at a work budget instead of freezing the editor',()=>{
 const before=performance.now();assert.throws(()=>textMatches('a'.repeat(20000),'*a*a*a*z',{wildcards:true}),/too complex/);assert.ok(performance.now()-before<2000);
 assert.equal(textMatches('𐐀𐐀 中文字','𐐀?',{wildcards:true})[0].length,4);
});
