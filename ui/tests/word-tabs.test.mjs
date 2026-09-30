import {test} from 'node:test';import assert from 'node:assert/strict';import {tabAdvance} from '../word-tabs.js';
test('tab stops align left, centred, right and decimal text, falling back to default stops',()=>{
 assert.equal(tabAdvance(20,[{kind:'left',at:100}],40),80);
 assert.equal(tabAdvance(20,[{kind:'center',at:100}],40),60);
 assert.equal(tabAdvance(20,[{kind:'right',at:100}],40),40);
 assert.equal(tabAdvance(20,[{kind:'decimal',at:100}],50,30),50);
 assert.equal(tabAdvance(110,[{kind:'left',at:100}],40),34);
});
