import {test} from 'node:test';
import assert from 'node:assert/strict';
import {inferFlashFill} from '../sheet-flash-fill.js';
const find=(rows,examples,expected)=>assert.ok(inferFlashFill(rows,examples).some(p=>JSON.stringify(p.values)===JSON.stringify(expected)));
test('flash fill extracts names and preserves leading zeroes',()=>{
 find([['Ada Lovelace'],['Grace Hopper'],['Alan Turing']],{0:'Ada',1:'Grace'},['Ada','Grace','Alan']);
 find([['CN-00123'],['CN-00456'],['CN-00789']],{0:'00123',1:'00456'},['00123','00456','00789']);
 find([['王小明'],['李小华'],['张伟']],{0:'王',1:'李'},['王','李','张']);
});
test('flash fill combines columns, changes case and leaves empty input rows blank',()=>{
 find([['Ada','Lovelace'],['Grace','Hopper'],['Alan','Turing'],['','']],{0:'ada.lovelace@example.com',1:'grace.hopper@example.com'},['ada.lovelace@example.com','grace.hopper@example.com','alan.turing@example.com',null]);
});
test('flash fill retains ambiguous previews and rejects inconsistent examples',()=>{
 assert.ok(inferFlashFill([['Ann Smith'],['Anna Jones']],{0:'Ann'}).length>=2);
 assert.throws(()=>inferFlashFill([['one'],['two']],{}));
 assert.throws(()=>inferFlashFill([['same'],['same']],{0:'yes',1:'no'}));
 assert.throws(()=>inferFlashFill(Array.from({length:20001},()=>['a']),{0:'a'}));
});
