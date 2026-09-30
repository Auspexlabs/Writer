import {test} from 'node:test';
import assert from 'node:assert/strict';
import {splitColumnText,textColumns,columnValue} from '../sheet-text-columns.js';
test('text columns respects escaped qualifiers, empty columns and Unicode fixed-width boundaries',()=>{
 assert.deepEqual(splitColumnText('"A,B",,"say ""yes"""'),['A,B','','say "yes"']);
 assert.deepEqual(splitColumnText('a,,;b',{delimiters:',;',consecutive:true}),['a','b']);
 assert.deepEqual(splitColumnText('😀中AB123',{mode:'fixed',cuts:'2,4'}),['😀中','AB','123']);
 assert.throws(()=>splitColumnText('a',{mode:'fixed',cuts:'4,2'}),/递增/);
 assert.throws(()=>splitColumnText('"missing'),/闭合/);
});
test('per-column import types preserve text codes, large integers, literal formulas and valid dates',()=>{
 const rows=textColumns(['00012,31/12/2026,1.234,ignore'],{types:['text','dmy','general','skip'],decimal:'.'});
 assert.equal(rows[0].length,3);assert.deepEqual(rows[0][0],{v:'00012',literal:true,s:{fmt:'text',code:'@'}});assert.equal(rows[0][2].v,1.234);
 assert.equal(columnValue('1.234,56','general',{decimal:',',thousands:'.'}).v,1234.56);
 assert.equal(columnValue('123456789012345678').v,'123456789012345678');assert.equal(columnValue('=SUM(A1)').literal,true);
 assert.throws(()=>columnValue('31/02/2026','dmy'),/不存在/);
 assert.equal(columnValue('1904-01-01','ymd',{date1904:true}).v,0);
});
