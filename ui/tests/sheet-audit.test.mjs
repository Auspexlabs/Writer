import {test} from 'node:test';import assert from 'node:assert/strict';
import {traceReferences,specialRanges} from '../sheet-audit.js';
const addresses=ranges=>ranges.map(b=>[b.si,b.r1,b.c1,b.r2,b.c2]);
test('trace finds ranges, names, cross-sheet and dynamic references without reading string literals',()=>{
 const doc={names:{Rate:"'Inputs'!B2"},sheets:[{name:'S',cells:{A1:{v:'=SUM(Inputs!A:A)+Rate+INDIRECT("Inputs!C3")+OFFSET(Inputs!D4,1,0)+LEN("Z99")'}}},{name:'Inputs',cells:{}}]};
 const refs=traceReferences(doc,0,0,0);
 assert.deepEqual(addresses(refs.ranges),[[1,0,0,1048575,0],[1,1,1,1,1],[1,2,2,2,2],[1,3,3,3,3],[1,4,3,4,3]]);assert.deepEqual(refs.unresolved,[]);
 assert.deepEqual(addresses(traceReferences(doc,1,4,0,'dependents').ranges),[[0,0,0,0,0]]);
 assert.deepEqual(traceReferences(doc,1,98,25,'dependents').ranges,[]);
});
test('go to special distinguishes formula empty text, errors, constants, notes and sparse blanks',()=>{
 const doc={sheets:[{name:'S',cells:{A1:{v:'=1/0'},B1:{v:2},A3:{v:'=""'},B3:{v:'x',s:{note:'Review'}},A1048576:{v:4}},dv:[{range:'A2:B3'},{range:'B3:C3'}]}]},box={r1:0,c1:0,r2:3,c2:2};
 assert.deepEqual(specialRanges(doc,0,box,'formulas'),[{r1:0,c1:0,r2:0,c2:0},{r1:2,c1:0,r2:2,c2:0}]);
 assert.deepEqual(specialRanges(doc,0,box,'errors'),[{r1:0,c1:0,r2:0,c2:0}]);
 assert.deepEqual(specialRanges(doc,0,box,'notes'),[{r1:2,c1:1,r2:2,c2:1}]);
 const blank=specialRanges(doc,0,{r1:0,c1:0,r2:0,c2:0},'blanks');assert.ok(blank.length<10);assert.ok(blank.some(b=>b.r2-b.r1>1000000));
 const validation=specialRanges(doc,0,box,'validation');assert.equal(validation.reduce((n,b)=>n+(b.r2-b.r1+1)*(b.c2-b.c1+1),0),5);
});

test('blank selections skip merge continuations and validation finds rules on otherwise empty sheets',()=>{
 const blank=specialRanges({sheets:[{cells:{},merges:[{r:0,c:0,rs:2,cs:2}]}]},0,{r1:0,c1:0,r2:1,c2:2},'blanks');
 assert.deepEqual(blank,[{r1:0,r2:0,c1:0,c2:0},{r1:0,r2:0,c1:2,c2:2},{r1:1,r2:1,c1:2,c2:2}]);
 assert.deepEqual(specialRanges({sheets:[{cells:{},dv:[{range:'D10:D20'}]}]},0,{r1:0,c1:0,r2:0,c2:0},'validation'),[{r1:9,c1:3,r2:19,c2:3}]);
});

test('auditing traces named tables and reports invalid formulas without throwing',()=>{
 const doc={sheets:[{name:'S',cells:{A1:{v:'=SUM(Data)'},A2:{v:'=SUM('}},tables:[{name:'Data',range:'D1:E4',columns:[{name:'X'},{name:'Y'}]}]}]};
 assert.deepEqual(addresses(traceReferences(doc,0,0,0).ranges),[[0,1,3,3,4]]);
 assert.ok(traceReferences(doc,0,1,0).unresolved.length);
});
test('protected hidden formulas do not reveal their precedents or dependent relation',()=>{
 const doc={sheets:[{name:'S',protected:true,cells:{A1:{v:8},B1:{v:'=A1*5',s:{formulaHidden:true}},C1:{v:'=A1*2'}}}]};
 assert.equal(traceReferences(doc,0,0,1).ranges.length,0);assert.deepEqual(addresses(traceReferences(doc,0,0,0,'dependents').ranges),[[0,0,2,0,2]]);
});
