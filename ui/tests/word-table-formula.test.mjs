import {test} from 'node:test';
import assert from 'node:assert/strict';
import {calculateWordTable,calculateWordTables,wordFormulaParts} from '../word-table-formula.js';
const n=text=>({text:String(text)}),f=formula=>({formula});
test('Word bookmarked table formulas calculate cross-table dependencies and reject cycles',()=>{
 const tables=[{names:['Sales'],rows:[[n(3),n(4)],[n(5),f('=PRODUCT(R1)')]]},{names:['Totals'],rows:[[f('=SUM(Sales R1C1:R2C2)')],[f('=Sales R2C2*2')]]}];
 assert.deepEqual(calculateWordTables(tables).map(list=>list.map(x=>x.text)),[['12'],['24','24']]);
 tables[0].rows[1][1]=f('=Totals R1C1');assert.ok(calculateWordTables(tables).flat().every(x=>x.error));
 assert.equal(calculateWordTable([[n(1),n(2)],[n(3),f('=SUM(R1,C1)')]])[0].text,'7');
 assert.equal(calculateWordTable([[f('=SUM(R0)')]])[0].error,'#REF!');
});
test('Word formula fields resolve positions, A1/RnCn references, dependencies and number formats',()=>{
 const rows=[[{...n(999),header:true},n('Qty')],[n(20),n(3),f('=A2*B2')],[n(10),n(2),f('=PRODUCT(LEFT)')],[f('=SUM(ABOVE) \\# "0.00"'),f('=AVERAGE(B2:B3)'),f('=SUM(R2C3:R3C3)')]];
 assert.deepEqual(calculateWordTable(rows).map(x=>x.text),['60','20','30.00','2.5','80']);
 rows[1][0].text='40';assert.deepEqual(calculateWordTable(rows).map(x=>x.text),['120','20','50.00','2.5','140']);
});
test('formula reference ranges exclude own cell and merged continuations, booleans render numerically',()=>{
 assert.deepEqual(calculateWordTable([[n(5),null,f('=SUM(A1:C1)')],[n(4),n(6),f('=IF(SUM(LEFT)>5,TRUE,FALSE)')]]).map(x=>x.text),['5','1']);
 assert.equal(calculateWordTable([[f('=1/0')]])[0].error,'#DIV/0!');
 assert.ok(calculateWordTable([[f('=B1') , f('=A1')]]).every(x=>x.error));
 assert.throws(()=>wordFormulaParts('=FETCH("https://example.com")'));
 assert.throws(()=>wordFormulaParts('=process.exit()'));
});

test('Word formulas resolve numeric range bookmarks, DEFINED and whole current rows/columns',()=>{
 const rows=[[n(2),n(3),f('=SUM(R)')],[n(4),f('=SUM(C)'),f('=总价 * Tax')],[f('=IF(DEFINED(Missing),1,2)'),f('=IF(DEFINED(Label),1,2)'),f('=SUM(C)')]];
 assert.deepEqual(calculateWordTable(rows,{总价:'1,250',Tax:'20%',Label:'present'}).map(x=>x.text),['5','4','250','2','1','255']);
 assert.equal(calculateWordTable([[f('=Label+1')]],{Label:'text'})[0].error,'#VALUE!');
});
