import {test} from 'node:test';import assert from 'node:assert/strict';import {Calc} from '../sheet-engine.js';import {validateDataTable} from '../sheet-whatif.js';
test('one-variable data tables evaluate multiple formulas using row or column substitutions',()=>{
 const doc={sheets:[{name:'S',cells:{F1:{v:10},B1:{v:'=F1*2'},C1:{v:'=F1^2'},A2:{v:3},A3:{v:5}},dataTables:[{range:'A1:C3',colInput:'F1'}]}]};validateDataTable(doc,0,doc.sheets[0].dataTables[0]);const calc=new Calc(doc);assert.equal(calc.value(0,1,1),6);assert.equal(calc.value(0,2,2),25);assert.equal(calc.value(0,0,5),10);
 const row={sheets:[{name:'S',cells:{F1:{v:8},A2:{v:'=F1*2'},A3:{v:'=F1^2'},B1:{v:3},C1:{v:5}},dataTables:[{range:'A1:C3',rowInput:'F1'}]}]};const c=new Calc(row);assert.equal(c.value(0,1,1),6);assert.equal(c.value(0,2,2),25);
});
test('two-variable data tables update dependencies, propagate errors and reject circular results',()=>{
 const doc={sheets:[{name:'S',cells:{F1:{v:7},F2:{v:9},F3:{v:'=F1*F2'},A1:{v:'=F3'},B1:{v:3},C1:{v:5},A2:{v:2},A3:{v:4},H1:{v:'=SUM(B2:C3)'}},dataTables:[{range:'A1:C3',rowInput:'F1',colInput:'F2'}]}]};const calc=new Calc(doc);assert.equal(calc.value(0,2,2),20);assert.equal(calc.value(0,0,7),48);assert.equal(doc.sheets[0].cells.F1.v,7);
 doc.sheets[0].cells.A1.v='=B2';assert.equal(new Calc(doc).value(0,1,1).err,'#CIRC!');assert.throws(()=>validateDataTable(doc,0,{range:'A1:C3',rowInput:'A2'}),/不能位于/);
});
