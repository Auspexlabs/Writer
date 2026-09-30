import {test} from 'node:test';
import assert from 'node:assert/strict';
import {evaluateFormula} from '../sheet-evaluate.js';
const book=cells=>({sheets:[{name:'S',cells}]});
test('formula evaluation follows dependencies, errors and actual short-circuit branches',()=>{
 const doc=book({A1:{v:3},B1:{v:'=A1*2'},C1:{v:'=IF(B1>5,B1+4,1/0)'},D1:{v:'=IFERROR(1/0,9)'}}),before=JSON.stringify(doc);
 const result=evaluateFormula(doc,0,'C1');assert.equal(result.result,'10');assert.ok(result.steps.some(s=>s.cell==='S!B1'&&s.value==='6'));assert.ok(!result.steps.some(s=>s.expression==='(1/0)'));assert.equal(result.steps.filter(s=>s.cell==='S!B1'&&s.expression==='(A1*2)').length,1);
 const error=evaluateFormula(doc,0,'D1');assert.ok(error.steps.some(s=>s.value==='#DIV/0!'));assert.equal(error.result,'9');assert.equal(JSON.stringify(doc),before);
});
test('volatile functions run once, scopes and arrays retain computed values',()=>{
 const random=Math.random;let calls=0;Math.random=()=>{calls++;return .25;};try{const r=evaluateFormula(book({A1:{v:'=RAND()+2'}}),0,'A1');assert.equal(calls,1);assert.equal(r.result,'2.25');assert.equal(r.steps.find(s=>s.expression==='RAND()').value,'0.25');}finally{Math.random=random;}
 assert.equal(evaluateFormula(book({A1:{v:'=LET(n,2,n+3)'}}),0,'A1').result,'5');
 const array=evaluateFormula(book({A1:{v:'=MAP({1,2},LAMBDA(x,x*3))'}}),0,'A1');assert.equal(array.result,'3');assert.ok(array.steps.some(s=>s.value==='1 × 2\n3 | 6'));
});
test('hidden protected formula internals stay hidden and step count is bounded',()=>{
 const doc=book({A1:{v:'=50+7',s:{formulaHidden:true}},B1:{v:'=A1*2'}});doc.sheets[0].protected=true;
 assert.throws(()=>evaluateFormula(doc,0,'A1'),/隐藏/);const r=evaluateFormula(doc,0,'B1');assert.equal(r.result,'114');assert.ok(r.steps.every(s=>s.cell==='S!B1'));assert.ok(!r.steps.some(s=>s.expression.includes('50')));
 const limited=evaluateFormula(book({A1:{v:'=1+2+3+4'}}),0,'A1',{maxSteps:2});assert.equal(limited.steps.length,2);assert.equal(limited.truncated,true);assert.equal(limited.result,'10');
 assert.throws(()=>evaluateFormula(book({A1:{v:'=1',literal:true}}),0,'A1'),/公式/);
});
