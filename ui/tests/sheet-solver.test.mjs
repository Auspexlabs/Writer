import {test} from 'node:test';import assert from 'node:assert/strict';import {linearProgram,solveWorkbook} from '../sheet-solver.js';
const close=(a,b)=>assert.ok(Math.abs(a-b)<1e-7,`${a} != ${b}`);
test('two-phase simplex handles feasibility, equality, bounds and degeneracy',()=>{
 const r=linearProgram([3,2],[{a:[1,1],op:'<=',b:4},{a:[1,0],op:'<=',b:2},{a:[0,1],op:'<=',b:3}]);assert.equal(r.status,'optimal');assert.deepEqual(r.values,[2,2]);close(r.value,10);
 const eq=linearProgram([-1,-1],[{a:[1,2],op:'>=',b:8},{a:[1,-1],op:'=',b:1}]);assert.equal(eq.status,'optimal');close(eq.values[0],10/3);close(eq.values[1],7/3);
 assert.equal(linearProgram([1],[{a:[1],op:'<=',b:-2}]).status,'infeasible');assert.equal(linearProgram([1],[]).status,'unbounded');
 assert.equal(linearProgram([1,1],[{a:[1,1],op:'=',b:4},{a:[2,2],op:'=',b:8}]).status,'optimal');
});
test('workbook solver supports integer branches and unrestricted variables without mutating source',()=>{
 const doc={sheets:[{name:'S',cells:{A1:{v:0},A2:{v:0},B1:{v:'=3*A1+2*A2'},B2:{v:'=2*A1+2*A2'}}}]},original=JSON.stringify(doc);
 const r=solveWorkbook(doc,0,{target:'B1',variables:'A1:A2',mode:'max',integer:'A1:A2',constraints:[{left:'B2',op:'<=',right:7},{left:'A1',op:'<=',right:2}]});assert.equal(r.ok,true);assert.deepEqual(r.values,{A1:2,A2:1});close(r.result,8);assert.equal(JSON.stringify(doc),original);
 const free=solveWorkbook(doc,0,{target:'B1',variables:'A1',nonnegative:false,mode:'min',constraints:[{left:'A1',op:'>=',right:-4}]});assert.equal(free.ok,true);close(free.values.A1,-4);
 const desired=solveWorkbook(doc,0,{target:'B1',variables:'A1',mode:'value',value:12});assert.equal(desired.ok,true);close(desired.values.A1,4);
 assert.throws(()=>solveWorkbook({sheets:[{name:'S',cells:{A1:{v:1},B1:{v:'=A1^2'}}}]},0,{target:'B1',variables:'A1'}),/非线性/);
});
test('nonlinear solving handles coupled variables, equality constraints and formula domain restrictions',()=>{
 const doc={sheets:[{name:'S',cells:{A1:{v:1},A2:{v:1},B1:{v:'=(A1-3)^2+(A2-2)^2'},B2:{v:'=A1+A2'}}}]};
 const r=solveWorkbook(doc,0,{target:'B1',variables:'A1:A2',method:'nonlinear',mode:'min',constraints:[{left:'B2',op:'=',right:6}]});assert.equal(r.ok,true);assert.ok(Math.abs(r.values.A1-3.5)<1e-4);assert.ok(Math.abs(r.values.A2-2.5)<1e-4);assert.equal(doc.sheets[0].cells.A1.v,1);
 const log={sheets:[{name:'S',cells:{A1:{v:1},B1:{v:'=LN(A1)'}}}]};const solved=solveWorkbook(log,0,{target:'B1',variables:'A1',method:'nonlinear',mode:'value',value:2});assert.equal(solved.ok,true);assert.ok(Math.abs(solved.values.A1-Math.exp(2))<1e-5);
});
