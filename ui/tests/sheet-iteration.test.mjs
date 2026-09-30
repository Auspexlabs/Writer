import test from 'node:test';
import assert from 'node:assert/strict';
import { Calc } from '../sheet-engine.js';
const documentOf = (cells, iteration = {enabled:true,count:200,delta:1e-10}) => ({iteration,sheets:[{name:'Sheet1',cells:Object.fromEntries(Object.entries(cells).map(([k,v])=>[k,{v}]))}]});
test('circular formulas are opt-in and iteration recomputes downstream formulas',()=>{
  assert.deepEqual(new Calc(documentOf({A1:'=(A1+1)/2'},null)).value(0,0,0),{err:'#CIRC!'});
  const calc=new Calc(documentOf({A1:'=(A1+1)/2',B1:'=A1*10'}));
  assert.ok(Math.abs(calc.value(0,0,1)-10)<1e-8);
  assert.ok(calc.iterationStatus.converged);assert.ok(calc.iterationStatus.iterations<200);
  assert.ok(Math.abs(calc.value(0,0,0)-1)<1e-9);
});
test('mutually dependent cells and cross-sheet cycles are independent of initial requested cell',()=>{
  const doc=documentOf({A1:'=(B1+2)/2',B1:'=A1/2',C1:'=A1+B1'});
  for(const first of [0,1,2]){const calc=new Calc(doc);calc.value(0,0,first);assert.ok(Math.abs(calc.value(0,0,0)-4/3)<1e-8);assert.ok(Math.abs(calc.value(0,0,1)-2/3)<1e-8);assert.ok(Math.abs(calc.value(0,0,2)-2)<1e-8);}
  const cross=documentOf({A1:'=(Other!A1+2)/2'});cross.sheets.push({name:'Other',cells:{A1:{v:'=Sheet1!A1/2'}}});assert.ok(Math.abs(new Calc(cross).value(1,0,0)-2/3)<1e-8);
});
test('iteration limit stops divergence, preserves errors and handles changing conditional dependencies',()=>{
  const calc=new Calc(documentOf({A1:'=A1+1'},{enabled:true,count:5,delta:0}));assert.equal(calc.value(0,0,0),5);assert.equal(calc.iterationStatus.converged,false);assert.equal(calc.iterationStatus.iterations,5);
  const errors=new Calc(documentOf({A1:'=A1+1/0'}));assert.deepEqual(errors.value(0,0,0),{err:'#DIV/0!'});
  const branch=new Calc(documentOf({A1:'=IF(A1<.5,.5,(B1+1)/2)',B1:'=A1/2'}));assert.ok(Math.abs(branch.value(0,0,0)-2/3)<1e-8);
  const nested=new Calc(documentOf({A1:'=(A1+1)/2',B1:'=(B1+2)/2',C1:'=IF(A1>.9,B1,0)'}));assert.ok(Math.abs(nested.value(0,0,2)-2)<1e-8);
  const unaffected=new Calc(documentOf({A1:'=(A1+1)/2',B1:'=3+4',C1:'=B1*2'}));assert.equal(unaffected.value(0,0,1),7);unaffected.value(0,0,0);assert.equal(unaffected.value(0,0,2),14);
});
