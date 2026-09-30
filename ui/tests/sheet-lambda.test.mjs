import {test} from 'node:test';import assert from 'node:assert/strict';import {Calc} from '../sheet-engine.js';
const make=(formula,names={})=>new Calc({names,sheets:[{name:'S',cells:{A1:{v:formula}}}]});
function result(formula,names){const c=make(formula,names),v=c.value(0,0,0);return c.arrays.get('0:0:0')||v;}
test('lambda calls, lexical scope, named recursion and arity validation',()=>{
 assert.equal(result('=LAMBDA(x,x+1)(2)'),3);
 assert.equal(result('=LET(n,10,f,LAMBDA(x,x+n),f(2))'),12);
 assert.equal(result('=FACTREC(5)',{FACTREC:'LAMBDA(n,IF(n<2,1,n*FACTREC(n-1)))'}),120);
 assert.deepEqual(result('=LAMBDA(x,x+1)'),{err:'#CALC!'});
 assert.deepEqual(result('=LAMBDA(x,x+1)(1,2)'),{err:'#VALUE!'});
 assert.equal(result('=LAMBDA(x,IF(ISOMITTED(x),7,x))(,)').err,'#VALUE!');
});
test('higher-order arrays spill, preserve dimensions and fold values',()=>{
 assert.deepEqual(result('=MAP({1,2;3,4},LAMBDA(x,x*2))'),[[2,4],[6,8]]);
 assert.deepEqual(result('=MAP({1,2},{3,4},LAMBDA(a,b,a+b))'),[[4,6]]);
 assert.deepEqual(result('=BYROW({1,2;3,4},LAMBDA(row,SUM(row)))'),[[3],[7]]);
 assert.deepEqual(result('=BYCOL({1,2;3,4},LAMBDA(col,SUM(col)))'),[[4,6]]);
 assert.equal(result('=REDUCE(0,{1,2;3,4},LAMBDA(acc,x,acc+x))'),10);
 assert.deepEqual(result('=SCAN(0,{1,2;3,4},LAMBDA(acc,x,acc+x))'),[[1,3],[6,10]]);
 assert.deepEqual(result('=MAKEARRAY(2,3,LAMBDA(r,c,r*c))'),[[1,2,3],[2,4,6]]);
});
test('reshape arrays support scan order, filtering and padding',()=>{
 assert.deepEqual(result('=TOCOL({1,2;3,4},0,TRUE)'),[[1],[3],[2],[4]]);
 assert.deepEqual(result('=TOROW({1,2;3,4})'),[[1,2,3,4]]);
 assert.deepEqual(result('=WRAPROWS({1,2,3},2,0)'),[[1,2],[3,0]]);
 assert.deepEqual(result('=WRAPCOLS({1,2,3},2,0)'),[[1,3],[2,0]]);
 assert.deepEqual(result('=EXPAND({1,2},2,3,0)'),[[1,2,0],[0,0,0]]);
});
