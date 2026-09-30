import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Calc, goalSeek, renameSheetRefs, shiftF } from '../sheet-engine.js';
const book = (formula, value=0) => ({ sheets: [{ name:'Data', cells:{ A1:{v:value,s:{fmt:'number'}}, B1:{v:formula} } }] });
test('goal seek handles smooth, nonlinear, domain-limited and unreachable formulas without mutating inputs', () => {
  for (const [formula, initial, desired, expected] of [['=A1*12+30',0,150,10],['=A1^2',1,25,5],['=A1^2',1,0,0],['=LN(A1)',1,2,Math.exp(2)],['=10/A1',0,2,5]]) {
    const doc=book(formula,initial), before=JSON.stringify(doc), result=goalSeek(doc,0,'B1','A1',desired);
    assert.equal(result.ok,true,formula);assert.ok(Math.abs(result.value-expected)<0.0002,formula);assert.equal(JSON.stringify(doc),before);
    assert.ok(result.iterations<=150);
  }
  assert.equal(goalSeek(book('=A1^2+1'),0,'B1','A1',0).ok,false);
  assert.equal(goalSeek(book('=IF(A1>0,1,-1)'),0,'B1','A1',0).ok,false,'a sign change across a discontinuity is not a root');
  assert.equal(goalSeek(book('=A1*2','=2'),0,'B1','A1',3).error,'formula');
});
test('explicit sheet-scoped names resolve references and lambdas on their owning sheet', () => {
  const doc={names:{Rate:'2'},sheets:[{name:'Summary',cells:{A1:{v:90}}},{name:"O'Brien Data",names:{Rate:'3',Base:'A1',Twice:'LAMBDA(x,x+Base)'},cells:{A1:{v:7}}}]},calc=new Calc(doc);
  for(const [formula,expected] of [["='O''Brien Data'!rate",3],["=SUM('O''Brien Data'!Base)",7],["='O''Brien Data'!Twice(5)",12],["=INDIRECT(\"'O''Brien Data'!Base\")",7]])assert.equal(calc.evaluate(formula),expected,formula);
  assert.equal(renameSheetRefs("='O''Brien Data'!Rate+SUM('O''Brien Data'!A1)+\"O'Brien Data!Rate\"","O'Brien Data",'Renamed'),"='Renamed'!Rate+SUM('Renamed'!A1)+\"O'Brien Data!Rate\"");
  assert.equal(shiftF("='O''Brien Data'!Rate+A1",1,1),"='O''Brien Data'!Rate+B2");
});
