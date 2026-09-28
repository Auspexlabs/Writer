import { test } from 'node:test';
import assert from 'node:assert/strict';
import { isSpillFormula } from '../sheet-model.js';
import { Calc, shiftF, adjF, subtotalSheet, renameTableColumnRefs, convertTableRefs } from '../sheet-engine.js';
const book = () => ({sheets:[{name:'Data',cells:{A1:{v:'Item'},B1:{v:'Amount'},A2:{v:'one'},A3:{v:'two'},B2:{v:'3'},B3:{v:'7'},A4:{v:'Total'},B4:{v:'10'}},tables:[{name:'Sales',range:'A1:B4',header:true,totals:true,columns:[{name:'Item'},{name:'Amount'}]}]}]});
test('structured references resolve data, headers, totals, this row, column spans and table names', () => {
  const c=new Calc(book());
  for(const [f,value] of [['SUM(Sales[Amount])',10],['SUM(Sales[[#All],[Amount]])',20],['Sales[[#Headers],[Amount]]','Amount'],['Sales[[#Totals],[Amount]]',10],['SUM(Sales)',10],['SUM(Sales[[Item]:[Amount]])',10]]) assert.equal(c.evaluate(f),value,f);
  assert.equal(c.evaluate('[@Amount]*2',0,1,1),6);
  assert.equal(c.evaluate('[@[Amount]]*2',0,2,1),14);
  assert.equal(c.evaluate('Sales[[#This Row],[Amount]]',0,1,0),3);
  assert.deepEqual(c.evaluate('[@Amount]',0,0,1),{err:'#VALUE!'});
  assert.deepEqual(c.evaluate('Sales[missing]'),{err:'#REF!'});
  const d=book();d.sheets.push({name:'Other',cells:{A1:{v:'=SUM(Sales[Amount])'}}});assert.equal(new Calc(d).value(1,0,0),10);
});
test('structured references spill and quoted special characters are kept', () => {
  const d=book();d.sheets[0].cells.D1={v:'=Sales[Amount]'};
  assert.equal(new Calc(d).value(0,1,3),7);
  d.sheets[0].tables[0].columns[1].name='Amount]';assert.equal(new Calc(d).evaluate("SUM(Sales[Amount']])"),10);
});
test('moving A1 references never rewrites structured column names', () => {
  assert.equal(shiftF('SUM(Sales[A1])+A1',1,1),'SUM(Sales[A1])+B2');
  assert.equal(adjF('SUM(Sales[[A1]:[B2]])+A1','Data','Data','r',0,1),'SUM(Sales[[A1]:[B2]])+A2');
});
test('subtotal excludes nested aggregates and respects filtered versus manually hidden rows', () => {
  const doc={sheets:[{name:'Data',cells:{A1:{v:1},A2:{v:2},A3:{v:3},A4:{v:'=SUBTOTAL(9,A1:A3)'}},hiddenRows:[1],frows:[2]}]};const c=new Calc(doc);
  assert.equal(c.evaluate('SUBTOTAL(9,A1:A4)'),3);assert.equal(c.evaluate('SUBTOTAL(109,A1:A4)'),1);
});

test('table conversions keep formulas usable and subtotal reports count detail only', () => {
  const d=book(), c=new Calc(d),t=d.sheets[0].tables[0];
  assert.equal(renameTableColumnRefs('SUM(Sales[Amount])+"Sales[Amount]"','Sales','Amount','Revenue'),'SUM(Sales[Revenue])+"Sales[Amount]"');
  assert.equal(convertTableRefs('SUM(Sales[Amount])',t,c,0,6,0),"SUM('Data'!$B$2:$B$3)");
  assert.equal(convertTableRefs('SUM(Sales)',t,c,0,6,0),"SUM('Data'!$A$2:$B$3)");
  assert.equal(convertTableRefs('[@Amount]',t,c,0,2,0),"'Data'!$B3");
  const report=subtotalSheet(d.sheets[0],{r1:0,r2:2,c1:0,c2:1},0,1,9,'Totals');d.sheets.push(report);
  const calc=new Calc(d);assert.equal(calc.value(1,5,1),10);assert.equal(report.outline.length,3);
});

test('scalar table calculated columns are not scanned as dynamic arrays', () => {
 for(const f of ['=[@Amount]*2','=[@[Amount]]*2','=Sales[[#This Row],[Amount]]*2'])assert.equal(isSpillFormula(f),false,f);
 assert.equal(isSpillFormula('=Sales[Amount]'),true);assert.equal(isSpillFormula('=[@Amount]*SEQUENCE(2)'),true);
});
