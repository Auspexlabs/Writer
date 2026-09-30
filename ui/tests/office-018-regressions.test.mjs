import {test} from 'node:test';
import assert from 'node:assert/strict';
import * as E from '../sheet-engine.js';
import {sheetPrint} from '../sheet-print.js';

test('COUNTA distinguishes blank cells from empty formula strings in small and sparse ranges',()=>{
 const calc=new E.Calc({sheets:[{name:'S',cells:{A1:{v:1},A2:{v:'=""'},A3:{v:2},A4:{v:'',s:{b:true}},A5:{v:'=1/0'},A1048576:{v:'=""'}}}]});
 for(const [formula,expected] of [['COUNTA(A1:A3)',3],['COUNTA(A1:A1048576)',5],['COUNTA(A4)',0],['COUNTA("")',1],['COUNT(A1:A5)',2]])assert.equal(calc.evaluate(formula,0,0,3),expected,formula);
 const spill=new E.Calc({sheets:[{name:'S',cells:{A1:{v:'=MAP({1,2},LAMBDA(x,""))'}}}]});
 assert.equal(spill.evaluate('COUNTA(A1:B1)',0,3,0),2);
 const blank=new E.Calc({sheets:[{name:'S',cells:{A1:{v:1},A3:{v:2}}}]});
 assert.equal(blank.evaluate('COUNTA(OFFSET(A1,0,0,3,1))'),2,'computed reference blanks must not become text values');
 assert.equal(blank.evaluate('COUNTA({"",2})'),2,'literal empty text is a value');
});

test('MAP keeps errors per element and IFERROR preserves the spill shape',()=>{
 const calc=new E.Calc({sheets:[{name:'S',cells:{A1:{v:'=IFERROR(MAP({1,0,2},LAMBDA(x,1/x)),99)'},A3:{v:'=MAP({1,0,2},LAMBDA(x,1/x))'}}}]});
 assert.equal(calc.value(0,0,0),1);
 assert.deepEqual(calc.arrays.get('0:0:0'),[[1,99,.5]]);
 assert.equal(calc.value(0,2,0),1);
 assert.deepEqual(calc.arrays.get('0:2:0'),[[1,{err:'#DIV/0!'},.5]]);
 assert.deepEqual(calc.evaluate('MAP({1,2},LAMBDA(a,b,a+b))'),{err:'#VALUE!'});
});

test('conditional sums use sparse full-column ranges, aligned cross-sheet criteria and spills',()=>{
 const cells={A1:{v:1},A2:{v:'=""'},A3:{v:2},A1048576:{v:4},C1:{v:'=SEQUENCE(3)'},E2:{v:9}},calc=new E.Calc({sheets:[{name:'S',cells},{name:'Flags',cells:{B1:{v:'yes'},B3:{v:'yes'},B1048576:{v:'no'}}}]});
 assert.equal(calc.evaluate('SUMIFS(A1:A1048576,A1:A1048576,">0")'),7);
 assert.equal(calc.evaluate('SUMIFS(A:A,Flags!B:B,"yes")'),3);
 assert.equal(calc.evaluate('AVERAGEIFS(A:A,Flags!B:B,"yes")'),1.5);
 assert.equal(calc.evaluate('SUMIFS(C1:C1048576,C1:C1048576,">1")'),5);
 assert.equal(calc.evaluate('SUMIFS(E2:E1048576,Flags!B1:B1048575,"yes")'),9);
 assert.deepEqual(calc.evaluate('SUMIFS(A:A,B1:B3,">0")'),{err:'#VALUE!'});
 assert.deepEqual(calc.evaluate('AVERAGEIFS(A:A,Flags!B:B,"absent")'),{err:'#DIV/0!'});
 assert.equal(calc.evaluate('SUMIFS({1;2},{3;4},">3")'),2,'array-expression fallback');
});

test('conditional format thresholds evaluate absolute, relative and text references; printing agrees',()=>{
 const sh={name:'S',cells:{A1:{v:20},A2:{v:5},B1:{v:10},B2:{v:3}},cf:[{type:'cellIs',range:'A1:A2',operator:'greaterThan',value:'B1',fill:'#123456'}]},calc=new E.Calc({sheets:[sh]});
 let cf=E.conditionalFormats(sh,calc,0,E);
 assert.equal(cf(0,0,20).bg,'#123456');assert.equal(cf(1,0,5).bg,'#123456');
 sh.cf[0].value='$B$1';cf=E.conditionalFormats(sh,calc,0,E);
 assert.equal(cf(1,0,5).bg,'');
 sh.cf[0].value='1/0';cf=E.conditionalFormats(sh,calc,0,E);assert.equal(cf(0,0,20).bg,'');
 sh.cf[0].value='B1';assert.match(sheetPrint({sheets:[sh]},E).body,/#123456/);
 sh.cf[0]={type:'cellIs',range:'A1',operator:'equal',value:'"Hello"',fill:'#123456'};
 assert.equal(E.conditionalFormats(sh,calc,0,E)(0,0,'hello').bg,'#123456');
 sh.cf[0].value='Hello';assert.equal(E.conditionalFormats(sh,calc,0,E)(0,0,'hello').bg,'#123456','legacy literal text rule');
});

test('colour scales include spill cells once even when ranges overlap or stored children exist',()=>{
 for(const cells of [{A1:{v:'=SEQUENCE(3)'}},{A1:{v:'=SEQUENCE(3)'},A2:{v:'2',spill:'A1'},A3:{v:'3',spill:'A1'}}]){
  const sh={name:'S',cells,cf:[{type:'colorScale',range:'A1:A3 A2:A3',colors:['#ff0000','#ffff00','#00ff00']}]},calc=new E.Calc({sheets:[sh]}),cf=E.conditionalFormats(sh,calc,0,E);
  assert.deepEqual([0,1,2].map(r=>cf(r,0,calc.value(0,r,0)).bg),['#ff0000','#ffff00','#00ff00']);
  const printed=sheetPrint({sheets:[sh]},E).body;for(const color of ['#ff0000','#ffff00','#00ff00'])assert.ok(printed.includes(color),color);
 }
});
