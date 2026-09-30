import {test} from 'node:test';
import assert from 'node:assert/strict';
import {Calc, fmt, dateSerial} from '../sheet-engine.js';
import {sheetPrint} from '../sheet-print.js';
import * as E from '../sheet-engine.js';
test('1904 calculations keep date serials, numeric arithmetic, date functions and formatting in one system',()=>{
 for(const date1904 of [false,true]){
  const serial=date1904?43894:45356, doc={date1904,sheets:[{name:'S',cells:{A1:{v:'2024-03-05',s:{fmt:'date'}},B1:{v:'=A1+1',s:{fmt:'date'}},C1:{v:serial}}}]},calc=new Calc(doc);
  assert.equal(calc.value(0,0,0),serial);assert.equal(calc.value(0,0,1),serial+1);
  for(const f of ['DATE(2024,3,5)','DATEVALUE("2024-03-05")'])assert.equal(calc.evaluate(f),serial);
  assert.equal(calc.evaluate('YEAR(A1)'),2024);assert.equal(calc.evaluate('WEEKDAY(A1)'),3);
  assert.equal(calc.evaluate('B1-A1'),1);assert.equal(calc.evaluate('A1=C1'),true);
  assert.equal(calc.evaluate('TEXT(B1,"yyyy-mm-dd")'),'2024-03-06');
  assert.equal(fmt(serial+1,{fmt:'date',date1904}),'2024/3/6');
  assert.equal(dateSerial('2024-03-05',date1904),serial);
  assert.match(sheetPrint(doc,E).body,/2024\/3\/6/);
 }
 assert.equal(dateSerial('2024-03-05'),45356,'context does not leak between workbooks');
 assert.equal(new Calc({date1904:true,sheets:[{cells:{}}]}).evaluate('YEAR(0)'),1904);
});
test('full explicit ranges aggregate sparsely, including edge cells and spills',()=>{
 const doc={sheets:[{name:'S',cells:{A1:{v:2},A1048576:{v:7},B1:{v:'=SEQUENCE(3)'},D1:{v:'=SUM(A1:A1048576)'}}}]}, c=new Calc(doc);
 assert.equal(c.value(0,0,3),9);assert.equal(c.evaluate('SUM(B1:B1048576)'),6);
 assert.equal(c.evaluate('AVERAGE(A1:A1048576)'),4.5);assert.equal(c.evaluate('COUNT(A:A)'),2);
 assert.equal(c.evaluate('MIN(A1:A1048576)'),2);assert.equal(c.evaluate('MAX(A1:A1048576)'),7);
 assert.equal(c.evaluate('PRODUCT(A1:A1048576)'),14);
 assert.equal(c.evaluate('SUM(3,TRUE,"4",A:A)'),17);
});
test('three-dimensional references aggregate sheets in workbook order, including quoted names',()=>{
 const c=new Calc({sheets:[{name:'Jan',cells:{A1:{v:1},A2:{v:2}}},{name:'Feb',cells:{A1:{v:10},A2:{v:20}}},{name:'Mar',cells:{A1:{v:100},A2:{v:200}}},{name:'Totals',cells:{}}]});
 assert.equal(c.evaluate('SUM(Jan:Mar!A1)',3),111);assert.equal(c.evaluate("SUM('Jan:Mar'!A1:A2)",3),333);
 assert.equal(c.evaluate("AVERAGE('Jan':'Mar'!A1)",3),37);assert.equal(c.evaluate('SUM(Mar:Jan!A1)',3),111);
});
test('sheet-local names override workbook names without leaking to other sheets',()=>{
 const c=new Calc({names:{Rate:'2'},sheets:[{name:'A',names:{rate:'3'},cells:{A1:{v:'=Rate*10'}}},{name:'B',cells:{A1:{v:'=Rate*10'}}}]});
 assert.equal(c.value(0,0,0),30);assert.equal(c.value(1,0,0),20);assert.equal(c.evaluate('INDIRECT("Rate")',0),3);assert.equal(c.evaluate('INDIRECT("Rate")',1),2);
});
test('date functions use the workbook epoch once, including workday and month arithmetic',()=>{
 const a=new Calc({date1904:false,sheets:[{cells:{}}]}),b=new Calc({date1904:true,sheets:[{cells:{}}]});
 for(const f of ['DATE(2024,3,5)','EDATE(DATE(2024,3,5),1)','EOMONTH(DATE(2024,3,5),0)','WORKDAY(DATE(2024,3,5),5)','WORKDAY.INTL(DATE(2024,3,5),-5,1)'])assert.equal(a.evaluate(f)-b.evaluate(f),1462,f);
 for(const f of ['YEAR(DATE(2024,3,5))','MONTH(DATE(2024,3,5))','DAY(DATE(2024,3,5))','WEEKNUM(DATE(2024,3,5),2)','ISOWEEKNUM(DATE(2024,3,5))','NETWORKDAYS(DATE(2024,3,5),DATE(2024,4,5))','YEARFRAC(DATE(2024,3,5),DATE(2025,3,5),1)'])assert.deepEqual(a.evaluate(f),b.evaluate(f),f);
});
test('three-dimensional reference endpoints survive copy and sheet renaming',()=>{
 assert.equal(E.shiftF('=SUM(Jan:Mar!A1)',1,1),'=SUM(Jan:Mar!B2)');
 assert.equal(E.renameSheetRefs('=SUM(Jan:Mar!A1)','Jan','January'),'=SUM(\'January:Mar\'!A1)');
 assert.equal(E.renameSheetRefs('=SUM(\'Jan:Mar\'!A1)','Mar','March'),'=SUM(\'Jan:March\'!A1)');
});
