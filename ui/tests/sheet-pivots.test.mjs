import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pivotOutput,refreshPivot,editWorkbook,Calc } from '../sheet-engine.js';
import {pivotFormulaSource} from '../sheet-pivot-fields.js';
const spec=()=>({name:'Sales',sourceSheet:'Data',sourceRange:'A1:C4',target:'A1',rows:[0],cols:[1],values:[{field:2,fn:'sum',name:'Sales'}]});
const book=()=>({sheets:[{name:'Data',cells:Object.fromEntries(Object.entries({A1:'Region',B1:'Channel',C1:'Value',A2:'East',B2:'Web',C2:3,A3:'East',B3:'Store',C3:7,A4:'West',B4:'Web',C4:5}).map(([a,v])=>[a,{v}]))},{name:'Result',cells:{},pivots:[spec()]}]});
test('calculated pivot fields use aggregate operands for detail and grand total, support quoted names and propagate errors',()=>{
 const d={sheets:[{name:'Data',cells:Object.fromEntries([['Region','Sales','Unit Cost'],['A',100,50],['A',200,80],['B',100,90]].flatMap((row,r)=>row.map((v,c)=>[String.fromCharCode(65+c)+(r+1),{v}])))}]},p={sourceSheet:'Data',sourceRange:'A1:C4',rows:[0],values:[{name:'Margin',formula:"=(Sales-'Unit Cost')/Sales"}]};
 let out=pivotOutput(d,p);assert.ok(Math.abs(out.cells.B2.v-170/300)<1e-14);assert.equal(out.cells.B3.v,.1);assert.equal(out.cells.B4.v,.45);
 p.values.push({name:'Profit',formula:"=Sales-'Unit Cost'"});out=pivotOutput(d,p);assert.equal(out.cells.C2.v,170);assert.equal(out.cells.C4.v,180);
 d.sheets[0].cells.B2.v='=1/0';out=pivotOutput(d,p);assert.equal(out.cells.B2.v,'#DIV/0!');assert.equal(out.cells.C4.v,'#DIV/0!');
 p.values[0].formula='IFERROR(Sales/0,42)';assert.equal(pivotOutput(d,p).cells.B2.v,42);
 for(const formula of ['A2+B2','SUM(A:A)','Data!A1','RAND()',"'missing'+1",'Sales+'])assert.throws(()=>pivotOutput(d,{...p,values:[{name:'Test',formula}]}));
 assert.equal(pivotFormulaSource("'Manager''s Sales'*2",["Manager's Sales"]).source,'A1*2');
 p.values=[{name:'Label',formula:'IF(TRUE,"=1+1",0)'}];assert.equal(pivotOutput(d,p).cells.B2.v,'#VALUE!','calculated fields require numeric results');
});
test('manual pivot groups combine typed members, retain ungrouped items and filter grouped labels',()=>{
 const d=book(),p=spec();p.groups=[{field:0,items:[{name:'All regions',items:['East','West']}]}];let out=pivotOutput(d,p);assert.equal(out.range,'A1:D3');assert.equal(out.cells.A2.v,'All regions');assert.equal(out.cells.D2.v,15);
 p.filters=[{field:0,items:['All regions']}];assert.equal(pivotOutput(d,p).cells.D3.v,15);
 p.groups[0].items[0].items=['East'];out=pivotOutput(d,p);assert.equal(out.cells.D3.v,10);
 delete p.filters;out=pivotOutput(d,p);assert.equal(out.cells.A3.v,'West');assert.equal(out.cells.D4.v,15);
 p.groups[0].items.push({name:'Other',items:['East']});assert.throws(()=>pivotOutput(d,p),/only one group/);
 p.groups=[{field:0,items:[{name:'=1+1',items:['East','West']}]}];out=pivotOutput(d,p);assert.equal(new Calc({sheets:[{name:'Result',cells:out.cells}]}).value(0,1,0),'=1+1','group labels must stay text');
});
test('literal text retains leading zeros and booleans in display and reference calculations',()=>{
 const c=new Calc({sheets:[{name:'Data',cells:{A1:{v:'00012',literal:true},A2:{v:'TRUE',literal:true},B1:{v:'=A1'},B2:{v:'=SUM(A1:A2)'}}}]});
 assert.equal(c.value(0,0,0),'00012');assert.equal(c.value(0,1,0),'TRUE');assert.equal(c.value(0,0,1),'00012');assert.equal(c.value(0,1,1),0);
});
test('pivot rows, columns and grand totals aggregate real source values',()=>{
 const d=book(),out=pivotOutput(d,spec());assert.equal(out.range,'A1:D4');assert.equal(out.cells.D2.v,10);assert.equal(out.cells.B4.v,8);assert.equal(out.cells.D4.v,15);
 const s=spec();s.values[0].fn='average';assert.equal(pivotOutput(d,s).cells.D2.v,5);
 s.cols=[];assert.equal(pivotOutput(d,s).range,'A1:B4');assert.equal(pivotOutput(d,s).cells.B4.v,5);
});
test('pivot refresh updates values, preserves format and rejects new cell collisions',()=>{
 let d=editWorkbook(book(),(sh,d)=>refreshPivot(d,d.sheets[1],d.sheets[1].pivots[0]));
 d=editWorkbook(d,(sh,d)=>{sh.cells.C2.v=30;d.sheets[1].cells.D4.s={color:'#ff0000'};refreshPivot(d,d.sheets[1],d.sheets[1].pivots[0]);});assert.equal(d.sheets[1].cells.D4.v,42);assert.equal(d.sheets[1].cells.D4.s.color,'#ff0000');
 assert.throws(()=>editWorkbook(d,(sh,d)=>{sh.cells.A5={v:'North'};sh.cells.B5={v:'Web'};sh.cells.C5={v:1};d.sheets[1].cells.A5={v:'keep'};d.sheets[1].pivots[0].sourceRange='A1:C5';refreshPivot(d,d.sheets[1],d.sheets[1].pivots[0]);}),/overwrite/);
});
test('multiple row/column/value fields keep independent aggregates, headers and totals',()=>{
 const data=[['Region','City','Channel','Year','Sales','Qty'],['East','A','Web',2025,3,1],['East','A','Store',2025,7,2],['West','B','Web',2026,5,3]],d={sheets:[{name:'Data',cells:Object.fromEntries(data.flatMap((row,r)=>row.map((v,c)=>[String.fromCharCode(65+c)+(r+1),{v}])))}]};
 const p={sourceSheet:'Data',sourceRange:'A1:F4',rows:[0,1],cols:[2,3],values:[{field:4,fn:'sum'},{field:4,fn:'average'},{field:5,fn:'sum'}]},out=pivotOutput(d,p);
 assert.equal(out.range,'A1:N6');assert.equal(out.cells.L4.v,10);assert.equal(out.cells.M4.v,5);assert.equal(out.cells.L6.v,15);assert.equal(out.cells.M6.v,5);assert.equal(out.cells.N6.v,6);assert.equal(out.cells.C2.v,'2025');assert.equal(out.cells.B3.v,'City');
 p.cols=[];p.values=[{field:4,fn:'stdDev'},{field:4,fn:'varP'}];const stats=pivotOutput(d,p);assert.equal(stats.cells.C2.v,Math.sqrt(8));assert.equal(stats.cells.D2.v,4);
});

test('pivot display calculations use filtered aggregates, preserve typed items and percentage formats',()=>{
 const d=book(),p=spec();p.values[0].showAs='percentOfRow';let out=pivotOutput(d,p);assert.equal(out.cells.B2.v,.3);assert.equal(out.cells.C2.v,.7);assert.equal(out.cells.D2.v,1);assert.equal(out.cells.B2.s.code,'0.00%');
 p.values[0].showAs='percentOfTotal';out=pivotOutput(d,p);assert.equal(out.cells.B2.v,.2);assert.equal(out.cells.D4.v,1);
 p.filters=[{field:0,items:['East']}];out=pivotOutput(d,p);assert.equal(out.range,'A1:D3');assert.equal(out.cells.B2.v,.3);
 p.filters=[{field:0,items:[]}];out=pivotOutput(d,p);assert.equal(out.range,'A1:C2');assert.equal(out.cells.C2.v,'#DIV/0!');
 delete p.filters;Object.assign(p.values[0],{showAs:'runTotal',baseField:0});out=pivotOutput(d,p);assert.equal(out.cells.B3.v,8);assert.equal(out.cells.D3.v,15);
 p.values[0].showAs='difference';out=pivotOutput(d,p);assert.equal(out.cells.B2.v,'');assert.equal(out.cells.B3.v,2);assert.equal(out.cells.D3.v,-5);
 p.values[0].showAs='percentDiff';assert.equal(pivotOutput(d,p).cells.D3.v,-.5);
 p.values[0].baseItem='next';assert.equal(pivotOutput(d,p).cells.D2.v,1);
});

test('changing pivot display mode applies and clears automatic percent formats without losing colors',()=>{
 const d=book(),sh=d.sheets[1],p=sh.pivots[0];refreshPivot(d,sh,p);sh.cells.B2.s={color:'#ff0000'};p.values[0].showAs='percentOfRow';refreshPivot(d,sh,p);assert.equal(sh.cells.B2.s.code,'0.00%');assert.equal(sh.cells.B2.s.color,'#ff0000');p.values[0].showAs='normal';refreshPivot(d,sh,p);assert.equal(sh.cells.B2.s.code,undefined);assert.equal(sh.cells.B2.s.color,'#ff0000');
});
