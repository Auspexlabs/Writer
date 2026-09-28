import { test } from 'node:test';
import assert from 'node:assert/strict';
import { pivotOutput,refreshPivot,editWorkbook } from '../sheet-engine.js';
const spec=()=>({name:'Sales',sourceSheet:'Data',sourceRange:'A1:C4',target:'A1',rows:[0],cols:[1],values:[{field:2,fn:'sum',name:'Sales'}]});
const book=()=>({sheets:[{name:'Data',cells:Object.fromEntries(Object.entries({A1:'Region',B1:'Channel',C1:'Value',A2:'East',B2:'Web',C2:3,A3:'East',B3:'Store',C3:7,A4:'West',B4:'Web',C4:5}).map(([a,v])=>[a,{v}]))},{name:'Result',cells:{},pivots:[spec()]}]});
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
