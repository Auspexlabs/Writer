import {test} from 'node:test';
import assert from 'node:assert/strict';
import {importQueryText,runQuery,transformQuery,refreshQueries} from '../sheet-query.js';
import {editWorkbook,Calc} from '../sheet-engine.js';
const data=()=>({columns:['Region','Sales','Qty'],rows:[['East',' 12 ',2],['East','18',3],['West','7',1]]});
test('CSV and JSON query imports preserve quoted text, types and empty values',()=>{
 assert.deepEqual(importQueryText('Region,Sales\r\n"East, North",12\r\nWest,'),{columns:['Region','Sales'],rows:[['East, North','12'],['West','']]});
 assert.deepEqual(importQueryText('[{"Name":"A","Value":2},{"Name":"B"}]',{format:'json'}).rows,[['A',2],['B',null]]);
 assert.throws(()=>importQueryText('A,A\n1,2'),/不重复/);
});
test('queries transform types, add formula columns and group after filtering',()=>{
 const doc={sheets:[],queries:[{name:'Clean',source:{data:data()},steps:[{type:'type',column:'Sales',as:'number'},{type:'addColumn',name:'Total',formula:'[Sales]*[Qty]'},{type:'filter',column:'Region',value:'East'},{type:'group',columns:['Region'],values:[{column:'Total',fn:'sum',name:'Revenue'}]}]}]};
 assert.deepEqual(runQuery(doc,'Clean'),{columns:['Region','Revenue'],rows:[['East',78]]});
 assert.deepEqual(data().rows[0],['East',' 12 ',2]);
});
test('append and many-to-many joins honor keys, unmatched rows and query dependencies',()=>{
 const a={columns:['Key','Value'],rows:[['A',1],['B',2]]},b={columns:['Key','Rate'],rows:[['A',3],['A',4],['C',5]]},doc={sheets:[],queries:[{name:'A',source:{data:a},steps:[]},{name:'B',source:{data:b},steps:[]},{name:'Joined',source:{query:'A'},steps:[{type:'merge',source:{query:'B'},left:['Key'],right:['Key'],join:'full'}]}]};
 assert.deepEqual(runQuery(doc,'Joined').rows,[['A',1,'A',3],['A',1,'A',4],['B',2,null,null],[null,null,'C',5]]);
 doc.queries[2].steps=[{type:'append',source:{query:'B'}}];assert.deepEqual(runQuery(doc,'Joined').rows,[['A',1,null],['B',2,null],['A',null,3],['A',null,4],['C',null,5]]);
 doc.queries[0].source={query:'Joined'};assert.throws(()=>runQuery(doc,'Joined'),/循环/);
});
test('pivot, unpivot, split, fill and distinct keep correct column order',()=>{
 const input={columns:['ID','Name'],rows:[[1,'A:B'],[2,'C:D']]};assert.deepEqual(transformQuery(input,{type:'split',column:'Name',delimiter:':'}),{columns:['ID','Name.1','Name.2'],rows:[[1,'A','B'],[2,'C','D']]});
 const wide={columns:['ID','Q1','Q2'],rows:[[1,2,3],[2,4,5]]},long=transformQuery(wide,{type:'unpivot',columns:['Q1','Q2'],attribute:'Quarter',valueName:'Sales'});assert.deepEqual(transformQuery(long,{type:'pivot',column:'Quarter',valueColumn:'Sales'}),wide);
 const filled=transformQuery({columns:['A'],rows:[['x'],[''],['y']]},{type:'fillDown',column:'A'});assert.deepEqual(transformQuery(filled,{type:'distinct',columns:['A']}).rows,[['x'],['y']]);
});
test('query refresh loads values, preserves outside cells, clears shrunk output and can be undone',()=>{
 const doc={sheets:[{name:'Data',cells:{A1:{v:'Region'},B1:{v:'Sales'},A2:{v:'East'},B2:{v:3},A3:{v:'West'},B3:{v:5}}}],queries:[{name:'Source',source:{sheet:'Data',range:'A1:B3'},steps:[],targetSheet:'Result',target:'A1'}]};
 const next=editWorkbook(doc,(_,d)=>refreshQueries(d));assert.equal(next.sheets[1].cells.B3.v,5);assert.equal(doc.sheets.length,1);
 const shrink=editWorkbook(next,(s,d)=>{d.sheets[1].cells.H8={v:'outside'};d.queries[0].steps=[{type:'filter',column:'Sales',op:'gt',value:3}];refreshQueries(d);});assert.equal(shrink.sheets[1].cells.A2.v,'West');assert.equal(shrink.sheets[1].cells.B3?.v??'','');assert.equal(shrink.sheets[1].cells.H8.v,'outside');
 assert.equal(next.sheets[1].cells.A2.v,'East');
 assert.throws(()=>editWorkbook(doc,(_,d)=>{d.queries[0].targetSheet='Data';refreshQueries(d);}),/数据源/);
});
test('imported values beginning with equals remain text when loaded',()=>{
 const doc={sheets:[],queries:[{name:'Text',source:{data:{columns:['Code'],rows:[['=1+2']]}},steps:[],targetSheet:'Result'}]};refreshQueries(doc);assert.equal(new Calc(doc).value(0,1,0),'=1+2');
});

test('nested JSON remains lossless text and formula references do not rewrite quoted strings',()=>{
 const imported=importQueryText('[{"Data":{"id":1},"List":[1,2]}]',{format:'json'});
 assert.deepEqual(imported.rows,[['{"id":1}','[1,2]']]);
 const result=transformQuery({columns:['Code'],rows:[['x']]},{type:'addColumn',name:'Label',formula:'"[Code]"&[Code]'});
 assert.equal(result.rows[0][1],'[Code]x');
 const doc={sheets:[{name:'Data',tables:[{name:'Input',range:'A1:B3'}],cells:{A1:{v:'A'},B1:{v:'B'},A2:{v:1},B2:{v:2}}}],queries:[{name:'Q',source:{table:'Input'},steps:[],targetSheet:'Data',target:'A1',range:'A1:B3'}]};
 assert.throws(()=>refreshQueries(doc),/数据源/);
});
