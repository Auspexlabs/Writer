import {test} from 'node:test';
import assert from 'node:assert/strict';
import {advancedCriterion,advancedFilter,copyAdvancedFilter,sortRowOrder} from '../sheet-data-tools.js';
const book=()=>({sheets:[{name:'Data',cells:{A1:{v:'Region'},B1:{v:'Sales'},A2:{v:'East'},B2:{v:100},A3:{v:'east end'},B3:{v:20},A4:{v:'West'},B4:{v:50},A5:{v:'East'},B5:{v:100},D1:{v:'Region'},E1:{v:'Sales'},D2:{v:'East'},E2:{v:'>30'},D3:{v:'West'},E3:{v:'<60'}}}]});
test('advanced criteria combine AND rows and OR groups, keep text prefixes and unique records',()=>{
 const doc=book(),spec={range:'A1:B5',criteria:'D1:E3'};
 assert.deepEqual(advancedFilter(doc,0,spec).rows,[1,3,4]);
 assert.deepEqual(advancedFilter(doc,0,{...spec,unique:true}).rows,[1,3]);
 doc.sheets[0].cells.E2.v='>0';assert.deepEqual(advancedFilter(doc,0,{...spec,criteria:'D1:E2'}).rows,[1,2,4]);
 doc.sheets[0].cells.D2={v:'=East',literal:true};assert.deepEqual(advancedFilter(doc,0,{...spec,criteria:'D1:E2'}).rows,[1,4]);
});
test('duplicate criteria headers express ranges and formula criteria shift from the first data row',()=>{
 const doc=book(),cells=doc.sheets[0].cells;Object.assign(cells,{D1:{v:'Sales'},E1:{v:'Sales'},D2:{v:'>20'},E2:{v:'<100'}});
 assert.deepEqual(advancedFilter(doc,0,{range:'A1:B5',criteria:'D1:E2'}).rows,[3]);
 cells.D1={v:'Condition'};cells.D2={v:'=AND(EXACT(A2,"East"),B2>$H$1)'};cells.H1={v:30};
 assert.deepEqual(advancedFilter(doc,0,{range:'A1:B5',criteria:'D1:D2'}).rows,[1,4]);
 cells.D2.v='bad criterion';assert.throws(()=>advancedFilter(doc,0,{range:'A1:B5',criteria:'D1:D2'}),/布尔公式/);
});
test('advanced criteria wildcard escaping, blank, numeric and logical types',()=>{
 for(const [criterion,value,expected]of [['a~*?', 'a*x',true],['a~*?','abc',false],['=East','East end',false],['East','EAST end',true],['<>',0,true],['=',null,true],['>20',21,true],['>20','21',false],['=TRUE',true,true],['=TRUE',1,false]])assert.equal(advancedCriterion(criterion)(value),expected,criterion+' '+value);
});
test('advanced extraction copies values and formats, leaves input untouched and rejects collisions atomically',()=>{
 const doc=book(),cells=doc.sheets[0].cells;cells.B2={v:'=100',s:{fmt:'money'}};
 const spec={range:'A1:B5',criteria:'D1:E3',unique:true,target:'G1'},before=JSON.stringify(doc),out=copyAdvancedFilter(doc,0,spec);
 assert.equal(JSON.stringify(doc),before);assert.deepEqual(out.cells.find(([a])=>a==='H2')[1],{v:100,s:{fmt:'money'}});assert.equal(out.cells.length,6);
 assert.throws(()=>copyAdvancedFilter(doc,0,{...spec,target:'B2'}),/覆盖列表/);
 cells.G1={v:'keep'};assert.throws(()=>copyAdvancedFilter(doc,0,spec),/已有单元格/);assert.equal(cells.G1.v,'keep');
 cells.G1={v:''};cells.F1={v:'=SEQUENCE(1,3)'};assert.throws(()=>copyAdvancedFilter(doc,0,spec),/已有单元格/);
});
test('sort reads conditional colors and icon ranks, custom lists, multiple levels and leaves blanks last',()=>{
 const doc=book(),sh=doc.sheets[0];sh.cells.A6={v:''};sh.cells.B6={v:0};sh.cells.A2.s={fill:'#FF0000'};sh.cells.A4.s={fill:'#FF0000'};
 assert.deepEqual(sortRowOrder(doc,0,{r1:0,r2:4,c1:0,c2:1},[{c:0,by:'fill',match:'#ff0000',asc:true},{c:1,asc:true}]),[3,1,2,4]);
 assert.deepEqual(sortRowOrder(doc,0,{r1:0,r2:5,c1:0,c2:1},[{c:0,by:'list',list:['West','East','east end'],asc:true}]),[3,1,4,2,5]);
 sh.cf=[{range:'B2:B5',type:'cellIs',operator:'greaterThan',value:'60',fill:'#00FF00'}];
 assert.deepEqual(sortRowOrder(doc,0,{r1:0,r2:4,c1:0,c2:1},[{c:1,by:'fill',match:'#00ff00',asc:true}]),[1,4,2,3]);
 sh.cf=[{range:'B2:B5',type:'iconSet',iconSet:'3Arrows'}];
 assert.deepEqual(sortRowOrder(doc,0,{r1:0,r2:4,c1:0,c2:1},[{c:1,by:'icon',match:'↑ #368251',asc:true}]),[1,4,2,3]);
 assert.deepEqual(sortRowOrder(doc,0,{r1:0,r2:5,c1:0,c2:1},[{c:0,asc:false}]).at(-1),5);
});
