import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as E from '../sheet-engine.js';
import { sheetPrint } from '../sheet-print.js';
test('print preserves merges, width, borders and excludes filtered rows and hidden sheets', () => {
 const doc = { sheets: [{ name: 'Report', cells: { A1: { v: 'Merged', s: { bd: { bottom: 'double' }, b: true } }, A2: { v: 'hidden' }, A3: { v: 'shown' } }, colW: { A: 180 }, merges: [{ r: 0, c: 0, rs: 1, cs: 2 }], frows: [1], print: { area: 'A1:B3', header: '&LReport&C&P/&N' } }, { name: 'Hidden', visibility: 'hidden', cells: { A1: { v: 'secret' } } }] };
 const out = sheetPrint(doc, E); assert.equal(out.pages.length, 1); assert.match(out.body, /colspan="2"/); assert.match(out.body, /width:180px/); assert.match(out.body, /border-bottom:3px double/); assert.match(out.body, /shown/); assert.doesNotMatch(out.body, />hidden<|>secret</); assert.match(out.body, /1\/1/);
});
test('print repeats rows and columns and fits to one page when requested', () => {
 const cells = {}; for (let r = 0; r < 90; r++) for (let c = 0; c < 8; c++) cells[E.A(r, c)] = { v: r === 0 ? 'Title' : c === 0 ? 'Label' : `${r},${c}` };
 const sh = { name: 'Rows', cells, print: { titles: '$1:$1,$A:$A' } }, doc = { sheets: [sh] }, out = sheetPrint(doc, E);
 assert.ok(out.pages.length > 3); for (const p of out.pages) { assert.match(p.content, /Title/); assert.match(p.content, /Label/); }
 sh.print.fitWidth = sh.print.fitHeight = 1; assert.equal(sheetPrint(doc, E).pages.length, 1);
});

test('print maps native image anchors through grid sizes and includes charts', () => {
 const sh={name:'Data',cells:{A1:{v:'Title'},B2:{v:8}},print:{area:'A1:F12'},images:[{src:'data:image/png;base64,AA',x:64,y:20,w:32,h:20}],charts:[{title:'Totals',type:'column',cat:'A1',ser:[{name:'Series',values:'B2'}],x:128,y:60,w:250,h:140}]};
 const html=sheetPrint({sheets:[sh]},E).body;assert.match(html,/left:100px;top:26px;width:32px/);assert.match(html,/<svg/);assert.match(html,/Totals/);
 sh.colW={A:180};sh.images[0].x=180;sh.hiddenRows=[0];assert.match(sheetPrint({sheets:[sh]},E).body,/left:180px;top:0px;width:32px/);
});
test('whole-column dimensions and blank counts use the full worksheet without allocating it', () => {
 const calc=new E.Calc({sheets:[{name:'S',cells:{A1:{v:1},A1000001:{v:'=IF(TRUE,"",2)'},B1:{v:3}}}]});
 assert.equal(calc.evaluate('=ROWS(A:A)'),1048576);assert.equal(calc.evaluate('=COLUMNS(1:1)'),16384);
 assert.equal(calc.evaluate('=COUNTBLANK(A:A)'),1048575);assert.equal(calc.evaluate('=COUNTBLANK(A1:A1048576)'),1048575);
});

test('print and grid share conditional styling and preserve base formatting',()=>{
 const sh={name:'S',cells:{A1:{v:20,s:{bd:'thin',st:true,rotate:45,code:'[Red]0.00'}},A2:{v:30,s:{bd:'double'}}},cf:[{range:'A1:A2',type:'cellIs',operator:'greaterThan',value:'10',fill:'#123456',bold:true}]},doc={font:'Georgia',fs:18,sheets:[sh]},out=sheetPrint(doc,E).body;
 assert.match(out,/background:#123456/);assert.match(out,/font-weight:700/);assert.match(out,/font-family:Georgia;font-size:18pt/);
 assert.match(out,/text-decoration:line-through/);assert.match(out,/rotate\(-45deg\)/);assert.match(out,/color:#FF0000/);
 for(const side of ['top','right','bottom','left']){assert.match(out,new RegExp('border-'+side+':1px solid'));assert.match(out,new RegExp('border-'+side+':3px double'));}
});
test('conditional rules honor priority and stopIfTrue on screen and in print',()=>{
 const sh={name:'S',cells:{A1:{v:20}},cf:[{range:'A1',type:'cellIs',operator:'greaterThan',value:'0',priority:20,fill:'#ff0000'}, {range:'A1',type:'cellIs',operator:'greaterThan',value:'10',priority:1,bold:true,stopIfTrue:true}]},doc={sheets:[sh]};
 const style=E.conditionalFormats(sh,new E.Calc(doc),0,E)(0,0,20,{});assert.equal(style.bold,true);assert.equal(style.bg,'');assert.doesNotMatch(sheetPrint(doc,E).body,/#ff0000/);
});

test('manual row and column breaks paginate with repeated titles and hidden break rows',()=>{
 const cells=Object.fromEntries(Array.from({length:8},(_,r)=>Array.from({length:4},(_,c)=>[E.A(r,c),{v:r===0?'TITLE':`${r}:${c}`}])).flat());
 const sh={name:'S',cells,hiddenRows:[3],print:{rowBreaks:[3,6],colBreaks:[2],titles:'$1:$1'}};const result=sheetPrint({sheets:[sh]},E);
 assert.equal(result.pages.length,6);for(const p of result.pages)assert.match(p.content,/TITLE/);assert.match(result.pages[0].content,/2:0/);assert.doesNotMatch(result.pages[0].content,/4:0/);assert.match(result.pages[2].content,/4:0/);assert.match(result.pages[4].content,/6:0/);
});
