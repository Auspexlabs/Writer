import {test} from 'node:test';import assert from 'node:assert/strict';import {Calc,usedRange} from '../sheet-engine.js';import {sparklineSvg,sparklineMap} from '../sheet-sparklines.js';
test('sparklines use formula values, blank gaps, negative colors, hidden data and far destinations',()=>{
 const group={type:'column',negative:true,colorNegative:'#C00000',sparklines:[{source:"'Data'!A1:D1",cell:'Z1000'}]},sh={name:'Data',cells:{A1:{v:4},B1:{v:-2},C1:{v:'=A1*2'}},sparklines:[group]},doc={sheets:[sh]},calc=new Calc(doc),entry=sparklineMap(sh).get('Z1000');let svg=sparklineSvg(entry,calc,0);assert.equal((svg.match(/<rect /g)||[]).length,3);assert.match(svg,/#C00000/);assert.deepEqual(usedRange(sh),{r1:0,c1:0,r2:999,c2:25});
 sh.hiddenCols=[1];svg=sparklineSvg(entry,calc,0);assert.equal((svg.match(/<rect /g)||[]).length,2);group.displayHidden=true;assert.equal((sparklineSvg(entry,calc,0).match(/<rect /g)||[]).length,3);
 group.type='line';group.markers=true;assert.equal((sparklineSvg(entry,calc,0).match(/<circle /g)||[]).length,3);group.type='stacked';assert.match(sparklineSvg(entry,calc,0),/data-sparkline="stacked"/);
});
