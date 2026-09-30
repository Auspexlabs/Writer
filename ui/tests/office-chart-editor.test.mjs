import {test} from 'node:test';
import assert from 'node:assert/strict';
import {chartTable,parseChartTable} from '../office-chart-editor.js';
test('chart data paste handles multiple series, quoted categories, blanks and combination kinds',()=>{
 const d=parseChartTable('Category,Actual,Target\n"East, North",12,15\nSouth,,20','combo',['area','line']);
 assert.deepEqual(d.labels,['East, North','South']);assert.deepEqual(d.series.map(s=>s.values),[[12,null],[15,20]]);assert.deepEqual(d.series.map(s=>s.kind),['area','line']);
 assert.deepEqual(parseChartTable(chartTable({cats:d.labels,series:d.series})).series.map(s=>s.values),[[12,null],[15,20]]);
 assert.throws(()=>parseChartTable('Name,A\nX,=1+2'),/数字/);
 assert.throws(()=>parseChartTable('Name,A\nX,1,2'),/列数/);
});
test('scatter table preserves individual X coordinates and series names',()=>{
 const chart={kind:'scatter',cats:['A','B'],series:[{name:'First',x:[1,2],values:[3,4]},{name:'Second',x:[10,20],values:[30,null]}]};
 const out=parseChartTable(chartTable(chart),'scatter');assert.deepEqual(out.series,chart.series);assert.deepEqual(out.labels,chart.cats);
});
