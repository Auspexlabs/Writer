import {test} from 'node:test';
import assert from 'node:assert/strict';
import {traceGeometry,traceSvg} from '../sheet-trace.js';
const b=(r,c,si=0)=>({si,r1:r,r2:r,c1:c,c2:c});
const geometry={xs:Array.from({length:31},(_,i)=>i*64),ys:Array.from({length:101},(_,i)=>i*20)};
test('arrows follow actual dependency direction, frozen panes, scroll and merged centers',()=>{
 const trace={source:b(0,0),ranges:[b(4,4)],direction:'precedents'};
 const options={width:500,height:300,scrollX:128,scrollY:40,freezeRows:1,freezeCols:1,merges:[{r:0,c:0,rs:1,cs:2}]};
 const [edge]=traceGeometry(trace,0,geometry,options);assert.equal(edge.from.x,204);assert.equal(edge.from.y,76);assert.equal(edge.to.x,76);assert.equal(edge.to.y,36);assert.equal(edge.dashed,false);
 const [dependent]=traceGeometry({...trace,direction:'dependents'},0,geometry,options);assert.deepEqual(dependent.from,edge.to);assert.deepEqual(dependent.to,edge.from);
});
test('offscreen and cross-sheet references have bounded dashed arrows without inflating the viewport',()=>{
 const trace={source:b(1,1),ranges:[b(1048575,16383),b(3,3,1)],direction:'precedents'};
 const edges=traceGeometry(trace,0,geometry,{width:500,height:300});assert.equal(edges.length,2);assert.ok(edges.every(e=>e.dashed));assert.equal(edges[1].portal.si,1);
 for(const e of edges)for(const p of [e.from,e.to]){assert.ok(p.x>=44&&p.x<=500);assert.ok(p.y>=26&&p.y<=300);}
 const svg=traceSvg(edges,500,300,'Dependencies',()=>'<img onerror="boom">');assert.ok(!svg.includes('<img'));assert.match(svg,/&lt;img/);assert.match(svg,/width="500" height="300"/);assert.ok(!svg.includes('NaN'));
 const reverse=traceGeometry(trace,1,geometry,{width:500,height:300});assert.equal(reverse.length,1);assert.equal(reverse[0].portal.si,0);
});
test('whole-column and duplicate references, self-loops and drawing cap stay bounded',()=>{
 const range={si:0,r1:0,r2:1048575,c1:0,c2:0},trace={source:b(1,2),ranges:[range,range,...Array.from({length:90},(_,r)=>b(r,3))],direction:'precedents'};
 const edges=traceGeometry(trace,0,geometry,{width:500,height:300,limit:8});assert.equal(edges.length,8);assert.equal(edges[0].from.y,163);
 const loop=traceGeometry({source:b(0,0),ranges:[b(0,0)],direction:'precedents'},0,geometry)[0];assert.equal(loop.loop,true);
});
