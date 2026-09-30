import test from 'node:test';
import assert from 'node:assert/strict';
import { inkGeometry,freeformPaths,freeformSvg,splitFreeformSegment } from '../slide-freeform.js';
import {objView,shape,THEMES,presenterHtml} from '../office-io.js';
test('Bezier paths render exact SVG commands and splitting preserves the curve',()=>{
 const path={points:[[0,0],[100,0],[200,0]],controls:{1:[[0,100],[100,100]],2:[[150,-100]]},fill:'none'};
 const data={w:200,h:100,paths:[path]},html=freeformSvg(data,'fill="none" stroke="#123456"');
 assert.match(html,/M0,0 C0,1000 500,1000 500,0 Q750,-1000 1000,0/);
 assert.equal(splitFreeformSegment(path,0),1);assert.deepEqual(path.points[1],[50,75]);assert.deepEqual(path.controls[1],[[0,50],[25,75]]);assert.deepEqual(path.controls[2],[[75,75],[100,50]]);assert.deepEqual(path.controls[3],[[150,-100]]);
 splitFreeformSegment(path,2);assert.deepEqual(path.points[3],[150,-50]);
 assert.deepEqual(freeformPaths({w:1,h:1,paths:[{points:[[0,0],[1,1]],controls:{0:[[1,1]]}}]}),[]);
});
test('ink remains independent vector paths with scale-invariant points, including a single tap',()=>{
 const geometry=inkGeometry([[[10,20],[40,80],[100,20]],[[50,50]]]);assert.equal(geometry.pathData.paths.length,2);assert.equal(geometry.pathData.paths[1].points.length,2);assert.equal(geometry.x,6);assert.equal(geometry.y,16);
 const path=freeformPaths(geometry.pathData)[0];assert.equal(path.closed,false);const view=objView(shape({...geometry,shape:'custom',stroke:'#FF3B30',sw:5}),THEMES.paper);assert.match(view.svg.__html,/polyline/);assert.match(view.svg.__html,/fill="none"/);assert.equal(view.bg,'transparent');assert.equal(view.border,'none');
});
test('closed and open paths keep native fill/stroke semantics and reject nonnumeric imported points',()=>{
 const data={w:10,h:20,paths:[{points:[[0,0],[10,20]],closed:true,fill:'norm',stroke:false}]};const html=freeformSvg(data,'fill="#123456" stroke="#ABCDEF"');assert.match(html,/<polygon points="0,0 1000,1000"/);assert.match(html,/stroke="none"/);assert.deepEqual(freeformPaths({w:1,h:1,paths:[{points:[['bad',2]]}]}),[]);
});
test('presenter preview and position follow custom routes, including a repeated final slide',()=>{
 const slides=['A','B','C'].map((label,i)=>({id:String(i),objs:[{...shape({html:label}),x:0,y:0,w:100,h:100}]})),doc={slides,theme:'paper',ratio:'16:9'},show={i:2,route:[2,0,2],routePos:0};
 const html=presenterHtml(doc,2,'00:00',show);assert.match(html,/幻灯片 1 \/ 3/);assert.ok(html.includes('>A<'));assert.ok(!html.includes('>B<'));
 assert.match(presenterHtml(doc,2,'00:00',{...show,routePos:2}),/放映结束/);
 assert.doesNotMatch(presenterHtml(doc,2,'00:00',{...show,routePos:2},true),/放映结束/);
});
