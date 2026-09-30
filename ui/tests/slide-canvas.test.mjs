import {test} from 'node:test';import assert from 'node:assert/strict';import {paintSlide,THEMES,txt,group} from '../office-io.js';import {richParagraphs} from '../slide-canvas.js';
test('PNG preserves rich text sizes, colours, emphasis and cropped picture effects',()=>{
 const calls=[],g=new Proxy({},{get:(o,k)=>k in o?o[k]:k==='measureText'?s=>({width:s.length*12}):(...args)=>calls.push({fn:k,args,font:o.font,color:o.fillStyle,filter:o.filter}),set:(o,k,v)=>(o[k]=v,true)});
 globalThis.Path2D??=class{roundRect(){}closePath(){}};
 const image={naturalWidth:400,naturalHeight:200};
 paintSlide(g,{objs:[txt({html:'<p>Normal <span style="font-size:64px;color:#ff0000;font-weight:700">RED</span></p>',x:0,y:0,w:800,h:200}),{t:'image',x:0,y:200,w:200,h:200,src:'pic',look:{crop:'25,0,25,0',grayscale:'true'}}]},THEMES.paper,900,new Map([['pic',image]]));
 const red=calls.find(c=>c.fn==='fillText'&&c.args[0]==='RED');assert.match(red.font,/700 64px/);assert.equal(red.color,'#ff0000');
 const draw=calls.find(c=>c.fn==='drawImage');assert.deepEqual(draw.args.slice(1),[100,0,200,200,0,0,200,200]);assert.match(draw.filter,/grayscale\(1\)/);
});
test('nested formatting restores outer styles after closing spans',()=>{
 const p=richParagraphs('<p><b>Bold <i>both</i> bold</b> plain &amp; &#20013;</p>',{fs:32});assert.equal(p[0].runs[1].italic,true);assert.equal(p[0].runs[2].italic,undefined);assert.equal(p[0].runs.at(-1).bold,undefined);assert.equal(p[0].runs.at(-1).text,' plain & 中');
});
test('PNG keeps imported list markers and explicit list starting numbers',()=>{
 assert.deepEqual(richParagraphs('<ol><li data-start="3" data-marker="III. ">First</li><li data-marker="IV. ">Second</li></ol>').map(p=>p.bullet),['III. ','IV. ']);
 assert.deepEqual(richParagraphs('<ol start="3"><li>First</li><li value="7">Second</li><li>Third</li></ol>').map(p=>p.bullet),['3.','7.','8.']);
});
test('PNG paints table cell rich text and composes nested group rotation and opacity',()=>{
 const calls=[],stack=[],g=new Proxy({globalAlpha:1,save(){stack.push(this.globalAlpha);},restore(){this.globalAlpha=stack.pop();},measureText:s=>({width:s.length*10}),fillText(text){calls.push({text,font:this.font,color:this.fillStyle,alpha:this.globalAlpha});},rotate(r){calls.push({rotate:r});}},{get:(o,k)=>k in o?o[k]:(()=>{})});
 const table=txt({t:'table',x:0,y:0,w:500,h:150,fs:24,rows:[['Rich']],cells:{'0:0':{html:'<p><span style="font-size:64px;color:#ff0000">Rich</span></p>'}}});
 const inner=group([table]);inner.rot=30;inner.op=.5;const outer=group([inner]);outer.rot=90;outer.op=.5;
 paintSlide(g,{objs:[outer,txt({html:'<p>Outside</p>',x:800,y:0,w:400,h:100})]},THEMES.paper,900,new Map());
 const rich=calls.find(c=>c.text==='Rich');assert.match(rich.font,/64px/);assert.equal(rich.color,'#ff0000');assert.equal(rich.alpha,.25);
 assert.ok(calls.some(c=>c.rotate===Math.PI/2));assert.ok(calls.some(c=>c.rotate===Math.PI/6));assert.equal(calls.find(c=>c.text==='Outside').alpha,1);
});
test('ungroup composes rotations and moves nested member coordinates without changing source',async()=>{
 const {ungroup}=await import('../office-io.js'),child=txt({x:0,y:0,w:100,h:50,rot:15,op:.5}),inner=group([child]);inner.rot=30;const outer=group([inner,txt({x:200,y:100,w:100,h:50})]);outer.rot=90;outer.op=.6;
 const original=JSON.stringify(outer),kids=ungroup(outer),n=kids[0];assert.equal(n.rot,120);assert.equal(n.op,.6);assert.equal(n.kids[0].x-child.x,n.x-inner.x);assert.equal(n.kids[0].y-child.y,n.y-inner.y);
 const flattened=ungroup(n)[0];assert.equal(flattened.rot,135);assert.equal(flattened.op,.3);assert.equal(JSON.stringify(outer),original);
});
