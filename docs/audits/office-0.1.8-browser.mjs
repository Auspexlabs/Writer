// Focused browser regression: actual Canvas pixels and contenteditable header serialization.
import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import {resolve,extname} from 'node:path';
import assert from 'node:assert/strict';
const {chromium}=await import(process.env.WRITER_PLAYWRIGHT_MODULE||'playwright');
const root=resolve('ui'),word=JSON.parse(await readFile('docs/audits/office-0.1.8-fixes.json','utf8')).word.headerTable;
const source=await readFile('ui/WordEditor.dc.html','utf8'),hfSource=source.slice(source.indexOf('const hfCells = (html, editing) =>'),source.indexOf('\nclass Component extends DCLogic'));
const server=createServer(async(req,res)=>{
 try { const pathname=new URL(req.url,'http://local').pathname;
  if(pathname==='/probe'){res.setHeader('Content-Type','text/html');res.end('<!doctype html><meta charset="utf-8"><body></body>');return;}
  const file=resolve(root,'.'+decodeURIComponent(pathname));if(!file.startsWith(root+'/'))throw Error('outside root');
  res.setHeader('Content-Type',extname(file)==='.js'?'text/javascript':'text/html');res.end(await readFile(file));
 }catch{res.writeHead(404).end();}
});
await new Promise(r=>server.listen(0,'127.0.0.1',r));let browser;
try{
 browser=await chromium.launch({channel:'chrome',headless:true});const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('http://127.0.0.1:'+server.address().port+'/probe');
 const result=await page.evaluate(async({word,hfSource})=>{
  const hfOut=new Function(hfSource+';return hfOut;')(),editor=document.createElement('div');editor.contentEditable='true';editor.innerHTML=word.before;document.body.append(editor);
  editor.querySelector('p').textContent='After';const header=hfOut(editor.innerHTML);
  const K=await import('/office-io.js'),canvas=document.createElement('canvas');canvas.width=1600;canvas.height=900;document.body.append(canvas);const g=canvas.getContext('2d');
  const table=K.txt({t:'table',x:0,y:0,w:500,h:150,fs:24,rows:[['Rich']],cells:{'0:0':{html:'<p><span style="font-size:64px;color:#ff0000">Rich</span></p>'}}});
  K.paintSlide(g,{objs:[table]},K.THEMES.paper,900,new Map());
  const pixels=g.getImageData(0,0,500,150).data;let redPixels=0,minY=150,maxY=0;
  for(let y=0;y<150;y++)for(let x=0;x<500;x++){const i=(y*500+x)*4;if(pixels[i]>180&&pixels[i+1]<60&&pixels[i+2]<60){redPixels++;minY=Math.min(minY,y);maxY=Math.max(maxY,y);}}
  const child=K.shape({x:100,y:100,w:200,h:100,fill:'#ff0000',html:''}),group=K.group([child]);group.rot=90;
  K.paintSlide(g,{objs:[group]},K.THEMES.paper,900,new Map());
  const rotated=Array.from(g.getImageData(200,75,1,1).data),vacated=Array.from(g.getImageData(110,110,1,1).data);
  const {richParagraphs}=await import('/slide-canvas.js'),numbering=richParagraphs('<ol><li data-start="3" data-marker="III. ">First</li><li data-marker="IV. ">Second</li></ol>').map(p=>p.bullet.trim());
  return {header,redPixels,redHeight:maxY-minY+1,rotated,vacated,numbering};
 },{word,hfSource});
 // HTML parsing inserts tbody; it must not alter the protected table marker or text.
 assert.equal(result.header.replace(/<\/?tbody>/g,''),word.serializedByEditor);assert.equal((result.header.match(/TABLE_SENTINEL/g)||[]).length,1);
 assert.ok(result.redPixels>300);assert.ok(result.redHeight>35,'64px table text should be taller than the former 24px fallback');
 assert.deepEqual(result.rotated,[255,0,0,255]);assert.notDeepEqual(result.vacated,[255,0,0,255]);assert.deepEqual(result.numbering,['III.','IV.']);assert.deepEqual(errors,[]);
 console.log(JSON.stringify({generatedAt:new Date().toISOString(),uiRoot:root,result,errors},null,2));
}finally{await browser?.close();await new Promise(r=>server.close(r));}
