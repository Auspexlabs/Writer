// WRITER_PLAYWRIGHT_MODULE=/path/to/playwright/index.mjs node docs/audits/pdf-offline-browser.mjs
import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import {resolve,extname} from 'node:path';
import assert from 'node:assert/strict';
const {chromium}=await import(process.env.WRITER_PLAYWRIGHT_MODULE||'playwright');
const root=resolve(process.env.WRITER_UI_ROOT||'ui'),requests=[];
const server=createServer(async(req,res)=>{
 try{const path=decodeURIComponent(new URL(req.url,'http://local').pathname);if(path==='/probe'){res.setHeader('Content-Type','text/html');res.end('<!doctype html><meta charset="utf-8"><canvas id="canvas"></canvas>');return;}
  const file=resolve(root,'.'+path);if(!file.startsWith(root+'/'))throw Error('outside root');const data=await readFile(file);res.setHeader('Content-Type',({'.mjs':'text/javascript','.js':'text/javascript','.html':'text/html','.bcmap':'application/octet-stream','.ttf':'font/ttf'})[extname(file)]||'application/octet-stream');res.end(data);requests.push(path);
 }catch{res.statusCode=404;res.end();}
});
await new Promise(r=>server.listen(0,'127.0.0.1',r));const origin='http://127.0.0.1:'+server.address().port;
let browser;
try{
 browser=await chromium.launch({headless:true,...(process.env.WRITER_CHROME?{executablePath:process.env.WRITER_CHROME}:{})});
 const context=await browser.newContext(),external=[],errors=[];
 await context.route('**/*',route=>{if(route.request().url().startsWith(origin+'/'))return route.continue();external.push(route.request().url());return route.abort();});
 const page=await context.newPage();page.on('console',m=>{console.error(m.text());});page.on('pageerror',e=>errors.push(e.message));await page.goto(origin+'/probe');
 const result=await page.evaluate(async()=>{
  const K=await import('/pdf-kit.js'),L=await import('/vendor/pdf-lib/dist/pdf-lib.esm.min.js');
  const js=await K.pdfjs();K.useLibs({...js,getDocument:opts=>js.getDocument({...opts,useSystemFonts:false})});
  const source=await L.PDFDocument.create(),font=await source.embedFont(L.StandardFonts.Helvetica);
  source.addPage([300,400]).drawText('PDF offline',{x:20,y:350,font,size:24});source.addPage([300,400]);
  const form=source.getForm(),field=form.createTextField('Name');field.setText('before');field.addToPage(source.getPage(0),{x:20,y:200,width:150,height:30});
  const bytes=await source.save({useObjectStreams:false}),pages=await K.importPdf({arrayBuffer:async()=>bytes},'browser-offline'),canvas=document.getElementById('canvas');
  await K.renderPage('browser-offline',pages[0],canvas,1);const pixels=canvas.getContext('2d').getImageData(0,0,canvas.width,canvas.height).data;
  let ink=0;for(let i=0;i<pixels.length;i+=4)if(pixels[i]<100&&pixels[i+1]<100&&pixels[i+2]<100)ink++;
  const widgets=await K.widgets('browser-offline',pages[0]);await K.setField('browser-offline',pages[0],widgets[0].id,'after');
  const doc={_gen:0,_n:2,pages,annots:[{id:'note',page:pages[0].id,t:'note',x:20,y:80,w:20,h:20,text:'离线批注'}, {id:'text',page:pages[0].id,t:'text',x:20,y:110,w:180,h:40,fs:16,color:'#000000',text:'中文离线文字'}],removed:[],form:1};
  const saved=await L.PDFDocument.load(await K.saveBytes('browser-offline',doc));
  const annotationContents=saved.getPage(0).node.lookup(L.PDFName.of('Annots'),L.PDFArray).asArray().map(ref=>saved.context.lookup(ref).lookup(L.PDFName.of('Contents'))?.decodeText?.()).filter(Boolean);
  const reordered=await L.PDFDocument.load(await K.saveBytes('browser-offline',{...doc,pages:[{...pages[1],rot:90},pages[0]]}));
  // A PDF with an unembedded Chinese CID font forces PDF.js to load the bundled CMap.
  const objs=['<< /Type /Catalog /Pages 2 0 R >>','<< /Type /Pages /Kids [3 0 R] /Count 1 >>','<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>','<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [5 0 R] >>','<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> /DW 1000 /FontDescriptor << /Type /FontDescriptor /FontName /STSong-Light /Flags 4 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 900 /Descent -200 /CapHeight 700 /StemV 80 >> >>'];
  const content='BT /F1 24 Tf 20 350 Td <4E2D6587> Tj ET';objs.push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`);let pdf='%PDF-1.4\n',offsets=[0];for(let i=0;i<objs.length;i++){offsets.push(pdf.length);pdf+=`${i+1} 0 obj\n${objs[i]}\nendobj\n`;}const start=pdf.length;pdf+=`xref\n0 ${objs.length+1}\n0000000000 65535 f \n`+offsets.slice(1).map(x=>String(x).padStart(10,'0')+' 00000 n \n').join('')+`trailer\n<< /Size ${objs.length+1} /Root 1 0 R >>\nstartxref\n${start}\n%%EOF\n`;
  const chinese=await K.importPdf({arrayBuffer:async()=>new TextEncoder().encode(pdf)},'chinese-offline');await K.renderPage('chinese-offline',chinese[0],canvas,1);const chineseText=await K.pageText('chinese-offline',chinese[0]);
  const {layoutTabs}=await import('/word-tabs.js'), positions=[];
  for(const kind of ['left','center','right','decimal']){
   const p=document.createElement('p');p.style.cssText='font:20px Arial;width:400px;margin:0;padding:0';p.setAttribute('data-w-tabs',kind+' 4cm');p.innerHTML='Label<span class="wd-tab">\t</span><span data-suffix>123.45</span>';document.body.appendChild(p);layoutTabs(p);
   const suffix=p.querySelector('[data-suffix]'),left=p.getBoundingClientRect().left,r=suffix.getBoundingClientRect(),range=document.createRange();range.setStart(suffix.firstChild,0);range.setEnd(suffix.firstChild,3);
   const x=kind==='left'?r.left:kind==='right'?r.right:kind==='center'?(r.left+r.right)/2:r.left+range.getBoundingClientRect().width;positions.push({kind,x:x-left});p.remove();
  }
  const slides=await import('/office-io.js'),sourceImage=document.createElement('canvas');sourceImage.width=400;sourceImage.height=200;const ig=sourceImage.getContext('2d');ig.fillStyle='#ff0000';ig.fillRect(0,0,100,200);ig.fillStyle='#00ff00';ig.fillRect(100,0,200,200);ig.fillStyle='#0000ff';ig.fillRect(300,0,100,200);
  canvas.width=1600;canvas.height=900;
  slides.paintSlide(canvas.getContext('2d'),{objs:[slides.txt({html:'<p>Normal <span style="font-size:64px;color:#ff0000;font-weight:700">RED</span></p>',x:0,y:0,w:800,h:200}),{t:'image',x:0,y:200,w:200,h:200,src:'striped',look:{crop:'25,0,25,0',grayscale:'true'}}]},slides.THEMES.paper,900,new Map([['striped',sourceImage]]));
  const grayPixel=Array.from(canvas.getContext('2d').getImageData(100,300,1,1).data),raster=canvas.getContext('2d').getImageData(0,0,800,200).data;let redPixels=0;for(let i=0;i<raster.length;i+=4)if(raster[i]>150&&raster[i+1]<80&&raster[i+2]<80)redPixels++;
  K.reset('browser-offline');K.reset('chinese-offline');
  return {ink,annotationContents,form:saved.getForm().getTextField('Name').getText(),pages:reordered.getPageCount(),rotation:reordered.getPage(0).getRotation().angle,chineseText,positions,grayPixel,redPixels};
 });
 for(const p of result.positions)assert.ok(Math.abs(p.x-4*96/2.54)<1,JSON.stringify(p));
 assert.ok(result.redPixels>200);assert.equal(result.grayPixel[0],result.grayPixel[1]);assert.equal(result.grayPixel[1],result.grayPixel[2]);assert.ok(result.grayPixel[0]>150);
 assert.ok(result.ink>100);assert.equal(result.form,'after');assert.equal(result.pages,2);assert.equal(result.rotation,90);assert.match(result.chineseText,/中文/);assert.deepEqual(external,[]);assert.deepEqual(errors,[]);
 assert.deepEqual(result.annotationContents,['离线批注','中文离线文字']);
 assert.ok(requests.some(p=>p.includes('pdf.worker')));assert.ok(requests.some(p=>p.includes('/standard_fonts/')));assert.ok(requests.some(p=>p.includes('/cmaps/')));
 console.log(JSON.stringify({uiRoot:root,generatedAt:new Date().toISOString(),result,externalRequests:external,worker:requests.filter(p=>p.includes('worker')),fonts:requests.filter(p=>p.includes('standard_fonts')),cmaps:requests.filter(p=>p.includes('cmaps'))},null,2));
}finally{await browser?.close();await new Promise(r=>server.close(r));}
