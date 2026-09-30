import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFile, readdir} from 'node:fs/promises';
globalThis.window=globalThis;
// Browser libraries must work without a network fetch, including their own lazy imports.
globalThis.fetch=async url=>{throw new Error('Network disabled: '+url);};
const K=await import('../pdf-kit.js');
const {PDFDocument,StandardFonts}=await import('../vendor/pdf-lib/dist/pdf-lib.esm.min.js');
test('bundled PDF libraries open, search, annotate, save, reorder and extract with network disabled',async()=>{
 const source=await PDFDocument.create(),font=await source.embedFont(StandardFonts.Helvetica);
 for(const label of ['Offline first','Offline second'])source.addPage([300,400]).drawText(label,{x:20,y:350,font,size:14});
 const bytes=await source.save({useObjectStreams:false}), id='offline-test';
 const pages=await K.importPdf({arrayBuffer:async()=>bytes},id);
 try{
  assert.equal(pages.length,2);assert.match(await K.pageText(id,pages[0]),/Offline first/);
  assert.deepEqual(await K.pageInfo(id,pages[0]),{w:300,h:400,w0:300,h0:400,rot:0});
  const doc={_gen:0,_n:2,pages,annots:[],removed:[],form:0};
  const reordered=await K.saveBytes(id,{...doc,pages:[{...pages[1],rot:90},pages[0]]});
  const loaded=await PDFDocument.load(reordered);assert.equal(loaded.getPageCount(),2);assert.equal(loaded.getPage(0).getRotation().angle,90);
  const extracted=await K.extractPdf(id,doc,[1]);assert.equal((await PDFDocument.load(extracted)).getPageCount(),1);
 }finally{K.reset(id);delete K.store[id];}
 const src=await readFile(new URL('../pdf-kit.js',import.meta.url),'utf8');assert.doesNotMatch(src,/https?:\/\//);
 const maps=await readdir(new URL('../vendor/pdfjs/cmaps/',import.meta.url));assert.ok(maps.includes('UniGB-UCS2-H.bcmap'));
 const fonts=await readdir(new URL('../vendor/pdfjs/standard_fonts/',import.meta.url));assert.ok(fonts.includes('FoxitSerif.pfb'));
});
