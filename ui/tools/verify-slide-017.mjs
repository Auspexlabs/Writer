// Real DOM + engine round trips. Run after dotnet build src/Writer.Cli.
// Set WRITER_NODE_MODULES when Playwright is supplied outside this repository.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, copyFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../..', import.meta.url));
const { chromium } = await import(process.env.WRITER_NODE_MODULES ? pathToFileURL(join(process.env.WRITER_NODE_MODULES, 'playwright/index.mjs')).href : 'playwright');
const dir = mkdtempSync(join(tmpdir(), 'writer-slide-017-'));
const engine = spawn('dotnet', [join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll'), 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
for (const name of ['charts-pie.pptx','decor.pptx']) copyFileSync(join(root,'tests/Writer.Tests/Fixtures/pptx',name),join(dir,name));
let browser, server;
try {
  const base = await new Promise((res, rej) => { let text = ''; engine.stderr.on('data', d => { text += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(text); if (m) res(m[1]); }); engine.once('exit', code => rej(new Error(`engine exit ${code}: ${text}`))); });
  server = createServer(async (req, res) => {
    try {
      if (req.url === '/favicon.ico') { res.writeHead(204).end(); return; }
      if (req.url === '/ui/test.dc.html') { res.setHeader('Content-Type', 'text/html'); res.end(`<!doctype html><meta charset="utf-8"><style>html,body,#dc-root{height:100%;margin:0;overflow:hidden}</style><script>window.__testBook = {id:'test',type:'pptx',ratio:'16:9',theme:'paper',slides:[{id:'s',objs:[],decor:[],bg:null,trans:'none'}]};window.__WRITER_LANG='zh';</script><script src="/ui/memo.js"></script><script src="/ui/i18n.js"></script><script src="/ui/vendor/offline.js"></script><script src="/ui/support.js"></script><body><x-dc><dc-import name="SlideEditor" doc="{{ doc }}" on-change="{{ onChange }}" toast="{{ toast }}" style="position:absolute;inset:0"></dc-import></x-dc><script type="text/x-dc" data-dc-script>class Component extends DCLogic { state = {doc:window.__testBook}; constructor(p){super(p);window.__pptParent=this;} renderVals(){return {doc:this.state.doc,onChange:doc=>this.setState({doc}),toast:message=>{window.__lastToast=message;}};} }</script></body>`); return; }
      if (req.url.startsWith('/ui/')) {
        const path = resolve(root, '.' + new URL(req.url, 'http://local').pathname);
        if (!path.startsWith(join(root, 'ui') + '/')) { res.writeHead(403).end(); return; }
        res.setHeader('Content-Type', path.endsWith('.js') ? 'text/javascript' : 'text/html'); res.end(readFileSync(path)); return;
      }
      const chunks = []; for await (const chunk of req) chunks.push(chunk);
      const response = await fetch(base + req.url, { method: req.method, headers: { 'Content-Type': req.headers['content-type'] || 'application/json' }, ...(chunks.length ? { body: Buffer.concat(chunks) } : {}) });
      res.writeHead(response.status, { 'Content-Type': response.headers.get('Content-Type') || 'application/json' }); res.end(Buffer.from(await response.arrayBuffer()));
    } catch (error) { res.writeHead(500).end(String(error)); }
  });
  await new Promise(res => server.listen(0, '127.0.0.1', res));
  browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage();
  page.on('requestfailed', r => console.error('REQUEST', r.url(), r.failure()?.errorText));
  page.on('console', m => { if (m.type() === 'error') console.error('BROWSER', m.text()); });
  page.on('pageerror', e => console.error('PAGE', e.message));
  await page.goto(`http://127.0.0.1:${server.address().port}/ui/test.dc.html`);
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await page.locator('[data-edroot]').waitFor();
  const results = await page.evaluate(async () => {
    const EN = await import('/ui/engine.js'), K = await import('/ui/office-io.js');
    const check = (yes, msg) => { if (!yes) throw new Error(msg); };
    const tick = () => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    const mount = async doc => { window.__pptParent.setState({doc}); await tick(); };
    const editor = () => { const el = document.querySelector('[data-edroot]'); let f = el[Object.keys(el).find(k=>k.startsWith('__reactFiber'))]; for (;f;f=f.return) if (f.stateNode?.logic?.patchSel) return f.stateNode.logic; throw new Error('editor not mounted'); };
    const docNow = () => window.__pptParent.state.doc;
    await EN.run(['create','deck.pptx']);
    for (let i=0;i<2;i++) await EN.run(['add','deck.pptx','/','--type','slide','--prop','layout=Blank']);
    await EN.run(['add','deck.pptx','/slide[1]','--type','shape','--prop','x=2cm','--prop','y=2cm','--prop','w=12cm','--prop','h=3cm','--prop','html=<p><b><span style="font-size:24pt;color:#ff0000">First</span></b></p><p><a href="https://example.com">Second</a></p>']);
    await EN.run(['set','deck.pptx','/slide[1]/shape[1]/paragraph[1]','--prop','align=right']);
    let doc = await EN.open({id:'deck',path:'deck.pptx',type:'pptx'}); await mount(doc);
    let c = editor(), o = doc.slides[0].objs[0]; c.select(o.id,false); await tick();
    let el = document.querySelector('[data-edit="'+o.id+'"]');
    check(Math.abs(parseFloat(getComputedStyle(el.querySelector('span')).fontSize) - 24*doc._orig.geo.ptPx)<.01,'24pt is not shown at actual slide scale');
    check(getComputedStyle(el.querySelector('p')).textAlign === 'right','paragraph alignment import');
    c.patchSel({fs:80,color:'#0000FF',bold:false,italic:true,underline:true,align:'center',va:'bottom'}); await tick();
    el=document.querySelector('[data-edit="'+o.id+'"]');
    check(getComputedStyle(el.querySelector('span')).fontSize === '80px','font size did not change immediately');
    check(getComputedStyle(el.querySelector('span')).color === 'rgb(0, 0, 255)','color did not change immediately');
    check(getComputedStyle(el).fontWeight === '400' && getComputedStyle(el).fontStyle === 'italic','box bold/italic override');
    check(getComputedStyle(el.querySelector('p')).textAlign === 'center','whole box alignment override');
    const beforeURL = location.href;
    check(!el.querySelector('a').dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true})) && location.href===beforeURL,'link navigated editor');
    doc=docNow(); await EN.save(doc); doc=await EN.open(doc); await mount(doc);
    o=doc.slides[0].objs[0];
    check(o.italic && o.underline && !o.bold && o.align==='center' && o.va==='bottom','box styles lost on save');
    check(Math.abs(o.fs-80)<1,'font size shrank on reopen');
    o.html=o.html.replace('First','Edited'); await EN.save(doc); doc=await EN.open(doc);
    check(Math.abs(doc.slides[0].objs[0].fs-80)<1,'editing rich text shrank the font');
    check(await EN.save(doc)===0,'unchanged deck rewrites itself');
    await mount(doc); c=editor(); c.select(doc.slides[0].objs[0].id,false); await tick(); c.copySels(false); c.pasteSels(); await tick();
    check(c.sels.length===1 && c.sels[0].id!==doc.slides[0].objs[0].id,'paste left original selected');
    c.deleteSels(); await tick(); check(docNow().slides[0].objs.length===1,'deleting paste removed original or retained copy');
    const table=K.txt({t:'table',rows:[['Nested table']],html:'',x:500,y:300,w:500,h:100,fs:25});
    const nested=K.group([K.group([table,K.txt({html:'<p>Nested text</p>',x:500,y:450,w:500,h:50})]),K.shape({html:'',x:400,y:260,w:30,h:30})]);
    c.insertObj(nested); await tick(); doc=docNow(); await EN.save(doc); doc=await EN.open(doc); await mount(doc);
    check(document.querySelector('[data-soid]') && document.querySelector('[data-edroot]').textContent.includes('Nested table'),'nested table is not rendered');
    check(doc.slides[0].objs.some(x=>x.t==='group' && x.kids.some(y=>y.t==='group' && y.kids.some(z=>z.t==='table'))),'nested group structure lost');
    c=editor(); const originalGroup=doc.slides[0].objs.find(x=>x.t==='group'); c.select(originalGroup.id,false); await tick(); c.dupSels(); await tick(); doc=docNow(); await EN.save(doc);
    const copiedGroup=doc.slides[0].objs.at(-1); const leaf=K.flatObjs([copiedGroup]).find(x=>x.t==='table');
    check((await EN.run(['get',doc.path,leaf.path])).kind==='table','copied nested group has stale member ids');
    copiedGroup.kids[0].kids.find(x=>x.t==='text').html='<p>Changed copied group</p>'; await EN.save(doc); doc=await EN.open(doc);
    check(K.flatObjs(doc.slides[0].objs).some(x=>x.html?.includes('Changed copied group')),'copied group cannot be edited');
    doc.slides[0].objs.reverse(); await EN.save(doc); let ordered=await EN.open(doc); check(ordered.slides[0].objs[0].path===doc.slides[0].objs[0].path,'z-order not persisted'); doc=ordered;
    K.resizeSlides(doc,'4:3'); await EN.save(doc); doc=await EN.open(doc);
    check(doc.ratio==='4:3' && Math.abs(doc._orig.geo.hcm/doc._orig.geo.wcm-.75)<.001,'slide dimensions not saved');
    await mount(doc); c=editor(); c.insertPageNums(); await tick(); doc=docNow(); await EN.save(doc); doc=await EN.open(doc); await mount(doc);
    c=editor(); c.insertPageNums(); await tick(); doc=docNow(); await EN.save(doc); doc=await EN.open(doc);
    check(doc.slides.every(s=>s.objs.filter(o=>o.field==='slideNumber').length===1),'page numbers duplicate after reopen');
    doc.slides.reverse(); await EN.save(doc); doc=await EN.open(doc);
    for (let i=0;i<doc.slides.length;i++) { const field=doc.slides[i].objs.find(o=>o.field==='slideNumber'); const raw=await EN.run(['get',doc.path,field.path,'--raw']); check(raw.includes('type="slidenum"') && field.html.includes(String(i+1)),'page number is not dynamic after reorder'); }
    await mount(doc); c=editor(); c.setState({show:{i:0,step:0,live:false}}); await tick();
    const e={key:'p',ctrlKey:true,preventDefault(){this.prevented=true},stopPropagation(){this.stopped=true}}; c.showKey(e); await tick();
    check(e.prevented && e.stopped && c.state.show.tool==='pen','presentation Ctrl+P leaked'); c.setState({show:null}); await tick();
    c.insertObj(K.txt({html:'<p><b>Hel</b><i>lo</i> HELLO shelloworld</p>',x:100,y:100,w:900,h:100})); await tick();
    doc=docNow(); let hits=K.findSlides(doc,'hello',{wholeWord:true}); check(hits.length===2,'search did not cross run formatting or respect word boundaries');
    K.replaceSlideHits(doc,hits,'World'); check(doc.slides[0].objs.at(-1).html.includes('<b>World</b>'),'replacement lost first-run formatting');
    await EN.save(doc); doc=await EN.open(doc); check(K.findSlides(doc,'World',{wholeWord:true}).length===2,'replacement not saved');
    check(K.findSlides(doc,'World',{wholeWord:true,caseSensitive:true}).length===2,'case-sensitive search');
    await EN.run(['create','target.pptx']); await EN.run(['add','target.pptx','/','--type','slide','--prop','layout=Blank']);
    let target=await EN.open({id:'target',type:'pptx',path:'target.pptx'}); const sourceGroup=doc.slides.flatMap(s=>s.objs).find(o=>o.t==='group');
    const importedGroup=structuredClone(sourceGroup); const stripPaths=o=>{delete o.path; delete o.from; o.id=K.oid(); (o.kids||[]).forEach(stripPaths);}; stripPaths(importedGroup);
    importedGroup.importSource={file:doc.path,path:sourceGroup.path}; target.slides[0].objs.push(importedGroup); await EN.save(target);
    const importedText=K.flatObjs(target.slides[0].objs).find(o=>o.t==='text'); importedText.html='<p>Imported nested text</p>'; await EN.save(target); target=await EN.open(target);
    check(K.flatObjs(target.slides[0].objs).some(o=>o.html?.includes('Imported nested text')),'cross-deck nested group has stale member ids');
    const changedShape=K.flatObjs(target.slides[0].objs).find(o=>o.t==='shape'); changedShape.shape='ellipse'; await EN.save(target); target=await EN.open(target);
    check(K.flatObjs(target.slides[0].objs).some(o=>o.shape==='ellipse'),'existing shape type not saved');
    const chartDoc=await EN.open({id:'charts',path:'charts-pie.pptx',type:'pptx'}), chart=chartDoc.slides.flatMap(s=>s.objs).find(o=>o.t==='object');
    check(chart && chart.data.chart,'chart cache not exposed'); check(await EN.save(chartDoc)===0,'reading chart rewrites deck');
    doc.slides[0].objs.push({...structuredClone(chart),id:K.oid(),path:undefined,importSource:{file:chartDoc.path,path:chart.path}}); await EN.save(doc); doc=await EN.open(doc);
    check(doc.slides[0].objs.some(o=>o.t==='object' && o.data.chart),'chart relationships lost during cross-deck copy');
    const background=await EN.open({id:'background',path:'decor.pptx',type:'pptx'}); background.slides[0].bg=null; await EN.save(background); const bg=await EN.open(background);
    check(bg.slides[0].bg===null && bg.slides[0].inheritedBg && bg.slides[0].inheritedBg!=='#FFFFFF','reset background overwrote master with white');
    await mount(doc); c=editor();
    const fillShape = K.shape({x:100,y:100,w:300,h:200,fill:'#FF0000',fillOpacity:.4,html:'<p>Opaque text</p>'}); c.insertObj(fillShape); await tick();
    c.commit((d,s) => { s.bgGradient='4472C4,FFFFFF,45'; s.advanceAfter=3000; s.advanceOnClick=false; }); await tick();
    check(getComputedStyle(document.querySelector(`[data-soid="${fillShape.id}"] > div`)).backgroundColor === 'rgba(255, 0, 0, 0.4)', 'shape fill transparency not drawn independently');
    doc=docNow(); await EN.save(doc); doc=await EN.open(doc); await mount(doc); c=editor();
    check(doc.slides[0].objs.find(o=>o.html?.includes('Opaque text')).fillOpacity===.4, 'fill opacity not persisted');
    check(doc.slides[0].bgGradient==='4472C4,FFFFFF,45' && doc.slides[0].advanceAfter===3000 && doc.slides[0].advanceOnClick===false, 'gradient or timing not persisted');
    check(await EN.save(doc)===0, 'background/timing no-op save rewrites deck');
    const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==', blob=new Blob([Uint8Array.from(atob(png),c=>c.charCodeAt(0))],{type:'image/png'}), file=new File([blob],'sample.png',{type:'image/png'});
    const count=c.slide.objs.length; const pasted=new DataTransfer(); pasted.items.add(file);
    c.stageRef.current.dispatchEvent(new ClipboardEvent('paste',{clipboardData:pasted,bubbles:true,cancelable:true})); await tick(); await tick();
    check(c.slide.objs.length===count+1 && c.obj.t==='image', 'image paste did not add and select picture: '+JSON.stringify({count:c.slide.objs.length,expected:count+1,selected:c.obj?.t,toast:window.__lastToast}));
    await c.insertImageFile(file,null,true); await tick(); doc=docNow(); await EN.save(doc); doc=await EN.open(doc); await mount(doc); c=editor();
    check(doc.slides[0].bgImage && !doc.slides[0].bgGradient && K.slideDecor(doc.slides[0])[0].src===doc.slides[0].bgImage, 'picture background did not reopen');
    check((await fetch(doc.slides[0].bgImage)).headers.get('content-type').includes('image/png'), 'background image binary inaccessible');
    const rich=K.txt({html:'<p>first <b>second</b></p><p>third</p>',x:50,y:100,w:500,h:200});c.insertObj(rich);await tick();c.inlineFormat({backgroundColor:'#FFFF00',verticalAlign:'super'});await tick();
    c.insertFooter();await tick();doc=docNow();await EN.save(doc);doc=await EN.open(doc);await mount(doc);c=editor();
    check(doc.slides[0].objs.some(o=>o.field==='footer'),'footer field missing');
    check(doc.slides[0].objs.find(o=>o.html?.includes('second')).html.includes('<sup>'), 'partial text formatting not persisted');
    c.setState({editing:null});await tick(); const newTable=K.txt({t:'table',rows:[['Bold cell','B'],['C','D']],html:'',x:100,y:200,w:600,h:180,fs:24});c.insertObj(newTable);await tick();
    c.setState({cell:{r:0,c:0},cellSel:null});await tick();c.formatCells({fontWeight:'700',color:'#FF0000',fontSize:'40px'});await tick();c.patchCells({valign:'bottom',line:'#0000FF'});await tick();c.setRowHeight(140);await tick();
    doc=docNow();await EN.save(doc);doc=await EN.open(doc);await mount(doc);c=editor();const savedTable=doc.slides[0].objs.find(o=>o.t==='table' && o.rows[0][0]==='Bold cell');
    check(savedTable?.rowH[0]===140 && savedTable.cells['0:0'].valign==='bottom' && Math.abs(parseFloat(/font-size:([\d.]+)/.exec(savedTable.cells['0:0'].html)?.[1])-40)<.01, 'table rich text, row height or vertical alignment did not persist: '+JSON.stringify(savedTable));
    check(await EN.save(doc)===0,'formatted table no-op save rewrites the deck');
    check(doc.layouts?.length>0 && doc.layouts.some(l=>l.objs.some(o=>o.ph)), 'deck layout gallery lacks its own placeholders');
    const ownLayout=doc.layouts.find(l=>l.objs.some(o=>o.ph));c.addSlide(ownLayout.name,doc.slides.length);await tick();doc=docNow();await EN.save(doc);doc=await EN.open(doc);await mount(doc);c=editor();
    check(doc.slides.at(-1).layout===ownLayout.name && doc.slides.at(-1).objs.some(o=>o.ph),'new slide did not use deck layout');
    c.setState({cur:0});await tick();
    const numbered=K.txt({html:'<p>Alpha</p><p>Beta</p>',x:50,y:100,w:500,h:200});c.insertObj(numbered);await tick();c.listFormat('romanUcPeriod',3);await tick();doc=docNow();await EN.save(doc);doc=await EN.open(doc);await mount(doc);c=editor();
    check(doc.slides[0].objs.some(o=>o.html?.includes('data-marker="III. "') && o.html?.includes('data-marker="IV. "')), 'numbering format or start was lost');
    c.rehearsal={}; c.setState({show:{i:0,step:0}}); await tick(); c.slideStarted=Date.now()-2500; c.exitShow(); await tick();
    check(c.slide.advanceAfter>=2500 && c.slide.advanceAfter<3500,'rehearsal time not recorded');
    return ['fill opacity, gradient and image backgrounds, auto-advance and rehearsal, image paste, rich text and footer fields', 'copied group member ids and z-order', 'cross-deck chart relationships and inherited background reset', 'cross-format find/replace and whole-word matching', '24pt scale and rich text size stability','paragraph alignment and immediate whole-box format changes','safe link clicks, paste selection and delete','nested groups and table render/save','aspect ratio save/reopen','dynamic page numbers after reopen and reorder','presentation keyboard isolation'];
  });
  assert.deepEqual(errors, []); for (const result of results) console.log('PASS', result);
} finally {
  await browser?.close(); server?.close(); engine.kill(); rmSync(dir, { recursive: true, force: true });
}
