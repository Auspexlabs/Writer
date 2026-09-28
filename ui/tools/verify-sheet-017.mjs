// Real DOM + engine round trips. Run after dotnet build src/Writer.Cli.
// Set WRITER_NODE_MODULES when Playwright is supplied outside this repository.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../..', import.meta.url));
const { chromium } = await import(process.env.WRITER_NODE_MODULES ? pathToFileURL(join(process.env.WRITER_NODE_MODULES, 'playwright/index.mjs')).href : 'playwright');
const dir = mkdtempSync(join(tmpdir(), 'writer-sheet-017-'));
const engine = spawn('dotnet', [join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll'), 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
let browser, server;
try {
  const base = await new Promise((res, rej) => { let text = ''; engine.stderr.on('data', d => { text += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(text); if (m) res(m[1]); }); engine.once('exit', code => rej(new Error(`engine exit ${code}: ${text}`))); });
  server = createServer(async (req, res) => {
    try {
      if (req.url === '/favicon.ico') { res.writeHead(204).end(); return; }
      if (req.url === '/ui/test.dc.html') { res.setHeader('Content-Type', 'text/html'); res.end(`<!doctype html><meta charset="utf-8"><style>html,body,#dc-root{height:100%;margin:0;overflow:hidden}</style><script>window.__testBook = {active:0,sheets:[{name:'Data',cells:{A1:{v:'Label'},A2:{v:'one'},AAA10001:{v:'far cell'}}}]};window.__WRITER_LANG='zh';</script><script src="/ui/memo.js"></script><script src="/ui/i18n.js"></script><script src="/ui/vendor/offline.js"></script><script src="/ui/support.js"></script><body><x-dc><dc-import name="SheetEditor" doc="{{ doc }}" on-change="{{ onChange }}" toast="{{ toast }}" style="position:absolute;inset:0"></dc-import></x-dc><script type="text/x-dc" data-dc-script>class Component extends DCLogic { state = {doc:window.__testBook}; constructor(p){super(p);window.__sheetParent=this;} renderVals(){return {doc:this.state.doc,onChange:doc=>this.setState({doc}),toast:message=>{window.__lastToast=message;}};} }</script></body>`); return; }
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
  page.on('console', m => { if (m.type() === 'error') console.error('BROWSER', m.text()); });
  page.on('pageerror', e => console.error('PAGE', e.message));
  await page.goto(`http://127.0.0.1:${server.address().port}/ui/test.dc.html`);
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  await page.locator('.sh-cell').first().waitFor();
  assert.ok(await page.locator('.sh-cell').count() < 1000, 'grid mounts a viewport, not all rows');
  const name = page.locator('input').first(); await name.fill('AAA10001'); await name.press('Enter');

  await page.getByText('far cell', {exact:true}).waitFor({timeout:3000});
  assert.ok(await page.locator('.sh-cell').count() < 1000, 'far selection stays virtualized');
  const results = await page.evaluate(async () => {
    const E = await import('/ui/sheet-engine.js'), EN = await import('/ui/engine.js');
    const check = (yes, msg) => { if (!yes) throw new Error(msg); };
    const tick=()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
    const mount=async doc=>{window.__sheetParent.setState({doc});await tick();};
    const editor=()=>{const el=document.querySelector('.sh-cell');let f=el[Object.keys(el).find(k=>k.startsWith('__reactFiber'))];for(;f;f=f.return)if(f.stateNode?.logic?.tableSetup)return f.stateNode.logic;throw new Error('editor not mounted');};
    const clip = E.readTableHTML('<style>.xl1{font-weight:700;color:#ff0000}</style><table><tr><td class="xl1" colspan="2">hello<br>world</td></tr><tr><td>001</td><td style="background:#00ff00">12</td></tr></table>');
    check(clip.cells[0][0].v === 'hello\nworld' && clip.cells[0][0].s.b, 'Office HTML text/format');
    check(clip.merges[0].cs === 2, 'Office HTML merged cell');
    await EN.run(['create', 'grid.xlsx']);
    let doc = await EN.open({id:'grid',path:'grid.xlsx',type:'xlsx'});
    doc = E.editWorkbook(doc, (sh,d) => { sh.cells.AAA10001 = {v:'far saved'}; sh.cells.A1={v:'7'}; sh.cells.B1={v:'=SEQUENCE(2,2)'}; d.names={Total:'Sheet1!$A$1'}; sh.cells.E1={v:'=Total*2'}; sh.print={orientation:'landscape',fitWidth:1,fitHeight:0,titles:'Sheet1!$1:$1',footer:'&P / &N'}; });
    await EN.save(doc); doc = await EN.open(doc);
    check(doc.sheets[0].cells.AAA10001.v === 'far saved', 'far cell round trip');
    check(new E.Calc(doc).value(0,0,4) === 14, 'defined name after reopen');
    check(doc.sheets[0].cells.C2.spill === 'B1', 'cached dynamic array children');
    check(await EN.save(doc) === 0, 'unchanged workbook rewrites nothing');
    const src = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jFq8AAAAASUVORK5CYII=';
    doc = E.editWorkbook(doc, sh => { sh.images = [{ id:'test-image',src,x:96,y:48,w:120,h:80,look:{} }]; }); await EN.save(doc); doc = await EN.open(doc);
    check(doc.sheets[0].images.length === 1, 'inserted image missing');
    doc = E.editWorkbook(doc, sh => { sh.images[0].x=192; sh.images[0].w=240; }); await EN.save(doc); doc = await EN.open(doc);
    check(Math.abs(doc.sheets[0].images[0].x - 192) < 1 && Math.abs(doc.sheets[0].images[0].w - 240) < 1, 'image move/resize lost');
    const image = {...doc.sheets[0].images[0],src:await EN.dataUrlOf(doc.sheets[0].images[0].src)};
    doc = E.editWorkbook(doc, sh => { sh.images=[]; }); await EN.save(doc); doc = await EN.open(doc); check(doc.sheets[0].images.length === 0, 'image removal lost');
    doc = E.editWorkbook(doc, sh => { sh.images=[image]; }); await EN.save(doc); doc = await EN.open(doc); check(doc.sheets[0].images.length === 1, 'deleted image undo lost bytes');
    await EN.run(['create','tables.xlsx']);let tables=await EN.open({id:'tables',path:'tables.xlsx',type:'xlsx'});
    tables=E.editWorkbook(tables,sh=>{for(const [a,v] of Object.entries({A1:'Region',B1:'Amount',C1:'Tax',A2:'East',B2:3,A3:'East',B3:7,A4:'West',B4:5,E1:'=SUM(Sales[Amount])'}))sh.cells[a]={v};});
    await mount(tables);let c=editor();c.setState({sel:{r:3,c:2},anc:{r:0,c:0}});await tick();c.tableSetup();await tick();c.state.dlg.ok({name:'Sales',range:'A1:C4',style:'TableStyleMedium2',action:'set'});c.setState({dlg:null});await tick();
    check(window.__sheetParent.state.doc.sheets[0].tables.length===1,'table command did not create a table');
    c.setState({edit:{r:1,c:2,val:'=[@Amount]*2'}});c.commitEdit(0,0);await tick();
    tables=window.__sheetParent.state.doc;check(new E.Calc(tables).value(0,2,2)===14,'table formula did not fill calculated column');
    await EN.save(tables);tables=await EN.open(tables);check(tables.sheets[0].tables[0].columns[1].name==='Amount','native table columns lost');check(await EN.save(tables)===0,'table no-op save rewrote file');
    await mount(tables);c=editor();c.commit(sh=>c.setCellRaw(sh,0,1,'Revenue'));await tick();tables=window.__sheetParent.state.doc;
    check(tables.sheets[0].cells.E1.v==='=SUM(Sales[Revenue])','table header rename broke references');await EN.save(tables);tables=await EN.open(tables);check(new E.Calc(tables).value(0,0,4)===15,'renamed table formula wrong');
    await mount(tables);c=editor();c.insDel('r',2,1);await tick();tables=window.__sheetParent.state.doc;check(tables.sheets[0].tables[0].range==='A1:C5','row insert failed to extend table');
    c.setState({anc:{r:1,c:0},sel:{r:4,c:0}});c.groupRows();await tick();c.outlineLevel(1);await tick();tables=window.__sheetParent.state.doc;await EN.save(tables);tables=await EN.open(tables);
    check(tables.sheets[0].outline[0].collapsed && tables.sheets[0].outline[0].end===4,'outline collapse did not round trip');check(await EN.save(tables)===0,'outline no-op save rewrote file');
    await mount(tables);c=editor();c.outlineLevel(8);await tick();tables=window.__sheetParent.state.doc;await EN.save(tables);tables=await EN.open(tables);check(!tables.sheets[0].outline[0].collapsed&&!tables.sheets[0].hiddenRows.length,'outline expansion left rows hidden');
    await mount(tables);c=editor();c.setState({anc:{r:0,c:0},sel:{r:4,c:2}});c.insertChart('column','percent');await tick();tables=window.__sheetParent.state.doc;let chart=tables.sheets[0].charts[0];c.chartOptions(chart);await tick();c.state.dlg.ok({xTitle:'Region',yTitle:'Share',labels:'yes',stack:'percent'});c.setState({dlg:null});await tick();tables=window.__sheetParent.state.doc;await EN.save(tables);tables=await EN.open(tables);chart=tables.sheets[0].charts[0];check(chart.percentStacked&&chart.dataLabels&&chart.xTitle==='Region'&&chart.yTitle==='Share','chart percentage/labels/titles missing');
    await mount(tables);c=editor();let vals=c.chartVals(chart,new E.Calc(tables));check(vals.labels.some(x=>x.text==='100%')&&vals.labels.some(x=>x.text==='Share'),'percentage chart not rendered');
    c.commit(sh=>{const ch=sh.charts[0];ch.type='combo';ch.stacked=false;ch.percentStacked=false;ch.ser.forEach((s,i)=>s.kind=i?'line':'column');});await tick();tables=window.__sheetParent.state.doc;await EN.save(tables);tables=await EN.open(tables);check(tables.sheets[0].charts[0].type==='combo','combo chart missing');await mount(tables);c=editor();vals=c.chartVals(tables.sheets[0].charts[0],new E.Calc(tables));check(vals.rects.length&&vals.lines.length,'combo chart needs columns and line');check(await EN.save(tables)===0,'combo no-op save rewrote file');
    await EN.run(['create','pivot.xlsx']);let pivotDoc=await EN.open({id:'pivot',path:'pivot.xlsx',type:'xlsx'});pivotDoc=E.editWorkbook(pivotDoc,sh=>{for(const [a,v] of Object.entries({A1:'Region',B1:'Channel',C1:'Value',A2:'East',B2:'Web',C2:3,A3:'East',B3:'Store',C3:7,A4:'West',B4:'Web',C4:5}))sh.cells[a]={v};});
    await mount(pivotDoc);c=editor();c.setState({anc:{r:0,c:0},sel:{r:3,c:2}});c.pivotSetup();await tick();c.state.dlg.ok({row:'0',col:'1',value:'2',fn:'sum'});c.setState({dlg:null});await tick();pivotDoc=window.__sheetParent.state.doc;
    check(pivotDoc.sheets.length===2&&pivotDoc.sheets[1].cells.D4.v===15,'pivot command did not aggregate');await EN.save(pivotDoc);pivotDoc=await EN.open(pivotDoc);check(pivotDoc.sheets[1].pivots[0].name==='PivotTable1'&&pivotDoc.sheets[1].cells.D4.v==='15','native pivot did not reopen');check(await EN.save(pivotDoc)===0,'pivot no-op save rewrote file');
    pivotDoc=E.editWorkbook(pivotDoc,(sh,d)=>{d.sheets[0].cells.C2.v='30';d.active=1;});await mount(pivotDoc);c=editor();c.refreshPivots();await tick();pivotDoc=window.__sheetParent.state.doc;check(pivotDoc.sheets[1].cells.D4.v===42,'pivot refresh failed');await EN.save(pivotDoc);pivotDoc=await EN.open(pivotDoc);check(pivotDoc.sheets[1].cells.D4.v==='42','refreshed native cache/output stale');check(await EN.save(pivotDoc)===0,'refreshed pivot no-op rewrote file');
    await mount(pivotDoc);c=editor();c.insDel('r',0,1);await tick();pivotDoc=window.__sheetParent.state.doc;check(pivotDoc.sheets[1].pivots[0].range==='A2:D5','pivot output anchor did not move');await EN.save(pivotDoc);pivotDoc=await EN.open(pivotDoc);check(pivotDoc.sheets[1].cells.D5.v==='42','moving pivot output lost cached cells');
    pivotDoc={...pivotDoc,active:0};await mount(pivotDoc);c=editor();c.insDel('c',1,1);await tick();c.commit(sh=>c.setCellRaw(sh,0,1,'Inserted'));await tick();pivotDoc=window.__sheetParent.state.doc;await EN.save(pivotDoc);pivotDoc=await EN.open(pivotDoc);check(pivotDoc.sheets[1].pivots[0].values[0].field===3,'inserted source column did not remap pivot field');
    pivotDoc={...pivotDoc,active:1};await mount(pivotDoc);c=editor();c.refreshPivots();await tick();pivotDoc=window.__sheetParent.state.doc;check(pivotDoc.sheets[1].cells.D5.v===42,'pivot refresh after source column insert wrong');await EN.save(pivotDoc);pivotDoc=await EN.open(pivotDoc);check(pivotDoc.sheets[1].cells.D5.v==='42','pivot source column round trip failed');
    return ['native pivot creation, row/column aggregation, refresh and save/reopen', 'percentage and combination charts, value labels and axis titles', 'native tables, calculated columns, header rename and structural edits', 'outline collapse/expand save and reopen', 'virtual grid navigation to AAA10001', 'Office HTML multiline, style and merge import', 'far cell, names, dynamic arrays and print settings save/reopen', 'image insert, move, resize, delete and undo round trips'];
  });
  assert.deepEqual(errors, []); for (const result of results) console.log('PASS', result);
  if(process.argv.includes('--perf')) {
    const perf=await page.evaluate(async()=>{
      const E=await import('/ui/sheet-engine.js'),tick=()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))),start=performance.now(),cells={};
      for(let r=0;r<10000;r++)for(let c=0;c<100;c++)cells[E.A(r,c)]={v:String(r+c)};
      const book={id:'million',active:0,sheets:[{name:'Million',cells:E.shareCells(cells)}]};const prepared=performance.now();window.__sheetParent.setState({doc:book});await tick();const mounted=performance.now();
      const el=document.querySelector('.sh-cell');let f=el[Object.keys(el).find(k=>k.startsWith('__reactFiber'))],c;for(;f;f=f.return)if(f.stateNode?.logic?.commitEdit){c=f.stateNode.logic;break;}if(!c)throw new Error('missing real editor');
      const samples=[];for(let i=0;i<6;i++){c.setState({edit:{r:i+1,c:0,val:String(5000+i)}});await tick();const t=performance.now();c.commitEdit(0,0);await tick();samples.push(+(performance.now()-t).toFixed(2));if(window.__sheetParent.state.doc.sheets[0].cells[E.A(i+1,0)].v!==String(5000+i))throw new Error('edit lost');}
      return {cells:1000000,prepareMs:+(prepared-start).toFixed(2),initialRenderMs:+(mounted-prepared).toFixed(2),editToPaintMs:samples,mountedCells:document.querySelectorAll('.sh-cell').length};
    });console.log('PERF',JSON.stringify(perf));
  }
} finally {
  await browser?.close(); server?.close(); engine.kill(); rmSync(dir, { recursive: true, force: true });
}
