// Real DOM + engine round trips. Run after dotnet build src/Writer.Cli.
// Set WRITER_NODE_MODULES when Playwright is supplied outside this repository.
import assert from 'node:assert/strict';
import { spawn, execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, copyFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { resolve, join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../..', import.meta.url));
const verifyUi=resolve(process.env.WRITER_UI_ROOT||join(root,'ui'));
const verifyCli=process.env.WRITER_VERIFY_CLI||join(root,'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
const cliCommand=verifyCli.endsWith('.dll')?'dotnet':verifyCli,cliPrefix=verifyCli.endsWith('.dll')?[verifyCli]:[];
const shellSource = readFileSync(join(root, 'ui/index.dc.html'), 'utf8');
const shellHistory = shellSource.slice(shellSource.indexOf('  setDoc(nd, record, opts) {'), shellSource.indexOf('  scheduleSave(id) {')) + shellSource.slice(shellSource.indexOf('  undo(redo) {'), shellSource.indexOf('  toggleThumbsAnim() {'));
const { chromium } = await import(process.env.WRITER_NODE_MODULES ? pathToFileURL(join(process.env.WRITER_NODE_MODULES, 'playwright/index.mjs')).href : 'playwright');
const dir = mkdtempSync(join(tmpdir(), 'writer-word-017-'));
copyFileSync(join(root, 'tests/Writer.Tests/Fixtures/docx/wmf-preview.docx'), join(dir, 'wmf-preview.docx'));
const hfFixture=join(dir,'header-table.docx');
execFileSync(cliCommand,[...cliPrefix,'create',hfFixture]);execFileSync(cliCommand,[...cliPrefix,'set',hfFixture,'/','--prop','header=Header before']);
execFileSync('python3',['-c',`import zipfile,sys,xml.etree.ElementTree as ET
p=sys.argv[1]
with zipfile.ZipFile(p) as z:d={n:z.read(n) for n in z.namelist()}
n=next(n for n in d if n.startswith('word/header') and n.endswith('.xml'));ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';r=ET.fromstring(d[n]);r.append(ET.fromstring('<w:tbl xmlns:w="'+ns+'"><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Table before</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'));r.append(ET.Element('{'+ns+'}p'));d[n]=ET.tostring(r)
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in d.items():z.writestr(n,b)`,hfFixture]);
const boundFixture=join(dir,'bound-controls.docx');execFileSync(cliCommand,[...cliPrefix,'create',boundFixture]);
execFileSync('python3',['-c',`import zipfile,sys,xml.etree.ElementTree as E
p=sys.argv[1]
with zipfile.ZipFile(p) as z:d={n:z.read(n) for n in z.namelist()}
w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; rel='http://schemas.openxmlformats.org/package/2006/relationships'; office='http://schemas.openxmlformats.org/officeDocument/2006/relationships/'; store='{397AD51F-C7C7-4D58-8283-BAB11518916A}'
root=E.fromstring(d['word/document.xml']);body=root.find('{'+w+'}body')
for i in (1,2):body.insert(i-1,E.fromstring('<w:p xmlns:w="'+w+'"><w:sdt><w:sdtPr><w:id w:val="'+str(i)+'"/><w:dataBinding w:storeItemID="'+store+'" w:xpath="/f:form/f:name" w:prefixMappings="xmlns:f=&quot;urn:writer-form&quot;"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Before</w:t></w:r></w:sdtContent></w:sdt></w:p>'))
d['word/document.xml']=E.tostring(root);root=E.fromstring(d['word/_rels/document.xml.rels']);E.SubElement(root,'{'+rel+'}Relationship',Id='rIdBoundFormTest',Type=office+'customXml',Target='../customXml/item1.xml');d['word/_rels/document.xml.rels']=E.tostring(root)
d['customXml/item1.xml']=b'<form xmlns="urn:writer-form"><name>Before</name><other>Keep</other></form>'
d['customXml/itemProps1.xml']=('<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="'+store+'"/>').encode()
d['customXml/_rels/item1.xml.rels']=('<Relationships xmlns="'+rel+'"><Relationship Id="rIdProps" Type="'+office+'customXmlProps" Target="itemProps1.xml"/></Relationships>').encode()
E.register_namespace('','http://schemas.openxmlformats.org/package/2006/content-types');root=E.fromstring(d['[Content_Types].xml']);E.SubElement(root,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/customXml/itemProps1.xml',ContentType='application/vnd.openxmlformats-officedocument.customXmlProperties+xml');d['[Content_Types].xml']=E.tostring(root)
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in d.items():z.writestr(n,b)`,boundFixture]);
const engine = spawn(cliCommand, [...cliPrefix, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
let browser, server;
try {
  const base = await new Promise((res, rej) => { let text = ''; engine.stderr.on('data', d => { text += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(text); if (m) res(m[1]); }); engine.once('exit', code => rej(new Error(`engine exit ${code}: ${text}`))); });
  server = createServer(async (req, res) => {
    try {
      if (req.url === '/') { res.setHeader('Content-Type', 'text/html'); res.end('<!doctype html><meta charset="utf-8"><div id="editor" contenteditable="true"></div>'); return; }
      if (req.url === '/ui/word-test.dc.html') {
        res.setHeader('Content-Type', 'text/html');
        res.end(`<!doctype html><meta charset="utf-8"><link rel="stylesheet" href="/ui/assets/fonts/fonts.css"><style>html,body,#dc-root{height:100%;margin:0;overflow:hidden}</style><script>window.__WRITER_LANG='zh';</script><script src="/ui/memo.js"></script><script src="/ui/i18n.js"></script><script src="/ui/vendor/offline.js"></script><script src="/ui/support.js"></script><body><x-dc><dc-import name="WordEditor" doc="{{ doc }}" on-change="{{ onChange }}" on-undo="{{ onUndo }}" on-redo="{{ onRedo }}" toast="{{ toast }}" style="position:absolute;inset:0"></dc-import></x-dc><script type="text/x-dc" data-dc-script>class Component extends DCLogic { state = { docs:[{id:'test',type:'docx',html:'<p>Test</p>',loaded:true}] }; hist={}; constructor(p){super(p);window.__wordParent=this;} get doc(){return this.state.docs[0];} scheduleSave(){} toastMsg(message){window.__lastToast=message;} ${shellHistory} renderVals(){return {doc:this.doc,onChange:(doc,opts)=>this.setDoc(doc,!opts?.silent),onUndo:()=>this.undo(false),onRedo:()=>this.undo(true),toast:message=>this.toastMsg(message)};} }</script></body>`); return;
      }
      if (req.url.startsWith('/ui/')) {
        const path = resolve(verifyUi, '.' + new URL(req.url, 'http://local').pathname.slice(3));
        if (!path.startsWith(verifyUi + '/')) { res.writeHead(403).end(); return; }
        res.setHeader('Content-Type', path.endsWith('.js') ? 'text/javascript' : path.endsWith('.css') ? 'text/css' : 'text/html'); res.end(readFileSync(path)); return;
      }
      const chunks = []; for await (const chunk of req) chunks.push(chunk);
      const response = await fetch(base + req.url, { method: req.method, headers: { 'Content-Type': req.headers['content-type'] || 'application/json' }, ...(chunks.length ? { body: Buffer.concat(chunks) } : {}) });
      res.writeHead(response.status, { 'Content-Type': response.headers.get('Content-Type') || 'application/json' }); res.end(Buffer.from(await response.arrayBuffer()));
    } catch (error) { res.writeHead(500).end(String(error)); }
  });
  await new Promise(res => server.listen(0, '127.0.0.1', res));
  browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage();
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const results = await page.evaluate(async () => {
    const EN = await import('/ui/engine.js'), ed = document.querySelector('#editor');
    const check = (condition, message) => { if (!condition) throw new Error(message); };
    const plain = () => { const clone = ed.cloneNode(true); clone.querySelectorAll('[data-office-object]').forEach(e => e.remove()); return clone.textContent; };
    const results = [];
    await EN.run(['create', 'comments.docx']);
    for (const html of ['first <b>bold</b> ending', 'second <i>italic</i> ending', 'third paragraph']) await EN.run(['add', 'comments.docx', '/body', '--type', 'paragraph', '--prop', 'html=' + html]);
    let doc = await EN.open({ id: 'comments', path: 'comments.docx', type: 'docx' });
    ed.innerHTML = doc.html;
    const before = ed.textContent, range = document.createRange();
    range.setStart(ed.children[0].firstChild, 3); range.setEnd(ed.children[2].firstChild, 5);
    EN.wrapCommentSelection(ed, range, 'c-test');
    check(ed.children.length === 3 && ed.textContent === before, 'cross-paragraph comment changed text or paragraph count');
    check(!ed.querySelector('span p'), 'paragraph nested inside a comment span');
    doc.comments = [{ id: 'c-test', text: 'Across three paragraphs' }]; doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.children.length === 3 && ed.textContent === before, 'comment round trip changed text');
    check(ed.querySelectorAll('[data-cid]').length >= 3, 'comment anchors did not reopen across paragraphs');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged cross-paragraph comment was rewritten');
    results.push('cross-paragraph formatted comment: text, anchors, save and reopen');

    await EN.run(['create', 'objects.docx']);
    const ole = '<w:object><v:shape id="_x0000_i1025" style="width:72pt;height:36pt"/><o:OLEObject ProgID="Equation.DSMT4"/></w:object>';
    await EN.run(['set', 'objects.docx', '/body', '--raw', '<w:body><w:p><w:r>' + ole + '</w:r></w:p><w:p><w:r><w:t>left</w:t></w:r><w:r>' + ole + '</w:r><w:r><w:t>right</w:t></w:r></w:p></w:body>']);
    doc = await EN.open({ id: 'objects', path: 'objects.docx', type: 'docx' }); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 2, 'object display missing');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged object save mutated file');
    ed.querySelector('p').append(' edited'); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 2, 'text edit lost objects');
    ed.querySelector('p [data-office-object]').remove(); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 1, 'deleted inline object came back');
    ed.querySelector('[data-office-object]').remove(); doc.html = ed.innerHTML;
    await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 0, 'deleted standalone object came back');
    check(ed.textContent === 'leftright edited', 'object deletion changed text');
    results.push('Office objects: display, no-op save, text edit, inline and block deletion');

    await EN.run(['create', 'object-history.docx']);
    await EN.run(['set', 'object-history.docx', '/body', '--raw', '<w:body><w:p><w:r><w:t>abc</w:t></w:r><w:r>' + ole + '</w:r><w:r><w:t>def</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:body>']);
    doc = await EN.open({ id: 'history', path: 'object-history.docx', type: 'docx' }); ed.innerHTML = doc.html;
    const history = ed.innerHTML;
    ed.querySelector('[data-office-object]').remove(); await EN.save(doc, { root: ed });
    ed.innerHTML = history; await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 1 && plain() === 'abcdefsecond', 'undo after saved deletion lost the object or changed text');
    let object = ed.querySelector('[data-office-object]');
    ed.children[1].prepend(object); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    object = ed.children[1].querySelector('[data-office-object]');
    check(object && EN.offsetIn(ed.children[1], object) === 0, 'cross-paragraph object move was lost');
    ed.children[1].prepend('prefix'); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html; object = ed.children[1].querySelector('[data-office-object]');
    check(EN.offsetIn(ed.children[1], object) === 6, 'typing before an object did not update its offset');
    const clone = object.cloneNode(true); ed.children[1].append(clone); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('[data-office-object]').length === 2, 'copying an inline object lost its original markup');
    check(await EN.save(doc, { root: ed }) === 0, 'object copy produced a repeated save diff');
    object = ed.querySelector('[data-office-object]'); ed.append(object); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.lastElementChild.hasAttribute('data-office-object'), 'inline object did not become a standalone block');
    ed.prepend(ed.lastElementChild); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.firstElementChild.hasAttribute('data-office-object'), 'standalone object reorder was lost');
    const blockHistory = ed.innerHTML; ed.firstElementChild.remove(); await EN.save(doc, { root: ed });
    ed.innerHTML = blockHistory; await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.firstElementChild.hasAttribute('data-office-object'), 'standalone object undo after save failed');
    ed.querySelector('p').append(ed.firstElementChild); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('p [data-office-object]').length === 2, 'standalone object move into a paragraph failed');
    const pair = document.createElement('p'); ed.querySelectorAll('[data-office-object]').forEach(o => pair.append(o)); ed.append(pair);
    await EN.save(doc, { root: ed }); doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.lastElementChild.querySelectorAll('[data-office-object]').length === 2, 'paragraph containing only two objects failed to save');
    results.push('Office object history: saved deletion undo, movement, copy, ordering and block/inline conversion');

    await EN.run(['create', 'table-objects.docx']);
    await EN.run(['set', 'table-objects.docx', '/body', '--raw', '<w:body><w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>left</w:t></w:r><w:r>' + ole + '</w:r><w:r><w:t>right</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:body>']);
    doc = await EN.open({ id: 'tableobjects', path: 'table-objects.docx', type: 'docx' }); ed.innerHTML = doc.html;
    let cell = ed.querySelector('td'), cellObject = cell.querySelector('[data-office-object]');
    check(cellObject && EN.offsetIn(cell, cellObject) === 10, 'object in a later cell paragraph is missing or at the wrong position');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged table object save changed the file');
    cell.prepend('prefix'); await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html; cell = ed.querySelector('td'); cellObject = cell.querySelector('[data-office-object]');
    check(cellObject && EN.offsetIn(cell, cellObject) === 16, 'editing a table cell lost or moved its object');
    const tableHistory = ed.innerHTML; cellObject.remove(); await EN.save(doc, { root: ed });
    ed.innerHTML = tableHistory; await EN.save(doc, { root: ed });
    doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelectorAll('td [data-office-object]').length === 1, 'undo after deleting a saved table object failed');
    check(await EN.save(doc, { root: ed }) === 0, 'restored table object produced a repeated save diff');
    results.push('Objects inside table cells: later paragraphs, text edits, deletion and saved undo');
    await EN.run(['create', 'fields.docx']);
    await EN.run(['add', 'fields.docx', '/body', '--type', 'paragraph', '--prop', 'style=Heading1', '--prop', 'text=Chapter']);
    doc = await EN.open({ id: 'fields', path: 'fields.docx', type: 'docx' }); ed.innerHTML = doc.html;
    ed.insertAdjacentHTML('beforeend', '<p data-style="Caption" data-w-caption="图" data-w-captionchapter="1" data-w-bookmark="FigureOne">图 1-1 Test image</p><p>See <span data-field="REF FigureOne \\h" contenteditable="false">old</span>. Dear <span data-field="MERGEFIELD Name" contenteditable="false">«Name»</span></p>');
    ed.insertAdjacentHTML('beforeend', EN.tocHtml({ caption: '图', title: '图目录', levels: '3', entries: EN.editorHeadings(ed, 3, '图') }));
    doc.html = ed.innerHTML; await EN.save(doc, { root: ed }); doc = await EN.open(doc); ed.innerHTML = doc.html;
    check(ed.querySelector('[data-w-caption]').textContent === '图 1-1 Test image', 'chapter caption duplicated: ' + ed.querySelector('[data-w-caption]').outerHTML);
    check(ed.querySelector('[data-field]').textContent === '图 1-1 Test image', 'REF cached text not refreshed');
    check(ed.querySelector('[data-toc-caption="图"]').textContent.includes('Test image'), 'figure directory missing caption');
    check(await EN.save(doc, { root: ed }) === 0, 'unchanged fields were rewritten');
    const MM = await import('/ui/word-mailmerge.js');
    const records = MM.mergeRecords('Name,Address\n张三,"第一行\n第二行"\n李四,上海', 'records.csv');
    const merged = await MM.mergeDocuments(EN, doc, records);
    check(merged.length === 2, 'mail merge output count');
    const first = await EN.open(merged[0]), second = await EN.open(merged[1]);
    check(first.html.includes('张三') && second.html.includes('李四'), 'mail merge recipient values');
    check((await EN.open(doc)).html.includes('MERGEFIELD'), 'mail merge changed template');
    check(records[0].Address === '第一行\n第二行', 'quoted CSV multiline changed');
    results.push('chapter captions, native REF, figure directory and separate mail merge outputs');

    return results;
  });
  assert.equal(results.length, 5);
  for (const result of results) console.log('PASS', result);
  await page.goto(`http://127.0.0.1:${server.address().port}/ui/word-test.dc.html`);
  await page.locator('.wd-ed[contenteditable="true"]').waitFor();
  const editorResults = await page.evaluate(async () => {
    const EN = await import('/ui/engine.js');
    const tick = () => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    const check = (yes, msg) => { if (!yes) throw new Error(msg); };
    const editor = () => { const el = document.querySelector('[data-edroot]'); let f = el[Object.keys(el).find(k => k.startsWith('__reactFiber'))]; for (; f; f = f.return) if (f.stateNode?.logic?.replaceAll) return f.stateNode.logic; throw new Error('editor not mounted'); };
    const parent = window.__wordParent; parent.EN = EN;
    await EN.run(['create','table-formula.docx']);
    await EN.run(['add','table-formula.docx','/body','--type','table','--prop','data=[[20,3,""],[10,2,""],["","",""]]']);
    parent.setState({docs:[await EN.open({id:'tableFormula',path:'table-formula.docx',type:'docx'})]});await tick();
    let formulaEditor=editor();formulaEditor.EN=EN;
    const total=formulaEditor.edRef.current.querySelector('table').rows[2].cells[0],formulaRange=document.createRange();formulaRange.selectNodeContents(total);formulaRange.collapse(true);formulaEditor.edRef.current.focus();getSelection().removeAllRanges();getSelection().addRange(formulaRange);formulaEditor.range=formulaRange;
    await formulaEditor.tableFormula();await tick();formulaEditor.state.dlg.ok({formula:'=SUM(ABOVE)'});formulaEditor.setState({dlg:null});formulaEditor.flush();await tick();
    check(formulaEditor.edRef.current.querySelector('[data-field]').textContent==='30','table formula result not inserted');
    await EN.save(parent.doc,{root:formulaEditor.edRef.current});check((await EN.open(parent.doc)).html.includes('SUM(ABOVE)'),'table formula not saved as field');
    check(await EN.save(parent.doc,{root:formulaEditor.edRef.current})===0,'formula save did not stabilize');
    formulaEditor.history(false);await tick();formulaEditor=editor();check(!formulaEditor.edRef.current.querySelector('[data-field]'),'formula insertion undo failed');
    await EN.save(parent.doc,{root:formulaEditor.edRef.current});check(!(await EN.open(parent.doc)).html.includes('SUM(ABOVE)'),'saved formula undo failed');
    formulaEditor.history(true);await tick();formulaEditor=editor();check(formulaEditor.edRef.current.querySelector('[data-field]').textContent==='30','formula redo failed');
    formulaEditor.edRef.current.querySelector('table').rows[0].cells[0].textContent='40';formulaEditor.after();formulaEditor.flush();await tick();
    await EN.save(parent.doc,{root:formulaEditor.edRef.current});check(formulaEditor.edRef.current.querySelector('[data-field]').textContent==='50','save did not update table field in editor');
    check(await EN.save(parent.doc,{root:formulaEditor.edRef.current})===0,'updated table cache caused repeated writes');
    const F=await import('/ui/word-table-formula.js'),formulaFixture=document.createElement('div');
    formulaFixture.innerHTML='<table><tr><td><span data-bookmark-start="Sales"></span>3</td><td>4</td></tr><tr><td>5</td><td><span data-field="=PRODUCT(R1)">0</span><span data-bookmark-end="Sales"></span></td></tr></table><table><tr><td><span data-field="=SUM(Sales R1C1:R2C2)">0</span></td></tr></table>';
    const formulaTables=F.wordFormulaTables(formulaFixture);check(formulaTables[0].names.includes('Sales'),'DOM table bookmark not associated with source table');
    check(F.calculateWordTables(formulaTables,F.wordFormulaBookmarks(formulaFixture))[1][0].text==='24','DOM cross-table formula failed');
    await EN.run(['create', 'design.docx']);
    await EN.run(['add', 'design.docx', '/body', '--type', 'heading', '--prop', 'text=Design']);
    parent.setState({ docs: [await EN.open({ id: 'design', path: 'design.docx', type: 'docx' })] }); await tick();
    const design = editor(); design.EN = EN;
    await design.nativeChange('theme', 'mist'); await tick();
    parent.undo(false);await tick();await EN.save(parent.doc);check((await EN.open(parent.doc)).theme !== 'Writer mist','native theme undo did not persist');
    parent.undo(true);await tick();await EN.save(parent.doc);check((await EN.open(parent.doc)).theme === 'Writer mist','native theme redo did not persist');
    await design.nativeChange('styleSet', 'formal'); await tick();
    await design.nativeChange('pageBorder', { style: 'double', color: '112233', width: 1.5, space: 24 }); await tick();
    check(parent.doc.theme === 'Writer mist' && design.edRef.current.parentElement.querySelector('.wd-page-border').style.border.includes('double'), 'theme or page border not shown');
    await design.nativeChange('chartData', { title: 'Sales', type: 'column', labels: ['Jan', 'Feb'], values: [4, -7] }); await tick();
    check(design.edRef.current.querySelector('[data-office-object] svg'), 'native Word chart preview missing');
    check(await EN.save(parent.doc, { root: design.edRef.current }) === 0, 'unchanged native chart rewritten');
    await design.nativeChange('protected', 'true'); await tick();
    check(design.edRef.current.contentEditable === 'false', 'protected document remains editable');
    const savedHtml = parent.doc.html; design.change({ html: '<p>lost</p>' }); await tick();
    check(parent.doc.html === savedHtml, 'protected document allowed toolbar mutation');
    await design.nativeChange('protected', 'false'); await tick();
    check(design.edRef.current.contentEditable === 'true', 'unprotected document remains readonly');
    const changes = EN.compareWord({ type: 'docx', html: '<p>First</p><p>Gone</p><p>Last</p>' }, { type: 'docx', html: '<p>First</p><p>Last</p><p>Added</p>' });
    check(changes.some(r => r.kind === '删除' && r.before === 'Gone') && changes.some(r => r.kind === '新增' && r.after === 'Added'), 'comparison lost insertion/deletion details');
    let formula = await EN.open({ id: 'formula', path: 'wmf-preview.docx', type: 'docx' });
    parent.setState({ docs: [formula] }); await tick();
    const MF = await import('/ui/metafile.js'); await MF.paintMetafiles(editor().edRef.current);
    let svg = editor().edRef.current.querySelector('[data-wmf-ready="yes"] svg');
    check(svg && svg.querySelector('line') && svg.textContent.includes('α+β'), 'MathType WMF preview did not render vectors and Symbol glyphs as SVG');
    const rawBefore = await EN.run(['get', formula.path, '/body/object[1]', '--raw']);
    editor().edRef.current.querySelector('p').append(' edited');
    await EN.save(formula, { root: editor().edRef.current });
    check((await EN.run(['get', formula.path, '/body/object[1]', '--raw'])) === rawBefore, 'drawing the WMF preview changed the embedded object XML');
    await EN.run(['create', 'history-editor.docx']);
    await EN.run(['add', 'history-editor.docx', '/body', '--type', 'paragraph', '--prop', 'html=<b>Hel</b><i>lo</i> HELLO shelloworld']);
    await EN.run(['add', 'history-editor.docx', '/body', '--type', 'table', '--prop', 'rows=2', '--prop', 'cols=2']);
    let doc = await EN.open({ id: 'editor', path: 'history-editor.docx', type: 'docx' });
    parent.setState({ docs: [doc] }); await tick(); let c = editor(); c.EN = EN;
    c.setState({ fq: 'hello', fr: 'World', findWhole: true }); await tick();
    check(c.hits().length === 2, 'Word search does not cross formatting or respect whole words');
    c.setState({ findCase: true }); await tick(); check(c.hits().length === 0, 'Word case-sensitive search'); c.setState({ findCase: false }); await tick();
    const original = c.edRef.current.innerHTML; c.replaceAll(); await tick();
    check(c.edRef.current.querySelector('b').textContent === 'World' && c.edRef.current.textContent.includes('World World shelloworld'), 'replacement lost formatting or replaced part of a whole word');
    await EN.save(parent.doc, { root: c.edRef.current });
    c.history(false); await tick(); c = editor();
    check(c.edRef.current.querySelector('b').textContent === 'Hel' && c.edRef.current.querySelector('i').textContent === 'lo', 'replace all undo lost run formatting');
    await EN.save(parent.doc, { root: c.edRef.current });
    c.history(true); await tick(); c = editor(); check(c.edRef.current.textContent.includes('World World'), 'replace all redo');
    const cell = c.edRef.current.querySelector('td'), range = document.createRange(); range.selectNodeContents(cell); range.collapse(true);
    c.edRef.current.focus(); window.getSelection().removeAllRanges(); window.getSelection().addRange(range); c.range = range;
    c.tbl('rowBelow'); c.flush(); await tick(); c = editor();
    check(c.edRef.current.querySelectorAll('tr').length === 3, 'insert table row failed');
    await EN.save(parent.doc, { root: c.edRef.current });
    c.history(false); await tick(); c = editor(); check(c.edRef.current.querySelectorAll('tr').length === 2, 'table row undo');
    await EN.save(parent.doc, { root: c.edRef.current });
    const reopened = await EN.open(parent.doc); check(new DOMParser().parseFromString(reopened.html, 'text/html').querySelectorAll('tr').length === 2, 'table undo after save did not reach the file');
    c.history(true); await tick(); c = editor(); check(c.edRef.current.querySelectorAll('tr').length === 3, 'table row redo');
    const p0 = c.edRef.current.querySelector('p'), selection = window.getSelection();
    const selectParagraph = () => { const range = document.createRange(); range.selectNodeContents(p0); selection.removeAllRanges(); selection.addRange(range); c.range = range; c.edRef.current.focus(); };
    selectParagraph(); p0.dispatchEvent(new KeyboardEvent('keydown', { key: 'j', ctrlKey: true, bubbles: true, cancelable: true })); c.flush(); await tick();
    check(p0.style.textAlign === 'justify', 'Ctrl+J did not justify the paragraph');
    selectParagraph(); p0.dispatchEvent(new KeyboardEvent('keydown', { key: '5', ctrlKey: true, bubbles: true, cancelable: true })); c.flush(); await tick();
    check(p0.getAttribute('data-w-linespacing') === '1.5', 'Ctrl+5 did not apply 1.5 line spacing');
    const zoom = c.state.zoom; selectParagraph(); p0.dispatchEvent(new KeyboardEvent('keydown', { key: '=', ctrlKey: true, bubbles: true, cancelable: true })); c.flush(); await tick();
    check(c.state.zoom === zoom && (p0.querySelector('sub') || p0.innerHTML.includes('vertical-align: sub')), 'Ctrl+= zoomed instead of applying subscript');
    selectParagraph(); p0.dispatchEvent(new KeyboardEvent('keydown', { key: '2', ctrlKey: true, altKey: true, bubbles: true, cancelable: true })); c.flush(); await tick();
    check(c.edRef.current.querySelector('h2'), 'Ctrl+Alt+2 did not apply Heading 2');
    await EN.run(['create', 'fonts-editor.docx']);
    await EN.run(['add', 'fonts-editor.docx', '/body', '--type', 'paragraph', '--prop', 'html=<b>Hello 中文</b> <i>World 汉字</i>']);
    parent.setState({ docs: [await EN.open({ id: 'fonts', path: 'fonts-editor.docx', type: 'docx' })] }); await tick(); c = editor();
    const fontRoot = c.edRef.current, fontRange = document.createRange(); fontRange.selectNodeContents(fontRoot.querySelector('p'));
    selection.removeAllRanges(); selection.addRange(fontRange); c.range = fontRange;
    c.setScriptFont('west', 'Georgia'); c.setScriptFont('ea', 'Noto Sans SC'); c.flush(); await tick();
    check(fontRoot.querySelector('b').textContent === 'Hello 中文' && fontRoot.querySelector('i').textContent === 'World 汉字', 'independent font controls changed text or formatting');
    check(c.fontCss.textContent.includes('PingFang') && c.fontCss.textContent.includes('unicode-range:'), 'script font aliases lost bundled font sources');
    await EN.save(parent.doc, { root: fontRoot });
    const fontDoc = await EN.open(parent.doc); fontRoot.innerHTML = fontDoc.html;
    check(fontRoot.querySelector('[data-font-west="Georgia"][data-font-ea="Noto Sans SC"]'), 'Chinese and Latin fonts did not reopen');
    check(await EN.save(fontDoc, { root: fontRoot }) === 0, 'independent font save was not stable');
    await EN.run(['create', 'lists-editor.docx']);
    await EN.run(['add', 'lists-editor.docx', '/body', '--type', 'paragraph', '--prop', 'text=First', '--prop', 'list=upperRoman', '--prop', 'listStart=4']);
    await EN.run(['add', 'lists-editor.docx', '/body', '--type', 'paragraph', '--prop', 'text=Second', '--prop', 'list=upperRoman']);
    await EN.run(['add', 'lists-editor.docx', '/body', '--type', 'heading', '--prop', 'text=Chapter', '--prop', 'level=1']);
    await EN.run(['add', 'lists-editor.docx', '/body', '--type', 'heading', '--prop', 'text=Part', '--prop', 'level=2']);
    parent.setState({ docs: [await EN.open({ id: 'lists', path: 'lists-editor.docx', type: 'docx' })] }); await tick(); c = editor();
    check(c.edRef.current.querySelector('ol').start === 4, 'custom numbering start did not open');
    const li = c.edRef.current.querySelector('li'), lr = document.createRange(); lr.selectNodeContents(li); selection.removeAllRanges(); selection.addRange(lr); c.range = lr;
    c.restartList(true, 7); c.flush(); await tick(); await EN.save(parent.doc, { root: c.edRef.current });
    let listDoc = await EN.open(parent.doc); const listHtml = new DOMParser().parseFromString(listDoc.html, 'text/html');
    check(listHtml.querySelector('ol').start === 7 && listHtml.querySelectorAll('li').length === 2, 'numbering start changed after save');
    check(await EN.save(listDoc) === 0, 'custom numbering did not stabilize');
    c.change({ headingNumbering: 'chapter' }); await tick();
    check(c.edRef.current.querySelector('h1').getAttribute('data-heading-label') === '第1章 ' && c.edRef.current.querySelector('h2').getAttribute('data-heading-label') === '1.1 ', 'heading numbering did not render');
    await EN.save(parent.doc, { root: c.edRef.current }); listDoc = await EN.open(parent.doc); check(listDoc.headingNumbering === 'chapter', 'heading numbering did not save');
    await EN.run(['create', 'revisions-editor.docx']);
    await EN.run(['add', 'revisions-editor.docx', '/body', '--type', 'paragraph', '--prop', 'html=before <ins data-author="Ann">added</ins> <del data-author="Bob">removed</del> after']);
    await EN.run(['set', 'revisions-editor.docx', '/', '--prop', 'track=true']);
    parent.setState({ docs: [await EN.open({ id: 'revisions', path: 'revisions-editor.docx', type: 'docx' })] }); await tick(); c = editor();
    check(c.edRef.current.querySelector('ins').getAttribute('data-rid'), 'revision identity was not exposed');
    c.stepRevision(1); await c.resolveOne(true); await tick(); c = editor();
    check(!c.edRef.current.querySelector('ins') && c.edRef.current.querySelector('del'), 'accepting one revision changed another');
    c.stepRevision(1); await c.resolveOne(false); await tick(); c = editor();
    check(!c.edRef.current.querySelector('del') && c.edRef.current.textContent.includes('added removed'), 'rejecting one deletion failed');
    const rr = document.createRange(); rr.selectNodeContents(c.edRef.current.querySelector('p')); selection.removeAllRanges(); selection.addRange(rr); c.range = rr; c.exec('bold'); c.flush(); await tick();
    await EN.save(parent.doc, { root: c.edRef.current }); check(parent.doc.formatRevisions.length > 0, 'formatting change was not tracked');
    const formats = [...parent.doc.formatRevisions]; for (const revision of formats) { await c.resolveFormat(revision.id, false); await tick(); c = editor(); }
    check(!c.edRef.current.querySelector('b,strong') && !c.edRef.current.querySelector('ins,del'), 'format rejection changed text or left bold formatting');
    await EN.run(['create', 'sections-editor.docx']);
    await EN.run(['add', 'sections-editor.docx', '/body', '--type', 'paragraph', '--prop', 'text=Preface', '--prop', 'sectionBreak=nextPage', '--prop', 'pageNumberFormat=lowerRoman', '--prop', 'pageNumberStart=1', '--prop', 'footer=Preface {page}']);
    await EN.run(['add', 'sections-editor.docx', '/body', '--type', 'paragraph', '--prop', 'text=Body text']);
    await EN.run(['set', 'sections-editor.docx', '/', '--prop', 'lastSection=' + JSON.stringify({ pageNumberStart: '1', footer: 'Body {page}' }), '--prop', 'evenAndOdd=true']);
    parent.setState({ docs: [await EN.open({ id: 'sections', path: 'sections-editor.docx', type: 'docx' })] }); await tick(); c = editor(); c.refreshInfo(); await tick();
    check(c.state.info.pages === 2, 'section break did not make two pages');
    check(c.pageMeta(0).label === 'i' && c.pageMeta(1).label === '1', 'section page numbers did not restart');
    c.setSectionFields({ evenFooter: 'Even body {page}' }, 1); await tick(); c.flush(); await EN.save(parent.doc, { root: c.edRef.current });
    const sectionDoc = await EN.open(parent.doc);
    check(sectionDoc.html.includes('Preface {page}') && sectionDoc.evenFooter === 'Even body {page}', 'section header save overwrote the preface');
    check(await EN.save(sectionDoc) === 0, 'section metadata created a repeated save diff');
    const printed = await c.edRef.current.__writerPrint(), frame = document.createElement('iframe'); document.body.append(frame);
    frame.contentDocument.open(); frame.contentDocument.write('<style>' + printed.css + '</style>' + printed.body); frame.contentDocument.close(); await tick();
    const pages = frame.contentDocument.querySelectorAll('.writer-print-page');
    check(pages.length === 2 && pages[0].textContent.includes('Preface i') && pages[1].textContent.includes('Even body 1'), 'print did not use section page fields');
    check(pages[0].querySelector('.wd-ed').textContent === 'Preface' && pages[1].querySelector('.wd-ed').textContent === 'Body text', 'print duplicated or lost page content');
    window.__wordPrint = printed; frame.remove();
    await EN.run(['create','native-dialogs.docx']);parent.setState({docs:[await EN.open({id:'nativeDialogs',path:'native-dialogs.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;
    const dialog=async selector=>{for(let i=0;i<200;i++){const el=document.querySelector(selector);if(el)return el;await new Promise(r=>setTimeout(r,10));}throw Error('dialog not shown: '+selector);};
    let action=c.insertChart(),modal=await dialog('[data-office-chart-editor]');modal.querySelector('[name=table]').value='Month\tActual\tPlan\nJan\t12\t10\nFeb\t20\t16';modal.querySelector('form').requestSubmit();await action;await tick();c=editor();
    let nativeObject=c.edRef.current.querySelector('[data-office-object]');check(nativeObject,'chart dialog did not insert an object: '+window.__lastToast);check(JSON.parse(nativeObject.dataset.officeProps).chart.series.length===2,'chart series from dialog lost');
    action=c.insertSmartArt();modal=await dialog('[data-smartart-editor]');modal.querySelector('input').value='Browser node';modal.querySelector('input').dispatchEvent(new Event('input'));[...modal.querySelectorAll('button')].find(b=>b.textContent==='保存 SmartArt').click();await action;await tick();c=editor();
    nativeObject=[...c.edRef.current.querySelectorAll('[data-office-object]')].find(o=>JSON.parse(o.dataset.officeProps).type==='smartart');check(nativeObject?.textContent.includes('Browser node'),'SmartArt dialog did not insert cached preview: '+window.__lastToast);
    c.contextTarget=nativeObject;action=c.insertSmartArt(true);modal=await dialog('[data-smartart-editor]');modal.querySelector('input').value='Edited browser node';modal.querySelector('input').dispatchEvent(new Event('input'));[...modal.querySelectorAll('button')].find(b=>b.textContent==='保存 SmartArt').click();await action;await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});check((await EN.open(parent.doc)).html.includes('Edited browser node'),'SmartArt dialog edit did not reopen');
    c.history(false);await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});check(!(await EN.open(parent.doc)).html.includes('Edited browser node'),'SmartArt native saved undo failed');
    check(nativeObject.querySelectorAll('line').length===2,'SmartArt process connectors absent in editor');
    await EN.run(['create','forms-editor.docx']);await EN.run(['add','forms-editor.docx','/body','--type','paragraph','--prop','text=Left Right']);parent.setState({docs:[await EN.open({id:'formsEditor',path:'forms-editor.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;
    let formRange=document.createRange();formRange.setStart(c.edRef.current.querySelector('p').firstChild,5);formRange.collapse(true);c.range=formRange;c.restore();
    action=c.contentControl();modal=await dialog('[data-word-control-editor]');modal.querySelector('[name=type]').value='checkbox';modal.querySelector('[name=type]').dispatchEvent(new Event('change'));modal.querySelector('[name=title]').value='Agreement';modal.querySelector('[name=checked]').checked=true;modal.querySelector('form').requestSubmit();await action;await tick();c=editor();c.flush();await EN.save(parent.doc,{root:c.edRef.current});
    let formDoc=await EN.open(parent.doc);check(formDoc.html.includes('☒')&&formDoc.html.includes('data-control'),'native checkbox lost on reopen');check(await EN.save(formDoc)===0,'content control no-op save changed document');parent.setState({docs:[formDoc]});await tick();c=editor();let formChip=c.edRef.current.querySelector('[data-control]');
    action=c.contentControl(formChip);modal=await dialog('[data-word-control-editor]');modal.querySelector('[name=checked]').checked=false;modal.querySelector('form').requestSubmit();await action;await tick();c=editor();c.flush();await EN.save(parent.doc,{root:c.edRef.current});check((await EN.open(parent.doc)).html.includes('☐'),'checkbox edit not persisted');
    c.history(false);await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});check((await EN.open(parent.doc)).html.includes('☒'),'saved checkbox edit undo failed');
    const controlBeforeDelete=structuredClone(parent.doc);formChip=c.edRef.current.querySelector('[data-control]');formChip.remove();c.after();await tick();c.flush();await EN.save(parent.doc,{root:c.edRef.current});const formDeleted=await EN.open(parent.doc);check(!formDeleted.html.includes('data-control'),'deleted control came back');const formRestored=EN.historyModel(controlBeforeDelete,formDeleted);await EN.save(formRestored);check((await EN.open(formRestored)).html.includes('☒'),'deleted control saved undo did not restore');
    parent.setState({docs:[await EN.open(formRestored)]});await tick();c=editor();formChip=c.edRef.current.querySelector('[data-control]');formChip.parentElement.append(formChip.cloneNode(true));c.after();await tick();c.flush();await EN.save(parent.doc,{root:c.edRef.current});formDoc=await EN.open(parent.doc);const formRoot=document.createElement('div');formRoot.innerHTML=formDoc.html;check(formRoot.querySelectorAll('[data-control]').length===2&&new Set([...formRoot.querySelectorAll('[data-control]')].map(el=>el.dataset.controlId)).size===2,'copied control IDs collided');
    parent.setState({docs:[await EN.open({id:'boundForms',path:'bound-controls.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;
    let boundChips=[...c.edRef.current.querySelectorAll('[data-control]')];check(boundChips.length===2&&JSON.parse(boundChips[0].dataset.control).bound,'bound form fixture missing');
    action=c.contentControl(boundChips[1]);modal=await dialog('[data-word-control-editor]');modal.querySelector('[name=value]').value='王 & Li';modal.querySelector('form').requestSubmit();await action;await tick();c=editor();c.flush();await EN.save(parent.doc,{root:c.edRef.current});
    check([...c.edRef.current.querySelectorAll('[data-control]')].every(el=>el.textContent==='王 & Li'),'bound mirrors did not update in editor');let boundDoc=await EN.open(parent.doc);const boundRoot=document.createElement('div');boundRoot.innerHTML=boundDoc.html;check([...boundRoot.querySelectorAll('[data-control]')].every(el=>el.textContent==='王 & Li'),'bound mirrors lost on native reopen');check(await EN.save(parent.doc,{root:c.edRef.current})===0,'bound controls rewrite on no-op save');
    c.history(false);await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});boundDoc=await EN.open(parent.doc);boundRoot.innerHTML=boundDoc.html;check([...boundRoot.querySelectorAll('[data-control]')].every(el=>el.textContent==='Before'),'bound controls saved undo failed');
    c.history(true);await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});boundRoot.innerHTML=(await EN.open(parent.doc)).html;check([...boundRoot.querySelectorAll('[data-control]')].every(el=>el.textContent==='王 & Li'),'bound controls saved redo failed');
    await EN.run(['create','range-bookmarks.docx']);
    await EN.run(['add','range-bookmarks.docx','/body','--type','paragraph','--prop','html=aa<b>selected</b>']);
    await EN.run(['add','range-bookmarks.docx','/body','--type','paragraph','--prop','text=morezz']);
    parent.setState({docs:[await EN.open({id:'bookmarks',path:'range-bookmarks.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;
    let br=document.createRange();br.setStart(c.edRef.current.children[0].firstChild,2);br.setEnd(c.edRef.current.children[1].firstChild,4);selection.removeAllRanges();selection.addRange(br);c.range=br;
    const oldPrompt=c.prompt;c.prompt=(_title,_fields,done)=>done({name:'Selection'});c.addBookmark();c.prompt=oldPrompt;await tick();c=editor();
    check(c.edRef.current.children.length===2&&c.edRef.current.textContent==='aaselectedmorezz','range bookmark changed paragraph structure or text');
    check(c.refTargets().find(t=>t.name==='Selection')?.text==='selectedmore','cross-reference target includes outside text');
    c.flush();await EN.save(parent.doc,{root:c.edRef.current});let bookmarked=await EN.open(parent.doc);
    check(bookmarked._orig.rangeBookmarks[0]?.startOffset===2&&bookmarked._orig.rangeBookmarks[0]?.endOffset===4,'bookmark offsets lost on reopen');
    check(await EN.save(bookmarked)===0,'range bookmark no-op save changed document');
    c.deleteBookmark('Selection');await tick();c=editor();c.flush();await EN.save(parent.doc,{root:c.edRef.current});check(!(await EN.open(parent.doc))._orig.rangeBookmarks.length,'deleted range bookmark remained in file');
    c.history(false);await tick();c=editor();await EN.save(parent.doc,{root:c.edRef.current});check((await EN.open(parent.doc))._orig.rangeBookmarks.length===1,'saved undo did not restore range bookmark');
    const MM=await import('/ui/word-mailmerge.js'),ME=await import('/ui/word-mailmerge-editor.js');
    const recipients=Array.from({length:23},(_,i)=>({Name:'Recipient '+i,Address:'Road '+i}));
    const wizard=ME.editMailMerge({records:recipients},{html:'<p><span data-field="MERGEFIELD Name">Name</span></p>'});
    const md=await dialog('[data-mail-merge-editor]');md.querySelector('[name=mode]').value='labels';md.querySelector('[name=mode]').dispatchEvent(new Event('input'));md.querySelector('[name=address]').value='{Name}\n{Address}';md.querySelector('[name=address]').dispatchEvent(new Event('input'));[...md.querySelectorAll('button')].find(b=>b.textContent==='生成文档').click();const mergeAnswer=await wizard;
    check(mergeAnswer.records.length===23&&mergeAnswer.options.mode==='labels','mail merge wizard lost recipients');
    let labelDoc=await EN.open(await EN.create('docx','label-test')),labelLayout=MM.mailLayout(mergeAnswer.records,mergeAnswer.options);labelDoc.html=labelLayout.html;labelDoc.page={...labelDoc.page,...labelLayout.page};await EN.save(labelDoc);labelDoc=await EN.open(labelDoc);
    parent.setState({docs:[labelDoc]});await tick();c=editor();c.refreshInfo();await tick();check(c.edRef.current.querySelectorAll('table').length===2&&c.edRef.current.textContent.includes('Recipient 22'),'native labels lost recipients');check(c.state.info.pages===2,'label layout made extra pages: '+c.state.info.pages);check(await EN.save(labelDoc)===0,'label no-op save changed file');
    let envelope=await EN.open(await EN.create('docx','envelope-test')),envelopeLayout=MM.mailLayout(recipients.slice(0,2),{mode:'envelopes',address:'{Name}\n{Address}',sender:'Sender',width:22,height:11});envelope.html=envelopeLayout.html;envelope.page={...envelope.page,...envelopeLayout.page};await EN.save(envelope);envelope=await EN.open(envelope);parent.setState({docs:[envelope]});await tick();c=editor();c.refreshInfo();await tick();
    check(Math.abs(c.edRef.current.closest('.wd-page')?.offsetWidth||0)>0||c.renderVals().pageW,'envelope editor failed to render');check(c.state.info.pages===2,'envelope pagination wrong: '+c.state.info.pages);check(envelope.page.size!=='A4','custom envelope dimensions fell back to A4');check(await EN.save(envelope)===0,'envelope no-op save changed file');
    parent.setState({docs:[{id:'tablePagination',type:'docx',loaded:false,html:'<p>Table heading</p><table>'+Array.from({length:70},(_,i)=>'<tr'+(i===0?' data-w-header="true"':'')+'><td><p>Row '+i+'</p></td><td><p>Cell '+i+'</p></td></tr>').join('')+'</table><p>After table</p>'}]});await tick();c=editor();c.refreshInfo();await tick();
    check(c.state.info.pages>=3,'long table did not paginate');
    const tableRoot=c.edRef.current,tableBox=tableRoot.getBoundingClientRect(),scale=tableBox.width/tableRoot.offsetWidth;
    for(const row of tableRoot.querySelectorAll('tr')){const box=row.querySelector('p').getBoundingClientRect(),top=(box.top-tableBox.top)/scale,bottom=(box.bottom-tableBox.top)/scale,step=1123+20,body=1123-96*2;check(Math.floor(top/step)===Math.floor((bottom-1)/step)&&bottom%step<=body+3,'table row crossed page body: '+JSON.stringify({text:row.textContent,top,bottom,body,step}));}
    check(tableRoot.querySelectorAll('tr').length===70,'pagination changed table structure');
    check(c.repeatedHeaders.length>=2&&c.repeatRef.current.textContent.includes('Row 0'),'table headers did not repeat');
    const tablePrint=await c.edRef.current.__writerPrint();check(tablePrint.body.includes('wd-repeat-header'),'printed table lost repeated headers');
    parent.setState({docs:[{id:'lineNumbering',type:'docx',loaded:false,html:'<p>'+('International documentation information. '.repeat(40))+'</p><p>Last line</p>',lineNumbers:true,hyphenation:true}]});await tick();c=editor();c.refreshInfo();await tick();
    check(c.lineRef.current.querySelectorAll('[data-line-number]').length>3,'rendered line numbers missing');check(getComputedStyle(c.edRef.current).hyphens==='auto'&&c.edRef.current.lang==='en','automatic English hyphenation missing');
    const numberedPrint=await c.edRef.current.__writerPrint();check(numberedPrint.body.includes('data-line-number='),'print lost line numbers');check(!c.edRef.current.innerHTML.includes('data-line-number'),'line numbers entered document text');
    let mixedDoc=await EN.open(await EN.create('docx','mixed-paper-test'));
    mixedDoc.html='<p>Small first section</p><p data-w-sectionbreak="nextPage" data-w-page="A5" data-w-orientation="portrait" data-w-margin="narrow" data-w-header="Small header">End first</p><p>'+('Landscape section text. '.repeat(450))+'</p><p data-w-sectionbreak="nextPage" data-w-page="A4" data-w-orientation="landscape" data-w-margin="normal" data-w-header="Landscape header">End landscape</p><p>Portrait final section</p>';
    mixedDoc.page={...mixedDoc.page,size:'Letter',orient:'portrait',margin:'normal'};mixedDoc.header='Letter header';await EN.save(mixedDoc);mixedDoc=await EN.open(mixedDoc);parent.setState({docs:[mixedDoc]});await tick();c=editor();c.refreshInfo();await tick();
    check(c.pageBoxes?.length>=4,'mixed-size sections did not create their own pages');
    const mbx=c.pageBoxes;check(mbx[0].w===559&&mbx[0].h===794,'A5 page dimensions wrong: '+JSON.stringify(mbx[0]));check(mbx[1].w===1123&&mbx[1].h===794,'landscape section dimensions wrong');check(mbx.at(-1).w===816&&mbx.at(-1).h===1056,'Letter final page dimensions wrong');
    const mixedRoot=c.edRef.current,mixedRect=mixedRoot.getBoundingClientRect(),mixedScale=mixedRect.width/mixedRoot.offsetWidth;
    for(const p of mixedRoot.children){const range=document.createRange();range.selectNodeContents(p);for(const r of [...range.getClientRects()].filter(r=>r.width>.5&&r.height>.5)){const top=(r.top-mixedRect.top)/mixedScale,bottom=(r.bottom-mixedRect.top)/mixedScale,left=(r.left-mixedRect.left)/mixedScale,right=(r.right-mixedRect.left)/mixedScale,g=mbx[c.physicalPage((top+bottom)/2)],x=g.left+g.m.left-96;check(top>=g.y-2&&bottom<=g.y+g.contentH+2&&left>=x-2&&right<=x+g.contentW+2,'mixed-section text outside page: '+JSON.stringify({top,bottom,left,right,g}));}}
    const mp=await mixedRoot.__writerPrint(),mpRoot=document.createElement('div');mpRoot.innerHTML=mp.body;const mixedPrinted=[...mpRoot.querySelectorAll('.writer-print-page')];check(mixedPrinted.length===mbx.length&&mixedPrinted[0].style.width==='559px'&&mixedPrinted[1].style.width==='1123px'&&mixedPrinted.at(-1).style.width==='816px','mixed paper print lost dimensions');check(mp.body.includes('Small header')&&mp.body.includes('Landscape header')&&mp.body.includes('Letter header'),'mixed-section headers lost');check(mp.css.includes('@page writer-page-1{size:1123px 794px'),'mixed-size print page rules missing');
    check(await EN.save(mixedDoc,{root:mixedRoot})===0,'mixed section layout changed saved content');
    const paragraphs = Array.from({ length: 900 }, (_, i) => `<p>Paragraph ${i}. ${'A long document needs reliable page boundaries. '.repeat(35)}</p>`).join('');
    parent.setState({ docs: [{ id: 'pagination', type: 'docx', html: paragraphs, loaded: false }] }); await tick();
    c = editor(); c.refreshInfo();
    check(c.state.info.pages > 200, 'long-document fixture did not cover enough pages');
    const perf = [];
    for (const index of [850, 400, 0, 700]) {
      const p = c.edRef.current.children[index], firstRule = c.pgCss.sheet.cssRules[0];
      if (index === 700) p.textContent = 'Short paragraph.'; else p.append(' This edit adds several new lines to an existing paragraph.'.repeat(12));
      const start = performance.now(); c.refreshInfo(true); const elapsed = performance.now() - start;
      const partial = c.state.info.at.map(a => a.slice()), css = c.pgRules.slice();
      if (index > 10) check(c.pgCss.sheet.cssRules[0] === firstRule, 'incremental pagination replaced an unaffected page rule');
      const fullStart = performance.now(); c.refreshInfo(); const fullElapsed = performance.now() - fullStart;
      const full = c.state.info.at;
      check(partial.length === full.length && partial.every((a, i) => a.every((v, j) => Math.abs(v - full[i][j]) <= 1)), `incremental pagination differs from a full layout at paragraph ${index}: ${JSON.stringify({ partial: partial.find((a, i) => a.some((v, j) => Math.abs(v - (full[i]?.[j] || 0)) > 1)), full: full[partial.findIndex((a, i) => a.some((v, j) => Math.abs(v - (full[i]?.[j] || 0)) > 1))] })}`);
      perf.push({ paragraph: index, pages: full.length, incrementalMs: Math.round(elapsed), fullMs: Math.round(fullElapsed) });
    }
    window.__wordPaginationPerf = perf;
    await EN.run(['create','wildcard-editor.docx']);await EN.run(['add','wildcard-editor.docx','/body','--type','paragraph','--prop','html=<b>Zhao Hong</b> and Wang Ming']);
    parent.setState({docs:[await EN.open({id:'wildcards',path:'wildcard-editor.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;c.setState({fq:'([A-Z][a-z]@) ([A-Z][a-z]@)',fr:'\\2, \\1',findWildcards:true,findCase:true,findFormat:{bold:true}});await tick();check(c.hits().length===1,'format wildcard search did not isolate bold name');c.replaceAll();await tick();check(c.edRef.current.textContent.includes('Hong, Zhao and Wang Ming'),'wildcard capture replacement incorrect');await EN.save(parent.doc,{root:c.edRef.current});c.history(false);await tick();await EN.save(parent.doc,{root:editor().edRef.current});check((await EN.open(parent.doc)).html.includes('Zhao Hong'),'saved wildcard replacement undo failed');
    await EN.run(['create','replace-format.docx']);let formatDoc=await EN.open({id:'formatReplace',path:'replace-format.docx',type:'docx'});formatDoc.html='<p><u><b><i>Left TARGET Right TARGET end</i></b></u></p><p>Other TARGET</p>';await EN.save(formatDoc);parent.setState({docs:[await EN.open(formatDoc)]});await tick();c=editor();c.EN=EN;
    const formatBefore=parent.doc;c.setState({fq:'TARGET',findWildcards:false,findFormat:{bold:true},replaceKeepText:true,replaceFormat:{bold:false,italic:false,underline:false,font:'Arial',size:18,color:'#CC0000'}});await tick();check(c.hits().length===2,'format-only replacement search count');c.replaceAll();await tick();check(c.edRef.current.textContent==='Left TARGET Right TARGET endOther TARGET','format-only replacement changed text');const targetRuns=JSON.parse(EN.runsOf(c.edRef.current.querySelector('p')));for(const run of targetRuns.filter(r=>r.t==='TARGET')){const style=JSON.parse(run.s);check(!style[0]&&!style[1]&&!style[2]&&style[6]==='CC0000'&&style[8]==='18','replacement format not reflected in native runs: '+run.s);}check(targetRuns.filter(r=>!r.t.includes('TARGET')).every(r=>JSON.parse(r.s)[2]),'format replacement stripped underline from surrounding text');
    await EN.save(parent.doc,{root:c.edRef.current});formatDoc=await EN.open(parent.doc);const formatRoot=document.createElement('div');formatRoot.innerHTML=formatDoc.html;check(JSON.parse(EN.runsOf(formatRoot)).filter(r=>r.t==='TARGET').length===2,'replacement formatting lost after save');c.history(false);await tick();await EN.save(parent.doc,{root:editor().edRef.current});check(JSON.parse(EN.runsOf((await EN.open(parent.doc)).html)).filter(r=>r.t.includes('Left')).some(r=>JSON.parse(r.s)[0]&&JSON.parse(r.s)[1]&&JSON.parse(r.s)[2]),'saved format replacement undo failed');

    parent.setState({docs:[await EN.open({id:'headerTable',path:'header-table.docx',type:'docx'})]});await tick();c=editor();c.EN=EN;c.enterHF('header',0);await tick();let hfCell=c.hfRef.current.querySelector('[data-w-hf-cell]');check(hfCell?.isContentEditable,'header table cell remains read-only');hfCell.innerHTML='<b>Table after</b><div>Second line</div>';c.hfInput();c.commitHF();await tick();await EN.save(parent.doc,{root:c.edRef.current});const headerOpen=await EN.open(parent.doc);check(headerOpen.header.includes('Table after')&&!headerOpen.header.includes('Table before'),'header cell edit lost on save');
    c.history(false);await tick();await EN.save(parent.doc,{root:editor().edRef.current});const headerUndo=await EN.open(parent.doc);check(headerUndo.header.includes('Table before')&&!headerUndo.header.includes('Table after'),'saved header-table undo failed: '+JSON.stringify({html:headerUndo.header,model:parent.doc.header,hist:parent.hist[parent.doc.id]}));
    return ['Bound Word controls: XML-backed mirrored fields, real edit dialog, no-op save and saved undo/redo', 'Mixed paper sizes, orientation, margins, headers, native reopen and per-page print geometry', 'Native Word controls: insert, edit, copy, delete, save/reopen and saved undo', 'Repeated table headers, visual line numbers and print; English automatic hyphenation', 'Mail merge recipient wizard, editable labels and custom-size envelopes: native save/reopen and pagination', 'Exact selected-range bookmarks: cross-paragraph endpoints, reference target, native save/reopen, deletion and saved undo', 'Native chart and SmartArt dialogs: multi-series data, node text, save/reopen and saved undo', 'Multi-page table rows preserve content and table structure', 'Header table cell editing, rich text, native save/reopen and saved undo', 'Wildcard and format search, capture replacement and saved undo', 'Word table formulas: bookmarked cross-table dependencies, whole rows, insertion, field save/reopen, saved undo/redo and stable no-op', 'Native chart preview/no-op save, theme/style set/page borders, edit protection and document comparison', 'MathType WMF: vector and Symbol glyph SVG preview, original object XML preserved after editing', 'WordEditor: cross-format find/replace, whole words, case sensitivity, undo and redo after saving', 'WordEditor table commands: insert row, save, undo, save/reopen and redo', 'Word shortcuts: justify, line spacing, subscript without zoom, Heading 2', 'Independent Chinese/Latin font controls: formatting, font sources, save/reopen and stable no-op', 'Section numbering, independent odd/even headers, paginated print and stable save', 'Native numbering start values and heading numbering; individual text and formatting revision resolution'];
  });
  for (const result of editorResults) console.log('PASS', result);
  console.log('PASS incremental page boundaries match full layout', JSON.stringify(await page.evaluate(() => window.__wordPaginationPerf)));
} finally {
  await browser?.close(); server?.close(); engine.kill(); rmSync(dir, { recursive: true, force: true });
}
