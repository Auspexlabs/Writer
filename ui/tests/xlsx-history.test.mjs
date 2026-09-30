// The real editor → save bridge → native XLSX validation/writer → reopened file.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import * as E from '../sheet-engine.js';
import { cloneWorkbook } from '../sheet-model.js';
globalThis.window = globalThis;
const EN = await import('../engine.js');
const cli = fileURLToPath(new URL('../../src/Writer.Cli/bin/Debug/net10.0/writer.dll', import.meta.url));
const skip = !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';
const script = readFileSync(new URL('../SheetEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
const ctx = {
  React: { createRef: () => ({ current: null }) }, structuredClone,
  DCLogic: class { setState(value, cb) { Object.assign(this.state, typeof value === 'function' ? value(this.state, this.props) : value); cb?.(); } forceUpdate() {} },
  $t: s => s, $lang: () => 'zh',
};
vm.runInNewContext(script + '\nglobalThis.Editor = Component;', ctx);
let server, dir;
const fetchOriginal = globalThis.fetch;
before(async () => {
  if (skip) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-xlsx-history-'));
  server = spawn('dotnet', [cli, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  const base = await new Promise((resolve, reject) => {
    let stderr = '';
    server.stderr.on('data', chunk => { stderr += chunk; const match = /"url"\s*:\s*"([^"]+)"/.exec(stderr); if (match) resolve(match[1]); });
    server.on('error', reject);
    server.on('exit', code => reject(new Error(`Engine exited ${code}: ${stderr}`)));
  });
  globalThis.fetch = (url, options) => fetchOriginal(new URL(url, base), options);
});
after(() => { globalThis.fetch = fetchOriginal; server?.kill(); if (dir) rmSync(dir, { recursive: true, force: true }); });
const plain = value => JSON.parse(JSON.stringify(value));
async function workbook(name) {
  const path = join(dir, name + '.xlsx');
  await EN.run(['create', path]);
  const source = { A1: { v: '123' }, B1: { v: '文字', s: { b: true, color: '#FF0000' } }, A2: { v: '=A1*2' }, B2: { v: '保留格式', s: { fill: '#FFFF00', bd: 'thin' } }, D9: { v: '区域外' } };
  const opened = await EN.open({ id: name, path, type: 'xlsx' });
  let doc = E.editWorkbook(opened, sheet => { sheet.cells = source; });
  await EN.save(doc);
  const editor = new ctx.Editor(); editor.E = E;
  editor.props = { get doc() { return doc; }, onChange(next) { doc = next; } };
  const select = () => editor.setState({ anc: { r: 0, c: 0 }, sel: { r: 1, c: 1 } });
  const save = async () => { await EN.save(doc); return EN.open({ id: name, path, type: 'xlsx' }); };
  return { editor, select, save, get doc() { return doc; }, restore(snapshot) { doc = EN.historyModel(snapshot, doc); } };
}

test('workbook structure protection blocks every sheet command, persists passwords and supports saved undo', {skip}, async()=>{
 const x=await workbook('book-protect');x.editor.EN=EN;
 x.editor.addSheet();const before=cloneWorkbook(x.doc),messages=[];x.editor.toast=t=>messages.push(t);
 x.editor.prompt=(title,fields,ok)=>{x.dialog={title,fields,ok};};
 x.editor.protectWorkbook();assert.equal(x.dialog.fields[0].type,'password');
 await x.dialog.ok({password:'secret',confirm:'secret'});
 assert.equal(x.doc.workbookProtection.structure,true);
 const protectedDoc=cloneWorkbook(x.doc);
 for(const action of [()=>x.editor.addSheet(),()=>x.editor.copySheet(0),()=>x.editor.renameSheet(0),()=>x.editor.moveSheet(0,1),()=>x.editor.hideSheet(0),()=>x.editor.deleteSheet(0)])action();
 assert.deepEqual(x.doc.sheets.map(s=>s.name),before.sheets.map(s=>s.name));assert.equal(messages.length,6);
 assert.equal(x.editor.commit((sh,d)=>d.sheets.push({name:'bypass',cells:{}})),false);
 x.editor.commit(sh=>{sh.cells.A3={v:'allowed'};});assert.equal(x.doc.sheets[1].cells.A3.v,'allowed');
 let reopened=await x.save();assert.equal(reopened.workbookProtection.structure,true);assert.ok(reopened.workbookProtection.verifier.hashValue);
 x.editor.protectWorkbook();await x.dialog.ok({password:'wrong'});assert.equal(x.doc.workbookProtection.structure,true);
 x.editor.protectWorkbook();await x.dialog.ok({password:'secret'});reopened=await x.save();assert.equal(reopened.workbookProtection.structure,false);
 x.restore(protectedDoc);reopened=await x.save();assert.equal(reopened.workbookProtection.structure,true);
 x.restore(before);reopened=await x.save();assert.equal(reopened.workbookProtection.structure,false,'undo before protection clears native lock');
});

test('protected pivot permissions allow only pivot output updates',()=>{
 const before={sheets:[{name:'Data',protected:true,protection:{pivotTables:true},pivots:[{name:'p',range:'B2:C3'}],cells:{A1:{v:'keep'},B2:{v:1}}}]};
 const after=E.editWorkbook(before,sh=>{sh.cells.B2={v:2,s:{b:true}};sh.pivots[0].refresh=1;});
 assert.equal(E.protectionViolation(before,after,'pivotTables'),null);
 assert.ok(E.protectionViolation(before,after));
 assert.ok(E.protectionViolation(before,E.editWorkbook(after,sh=>{sh.cells.A1={v:'forbidden'};}),'pivotTables'));
 const blocked=E.editWorkbook(before,sh=>{sh.protection.pivotTables=false;});assert.ok(E.protectionViolation(blocked,after,'pivotTables'));
});

test('grouped sheets apply input and formatting atomically, preserve individual contents and saved undo', {skip}, async()=>{
 const x=await workbook('grouped'),e=x.editor;e.addSheet();e.selectSheet(0,{ctrlKey:true});
 assert.deepEqual(plain(e.groupedSheets()),[0,1]);e.setState({anc:{r:3,c:0},sel:{r:3,c:0}});
 const prior=cloneWorkbook(x.doc);e.startEdit('17','type');e.commitEdit(0,0);
 assert.deepEqual(x.doc.sheets.map(s=>s.cells.A4.v),['17','17']);
 e.style({b:true});assert.ok(x.doc.sheets.every(s=>s.cells.A4.s.b));
 let reopened=await x.save();assert.ok(reopened.sheets.every(s=>s.tabSelected));assert.ok(reopened.sheets.every(s=>s.cells.A4.v==='17'));
 const saved=cloneWorkbook(x.doc);x.restore(prior);reopened=await x.save();assert.ok(reopened.sheets.every(s=>!s.cells.A4?.v));
 x.restore(saved);e.selectSheet(1);e.commit(sh=>{sh.protected=true;});e.selectSheet(0,{ctrlKey:true});
 const beforeBlocked=x.doc;e.setState({anc:{r:3,c:0},sel:{r:3,c:0}});e.style({i:true});assert.equal(x.doc,beforeBlocked,'protected second sheet must refuse the entire group edit');
 x.restore(saved);e.selectSheet(1);e.commit(sh=>{sh.dv=[{range:'A4',type:'whole',operator:'greaterThan',value:'100'}];});e.selectSheet(0,{ctrlKey:true});
 const beforeValidation=x.doc;e.startEdit('20','type');e.commitEdit(0,0);assert.equal(x.doc,beforeValidation,'validation on another grouped sheet must be checked');
 x.restore(saved);e.setState({edit:null,anc:{r:0,c:0},sel:{r:0,c:0}});e.clear('v');assert.equal(x.doc.sheets[0].cells.A1?.v??'','');assert.equal(x.doc.sheets[1].cells.A1?.v??'','');
 e.selectSheet(1);e.startEdit('only second','type');e.commitEdit(0,0);assert.equal(x.doc.sheets[0].cells.A1?.v??'','');assert.equal(x.doc.sheets[1].cells.A1.v,'only second');
});

for (const kind of ['center', 'across']) test(`native XLSX: merge ${kind}, undo, redo, unmerge and undo survive saves`, { skip }, async () => {
  const x = await workbook(kind), original = cloneWorkbook(x.doc), originalCells = plain(original.sheets[0].cells);
  x.select(); x.editor.mergeKind(kind);
  const merged = cloneWorkbook(x.doc);
  let reopened = await x.save();
  assert.equal(reopened.sheets[0].merges.length, kind === 'across' ? 2 : 1);
  assert.equal(reopened.sheets[0].cells.B1, undefined);
  if (kind === 'center') assert.equal(reopened.sheets[0].cells.A1.s.align, 'center');
  assert.equal(reopened.sheets[0].cells.D9.v, '区域外');
  x.restore(original); reopened = await x.save();
  assert.deepEqual(reopened.sheets[0].merges, []);
  assert.deepEqual(plain(reopened.sheets[0].cells), originalCells, 'undo restores text, formula, formatting and default alignment');
  x.restore(merged); reopened = await x.save();
  assert.equal(reopened.sheets[0].merges.length, kind === 'across' ? 2 : 1);
  x.select(); x.editor.mergeKind('un'); reopened = await x.save();
  assert.deepEqual(reopened.sheets[0].merges, []);
  x.restore(merged); reopened = await x.save();
  assert.equal(reopened.sheets[0].merges.length, kind === 'across' ? 2 : 1);
  assert.equal(await EN.save(x.doc), 0, 'an unchanged save emits no mutations');
});

test('native XLSX: undo all cell format defaults, including a batched range, keeps values', { skip }, async () => {
  const x = await workbook('defaults');
  x.editor.commit(sh => { for (let i = 1; i <= 8; i++) sh.cells['A' + i] = { v: String(i) }; });
  await x.save();
  const original = cloneWorkbook(x.doc);
  x.editor.setState({ anc: { r: 0, c: 0 }, sel: { r: 7, c: 0 } });
  x.editor.style({ align: 'center', va: 'top', b: true, i: true, u: true, st: true, wrap: true, indent: 2, rotate: 45, fs: 14, font: 'Georgia', fill: '#FF0000', color: '#FFFFFF', bd: 'thin', bdc: '#00FF00', fmt: 'pct', dec: 2, note: 'Note', link: 'https://example.com' });
  const formatted = cloneWorkbook(x.doc); await x.save();
  x.restore(original);
  const reopened = await x.save();
  for (let i = 1; i <= 8; i++) {
    const cell = reopened.sheets[0].cells['A' + i];
    assert.equal(cell.v, String(i));
    for (const prop of ['align', 'va', 'b', 'i', 'u', 'st', 'wrap', 'indent', 'rotate', 'fill', 'color', 'bd', 'bdc', 'fmt', 'note', 'link']) assert.ok(!cell.s?.[prop], `A${i}.${prop} cleared`);
  }
  assert.deepEqual(plain(reopened.sheets[0].cells.B1), plain(original.sheets[0].cells.B1));
  x.restore(formatted); const redone = await x.save();
  assert.equal(redone.sheets[0].cells.A8.s.align, 'center');
  assert.equal(redone.sheets[0].cells.A8.s.note, 'Note');
});
