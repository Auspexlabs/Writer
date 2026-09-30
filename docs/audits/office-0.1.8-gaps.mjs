// Targeted audit of the current 0.1.8 implementation. Temporary Office files only; no product mutations.
import {readFileSync,mkdtempSync,rmSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import vm from 'node:vm';
import * as E from '../../ui/sheet-engine.js';
import * as K from '../../ui/office-io.js';
import {richParagraphs} from '../../ui/slide-canvas.js';
const root=resolve(new URL('../..',import.meta.url).pathname),dir=mkdtempSync(join(tmpdir(),'writer-018-gaps-'));
const cli=process.env.WRITER_AUDIT_CLI||join(root,'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
const run=(...args)=>{const out=execFileSync(cli.endsWith('.dll')?'dotnet':cli,cli.endsWith('.dll')?[cli,...args]:args,{encoding:'utf8'}).trim();return /^[{[]/.test(out)?JSON.parse(out):out;};
const py=(source,...args)=>execFileSync('python3',['-c',source,...args],{encoding:'utf8'}).trim();
const output={version:run('--version'),engineCli:resolve(cli),baseCommit:execFileSync('git',['rev-parse','HEAD'],{cwd:root,encoding:'utf8'}).trim(),workingTreeModified:true,generatedAt:new Date().toISOString(),scope:'Actual Writer formula/render logic and native save/reopen. Expected results use documented semantics; Microsoft Office desktop was not run.',excel:[],conditionalFormatting:[],word:{},ppt:{}};
try{
 for(const [formula,expected] of [['=COUNTA(A1:A3)',3],['=IFERROR(MAP({1,0,2},LAMBDA(x,1/x)),99)',[[1,99,.5]]],['=SUMIFS(A1:A1048576,A1:A1048576,">0")',3],['=WRAPROWS({1,2},5,0)',[[1,2]]]]){
  const calc=new E.Calc({sheets:[{name:'S',cells:{A1:{v:1},A2:{v:'=""'},A3:{v:2},D1:{v:formula}}}]}),head=calc.value(0,0,3),actual=calc.arrays.get('0:0:3')||head;
  output.excel.push({formula,expected,actual,matches:JSON.stringify(expected)===JSON.stringify(actual)});
 }
 const book=join(dir,'formula.xlsx');run('create',book);for(const [ref,prop,value] of [['A1','value','1'],['A2','formula','""'],['A3','value','2'],['D1','formula','COUNTA(A1:A3)'],['D2','formula','IFERROR(MAP({1,0,2},LAMBDA(x,1/x)),99)']])run('set',book,`/sheet[1]/cell[${ref}]`,'--prop',`${prop}=${value}`);
 output.nativeFormulaCache=Object.fromEntries(['D1','D2','E2','F2'].map(ref=>[ref,run('get',book,`/sheet[1]/cell[${ref}]`).props.value]));
 const cf=(name,cells,rules)=>{const sh={name:'S',cells,cf:rules},calc=new E.Calc({sheets:[sh]}),apply=E.conditionalFormats(sh,calc,0,E);return {name,values:[0,1,2].map(r=>calc.value(0,r,0)),backgrounds:[0,1,2].map(r=>apply(r,0,calc.value(0,r,0),{}).bg)};};
 output.conditionalFormatting.push(cf('cellIs threshold reference',{A1:{v:20},B1:{v:10}},[{type:'cellIs',range:'A1',operator:'greaterThan',value:'$B$1',fill:'#123456'}]));
 const scale=[{type:'colorScale',range:'A1:A3',colors:['#ff0000','#ffff00','#00ff00']}];
 output.conditionalFormatting.push(cf('dynamic spill scale',{A1:{v:'=SEQUENCE(3)'}},scale));
 output.conditionalFormatting.push(cf('stored values scale control',{A1:{v:1},A2:{v:2},A3:{v:3}},scale));
 const html=readFileSync(join(root,'ui/WordEditor.dc.html'),'utf8'),script=html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
 const ctx={React:{createRef:()=>({current:null})},DCLogic:class{},$t:s=>s,$lang:()=> 'zh',setTimeout,clearTimeout};
 vm.runInNewContext(script+'\nglobalThis.hfOut=hfOut;',ctx);
 const word=join(dir,'header.docx');run('create',word);run('set',word,'/','--prop','header=Before');
 py(`import zipfile,sys,xml.etree.ElementTree as ET\np=sys.argv[1]\nwith zipfile.ZipFile(p) as z:d={n:z.read(n) for n in z.namelist()}\nn=next(n for n in d if n.startswith('word/header') and n.endswith('.xml'));ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main';r=ET.fromstring(d[n]);r.append(ET.fromstring('<w:tbl xmlns:w="'+ns+'"><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>TABLE_SENTINEL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'));d[n]=ET.tostring(r)\nwith zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:\n for n,b in d.items():z.writestr(n,b)`,word);
 const header=run('get',word,'/').props.header,edited=ctx.hfOut(header.replace('Before','After'));run('set',word,'/','--prop','header='+edited);
 const reopened=run('get',word,'/').props.header;
 output.word.headerTable={before:header,serializedByEditor:edited,afterSave:reopened,sentinelBefore:(header.match(/TABLE_SENTINEL/g)||[]).length,sentinelAfter:(reopened.match(/TABLE_SENTINEL/g)||[]).length};
 output.ppt.numbering={input:'III. / IV.',actual:richParagraphs('<ol><li data-num-format="romanUcPeriod" data-start="3" data-marker="III. ">First</li><li data-marker="IV. ">Second</li></ol>',{fs:32}).map(p=>p.bullet)};
 globalThis.Path2D=class{roundRect(){}closePath(){}ellipse(){}lineTo(){}moveTo(){}};
 const calls=[],g=new Proxy({measureText:text=>({width:text.length*10}),rotate(n){calls.push({rotate:n});},fillText(text){calls.push({text,font:this.font,color:this.fillStyle});}},{get:(o,k)=>k in o?o[k]:(()=>{})});
 K.paintSlide(g,{objs:[K.txt({t:'table',x:0,y:0,w:500,h:100,fs:24,rows:[['Rich table']],cells:{'0:0':{html:'<p><span style="font-size:64px;color:#ff0000">Rich table</span></p>'}}})]},K.THEMES.paper,900,new Map());output.ppt.tablePng=calls.splice(0);
 const group=K.group([K.txt({html:'<p>Rotated group</p>',x:10,y:20,w:300,h:100})]);group.rot=90;K.paintSlide(g,{objs:[group]},K.THEMES.paper,900,new Map());output.ppt.rotatedGroupPng=calls.splice(0);
 console.log(JSON.stringify(output,null,2));
}finally{rmSync(dir,{recursive:true,force:true});}
