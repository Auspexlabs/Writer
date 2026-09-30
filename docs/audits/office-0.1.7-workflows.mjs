// Audit only: actual editor/render functions and native temporary files. No user documents are modified.
// Requires the built CLI and python3. Run: node docs/audits/office-0.1.7-workflows.mjs
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';
import * as E from '../../ui/sheet-engine.js';
import * as K from '../../ui/office-io.js';
import { sheetPrint } from '../../ui/sheet-print.js';
const root = fileURLToPath(new URL('../../', import.meta.url));
const cli = process.env.WRITER_AUDIT_CLI || join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
const dir = mkdtempSync(join(tmpdir(), 'writer-office-audit-'));
const run = (...args) => { const out = execFileSync(cli.endsWith('.dll') ? 'dotnet' : cli, cli.endsWith('.dll') ? [cli, ...args] : args, { encoding: 'utf8' }); return /^[{[]/.test(out.trim()) ? JSON.parse(out) : out.trim(); };
const py = (code, ...args) => execFileSync('python3', ['-c', code, ...args], { encoding: 'utf8' }).trim();
const patchZip = (path, code) => py(`import sys,zipfile,xml.etree.ElementTree as ET
path=sys.argv[1]
with zipfile.ZipFile(path) as z: data={n:z.read(n) for n in z.namelist()}
${code}
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in data.items(): z.writestr(n,b)
`, path);
const output = { sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(), scope: 'Native file and Node render/Canvas command probes; not desktop pixel or Microsoft Office acceptance', spreadsheetPrint: [], dateSystems: [], wordHeader: {}, slidePng: {} };
Object.assign(output,{workingTreeModified:!!execFileSync('git',['status','--porcelain'],{cwd:root,encoding:'utf8'}).trim(),generatedAt:new Date().toISOString(),cliVersion:run('--version')});
try {
  const script = readFileSync(new URL('../../ui/SheetEditor.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { React: { createRef: () => ({ current: null }), createElement: (...args) => args }, DCLogic: class {}, structuredClone, document: { documentElement: { dataset: {} } }, $t: s => s, $lang: () => 'zh' };
  vm.runInNewContext(script + '\nglobalThis.Editor = Component;', ctx);
  const editor = new ctx.Editor(); editor.E = E;
  for (const [name, style] of [['borderThin', { bd: 'thin' }], ['borderDouble', { bd: 'double' }], ['perSideBorderControl', { bd: { bottom: 'double' } }], ['strike', { st: true }], ['rotatedText', { rotate: 45 }], ['numberFormatColor', { code: '[Red]0.00' }], ['workbookDefaultFont', {}]]) {
    const doc = { active: 0, ...(name === 'workbookDefaultFont' ? { font: 'Georgia', fs: 18 } : {}), sheets: [{ name: 'S', cells: { A1: { v: '123', s: style } } }] };
    editor.props = { doc };
    const rendered = Array.from(editor.renderVals().cells).find(c => c.gr === '2' && c.gc === '2');
    const html = sheetPrint(doc, E).body;
    output.spreadsheetPrint.push({ name, style, editor: Object.fromEntries(Object.entries(rendered).filter(([k,v]) => typeof v === 'string' && /border|^b[tlrb]$|^td$|^tr$|^ff$|^fs$|dec|rot|transform|color|text/i.test(k))), printedCellStyle: /<td[^>]*style="([^"]*)"/.exec(html)?.[1] });
  }
  for (const date1904 of [false, true]) {
    const path = join(dir, `dates-${date1904}.xlsx`);
    run('create', path);
    run('set', path, '/sheet[1]/range[A1:B1]', '--prop', 'format=yyyy-mm-dd');
    patchZip(path, `ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
wb=ET.fromstring(data['xl/workbook.xml']); prop=wb.find('{'+ns+'}workbookPr')
if prop is None:
 prop=ET.Element('{'+ns+'}workbookPr'); wb.insert(0,prop)
prop.set('date1904','${date1904 ? 1 : 0}'); data['xl/workbook.xml']=ET.tostring(wb)
ws=ET.fromstring(data['xl/worksheets/sheet1.xml']); cell=ws.find('.//{'+ns+'}c[@r="A1"]')
cell.attrib.pop('t',None)
v=cell.find('{'+ns+'}v')
if v is None: v=ET.SubElement(cell,'{'+ns+'}v')
v.text='${date1904 ? 43894 : 45356}'; data['xl/worksheets/sheet1.xml']=ET.tostring(ws)`);
    const before = run('get', path, '/sheet[1]/cell[A1]').props;
    run('set', path, '/sheet[1]/cell[B1]', '--prop', 'formula=A1+1');
    const after = run('get', path, '/sheet[1]/cell[B1]').props;
    output.dateSystems.push({ date1904, input: before.value, formula: 'A1+1', expected: '2024-03-06', savedAndReopened: after.value });
  }
  const word = join(dir, 'header.docx');
  run('create', word); run('set', word, '/', '--prop', 'header=Visible paragraph');
  patchZip(word, `ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
name=next(n for n in data if n.startswith('word/header') and n.endswith('.xml'))
h=ET.fromstring(data[name]); h.append(ET.fromstring('<w:tbl xmlns:w="'+ns+'"><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>HEADER_TABLE_SENTINEL</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'))
data[name]=ET.tostring(h)`);
  const header = run('get', word, '/').props.header;
  run('set', word, '/', '--prop', 'header=Changed paragraph');
  const retained = py(`import zipfile,sys
with zipfile.ZipFile(sys.argv[1]) as z: print(any(b'HEADER_TABLE_SENTINEL' in z.read(n) for n in z.namelist() if n.startswith('word/header')))
`, word);
  output.wordHeader = { editorHtml: header, tableVisibleToEditor: header?.includes('HEADER_TABLE_SENTINEL') || false, tableRetainedAfterHeaderTextEdit: retained === 'True' };
  // Record commands sent to Canvas by the production PNG renderer. This is not raster pixel testing.
  globalThis.Path2D = class { roundRect() {} closePath() {} ellipse() {} lineTo() {} moveTo() {} };
  const draws = [], graphics = new Proxy({ measureText: text => ({ width: text.length * 10 }), fillText(text) { draws.push({ text, font: this.font, color: this.fillStyle }); }, drawImage(...args) { draws.push({ imageArgumentCount: args.length }); } }, { get(target, key) { return key in target ? target[key] : () => {}; } });
  const text = { t: 'text', x: 0, y: 0, w: 1000, h: 200, fs: 32, html: '<p>Normal <span style="font-size:64px;color:#ff0000;font-weight:700">RED</span></p>' };
  K.paintSlide(graphics, { objs: [text] }, K.THEMES.paper, 900, new Map());
  output.slidePng.richText = { inputHtml: text.html, canvasTextDraws: draws.splice(0) };
  K.paintSlide(graphics, { bgImage: 'background-probe', objs: [] }, K.THEMES.paper, 900, new Map([['background-probe', {}]]));
  output.slidePng.backgroundImage = { input: 'background-probe', canvasImageDraws: draws.splice(0) };
  const picture = { t: 'image', src: 'picture-probe', x: 0, y: 0, w: 200, h: 100, look: { crop: '25,0,25,0', grayscale: 'true' } };
  K.paintSlide(graphics, { objs: [picture] }, K.THEMES.paper, 900, new Map([['picture-probe', {}]]));
  const pictureView = K.objView(picture, K.THEMES.paper);
  output.slidePng.croppedPicture = { inputLook: picture.look, editorPictureStyle: pictureView.picImage, canvasImageDraws: draws.splice(0), canvasFilter: typeof graphics.filter === 'function' ? 'unset' : graphics.filter };
  output.slidePng.scope = 'Actual paintSlide invocation; fake Canvas records commands without drawing pixels';
  output.commonFunctions = Object.fromEntries(['SUM','SUMIF','SUMIFS','SUMPRODUCT','COUNTIF','COUNTIFS','AVERAGEIFS','XLOOKUP','XMATCH','VLOOKUP','INDEX','MATCH','IFERROR','IFS','SWITCH','TEXTJOIN','FILTER','SORT','UNIQUE','LET','SEQUENCE','XIRR','XNPV','NETWORKDAYS','WORKDAY','DATEDIF','SUBTOTAL','AGGREGATE','LAMBDA','GETPIVOTDATA'].map(name => [name, E.FUNCS.includes(name)]));
  console.log(JSON.stringify(output, null, 2));
} finally { rmSync(dir, { recursive: true, force: true }); }
