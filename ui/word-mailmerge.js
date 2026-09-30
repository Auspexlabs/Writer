import { readTSV } from './sheet-clipboard.js';
import { Calc, usedRange, A, parseA, fmt } from './sheet-engine.js';

export function recordsFromRows(rows) {
  const headers=(rows[0]||[]).map(x=>String(x??'').trim());
  if(!headers.length||new Set(headers).size!==headers.length||headers.some(h=>!h))throw Error('第一行需要非空且不重复的字段名。');
  const records=rows.slice(1).filter(r=>r.some(v=>v!=null&&v!=='')).map(row=>Object.fromEntries(headers.map((h,i)=>[h,row[i]??''])));
  if(!records.length)throw Error('数据源没有收件人记录。');return records;
}

export function workbookRecords(doc, si=0, range='') {
  const sh=doc.sheets[si];if(!sh)throw Error('请选择数据工作表');let box=usedRange(sh);
  if(range){const [a,b=a]=range.split(':').map(parseA);if(!a||!b||a.r>b.r||a.c>b.c)throw Error('数据区域地址无效');box={r1:a.r,c1:a.c,r2:b.r,c2:b.c};}
  if(!box||(box.r2-box.r1+1)*(box.c2-box.c1+1)>500000)throw Error('请选择最多 50 万个单元格的数据区域');
  const calc=new Calc(doc),rows=[];
  for(let r=box.r1;r<=box.r2;r++){const row=[];for(let c=box.c1;c<=box.c2;c++){const value=calc.value(si,r,c);if(value?.err)throw Error(A(r,c)+': '+value.err);row.push(fmt(value,{...sh.cells[A(r,c)]?.s,date1904:!!doc.date1904}));}rows.push(row);}
  return recordsFromRows(rows);
}

export function mergeFields(html) {
  const names=new Set();for(const m of String(html||'').matchAll(/data-field=["']MERGEFIELD\s+([\p{L}_][\p{L}\p{N}_]*)/gu))names.add(m[1]);return [...names];
}
export const addressText=(pattern,record)=>String(pattern).replace(/\{([^{}]+)\}/g,(_,field)=>String(record[field]??''));
const escapeHtml=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

/** Labels and envelopes are ordinary editable DOCX pages; no browser-only print artifact. */
export function mailLayout(records,options) {
  const address=r=>escapeHtml(addressText(options.address,r)).replace(/\n/g,'<br>');
  if(options.mode==='envelopes'){if(![+options.width||22,+options.height||11].every(v=>v>=9&&v<=55.88))throw Error('信封宽高须为 9 至 55.88 厘米');return {page:{size:`${options.width||22}cm x ${options.height||11}cm`,orient:'portrait',margin:'1cm 1cm 1cm 1cm',cols:1},html:records.map(r=>`<p data-w-spaceafter="0pt">${escapeHtml(options.sender||'').replace(/\n/g,'<br>')}</p><p data-w-spacebefore="60pt" data-w-indentleft="7cm" style="margin-top:60pt;margin-left:7cm">${address(r)}</p>`).join('<hr data-pb="1">')};}
  const cols=+options.cols||3,rows=+options.rows||7,width=+options.labelWidth||6.35,height=+options.labelHeight||3.81,margin=options.margin==null?.7:+options.margin;
  if(!Number.isInteger(cols)||!Number.isInteger(rows)||cols<1||rows<1||cols>20||rows>50||!Number.isFinite(margin)||margin<0||width<=0||height<=0||width*cols+2*margin>21||height*rows+2*margin>29.7)throw Error('标签尺寸、数量和页边距超出 A4 纸张');
  const pages=[];for(let first=0;first<records.length;first+=cols*rows){let html=`<table data-w-borders="none" data-w-style="none" data-w-width="${cols*width}cm" data-w-widths="${escapeHtml(JSON.stringify(Array(cols).fill(width+'cm')))}" style="border-collapse:collapse;table-layout:fixed;width:${cols*width}cm"><tbody>`;
    for(let row=0;row<rows;row++){html+=`<tr data-w-height="${height}cm" style="height:${height}cm">`;for(let col=0;col<cols;col++){const record=records[first+row*cols+col];html+=`<td data-w-valign="middle" style="vertical-align:middle;border:0"><p data-w-spaceafter="0pt" style="margin:0">${record?address(record):'<br>'}</p></td>`;}html+='</tr>';}
    pages.push(html+'</tbody></table>');}
  return {page:{size:'A4',orient:'portrait',margin:`${margin}cm ${margin}cm ${margin}cm ${margin}cm`,cols:1},html:pages.join('<hr data-pb="1">')};
}

export function mergeRecords(text, name = '') {
  text = text.replace(/^\uFEFF/, '');
  let records;
  if (/\.json$/i.test(name) || /^\s*\[/.test(text)) records = JSON.parse(text);
  else {
    const rows = readTSV(text, /\.tsv$/i.test(name) || text.split(/\r?\n/, 1)[0].includes('\t') ? '\t' : ',');
    records = recordsFromRows(rows);
  }
  if (!Array.isArray(records) || !records.length || records.some(r => !r || typeof r !== 'object' || Array.isArray(r) || Object.values(r).some(v => v != null && typeof v === 'object'))) throw new Error('请选择包含数据记录的 CSV、TSV 或 JSON 数组。');
  return records;
}

/** Each recipient starts from the same saved template; generated files are separate documents. */
export async function mergeDocuments(EN, template, records, progress) {
  const result = [];
  for (let i = 0; i < records.length; i++) {
    const doc = await EN.copy({ ...template, title: template.title + '-' + (i + 1) });
    try { await EN.run(['set', doc.path, '/', '--prop', 'mergeData=' + JSON.stringify(records[i])]); }
    catch (e) { await EN.discard(doc); e.message += `（已完成 ${result.length} 份）`; throw e; }
    result.push(doc); progress?.(i + 1, records.length);
  }
  return result;
}
