import {workbookRecords,mergeFields,addressText,mailLayout} from './word-mailmerge.js';

export function editMailMerge(source,template) {return new Promise(resolve=>{
  const dialog=document.createElement('dialog');dialog.dataset.mailMergeEditor='1';dialog.style.cssText='width:880px;max-width:94vw;max-height:90vh;overflow:auto;padding:24px;border:1px solid #bbb;border-radius:16px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
  const title=document.createElement('h3');title.textContent='邮件合并';dialog.append(title);
  const row=()=>{const r=document.createElement('div');r.style.cssText='display:flex;gap:12px;flex-wrap:wrap;margin:12px 0';dialog.append(r);return r;};
  const field=(parent,key,label,value,choices,multiline=false)=>{const wrap=document.createElement('label');wrap.style.cssText='display:grid;gap:6px;flex:1;min-width:160px';wrap.append(document.createTextNode(label));const input=document.createElement(choices?'select':multiline?'textarea':'input');input.name=key;input.style.cssText='padding:7px;border:1px solid #bbb;border-radius:6px;background:transparent;color:inherit;font:inherit;min-width:0';if(choices)for(const [v,label]of choices){const o=document.createElement('option');o.value=v;o.textContent=label;input.append(o);}input.value=value??'';if(multiline)input.rows=4;wrap.append(input);parent.append(wrap);return input;};
  const controls=row(),mode=field(controls,'mode','生成内容','letters',[['letters','当前文档：每人一份'],['labels','邮寄标签'],['envelopes','信封']]);let sheet,range;
  if(source.workbook){sheet=field(controls,'sheet','工作表','0',source.workbook.sheets.map((s,i)=>[String(i),s.name]));range=field(controls,'range','数据区域（空白为已用区域）','');}
  const filterRow=row(),filterColumn=field(filterRow,'filterColumn','筛选字段','',[['','全部记录']]),filter=field(filterRow,'filter','包含文字',''),sort=field(filterRow,'sort','排序字段','',[['','数据源顺序']]);
  const selected=field(filterRow,'selected','记录编号（留空为全部，例如 1-10,15）','');
  const mapping=row(),fields=mergeFields([template.html,...['header','footer','firstHeader','firstFooter','evenHeader','evenFooter'].map(k=>template[k]||'')].join('')),mapped=new Map();
  const addressRow=row(),address=field(addressRow,'address','收件地址（用 {字段名} 引用）','',null,true),sender=field(addressRow,'sender','寄件人','',null,true);
  const layoutRow=row(),cols=field(layoutRow,'cols','标签列数',3),rows=field(layoutRow,'rows','标签行数',7),labelWidth=field(layoutRow,'labelWidth','标签宽度（厘米）',6.35),labelHeight=field(layoutRow,'labelHeight','标签高度（厘米）',3.81),margin=field(layoutRow,'margin','标签页边距（厘米）',.7);
  const envelopeRow=row(),width=field(envelopeRow,'width','信封宽度（厘米）',22),height=field(envelopeRow,'height','信封高度（厘米）',11);
  const error=document.createElement('p');error.setAttribute('role','alert');error.style.color='#b3261e';const info=document.createElement('p'),preview=document.createElement('div');preview.style.cssText='max-height:260px;overflow:auto;border:1px solid #ddd;padding:10px';dialog.append(error,info,preview);
  let records=[],columns=[],chosen=[];
  const options=()=>({mode:mode.value,address:address.value,sender:sender.value,cols:+cols.value,rows:+rows.value,labelWidth:+labelWidth.value,labelHeight:+labelHeight.value,margin:+margin.value,width:+width.value,height:+height.value});
  function update(){try{
    let list=records.map((record,i)=>({record,i:i+1}));if(filterColumn.value)list=list.filter(x=>String(x.record[filterColumn.value]??'').toLocaleLowerCase().includes(filter.value.toLocaleLowerCase()));
    if(selected.value.trim()){const ids=new Set();for(const chunk of selected.value.split(/[,，]/)){const m=/^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(chunk);if(!m||+m[1]<1||+(m[2]||m[1])>records.length||+(m[2]||m[1])<+m[1])throw Error('收件人编号无效');for(let i=+m[1];i<=+(m[2]||m[1]);i++)ids.add(i);}list=list.filter(x=>ids.has(x.i));}
    if(sort.value)list.sort((a,b)=>String(a.record[sort.value]??'').localeCompare(String(b.record[sort.value]??''),undefined,{numeric:true}));
    chosen=list.map(x=>({...x.record,...Object.fromEntries([...mapped].map(([field,input])=>[field,x.record[input.value]??'']))}));
    mapping.hidden=mode.value!=='letters';addressRow.hidden=mode.value==='letters';layoutRow.hidden=mode.value!=='labels';envelopeRow.hidden=mode.value!=='envelopes';sender.parentElement.hidden=mode.value!=='envelopes';
    if(mode.value!=='letters'&&chosen.length)mailLayout(chosen.slice(0,1),options());
    preview.replaceChildren();if(mode.value==='letters'){const table=document.createElement('table');table.style.cssText='border-collapse:collapse;white-space:nowrap';[columns,...chosen.slice(0,20).map(r=>columns.map(c=>r[c]))].forEach((values,i)=>{const tr=document.createElement('tr');for(const value of values){const cell=document.createElement(i?'td':'th');cell.style.cssText='padding:5px 10px;border:1px solid #ddd';cell.textContent=value??'';tr.append(cell);}table.append(tr);});preview.append(table);}else{const p=document.createElement('pre');p.style.cssText='white-space:pre-wrap;font:inherit;line-height:1.8';p.textContent=chosen[0]?addressText(address.value,chosen[0]):'';preview.append(p);}
    info.textContent=`${records.length} 条记录，已选 ${chosen.length} 条；预览前 ${mode.value==='letters'?20:1} 条。`;error.textContent='';return true;
  }catch(e){chosen=[];error.textContent=e.message;return false;}}
  function load(){try{records=source.workbook?workbookRecords(source.workbook,+sheet.value,range.value.trim()):source.records;columns=[...new Set(records.flatMap(Object.keys))];
    for(const input of [filterColumn,sort]){const old=input.value;input.replaceChildren();for(const name of ['',...columns]){const o=document.createElement('option');o.value=name;o.textContent=name||'全部 / 原始顺序';input.append(o);}input.value=columns.includes(old)?old:'';}
    mapping.replaceChildren();mapped.clear();for(const name of fields)mapped.set(name,field(mapping,'map-'+name,'文档字段：'+name,columns.includes(name)?name:'',[['','选择对应字段'],...columns.map(c=>[c,c])]));
    for(const input of mapped.values())input.onchange=update;if(!address.value)address.value=columns.slice(0,4).map(c=>'{'+c+'}').join('\n');update();
  }catch(e){records=[];chosen=[];error.textContent=e.message;}}
  for(const input of [mode,filterColumn,filter,sort,selected,address,sender,cols,rows,labelWidth,labelHeight,margin,width,height])input.addEventListener('input',update);
  if(sheet)sheet.onchange=load;if(range)range.onchange=load;
  const actions=row();actions.style.justifyContent='flex-end';let ended=false;const done=v=>{if(ended)return;ended=true;dialog.close();dialog.remove();resolve(v);};for(const [text,fn]of [['取消',()=>done(null)],['生成文档',()=>{if(mode.value==='letters'&&[...mapped.values()].some(input=>!input.value)){error.textContent='请为每个文档字段选择对应的数据字段';return;}if(update()&&chosen.length)done({records:chosen,options:options()});else if(!error.textContent)error.textContent='请选择至少一条记录';}]]){const button=document.createElement('button');button.textContent=text;button.onclick=fn;actions.append(button);}
  dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});document.body.append(dialog);load();dialog.showModal();
});}
