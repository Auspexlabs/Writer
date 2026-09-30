import {dateSerial} from './sheet-engine.js';

export function splitColumnText(value,options={}) {
  const text=String(value??'');
  if(options.mode==='fixed') {
    const cuts=String(options.cuts||'').split(/[,，\s]+/).filter(Boolean).map(Number);
    if(cuts.some((n,i)=>!Number.isInteger(n)||n<1||i>0&&n<=cuts[i-1]))throw Error('固定宽度分界位置须为递增的正整数');
    const chars=Array.from(text),out=[];let last=0;for(const at of cuts){out.push(chars.slice(last,at).join(''));last=at;}out.push(chars.slice(last).join(''));return out;
  }
  const delimiters=new Set(String(options.delimiters??',').replace(/\\t/g,'\t')),quote=options.quote===undefined?'"':options.quote;
  if(!delimiters.size)throw Error('请选择至少一个分隔字符');let quoted=false,valueOut='',after=false;const out=[];
  for(let i=0;i<text.length;i++){const c=text[i];if(quote&&c===quote){if(quoted&&text[i+1]===quote){valueOut+=quote;i++;}else if(quoted){quoted=false;after=true;}else if(!valueOut&&!after)quoted=true;else valueOut+=c;}
    else if(!quoted&&delimiters.has(c)){out.push(valueOut);valueOut='';after=false;if(options.consecutive)while(delimiters.has(text[i+1]))i++;}
    else {valueOut+=c;after=false;}}
  if(quoted)throw Error('文本限定符没有闭合');out.push(valueOut);return out;
}
export function columnValue(text,type='general',options={}) {
  const raw=options.trim===false?String(text):String(text).trim();if(type==='skip')return null;
  if(type==='text')return {v:raw,literal:true,s:{fmt:'text',code:'@'}};
  if(['ymd','mdy','dmy'].includes(type)&&raw){const parts=raw.split(/[\/.\-\s]+/).map(Number);if(parts.length!==3||parts.some(n=>!Number.isInteger(n)))throw Error('无法解析日期：'+raw);const order=type.split(''),part=Object.fromEntries(order.map((c,i)=>[c,parts[i]]));if(part.y<100)part.y+=part.y<30?2000:1900;
    const date=new Date(0);date.setUTCFullYear(part.y,part.m-1,part.d);if(date.getUTCFullYear()!==part.y||date.getUTCMonth()!==part.m-1||date.getUTCDate()!==part.d)throw Error('日期不存在：'+raw);
    const serial=dateSerial(`${part.y}-${String(part.m).padStart(2,'0')}-${String(part.d).padStart(2,'0')}`,options.date1904);if(typeof serial!=='number'||!Number.isFinite(serial))throw Error('日期超出支持范围：'+raw);return {v:serial,s:{fmt:'date',code:'yyyy-mm-dd'}};
  }
  const grouping=options.thousands??'',decimal=options.decimal??'.';if(grouping&&grouping===decimal)throw Error('小数分隔符与千位分隔符不能相同');
  const normalized=(grouping?raw.split(grouping).join(''):raw).replace(decimal,'.');
  if(/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(normalized)&&Number.isFinite(+normalized)&&normalized.replace(/[^0-9]/g,'').length<=15)return {v:+normalized};
  return {v:raw,...(raw.startsWith('=')?{literal:true}:{})};
}
export function textColumns(values,options={}) {
  const rows=values.map(v=>splitColumnText(v,options)),count=rows.reduce((n,r)=>Math.max(n,r.length),0);
  return rows.map(row=>Array.from({length:count},(_,i)=>columnValue(row[i]??'',options.types?.[i]||'general',options)).filter(c=>c!==null));
}

export function editTextColumns(values,initial={}) {return new Promise(resolve=>{
  const dialog=document.createElement('dialog');dialog.dataset.textColumnsEditor='1';dialog.style.cssText='width:800px;max-width:94vw;max-height:90vh;overflow:auto;padding:24px;border:1px solid #bbb;border-radius:16px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
  const title=document.createElement('h3');title.textContent='文本分列';dialog.append(title);const controls=document.createElement('div');controls.style.cssText='display:flex;gap:10px;flex-wrap:wrap';dialog.append(controls);const inputs={};
  const field=(key,label,value,choices)=>{const wrap=document.createElement('label');wrap.style.cssText='display:grid;gap:5px;min-width:145px;flex:1;margin-bottom:12px';wrap.append(document.createTextNode(label));const input=document.createElement(choices?'select':'input');input.name=key;input.style.cssText='padding:7px;border:1px solid #bbb;border-radius:6px;background:transparent;color:inherit;font:inherit;min-width:0';if(choices)for(const [v,label]of choices){const o=document.createElement('option');o.value=v;o.textContent=label;input.append(o);}input.value=value;wrap.append(input);controls.append(wrap);inputs[key]=input;};
  field('mode','分列方式','delimiter',[['delimiter','按分隔字符'],['fixed','固定宽度']]);field('delimiters','分隔字符（制表符写 \\t）',',');field('quote','文本限定符','"',[['"','双引号'],["'",'单引号'],['','无']]);field('consecutive','连续分隔符','no',[['no','保留空列'],['yes','视为一个']]);field('cuts','固定宽度分界（例如 4,10）','4,10');field('decimal','小数分隔符','.');field('thousands','千位分隔符','');field('target','目标起始单元格',initial.target||'A1');field('overwrite','已有数据','no',[['no','遇到已有数据时停止'],['yes','覆盖目标区域']]);
  const types=[],typeRow=document.createElement('div'),preview=document.createElement('div'),error=document.createElement('p');typeRow.style.cssText='display:flex;gap:10px;overflow:auto;margin:10px 0';preview.style.cssText='max-height:250px;overflow:auto';error.setAttribute('role','alert');error.style.color='#b3261e';dialog.append(typeRow,preview,error);
  const options=()=>({mode:inputs.mode.value,delimiters:inputs.delimiters.value,quote:inputs.quote.value,consecutive:inputs.consecutive.value==='yes',cuts:inputs.cuts.value,decimal:inputs.decimal.value,thousands:inputs.thousands.value,date1904:!!initial.date1904,types:[...types]});
  function update(rebuild=true){try{const opts=options(),source=values.slice(0,20),parts=source.map(v=>splitColumnText(v,opts)),n=Math.max(0,...parts.map(r=>r.length));if(rebuild){typeRow.replaceChildren();for(let i=0;i<n;i++){const wrap=document.createElement('label');wrap.textContent='第 '+(i+1)+' 列 ';const select=document.createElement('select');select.dataset.column=String(i);for(const [v,text]of [['general','常规'],['text','文本'],['ymd','日期 YMD'],['mdy','日期 MDY'],['dmy','日期 DMY'],['skip','不导入']]){const o=document.createElement('option');o.value=v;o.textContent=text;select.append(o);}select.value=types[i]||'general';select.onchange=()=>{types[i]=select.value;update(false);};wrap.append(select);typeRow.append(wrap);}}
    const result=textColumns(source,options());preview.replaceChildren();const table=document.createElement('table');table.style.cssText='border-collapse:collapse;white-space:pre';for(const row of result){const tr=document.createElement('tr');for(const cell of row){const td=document.createElement('td');td.style.cssText='border:1px solid #ddd;padding:6px 10px';td.textContent=cell.v;tr.append(td);}table.append(tr);}preview.append(table);error.textContent='';return true;
  }catch(e){error.textContent=e.message;return false;}}
  for(const input of Object.values(inputs))input.oninput=()=>update();const actions=document.createElement('div');actions.style.cssText='display:flex;gap:10px;justify-content:flex-end;margin-top:14px';dialog.append(actions);let ended=false;const done=value=>{if(ended)return;ended=true;dialog.close();dialog.remove();resolve(value);};for(const [label,fn]of [['取消',()=>done(null)],['完成分列',()=>{try{if(!update(false))return;done({cells:textColumns(values,options()),target:inputs.target.value,overwrite:inputs.overwrite.value==='yes'});}catch(e){error.textContent=e.message;}}]]){const button=document.createElement('button');button.textContent=label;button.onclick=fn;actions.append(button);}
  dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});document.body.append(dialog);update();dialog.showModal();
});}
