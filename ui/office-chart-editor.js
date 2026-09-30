// A shared data editor for native Word and PowerPoint charts. Cells are plain data;
// text beginning with '=' is never interpreted as an embedded workbook formula.
export const chartTypes = [['column','柱形图'],['bar','条形图'],['line','折线图'],['area','面积图'],['pie','饼图'],['doughnut','圆环图'],['scatter','散点图'],['combo','组合图']];
export function chartTable(chart) {
  const series=chart?.series?.length?chart.series:[{name:'系列 1',values:[12,18,15]}],scatter=chart?.kind==='scatter',labels=chart?.cats?.length?chart.cats:Array.from({length:Math.max(...series.map(s=>s.values?.length||0))},(_,i)=>String(i+1));
  const rows=[['类别',...series.flatMap((s,i)=>scatter?[`${s.name||'系列 '+(i+1)} X`,`${s.name||'系列 '+(i+1)} Y`]:[s.name||'系列 '+(i+1)])]],quote=v=>/[\t\n\r"]/.test(String(v??''))?'"'+String(v).replace(/"/g,'""')+'"':String(v??'');
  for(let i=0;i<labels.length;i++)rows.push([labels[i],...series.flatMap(s=>scatter?[s.x?.[i]??i+1,s.values?.[i]??'']:[s.values?.[i]??''])]);
  return rows.map(r=>r.map(quote).join('\t')).join('\n');
}
export function parseChartTable(text, type='column', kinds=[]) {
  const delimiter=text.includes('\t')?'\t':',',rows=[],row=[];let cell='',quoted=false;
  for(let i=0;i<=text.length;i++) { const c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"'){cell+='"';i++;}else if(quoted||!cell)quoted=!quoted;else cell+=c;}else if(!quoted&&(c===delimiter||c==='\n'||c==='\r'||c===undefined)){row.push(cell);cell='';if(c!==delimiter){if(row.some(v=>v!==''))rows.push(row.splice(0));else row.length=0;if(c==='\r'&&text[i+1]==='\n')i++;}}else if(c!==undefined)cell+=c; }
  if(quoted)throw Error('数据中的引号没有闭合');
  const headers=rows.shift();if(!headers||headers.length<2||!rows.length||rows.some(r=>r.length!==headers.length))throw Error('请输入标题行和至少一行数据，每行列数需一致');
  const numeric=v=>{if(!v.trim())return null;const n=Number(v);if(!Number.isFinite(n))throw Error('数值列只能填写数字或留空');return n;},series=[];
  for(let c=1;c<headers.length;c+=type==='scatter'?2:1){if(type==='scatter'&&c+1>=headers.length)throw Error('散点图每个系列需要 X、Y 两列');series.push({name:headers[c].replace(type==='scatter'?/ X$/:/$^/,'')||'系列 '+series.length,values:rows.map(r=>numeric(r[c+(type==='scatter'?1:0)])),...(type==='scatter'?{x:rows.map(r=>{const n=numeric(r[c]);if(n===null)throw Error('散点图的 X 坐标不能为空');return n;})}:{}),...(type==='combo'?{kind:kinds[series.length]|| (series.length?'line':'column')}: {})});}
  return {labels:rows.map(r=>r[0]),series};
}
export function editChart(chart=null) {
  return new Promise(resolve=>{
    const dialog=document.createElement('dialog');dialog.dataset.officeChartEditor='1';dialog.style.cssText='width:660px;max-width:calc(100vw - 48px);max-height:calc(100vh - 64px);overflow:auto;border:1px solid #ccc;border-radius:16px;padding:24px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui;box-shadow:0 20px 80px #0004';
    const form=document.createElement('form');form.style.cssText='display:grid;gap:12px';dialog.append(form);
    const heading=document.createElement('strong');heading.textContent=chart?'编辑图表数据':'插入图表';form.append(heading);
    const field=(name,label,value,options)=>{const wrap=document.createElement('label');wrap.style.cssText='display:grid;gap:5px';wrap.append(document.createTextNode(label));const input=document.createElement(options?'select':name==='table'?'textarea':'input');input.name=name;input.style.cssText='padding:8px;border:1px solid #bbb;border-radius:6px;background:transparent;color:inherit;font:inherit';if(options)for(const [v,l] of options){const option=document.createElement('option');option.value=v;option.textContent=l;input.append(option);}input.value=value??'';wrap.append(input);form.append(wrap);return input;};
    const combo=chart?.series?.some(s=>s.kind&&s.kind!==chart.kind),kind=combo?'combo':chart?.kind==='bar'?(chart.dir==='bar'?'bar':'column'):chart?.kind||'column';
    const title=field('title','标题',chart?.title||''),type=field('type','图表类型',kind,chartTypes),table=field('table','数据：首行为系列名称，首列为类别。可从表格复制粘贴，散点图每个系列填写 X、Y 两列。',chartTable(chart));table.rows=9;table.wrap='off';table.style.fontFamily='monospace';
    const legend=field('legend','图例位置',({b:'bottom',t:'top',l:'left',r:'right',tr:'topRight','':'none'})[chart?.legend]||'bottom',[['bottom','下方'],['top','上方'],['left','左侧'],['right','右侧'],['none','隐藏']]);
    const grouping=field('grouping','系列排列',chart?.grouping||'clustered',[['clustered','并列'],['stacked','堆积'],['percentStacked','百分比堆积']]),labels=field('labels','数据标签',chart?.labels?'yes':'no',[['no','隐藏'],['yes','显示']]);
    const xt=field('xTitle','横轴标题',chart?.xTitle||''),yt=field('yTitle','纵轴标题',chart?.yTitle||'');
    const kinds=document.createElement('div');kinds.style.cssText='display:flex;gap:8px;flex-wrap:wrap';form.append(kinds);let seriesKinds=[];
    const updateKinds=()=>{const previous=seriesKinds.map(s=>s.value);kinds.replaceChildren();seriesKinds=[];if(type.value!=='combo')return;try{parseChartTable(table.value,'combo').series.forEach((s,i)=>{const label=document.createElement('label');label.append(document.createTextNode(s.name+' '));const select=document.createElement('select');for(const [v,t] of chartTypes.filter(([k])=>['column','line','area'].includes(k))){const opt=document.createElement('option');opt.value=v;opt.textContent=t;select.append(opt);}select.value=previous[i]||(chart?.series?.[i]?.kind==='bar'?'column':chart?.series?.[i]?.kind)|| (i?'line':'column');label.append(select);kinds.append(label);seriesKinds.push(select);});}catch{}};
    let priorType=kind;type.onchange=()=>{try{if((type.value==='scatter')!==(priorType==='scatter')){const data=parseChartTable(table.value,priorType,seriesKinds.map(s=>s.value));table.value=chartTable({...data,kind:type.value,cats:data.labels});}priorType=type.value;updateKinds();error.textContent='';}catch(e){type.value=priorType;error.textContent=e.message;}};table.addEventListener('change',updateKinds);updateKinds();
    const error=document.createElement('div');error.setAttribute('role','alert');error.style.color='#b3261e';form.append(error);
    const buttons=document.createElement('div');buttons.style.cssText='display:flex;justify-content:flex-end;gap:10px';const cancel=document.createElement('button'),save=document.createElement('button');cancel.type='button';cancel.textContent='取消';save.type='submit';save.textContent='保存图表';buttons.append(cancel,save);form.append(buttons);
    let finished=false;const done=value=>{if(finished)return;finished=true;dialog.close();dialog.remove();resolve(value);};cancel.onclick=()=>done(null);dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});
    form.onsubmit=e=>{e.preventDefault();try{const data=parseChartTable(table.value,type.value,seriesKinds.map(s=>s.value));done({...data,title:title.value,type:type.value,legend:legend.value,stacked:grouping.value!=='clustered',percentStacked:grouping.value==='percentStacked',dataLabels:labels.value==='yes',xTitle:xt.value,yTitle:yt.value});}catch(e){error.textContent=e.message;}};
    document.body.append(dialog);dialog.showModal();title.focus();
  });
}
