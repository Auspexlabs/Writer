import {Calc,A,colName,parseA} from './sheet-engine.js';
import {formulaHidden} from './sheet-protection.js';

function expression(node,depth=0){
  if(depth>24)return '…';const sub=n=>expression(n,depth+1),sheet=node.sh?"'"+node.sh.replace(/'/g,"''")+"'!":'';
  switch(node.t){
    case 'n':return String(node.v);case 's':return '"'+node.v.replace(/"/g,'""')+'"';case 'b':return node.v?'TRUE':'FALSE';case 'empty':return '';
    case 'err':return node.e;case 'name':return node.v;case 'structured':return node.name+node.spec;
    case 'ref':return sheet+A(node.r,node.c);
    case 'rng':return sheet+(node.r2==null?colName(node.c1)+':'+colName(node.c2):node.c2==null?(node.r1+1)+':'+(node.r2+1):A(node.r1,node.c1)+':'+A(node.r2,node.c2));
    case 'op':return '('+sub(node.a)+node.op+sub(node.b)+')';case 'neg':return '-'+sub(node.a);case 'pct':return sub(node.a)+'%';case 'spill':return sub(node.a)+'#';
    case 'fn':return node.name+'('+node.args.map(sub).join(',')+')';case 'call':return sub(node.callee)+'('+node.args.map(sub).join(',')+')';
    case 'arr':return '{'+node.v.slice(0,6).map(row=>row.slice(0,6).map(sub).join(',')).join(';')+'}';default:return node.t;
  }
}
function preview(value){
  if(value?.err)return value.err;if(value?.lambda)return 'LAMBDA';
  if(value?.range){const rows=value.range,cols=rows[0]?.length||0;return rows.length+' × '+cols+'\n'+rows.slice(0,6).map(row=>row.slice(0,6).map(preview).join(' | ')+(cols>6?' …':'')).join('\n')+(rows.length>6?'\n…':'');}
  return typeof value==='string'?JSON.stringify(value).slice(0,2000):value==null?'空白':typeof value==='boolean'?value?'TRUE':'FALSE':String(value);
}
/** Trace the actual evaluator once, preserving short-circuit decisions, scopes,
 * errors, volatile values and dependency caches. Never rerun nodes to explain them. */
export function evaluateFormula(doc,si,ref,{maxSteps=1000}={}){
  const cell=doc.sheets[si]?.cells[ref],position=parseA(ref);
  if(!position||cell?.literal||!String(cell?.v||'').startsWith('='))throw Error('请选择含公式的单元格');
  if(formulaHidden(doc.sheets[si],ref))throw Error('受保护的隐藏公式不能求值查看');
  const steps=[];let depth=0,truncated=false;
  class Evaluator extends Calc{
    ev(node,sheet){
      const level=depth++,address=A(this.cur.r,this.cur.c),hidden=this.privateTrace>0||formulaHidden(doc.sheets[sheet],address);if(hidden)this.privateTrace=(this.privateTrace||0)+1;
      const record=value=>{if(hidden||['n','s','b','empty'].includes(node.t))return;if(steps.length>=maxSteps){truncated=true;return;}steps.push({cell:doc.sheets[sheet].name+'!'+address,depth:level,expression:expression(node).slice(0,4000),value:preview(value)});};
      try{const value=super.ev(node,sheet);record(value);return value;}catch(error){record({err:error.e||'#VALUE!'});throw error;}finally{depth--;if(hidden)this.privateTrace--;}
    }
  }
  const calc=new Evaluator(doc),value=calc.value(si,position.r,position.c);
  if(!steps.length)steps.push({cell:doc.sheets[si].name+'!'+ref,depth:0,expression:cell.v.slice(1),value:preview(value)});
  return {formula:cell.v,cell:doc.sheets[si].name+'!'+ref,result:preview(value),steps,truncated,iteration:calc.iterationStatus||null};
}

export function showEvaluation(doc,si,ref){return new Promise(resolve=>{
  const dialog=document.createElement('dialog');dialog.dataset.formulaEvaluation='1';dialog.style.cssText='width:650px;max-width:90vw;max-height:85vh;overflow:auto;padding:24px;border:1px solid #bbb;border-radius:16px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
  const title=document.createElement('h3');title.textContent='公式求值';const formula=document.createElement('pre'),where=document.createElement('p'),expr=document.createElement('pre'),value=document.createElement('pre'),status=document.createElement('p'),actions=document.createElement('div');
  for(const el of [formula,expr,value])el.style.cssText='white-space:pre-wrap;overflow-wrap:anywhere;padding:12px;background:#8881;border-radius:8px';
  status.setAttribute('role','status');actions.style.cssText='display:flex;gap:10px;flex-wrap:wrap';dialog.append(title,formula,where,expr,value,status,actions);
  let worker,timer,result,index=0,ended=false;
  const close=()=>{if(ended)return;ended=true;worker?.terminate();clearTimeout(timer);dialog.close();dialog.remove();resolve();};
  const render=()=>{const step=result?.steps[index];formula.textContent=result?.formula||doc.sheets[si].cells[ref]?.v||'';where.textContent=step?`${index+1} / ${result.steps.length} · ${step.cell}`:'';expr.textContent=step?.expression||'';value.textContent=step?.value||'';status.textContent=result?'结果：'+result.result+(result.truncated?' · 仅显示前 1000 步':'')+(result.iteration?' · 迭代 '+result.iteration.iterations+' 次':''):'正在计算…';previous.disabled=!result||index===0;next.disabled=!result||index>=result.steps.length-1;last.disabled=!result||!result.steps.length;};
  const button=(text,fn)=>{const b=document.createElement('button');b.textContent=text;b.onclick=fn;actions.append(b);return b;};
  const previous=button('上一步',()=>{index=Math.max(0,index-1);render();}),next=button('下一步',()=>{index=Math.min(result.steps.length-1,index+1);render();}),last=button('最后一步',()=>{index=Math.max(0,result.steps.length-1);render();});button('关闭',close);
  dialog.addEventListener('cancel',e=>{e.preventDefault();close();});document.body.append(dialog);dialog.showModal();render();
  try{worker=new Worker(new URL('./sheet-evaluate-worker.js',import.meta.url),{type:'module'});timer=setTimeout(()=>{worker.terminate();status.textContent='计算超过 10 秒，请缩小公式依赖范围';},10000);
    worker.onmessage=e=>{clearTimeout(timer);worker.terminate();if(e.data.error){status.textContent=e.data.error;return;}result=e.data.result;render();};
    worker.onerror=e=>{clearTimeout(timer);worker.terminate();status.textContent=e.message;};
    const snapshot=JSON.parse(JSON.stringify({sheets:doc.sheets.map(s=>({name:s.name,cells:s.cells,names:s.names,tables:s.tables,dataTables:s.dataTables,hiddenRows:s.hiddenRows,frows:s.frows,outline:s.outline,protected:s.protected,protection:s.protection})),names:doc.names,iteration:doc.iteration,date1904:doc.date1904}));
    worker.postMessage({doc:snapshot,si,ref});
  }catch(e){clearTimeout(timer);worker?.terminate();status.textContent=e.message;}
});}
