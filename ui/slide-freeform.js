const T=s=>typeof globalThis.$t==='function'?globalThis.$t(s):s;
export function freeformPaths(data) {
  const w=Number(data?.w),h=Number(data?.h);
  if(!(w>0&&h>0&&Number.isFinite(w)&&Number.isFinite(h))||!Array.isArray(data.paths)||!data.paths.length)return [];
  const scale=point=>[Number(point?.[0])/w*1000,Number(point?.[1])/h*1000];
  return data.paths.map(path=>({...path,points:(path.points||[]).map(scale),...(path.controls?{controls:Object.fromEntries(Object.entries(path.controls).map(([i,controls])=>[i,Array.isArray(controls)?controls.map(scale):[]]))}:{})})).filter(path=>path.points.length&&path.points.every(point=>point.every(Number.isFinite))&&Object.entries(path.controls||{}).every(([i,controls])=>/^[1-9]\d*$/.test(i)&&+i<path.points.length&&controls.length>=1&&controls.length<=2&&controls.every(p=>p.every(Number.isFinite))));
}
export function freeformPathD(path) {return path.points.map((point,i)=>{const controls=path.controls?.[i];return (i===0?'M':controls?.length===2?'C':controls?.length===1?'Q':'L')+(controls||[]).concat([point]).map(p=>p.map(n=>+n.toFixed(4)).join(',')).join(' ');}).join(' ')+(path.closed?' Z':'');}
export function freeformSvg(data,attrs) {
  return freeformPaths(data).map(path=>`${Object.keys(path.controls||{}).length?'<path d="'+freeformPathD(path)+'"':'<'+(path.closed?'polygon':'polyline')+' points="'+path.points.map(point=>point.map(n=>+n.toFixed(4)).join(',')).join(' ')+'"'} ${attrs.replace(/fill="[^"]*"/,s=>path.fill==='none'?'fill="none"':s).replace(/stroke="[^"]*"/,s=>path.stroke===false?'stroke="none"':s)} stroke-linecap="round"/>`).join('');
}
export function splitFreeformSegment(path,index) {
  const next=index+1,a=path.points[index],b=path.points[next]||path.points[0],controls=path.controls?.[next],mid=(a,b)=>a.map((v,i)=>(v+b[i])/2);
  let point=mid(a,b),left,right;
  if(controls?.length===1){const x=mid(a,controls[0]),y=mid(controls[0],b);point=mid(x,y);left=[x];right=[y];}
  if(controls?.length===2){const x=mid(a,controls[0]),y=mid(controls[0],controls[1]),z=mid(controls[1],b),u=mid(x,y),v=mid(y,z);point=mid(u,v);left=[x,u];right=[v,z];}
  path.points.splice(next,0,point);const shifted={};for(const [i,control] of Object.entries(path.controls||{}))shifted[+i>=next?+i+1:i]=control;
  if(left){shifted[next]=left;shifted[next+1]=right;}path.controls=shifted;return next;
}
export function inkGeometry(lines) {
  const paths=lines.filter(line=>line.length).map(line=>line.length===1?[line[0],[line[0][0]+.1,line[0][1]]]:line);
  if(!paths.length)return null;
  let left=Infinity,top=Infinity,right=-Infinity,bottom=-Infinity;
  for(const path of paths)for(const [x,y]of path){left=Math.min(left,x);top=Math.min(top,y);right=Math.max(right,x);bottom=Math.max(bottom,y);}
  const x=left-4,y=top-4,w=Math.max(8,right-left+8),h=Math.max(8,bottom-top+8);
  return {x,y,w,h,pathData:{w:1000,h:1000,paths:paths.map(points=>({closed:false,fill:'none',stroke:true,points:points.map(([px,py])=>[(px-x)/w*1000,(py-y)/h*1000])}))}};
}
/** Edits a copy; applying it is one normal slide history operation. */
export function editFreeform(data) {
  const paths=freeformPaths(data);if(!paths.length)paths.push({points:[[100,800],[500,150],[900,800]],closed:false,fill:'none',stroke:true});
  let pathIndex=0,pointIndex=0,drag=false,dragControl=null;
  return new Promise(resolve=>{
    const dialog=document.createElement('dialog');dialog.dataset.freeformEditor='1';dialog.style.cssText='width:min(720px,90vw);padding:22px;border:1px solid #ccc;border-radius:16px;background:#fff;color:#222;font:14px sans-serif';
    dialog.innerHTML=`<form><h3>${T('编辑顶点')}</h3><div style="display:flex;gap:12px;align-items:center"><select name="path" aria-label="${T('路径')}"></select><label><input type="checkbox" name="closed"> ${T('闭合路径')}</label><button type="button" data-action="new">${T('新路径')}</button><button type="button" data-action="removePath">${T('删除路径')}</button></div><svg viewBox="-50 -50 1100 1100" style="display:block;width:100%;height:340px;background:#f5f5f7;border:1px solid #ddd;touch-action:none;margin:12px 0" aria-label="${T('拖动顶点调整形状')}"></svg><div style="display:flex;gap:10px;align-items:center"><label>${T('顶点')} <input type="number" name="point" min="1" style="width:60px"></label><label>${T('线段')} <select name="segment"><option value="line">${T('直线')}</option><option value="quadratic">${T('二次曲线')}</option><option value="cubic">${T('三次曲线')}</option></select></label><label>X <input type="number" name="x" step="any" style="width:80px"></label><label>Y <input type="number" name="y" step="any" style="width:80px"></label><button type="button" data-action="insert">${T('增加顶点')}</button><button type="button" data-action="remove">${T('删除顶点')}</button></div><p role="alert" style="color:#a33"></p><div style="display:flex;justify-content:flex-end;gap:10px;margin-top:16px"><button type="button" data-action="cancel">${T('取消')}</button><button type="submit">${T('应用')}</button></div></form>`;
    const svg=dialog.querySelector('svg'),form=dialog.querySelector('form'),select=form.elements.path,x=form.elements.x,y=form.elements.y,closed=form.elements.closed,segment=form.elements.segment,vertex=form.elements.point;
    const current=()=>paths[pathIndex],point=()=>current().points[pointIndex],finish=value=>{dialog.close();dialog.remove();resolve(value);};
    const paint=()=>{svg.innerHTML=freeformSvg({w:1000,h:1000,paths},'fill="#DCEBFF" stroke="#2F5D8A" stroke-width="3" vector-effect="non-scaling-stroke"')+(current().controls?.[pointIndex]||[]).map(([px,py],i)=>{const anchor=current().points[i===0?pointIndex-1:pointIndex];return `<line x1="${anchor[0]}" y1="${anchor[1]}" x2="${px}" y2="${py}" stroke="#888" stroke-dasharray="8 8"/><circle data-control="${i}" cx="${px}" cy="${py}" r="11" fill="#FFC966" stroke="#965500"/>`;}).join('')+current().points.map(([px,py],i)=>`<circle data-point="${i}" cx="${px}" cy="${py}" r="${i===pointIndex?13:10}" fill="${i===pointIndex?'#FF7A00':'#fff'}" stroke="#2F5D8A" stroke-width="2" vector-effect="non-scaling-stroke"/>`).join('');x.value=+point()[0].toFixed(2);y.value=+point()[1].toFixed(2);closed.checked=!!current().closed;vertex.value=pointIndex+1;vertex.max=current().points.length;segment.disabled=pointIndex===0;segment.value=current().controls?.[pointIndex]?.length===2?'cubic':current().controls?.[pointIndex]?.length===1?'quadratic':'line';};
    const list=()=>{select.replaceChildren(...paths.map((p,i)=>new Option(T('路径')+' '+(i+1),String(i))));select.value=pathIndex;paint();};
    vertex.onchange=()=>{pointIndex=Math.max(0,Math.min(current().points.length-1,Math.trunc(+vertex.value||1)-1));paint();};
    segment.onchange=()=>{if(!pointIndex)return;const a=current().points[pointIndex-1],b=point(),at=t=>a.map((v,i)=>v+(b[i]-v)*t);current().controls||={};if(segment.value==='line')delete current().controls[pointIndex];else current().controls[pointIndex]=segment.value==='cubic'?[at(1/3),at(2/3)]:[at(.5)];paint();};
    select.onchange=()=>{pathIndex=+select.value;pointIndex=0;paint();};closed.onchange=()=>{current().closed=closed.checked;current().fill=closed.checked?'norm':'none';paint();};
    const coordinate=event=>{const pt=svg.createSVGPoint();pt.x=event.clientX;pt.y=event.clientY;const p=pt.matrixTransform(svg.getScreenCTM().inverse());return [Math.max(-10000,Math.min(10000,p.x)),Math.max(-10000,Math.min(10000,p.y))];};
    svg.onpointerdown=event=>{const node=event.target.closest('[data-point],[data-control]');if(!node)return;event.preventDefault();dragControl=node.dataset.control==null?null:+node.dataset.control;if(dragControl==null)pointIndex=+node.dataset.point;drag=true;svg.setPointerCapture(event.pointerId);paint();};
    svg.onpointermove=event=>{if(drag){if(dragControl==null)current().points[pointIndex]=coordinate(event);else current().controls[pointIndex][dragControl]=coordinate(event);paint();}};svg.onpointerup=svg.onpointercancel=()=>{drag=false;};
    svg.ondblclick=event=>{if(event.target.closest('[data-point],[data-control]'))return;pointIndex=splitFreeformSegment(current(),pointIndex);current().points[pointIndex]=coordinate(event);paint();};
    for(const input of [x,y])input.onchange=()=>{const values=[+x.value,+y.value];if(values.every(n=>Number.isFinite(n)&&Math.abs(n)<=10000))current().points[pointIndex]=values;paint();};
    dialog.onclick=event=>{const action=event.target.closest('[data-action]')?.dataset.action;if(action==='cancel')finish(null);else if(action==='insert'){pointIndex=splitFreeformSegment(current(),pointIndex);paint();}else if(action==='remove'&&current().points.length>1){current().points.splice(pointIndex,1);current().controls=Object.fromEntries(Object.entries(current().controls||{}).filter(([i])=>+i!==pointIndex&&+i!==pointIndex+1).map(([i,c])=>[+i>pointIndex?+i-1:i,c]));pointIndex=Math.min(pointIndex,current().points.length-1);paint();}else if(action==='new'){paths.push({points:[[150,500],[850,500]],fill:'none',stroke:true,closed:false});pathIndex=paths.length-1;pointIndex=0;list();}else if(action==='removePath'&&paths.length>1){paths.splice(pathIndex,1);pathIndex=0;pointIndex=0;list();}};
    form.onsubmit=event=>{event.preventDefault();if(paths.length>1000||paths.reduce((n,p)=>n+p.points.length,0)>100000){dialog.querySelector('[role=alert]').textContent=T('路径或顶点过多');return;}finish({w:1000,h:1000,paths});};
    dialog.oncancel=event=>{event.preventDefault();finish(null);};document.body.append(dialog);list();dialog.showModal();
  });
}
