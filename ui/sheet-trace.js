// Viewport-sized dependency arrows. World coordinates stay outside the SVG, so
// a reference to row 1,048,576 never allocates a worksheet-sized drawing surface.
const escape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const clamp=(x,a,b)=>Math.max(a,Math.min(b,x));
const overlap=(a,b)=>a.si===b.si&&a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;

export function traceGeometry(trace,si,geometry,{width=1000,height=700,scrollX=0,scrollY=0,headerX=44,headerY=26,freezeRows=0,freezeCols=0,merges=[],limit=64}={}){
  if(!trace||width<headerX+20||height<headerY+20)return [];
  const {xs,ys}=geometry,axisAt=(coords,i,step)=>i<coords.length?coords[Math.max(0,i)]:coords.at(-1)+(i-coords.length+1)*step;
  const point=(b,slot)=>{
    if(b.si!==si)return {x:width-12,y:Math.min(height-12,headerY+22+slot*26),external:true,visible:false};
    if(b.r1===b.r2&&b.c1===b.c2){const m=merges.find(m=>b.r1>=m.r&&b.r1<m.r+m.rs&&b.c1>=m.c&&b.c1<m.c+m.cs);if(m)b={...b,r1:m.r,r2:m.r+m.rs-1,c1:m.c,c2:m.c+m.cs-1};}
    const axis=(a,b,coords,freeze,scroll,header,size,step)=>{
      // The visible intersection of a range supplies its attachment point.
      const positions=[];
      for(const [start,end,frozen] of [[a,Math.min(b,freeze-1),true],[Math.max(a,freeze),b,false]]){
        if(start>end)continue;
        const lo=header+(frozen?0:axisAt(coords,freeze,step)),hi=frozen?Math.min(size,header+axisAt(coords,freeze,step)):size;
        const p=header+axisAt(coords,start,step)-(frozen?0:scroll),q=header+axisAt(coords,end+1,step)-(frozen?0:scroll);
        if(q>lo&&p<hi&&q>p)positions.push([Math.max(lo,p),Math.min(hi,q)]);
      }
      if(positions.length){const p=positions[0];return {value:(p[0]+p[1])/2,visible:true};}
      const raw=header+axisAt(coords,a,step)-(a<freeze?0:scroll);
      return {value:clamp(raw,header+4,size-5),visible:false};
    };
    const x=axis(b.c1,b.c2,xs,freezeCols,scrollX,headerX,width,64),y=axis(b.r1,b.r2,ys,freezeRows,scrollY,headerY,height,20);
    return {x:x.value,y:y.value,visible:x.visible&&y.visible,external:false};
  };
  const result=[],seen=new Set();
  for(const [index,b] of trace.ranges.entries()){
    if(trace.source.si!==si&&b.si!==si)continue;
    const key=JSON.stringify(b);if(seen.has(key))continue;seen.add(key);
    const source=point(trace.source,result.length),target=point(b,result.length);
    if(!source.visible&&!target.visible)continue;
    const from=trace.direction==='dependents'?source:target,to=trace.direction==='dependents'?target:source;
    const related=b.si===si?trace.source:b;
    result.push({index,from,to,jump:b,portal:source.external||target.external?related:null,dashed:!source.visible||!target.visible,loop:overlap(trace.source,b)&&Math.hypot(to.x-from.x,to.y-from.y)<8});
    if(result.length>=limit)break;
  }
  return result;
}

export function traceSvg(edges,width,height,label,rangeName){
  if(!edges.length)return '';
  const color='#0A5BD7',path=(from,to,loop)=>{
    if(loop)return `M ${from.x} ${from.y} c 28 -30 38 24 4 4`;
    const dx=to.x-from.x,dy=to.y-from.y,k=1-7/Math.max(7,Math.hypot(dx,dy));
    return `M ${from.x} ${from.y} L ${from.x+dx*k} ${from.y+dy*k}`;
  };
  return `<svg xmlns="http://www.w3.org/2000/svg" data-sheet-trace="1" role="img" aria-label="${escape(label)}" width="${width}" height="${height}" style="display:block;overflow:hidden;pointer-events:none"><defs><marker id="sheet-trace-tip" markerWidth="7" markerHeight="7" refX="6" refY="3.5" orient="auto"><path d="M0 0 L7 3.5 L0 7 Z" fill="${color}"/></marker></defs>${edges.map(edge=>{
    const {from,to,index,portal}=edge,d=path(from,to,edge.loop),title=rangeName(edge.jump),p=portal?(from.external?from:to):null,text=portal?rangeName(portal):'';
    return `<g data-trace-range="${index}" tabindex="0" role="button" aria-label="${escape(title)}" style="cursor:pointer"><title>${escape(title)}</title><path d="${d}" fill="none" stroke="white" stroke-width="4" opacity=".85"/><path data-trace-line="1" d="${d}" fill="none" stroke="${color}" stroke-width="1.7" ${edge.dashed?'stroke-dasharray="5 4"':''} marker-end="url(#sheet-trace-tip)" style="pointer-events:stroke"/><circle cx="${from.x}" cy="${from.y}" r="3" fill="${color}"/>${p?`<text x="${p.x-7}" y="${p.y-6}" text-anchor="end" fill="${color}" stroke="white" stroke-width="3" paint-order="stroke" font-size="11" style="pointer-events:auto">${escape(text)}</text>`:''}</g>`;
  }).join('')}</svg>`;
}
