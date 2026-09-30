import {Calc,parseA,A,usedRange} from './sheet-engine.js';
import {formulaHidden} from './sheet-protection.js';
const inside=(b,r,c)=>r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2;
export function traceReferences(doc,si,r,c,direction='precedents') {
 const calc=new Calc(doc),source={si,r1:r,c1:c,r2:r,c2:c},ranges=[],unresolved=[];
 if(direction==='precedents'&&formulaHidden(doc.sheets[si],A(r,c)))return {source,direction,ranges,unresolved:['Protected formula']};
 if(direction==='precedents')return {...calc.references(doc.sheets[si].cells[A(r,c)]?.literal?'':doc.sheets[si].cells[A(r,c)]?.v||'',si,r,c),source,direction};
 for(const [sheet,sh] of doc.sheets.entries())for(const [address,cell] of Object.entries(sh.cells)){
  if(cell.literal||typeof cell.v!=='string'||!cell.v.startsWith('='))continue;
  if(formulaHidden(sh,address))continue;
  const p=parseA(address);if(!p)continue;const refs=calc.references(cell.v,sheet,p.r,p.c);
  if(refs.ranges.some(b=>b.si===si&&inside(b,r,c)))ranges.push({si:sheet,r1:p.r,c1:p.c,r2:p.r,c2:p.c});
  if(refs.unresolved.length)unresolved.push(`${sh.name}!${address}`);
 }
 return {source,direction,ranges,unresolved};
}
// Rectangles keep selections compact, including a million blank cells between two stored rows.
export function specialRanges(doc,si,selection,kind) {
 const sh=doc.sheets[si],calc=new Calc(doc);calc.ensureSpills(si);
 const single=selection.r1===selection.r2&&selection.c1===selection.c2;
 let region=single?(kind==='validation'?{r1:0,c1:0,r2:1048575,c2:16383}:usedRange(sh)):selection;if(!region)return [];
 region={...region};
 if(single)for(const [key,m] of calc.arrays){const [s,r,c]=key.split(':').map(Number);if(s===si){region.r2=Math.max(region.r2,r+m.length-1);region.c2=Math.max(region.c2,c+m[0].length-1);}}
 const rows=new Map(),put=(r,c)=>{if(!inside(region,r,c))return;let cols=rows.get(r);if(!cols)rows.set(r,cols=new Set());cols.add(c);};
 if(kind==='validation'){
  const result=[];
  for(const rule of sh.dv||[])for(const ref of String(rule.range||'').split(/\s+/)){
   const [a,b=a]=ref.split(':').map(parseA);if(!a||!b)continue;
   let pieces=[{r1:Math.max(region.r1,a.r),c1:Math.max(region.c1,a.c),r2:Math.min(region.r2,b.r),c2:Math.min(region.c2,b.c)}].filter(x=>x.r1<=x.r2&&x.c1<=x.c2);
   for(const old of result)pieces=pieces.flatMap(p=>{const r1=Math.max(p.r1,old.r1),r2=Math.min(p.r2,old.r2),c1=Math.max(p.c1,old.c1),c2=Math.min(p.c2,old.c2);return r1>r2||c1>c2?[p]:[{...p,r2:r1-1},{...p,r1:r2+1},{r1,r2,c1:p.c1,c2:c1-1},{r1,r2,c1:c2+1,c2:p.c2}].filter(x=>x.r1<=x.r2&&x.c1<=x.c2);});
   result.push(...pieces);
  }return result.sort((a,b)=>a.r1-b.r1||a.c1-b.c1);
 }
 for(const [address,cell] of Object.entries(sh.cells)){
  const p=parseA(address);if(!p||!inside(region,p.r,p.c))continue;
  const formula=!cell.literal&&typeof cell.v==='string'&&cell.v.startsWith('='),filled=cell.v!==''&&cell.v!=null;
  const match=kind==='blanks'?filled:kind==='formulas'?formula&&!cell.spill:kind==='constants'?filled&&!formula&&!cell.spill:kind==='errors'?!!calc.value(si,p.r,p.c)?.err:kind==='notes'?!!cell.s?.note:kind==='validation'?!!cell.s?.dv:false;
  if(match)put(p.r,p.c);
 }
 if(kind==='blanks'||kind==='errors')for(const [key,spill] of calc.spills){const [s,r,c]=key.split(':').map(Number);if(s===si&&(kind==='blanks'||spill.value?.err))put(r,c);}
 const out=[],active=new Map(),emit=(r1,c1,r2,c2)=>{if(c2<c1||r2<r1)return;const key=c1+':'+c2,last=active.get(key);if(last&&last.r2+1===r1)last.r2=r2;else{const box={r1,c1,r2,c2};out.push(box);active.set(key,box);}};
 let row=region.r1;
 for(const [r,columns] of [...rows].sort((a,b)=>a[0]-b[0])){
  const cs=[...columns].sort((a,b)=>a-b);
  if(kind==='blanks'){
   emit(row,region.c1,r-1,region.c2);let start=region.c1;
   for(const c of cs){emit(r,start,r,c-1);start=c+1;}emit(r,start,r,region.c2);row=r+1;
  }else{let start=cs[0],end=start;for(const c of cs.slice(1)){if(c===end+1)end=c;else{emit(r,start,r,end);start=end=c;}}emit(r,start,r,end);}
 }
 if(kind==='blanks')emit(row,region.c1,region.r2,region.c2);
 if(kind==='blanks') {
  // Covered merge slots must never receive independent content or formatting.
  let result=out;
  for(const m of sh.merges||[]) for(const cut of [{r1:m.r,c1:m.c+1,r2:m.r,c2:m.c+(m.cs||1)-1},{r1:m.r+1,c1:m.c,r2:m.r+(m.rs||1)-1,c2:m.c+(m.cs||1)-1}]) {
   if(cut.r1>cut.r2||cut.c1>cut.c2)continue;
   result=result.flatMap(p=>{const r1=Math.max(p.r1,cut.r1),r2=Math.min(p.r2,cut.r2),c1=Math.max(p.c1,cut.c1),c2=Math.min(p.c2,cut.c2);return r1>r2||c1>c2?[p]:[{...p,r2:r1-1},{...p,r1:r2+1},{r1,r2,c1:p.c1,c2:c1-1},{r1,r2,c1:c2+1,c2:p.c2}].filter(x=>x.r1<=x.r2&&x.c1<=x.c2);});
  }return result.sort((a,b)=>a.r1-b.r1||a.c1-b.c1);
 }
 return out;
}
