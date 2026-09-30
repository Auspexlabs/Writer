import { changedCellKeys } from './sheet-model.js';

export const protectionOptions = [
  ['formatCells','设置单元格格式'], ['formatRows','设置行格式'], ['formatColumns','设置列格式'],
  ['insertRows','插入行'], ['insertColumns','插入列'], ['deleteRows','删除行'], ['deleteColumns','删除列'],
  ['sort','排序'], ['autoFilter','使用自动筛选'], ['insertHyperlinks','插入超链接'],
  ['objects','编辑对象和批注'], ['scenarios','编辑方案'], ['pivotTables','使用透视表'],
  ['selectLockedCells','选择锁定单元格'], ['selectUnlockedCells','选择未锁定单元格']
];
const empty = x => x&&typeof x==='object'&&Object.keys(x).length===0;
const same = (a,b) => a===b || a==null&&empty(b) || b==null&&empty(a) || JSON.stringify(a??null)===JSON.stringify(b??null);
const position=ref=>{const m=/^([A-Z]+)(\d+)$/.exec(ref);return {col:m[1],row:m[2]};};
const colName=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const colIndex=s=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;
export function inheritedStyle(sheet,ref){const p=position(ref),d=sheet.rowColumnStyles;return d?.styles?.[d.rows?.[p.row]??d.cols?.[p.col]]||d?.base||{};}
export function effectiveCellStyle(sheet,ref){const cell=sheet.cells[ref];return cell?.styleExplicit?cell.s||{}:{...inheritedStyle(sheet,ref).ui,...cell?.s};}
export function cellProtection(sheet,ref){const s=sheet.cells[ref]?.s||{},base=inheritedStyle(sheet,ref);return {locked:s.locked??base.locked??true,formulaHidden:s.formulaHidden??base.formulaHidden??false};}
export const cellLocked = (sheet,ref) => !!sheet.protected && cellProtection(sheet,ref).locked;
export const formulaHidden = (sheet,ref) => !!sheet.protected && cellProtection(sheet,ref).formulaHidden;
export const protectionAllows = (sheet,action) => !sheet.protected || sheet.protection?.[action]===true;

export function workbookStructureViolation(before,after) {
  if(!before.workbookProtection?.structure)return false;
  return before.sheets.length!==after.sheets.length || before.sheets.some((s,i)=>{
    const next=after.sheets[i];return s.name!==next.name || s.path!==next.path || (s.visibility||'visible')!==(next.visibility||'visible');
  });
}

/** True if any real or unmaterialized cell in a rectangle remains locked. Splits
 * at sparse defaults/overrides instead of enumerating a million-row column. */
export function areaHasLocked(sheet,b,wanted=true){
  if((b.r2-b.r1+1)*(b.c2-b.c1+1)<=4096){for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)if(cellProtection(sheet,colName(c)+(r+1)).locked===wanted)return true;return false;}
  const d=sheet.rowColumnStyles||{},rows=d.rows||{},cols=d.cols||{},xs=new Set([b.c1,b.c2+1]),ys=new Set([b.r1,b.r2+1]);
  for(const row of Object.keys(rows)){const r=+row-1;if(r>=b.r1&&r<=b.r2){ys.add(r);ys.add(r+1);}}
  for(const col of Object.keys(cols)){const c=colIndex(col);if(c>=b.c1&&c<=b.c2){xs.add(c);xs.add(c+1);}}
  // First check explicit cells. Then find an uncovered position in each uniform
  // default rectangle; unlocked explicit cells may fully cover a small block.
  const exceptions=[];
  for(const [ref,cell] of Object.entries(sheet.cells)){const p=position(ref),r=+p.row-1,c=colIndex(p.col);if(r<b.r1||r>b.r2||c<b.c1||c>b.c2)continue;if(cellProtection(sheet,ref).locked===wanted)return true;if(cell.s?.locked===!wanted)exceptions.push({r,c});}
  const xx=[...xs].sort((a,b)=>a-b),yy=[...ys].sort((a,b)=>a-b);
  for(let yi=0;yi<yy.length-1;yi++)for(let xi=0;xi<xx.length-1;xi++){
    const r1=yy[yi],r2=yy[yi+1],c1=xx[xi],c2=xx[xi+1];if((inheritedStyle(sheet,colName(c1)+(r1+1)).locked!==false)!==wanted)continue;
    const size=(r2-r1)*(c2-c1);if(size>exceptions.length||exceptions.filter(p=>p.r>=r1&&p.r<r2&&p.c>=c1&&p.c<c2).length<size)return true;
  }
  return false;
}
export function selectionAllowed(sheet,areas){return !sheet.protected||areas.every(b=>(sheet.protection?.selectLockedCells!==false||!areaHasLocked(sheet,b))&&(sheet.protection?.selectUnlockedCells!==false||!areaHasLocked(sheet,b,false)));}
export function nextSelectable(sheet,point,axis,direction,limit,visible=()=>true){
  if(point.r<0||point.c<0||(axis==='r'?point.r:point.c)>=limit)return null;
  if(visible(point.r,point.c)&&selectionAllowed(sheet,[{r1:point.r,r2:point.r,c1:point.c,c2:point.c}]))return point;
  const value=axis==='r'?point.r:point.c,candidates=new Set([value]);
  const add=i=>{for(const n of [i-1,i,i+1])if(n>=0&&n<limit&&(n-value)*direction>=0)candidates.add(n);};
  for(const key of Object.keys(sheet.rowColumnStyles?.[axis==='r'?'rows':'cols']||{}))add(axis==='r'?+key-1:colIndex(key));
  for(const ref of Object.keys(sheet.cells)){const p=position(ref),r=+p.row-1,c=colIndex(p.col);if(axis==='r'?c===point.c:r===point.r)add(axis==='r'?r:c);}
  for(const i of [...candidates].sort((a,b)=>(a-b)*direction)){const r=axis==='r'?i:point.r,c=axis==='c'?i:point.c;if(visible(r,c)&&selectionAllowed(sheet,[{r1:r,r2:r,c1:c,c2:c}]))return {r,c};}
  return null;
}

export function setAxisProtection(sheet,axis,start,end,patch){
  const d=sheet.rowColumnStyles||={rows:{},cols:{},styles:{},base:{format:'<x:xf xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" numFmtId="0" fontId="0" fillId="0" borderId="0"/>',locked:true,formulaHidden:false}};
  const entries=d[axis==='r'?'rows':'cols'];
  const changed=new Map();
  for(let i=start;i<=end;i++){
    const key=axis==='r'?String(i+1):colName(i),old=entries[key],source=d.styles[old]||d.base,signature=JSON.stringify([old,patch]);
    let id=changed.get(signature);if(!id){let n=0;while(d.styles['p'+n])n++;id='p'+n;d.styles[id]={...source,...patch};changed.set(signature,id);}entries[key]=id;
  }
  for(const [ref,cell] of Object.entries(sheet.cells)){const p=position(ref),i=axis==='r'?+p.row-1:colIndex(p.col);if(i>=start&&i<=end)cell.s={...cell.s,...patch};}
  pruneDefaults(d);
}
const pruneDefaults=d=>{const used=new Set([...Object.values(d.rows),...Object.values(d.cols)]);for(const id of Object.keys(d.styles))if(!used.has(id))delete d.styles[id];};
export function shiftStyleDefaults(sheet,axis,at,n){
  const d=sheet.rowColumnStyles;if(!d)return;const key=axis==='r'?'rows':'cols',old=d[key],next={},limit=axis==='r'?1048576:16384;
  const name=i=>axis==='r'?String(i+1):colName(i),inherited=old[name(Math.max(0,at-1))];
  for(const [key,value] of Object.entries(old)){const i=axis==='r'?+key-1:colIndex(key);if(n<0&&i>=at&&i<at-n)continue;const to=i>=at?i+n:i;if(to>=0&&to<limit)next[name(to)]=value;}
  if(n>0&&inherited!=null)for(let i=at;i<at+n;i++)next[name(i)]=inherited;
  d[key]=next;pruneDefaults(d);
}
/** Pasting into an unlocked input must not accidentally re-lock it or change
 * forbidden formatting. Protected destination permissions remain in place. */
export function protectedPaste(sheet,ref,cell) {
  if(!sheet.protected)return cell;
  const old=sheet.cells[ref],style={...(protectionAllows(sheet,'formatCells')?cell?.s:old?.s)};
  for(const key of ['locked','formulaHidden',...(!protectionAllows(sheet,'objects')?['note']:[]),...(!protectionAllows(sheet,'insertHyperlinks')?['link']:[])]){
    if(old?.s?.[key]===undefined)delete style[key];else style[key]=old.s[key];
  }
  return {...cell,v:cell?.v??'',s:style};
}
const sheetPermissions = { rowH:'formatRows',autoH:'formatRows',hiddenRows:'formatRows',colW:'formatColumns',hiddenCols:'formatColumns',filters:'autoFilter',frows:'autoFilter',images:'objects',charts:'objects',scenarios:'scenarios',pivots:'pivotTables' };
const viewKeys = new Set(['cells','path','id','name','color','visibility','frR','frC','tabSelected','split','noGrid','showF']);

/** Check the whole immutable candidate, including other sheets changed by cut/replace.
 * No mutation reaches history when one part of the operation is forbidden. */
export function protectionViolation(before,after,action) {
  for(let i=0;i<before.sheets.length;i++){
    const old=before.sheets[i];if(!old.protected)continue;
    const next=after.sheets.find(s=>s===old||(old.path?s.path===old.path:s.name===old.name));
    if(!next||next===old)continue; // worksheet deletion/rename is governed by workbook protection
    if(action && !protectionAllows(old,action))return {sheet:old.name,permission:action};
    const pivotRanges=action==='pivotTables'?[...(old.pivots||[]),...(next.pivots||[])].map(p=>p.range).filter(Boolean):[];
    for(const ref of changedCellKeys(old.cells,next.cells)){
      const a=old.cells[ref],b=next.cells[ref],sa=a?.s||{},sb=b?.s||{};
      const pos=position(ref),inPivot=pivotRanges.some(range=>{const [a,b=a]=range.split(':').map(position),c=colIndex(pos.col);return +pos.row>=+a.row&&+pos.row<=+b.row&&c>=colIndex(a.col)&&c<=colIndex(b.col);});
      if((a?.v??'')!==(b?.v??'') || !!a?.literal!==!!b?.literal)
        if(cellLocked(old,ref)&&!inPivot)return {sheet:old.name,ref};
      const pa=cellProtection(old,ref),pb=cellProtection(next,ref);
      if(pa.locked!==pb.locked||pa.formulaHidden!==pb.formulaHidden)return {sheet:old.name,ref,permission:'protection'};
      for(const key of new Set([...Object.keys(sa),...Object.keys(sb)])){
        if(key==='locked'||key==='formulaHidden'||same(sa[key],sb[key]))continue;
        const permission=key==='note'?'objects':key==='link'?'insertHyperlinks':'formatCells';
        if(!(inPivot&&permission==='formatCells')&&action!=='sort'&&(!protectionAllows(old,permission)||(key==='link'&&cellLocked(old,ref))))return {sheet:old.name,ref,permission};
      }
    }
    for(const key of new Set([...Object.keys(old),...Object.keys(next)])){
      if(viewKeys.has(key)||same(old[key],next[key]))continue;
      if(key==='autoH'||key==='frows'&&action==='sort')continue; // calculated geometry/visibility follows an allowed edit
      const permission=sheetPermissions[key];
      if(!permission||!protectionAllows(old,permission))return {sheet:old.name,permission:permission||key};
    }
  }
  return null;
}
