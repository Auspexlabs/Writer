import { Calc, A, parseA, fmt } from './sheet-engine.js';

// Word formula fields use the same parsed arithmetic as worksheets, with table-local references.
export function wordFormulaParts(instruction) {
  const m=/^\s*=\s*(.*?)\s*(?:\\#\s*"([^"\r\n]*)")?\s*(?:\\\*\s+MERGEFORMAT)?\s*$/i.exec(instruction||'');
  if(!m||!m[1]||m[1].length>4096)throw new Error('Invalid table formula');
  const expression=m[1].toUpperCase();
  if(/[^\p{L}\p{N}_.$,+\-*/^%=<>():;\s]/u.test(expression))throw new Error('Unsupported table formula');
  for(const f of expression.matchAll(/([A-Z_]+)\s*\(/g))if(!/^(SUM|AVERAGE|COUNT|MIN|MAX|PRODUCT|ABS|INT|MOD|ROUND|SIGN|IF|AND|OR|NOT|TRUE|FALSE|DEFINED)$/.test(f[1]))throw new Error('Unsupported table formula: '+f[1]);
  const format=m[2]||'';if(format&&!/^[#0,.%+\-; ()]+$/.test(format))throw new Error('Unsupported number format');
  return {expression,format};
}
export function calculateWordTable(rows,bookmarks={}) { return calculateWordTables([{rows}],bookmarks)[0]; }

/** Resolve all tables together so bookmarked references participate in normal
 * dependency calculation, including cycles and formulas in source tables. */
export function calculateWordTables(tables,bookmarks={}) {
  const aliases=new Map();tables.forEach((t,i)=>(t.names||[]).forEach(name=>aliases.set(name.toUpperCase(),i)));
  const numeric=text=>{const s=String(text??'').trim().replace(/[,，]/g,'');return s&&/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?%?$/.test(s)?String(parseFloat(s)/(s.endsWith('%')?100:1)):String(text??'');};
  const fields=tables.map(()=>[]),sheets=tables.map((t,si)=>({name:'WordTable'+si,cells:{}}));
  tables.forEach((table,si)=>{
    const rows=table.rows,named=new Map(Object.entries(table.bookmarks||bookmarks).map(([k,v])=>[k.toUpperCase(),v]));
    const compile=(instruction,r,c)=>{
      const {expression,format}=wordFormulaParts(instruction),pending=[];
      const token=value=>{const i=pending.push(value)-1;return '\uE000'+i+'\uE001';};
      const reference=(text,target=si,headers=false)=>{
        const tr=tables[target].rows,height=tr.length,width=tr.reduce((n,row)=>Math.max(n,row.length),0);
        const all=(r1,c1,r2,c2)=>{
          if(r1<0||c1<0||r2>=height||c2>=width)return '#REF!';
          const refs=[];for(let rr=r1;rr<=r2;rr++)for(let cc=c1;cc<=c2;cc++)if(tr[rr]?.[cc]&&(!headers||!tr[rr][cc].header)&&(target!==si||rr!==r||cc!==c))refs.push((target===si?'':"'WordTable"+target+"'!")+A(rr,cc));
          return refs.join(',')||'""';
        };
        const direction=text.match(/^(ABOVE|BELOW|LEFT|RIGHT)$/)?.[1];
        if(direction)return direction==='ABOVE'?r?all(0,c,r-1,c):'""':direction==='BELOW'?r+1<height?all(r+1,c,height-1,c):'""':direction==='LEFT'?c?all(r,0,r,c-1):'""':c+1<width?all(r,c+1,r,width-1):'""';
        // Explicit RnCn or Rn selects Word's R/C convention. A1 ranges such as
        // C1:C3 remain A1; this avoids changing existing imported formulas.
        const rc=/\bR(?:\d+C\d+|\d+|C)\b/.test(expression)||/^[RC]$/.test(text);
        const point=x=>{
          const m=/^R(\d+)C(\d+)$/.exec(x);if(m)return {r1:+m[1]-1,r2:+m[1]-1,c1:+m[2]-1,c2:+m[2]-1};
          const axis=/^([RC])(\d*)$/.exec(x);if(axis&&(rc||!axis[2])){const n=axis[2]?+axis[2]-1:axis[1]==='R'?r:c;return axis[1]==='R'?{r1:n,r2:n,c1:0,c2:width-1}:{r1:0,r2:height-1,c1:n,c2:n};}
          const p=parseA(x.replace(/\$/g,''));return p?{r1:p.r,r2:p.r,c1:p.c,c2:p.c}:null;
        };
        const [a,b=a]=text.replace(/\s/g,'').split(':').map(point);if(!a||!b)return '#REF!';
        return all(Math.min(a.r1,b.r1),Math.min(a.c1,b.c1),Math.max(a.r2,b.r2),Math.max(a.c2,b.c2));
      };
      const refPattern='(?:R[0-9]+C[0-9]+|\\$?[A-Z]+\\$?[0-9]+|[RC])';
      let formula=expression.replace(new RegExp('([\\p{L}_][\\p{L}\\p{N}_]*)\\s+('+refPattern+'(?:\\s*:\\s*'+refPattern+')?)(?![\\p{L}\\p{N}_])','gu'),(full,name,ref)=>aliases.has(name)?token(reference(ref,aliases.get(name))):full);
      formula=formula.replace(/\bDEFINED\(\s*([\p{L}_][\p{L}\p{N}_]*)\s*\)/gu,(_,name)=>named.has(name)?'1':'0');
      formula=formula.replace(/[\p{L}_][\p{L}\p{N}_]*/gu,(name,offset)=>{if(!named.has(name)||/^\s*\(/.test(formula.slice(offset+name.length)))return name;const v=numeric(named.get(name));return v!==''&&Number.isFinite(Number(v))?'('+Number(v)+')':'#VALUE!';});
      formula=formula.replace(new RegExp('(?<![\\p{L}\\p{N}_])('+refPattern+'(?:\\s*:\\s*'+refPattern+')?)(?![\\p{L}\\p{N}_]|\\s*\\()','gu'),ref=>token(reference(ref)));
      formula=formula.replace(/\b(ABOVE|BELOW|LEFT|RIGHT)\b(?!\s*\()/g,dir=>token(reference(dir,si,true)));
      formula=formula.replace(/\uE000(\d+)\uE001/g,(_,i)=>pending[+i]).replace(/;/g,',');
      return {formula,format};
    };
    rows.forEach((row,r)=>row.forEach((cell,c)=>{if(!cell)return;const key=A(r,c);sheets[si].cells[key]={v:numeric(cell.text)};if(cell.formula){const field={r,c,instruction:cell.formula};try{const p=compile(cell.formula,r,c);sheets[si].cells[key]={v:'='+p.formula};field.format=p.format;}catch(e){sheets[si].cells[key]={v:'=#VALUE!'};field.error=e.message;}fields[si].push(field);}}));
  });
  const calc=new Calc({sheets});
  return fields.map((list,si)=>list.map(field=>{let value=calc.value(si,field.r,field.c);if(typeof value==='boolean')value=+value;const error=field.error||value?.err;return {...field,value:typeof value==='number'?value:null,error:error||null,text:error?String(value?.err||'#VALUE!'):field.format?fmt(value,{code:field.format}):String(value??'')};}));
}

export function wordTableModel(table) {
  const rows=[],positions=new Map();
  [...table.rows].filter(row=>row.closest('table')===table).forEach((row,r)=>{
    const out=rows[r]||=[];let c=0;
    for(const cell of row.cells){while(c in out)c++;positions.set(cell,{r,c});const field=[...cell.querySelectorAll('[data-field]')].find(f=>f.closest('td,th')===cell&&/^\s*=/.test(f.getAttribute('data-field')));
      out[c]={text:cell.textContent||'',formula:field?.getAttribute('data-field'),header:row.getAttribute('data-w-header')==='true',element:cell,field};
      for(let rr=r;rr<r+Math.max(1,cell.rowSpan);rr++){rows[rr]||=[];for(let cc=c;cc<c+Math.max(1,cell.colSpan);cc++)if(rr!==r||cc!==c)rows[rr][cc]=null;}c+=Math.max(1,cell.colSpan);
    }
  });
  return {rows,positions};
}

export function wordFormulaTables(root) {
 const tables=[...root.querySelectorAll('table')].map(element=>({element,...wordTableModel(element),names:(element.getAttribute('data-w-bookmark')||'').split(/[,;\s]+/).filter(Boolean)}));
 const ends=new Map([...root.querySelectorAll('[data-bookmark-end]')].map(el=>[el.dataset.bookmarkEnd,el]));
 for(const start of root.querySelectorAll('[data-bookmark-start]')){
   const end=ends.get(start.dataset.bookmarkStart);if(!end||!(start.compareDocumentPosition(end)&4))continue;
   const range=root.ownerDocument.createRange();range.setStartAfter(start);range.setEndBefore(end);
   for(const table of tables){const textRange=root.ownerDocument.createRange();textRange.selectNodeContents(table.element);
     // A bookmark encloses the table, or its boundary anchors are within the
     // first/last cells and its selected text covers the entire table.
     if(range.intersectsNode(table.element)&&((start.compareDocumentPosition(table.element)&4)&&(table.element.compareDocumentPosition(end)&4)||start.closest('table')===table.element&&end.closest('table')===table.element&&range.toString()===textRange.toString()))table.names.push(start.dataset.bookmarkStart);
   }
 }
 return tables;
}

/** Bookmark spans are empty anchors, so Range extracts exactly the selected visible text. */
export function wordFormulaBookmarks(root) {
 const result={},ends=new Map([...root.querySelectorAll('[data-bookmark-end]')].map(el=>[el.dataset.bookmarkEnd,el]));
 for(const start of root.querySelectorAll('[data-bookmark-start]')){const end=ends.get(start.dataset.bookmarkStart);if(!end||!(start.compareDocumentPosition(end)&4))continue;const range=root.ownerDocument.createRange();range.setStartAfter(start);range.setEndBefore(end);const fragment=range.cloneContents();fragment.querySelectorAll('del,[data-office-object],[data-eq],[data-cite]').forEach(el=>el.remove());result[start.dataset.bookmarkStart]=fragment.textContent;}
 for(const block of root.querySelectorAll('[data-w-bookmark]'))for(const name of block.getAttribute('data-w-bookmark').split(/[,;\s]+/).filter(Boolean))if(!(name in result))result[name]=block.textContent;
 return result;
}
