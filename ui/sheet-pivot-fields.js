// A pivot calculation runs on each group's source-field sums, including totals.
// It is deliberately independent of worksheet addresses and workbook names.
export function pivotFormulaSource(formula, fields) {
  const names=new Map(fields.map((name,i)=>[String(name).toLocaleLowerCase(),i]));
  const ref=i=>{let n=i+1,s='';while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26);}return s+'1';};
  const deps=new Set(),source=String(formula).trim().replace(/^=/,'');let out='',at=0;
  if(!source||source.length>8192)throw new Error('Enter a pivot field formula of at most 8192 characters');
  while(at<source.length){
    const tail=source.slice(at);let m;
    if((m=/^\s+/.exec(tail))||(m=/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(tail))||(m=/^"(?:[^"]|"")*"/.exec(tail))||(m=/^(?:<=|>=|<>|[-+*/^&=<>%(),])/.exec(tail))){out+=m[0];at+=m[0].length;continue;}
    const quoted=/^'((?:[^']|'')+)'/.exec(tail),bare=/^[\p{L}_][\p{L}\p{N}_.]*/u.exec(tail);m=quoted||bare;
    if(!m)throw new Error('Pivot formulas must use source field names, not cell or range references');
    const name=quoted?m[1].replace(/''/g,"'"):m[0],next=source.slice(at+m[0].length);
    if(!quoted&&/^\s*\(/.test(next)){
      if(!/^(?:IF|IFERROR|AND|OR|NOT|ABS|ROUND|ROUNDUP|ROUNDDOWN|INT|MOD|POWER|SQRT|SUM|MIN|MAX|AVERAGE|COUNT|SIGN|EXP|LN|LOG|LOG10)$/i.test(name))throw new Error('Unsupported pivot formula function: '+name);
      out+=name;
    }else if(!quoted&&/^(?:TRUE|FALSE)$/i.test(name))out+=name;
    else {const i=names.get(name.toLocaleLowerCase());if(i==null)throw new Error('Unknown pivot field: '+name);deps.add(i);out+=ref(i);}
    at+=m[0].length;
  }
  return {source:out,fields:[...deps]};
}

export function pivotGroupValue(groups, field, value) {
  const g=(groups||[]).find(g=>g.field===field);
  const item=g?.items.find(g=>g.items.some(v=>typeof v===typeof value&&JSON.stringify(v)===JSON.stringify(value)));
  return item?item.name:value;
}

export function validatePivotGroups(groups, axes) {
  const seen=new Set();
  for(const group of groups||[]){
    if(!axes.includes(group.field)||seen.has(group.field)||!Array.isArray(group.items))throw new Error('Choose one grouping per row or column field');
    seen.add(group.field);const items=new Set(),names=new Set();
    for(const item of group.items){
      if(typeof item.name!=='string'||!item.name.trim()||item.name.length>255||names.has(item.name.toLocaleLowerCase())||!Array.isArray(item.items)||!item.items.length)throw new Error('Group names must be unique and contain at least one item');
      names.add(item.name.toLocaleLowerCase());
      for(const value of item.items){const key=JSON.stringify(value);if(value==null||!['string','number','boolean'].includes(typeof value)||typeof value==='number'&&!Number.isFinite(value)||items.has(key))throw new Error('Each source item may belong to only one group');items.add(key);}
    }
  }
}
