// Example-based text extraction/combination. Candidate programs are small,
// deterministic and bounded; ambiguous results are returned for preview.
export function inferFlashFill(rows, examples) {
  if(!rows.length||rows.length>20000||rows.some(r=>r.length>16))throw Error('快速填充支持最多 20000 行、16 个源列');
  const samples=Object.entries(examples).map(([r,value])=>({r:+r,value:String(value)})).filter(s=>s.value!=='');
  if(!samples.length)throw Error('请先在结果列填写至少一个示例');
  if(samples.some(s=>!rows[s.r]||s.value.length>512)||rows.some(row=>row.some(v=>String(v??'').length>2048)))throw Error('快速填充示例或源文本过长');
  const data=rows.map(row=>row.map(v=>String(v??''))),first=samples[0],atoms=[],signatures=new Set();
  const add=(column,part,transform=x=>x)=>{
    const run=row=>{const value=part(row[column]||'');return value==null?null:transform(value);};
    const values=samples.map(s=>run(data[s.r]));if(values.some(v=>v==null))return;
    const initial=values[0];if(!initial||!first.value.includes(initial))return;
    const signature=JSON.stringify(data.map(run));if(signatures.has(signature))return;signatures.add(signature);
    atoms.push({run,initial});
  };
  const transforms=[x=>x,x=>x.toLowerCase(),x=>x.toUpperCase(),x=>x.replace(/\p{L}[\p{L}\p{M}]*/gu,s=>s[0].toUpperCase()+s.slice(1).toLowerCase())];
  for(let col=0;col<(data[0]?.length||0);col++){
    const parts=[s=>s,s=>s.trim()];
    for(const separator of [' ','-','_','.','@','/',',',':','\\'])for(const index of [0,1,2,-1,-2])parts.push(s=>{if(!s.includes(separator))return null;const p=s.split(separator).filter(Boolean);return p[index<0?p.length+index:index]??null;});
    for(const index of [0,1,2,-1,-2])parts.push(s=>{const p=s.match(/\p{L}+|\p{N}+/gu)||[];return p[index<0?p.length+index:index]??null;});
    for(let n=1;n<=Math.min(first.value.length,8);n++){parts.push(s=>[...s].length>=n?[...s].slice(0,n).join(''):null);parts.push(s=>[...s].length>=n?[...s].slice(-n).join(''):null);}
    for(const part of parts)for(const transform of transforms)add(col,part,transform);
  }
  const programs=[],seen=new Set();let attempts=0;
  const search=(offset,pieces,used)=>{
    if(++attempts>12000||programs.length>=100)return;
    if(offset===first.value.length){if(!used)return;const key=JSON.stringify(pieces.map(p=>typeof p==='string'?p:atoms.indexOf(p)));if(seen.has(key))return;seen.add(key);
      const run=row=>{let value='';for(const part of pieces){const v=typeof part==='string'?part:part.run(row);if(v==null)return null;value+=v;}return value;};
      if(samples.every(s=>run(data[s.r])===s.value))programs.push(run);return;
    }
    if(used>=3)return;
    for(const atom of atoms){const next=first.value.indexOf(atom.initial,offset);if(next<0||next-offset>64)continue;
      const prefix=first.value.slice(offset,next);search(next+atom.initial.length,pieces.concat(prefix?[prefix,atom]:[atom]),used+1);
    }
    // A suffix is useful for email domains and fixed labels, but only after at
    // least one input-derived part. A constant-only rule is never proposed.
    if(used&&first.value.length-offset<=64)search(first.value.length,pieces.concat(first.value.slice(offset)),used);
  };
  search(0,[],0);
  const results=[],outputs=new Set();
  for(const run of programs){const values=data.map(row=>row.every(v=>v==='')?null:run(row));const key=JSON.stringify(values);if(outputs.has(key))continue;outputs.add(key);results.push({values});if(results.length===12)break;}
  if(!results.length)throw Error('无法从示例识别一致规律，请补充或修改示例');
  return results;
}
