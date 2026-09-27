function runsOf(html) {
  const root = new DOMParser().parseFromString(html || '', 'text/html').body, runs = [];
  let text = '';
  const walk = n => {
    if (n.nodeType === 3) { runs.push({node:n,start:text.length,end:text.length+n.nodeValue.length}); text += n.nodeValue; return; }
    if (/^(SCRIPT|STYLE)$/.test(n.tagName)) return;
    if (n.tagName === 'BR') { text += '\n'; return; }
    for (const c of n.childNodes) walk(c);
    if (/^(P|DIV|LI)$/.test(n.tagName)) text += '\n';
  };
  walk(root); return {root,runs,text};
}
function matches(text, query, options = {}) {
  if (!query) return [];
  const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.caseSensitive ? 'gu' : 'giu');
  const word = s => /[\p{L}\p{N}_]/u.test(s || '');
  return [...text.matchAll(re)].filter(m => !options.wholeWord || (!word(text[m.index-1]) && !word(text[m.index+m[0].length]))).map(m=>({index:m.index,length:m[0].length}));
}
export function findSlides(doc, query, options) {
  const hits = [];
  const walk = (o, slide, rootId) => {
    if ((!o.field || o.field === 'footer') && (o.t==='text' || o.t==='shape')) {
      const {text} = runsOf(o.html); for (const m of matches(text,query,options)) hits.push({...m,slide,id:o.id,rootId,text});
    }
    if (o.t==='table') (o.rows || []).forEach((row,r)=>row.forEach((text,c)=>{ for(const m of matches(String(text),query,options)) hits.push({...m,slide,id:o.id,rootId,r,c,text}); }));
    for (const k of o.kids || []) walk(k,slide,rootId);
  };
  doc.slides.forEach((s,i)=>s.objs.forEach(o=>walk(o,i,o.id))); return hits;
}
export function replaceSlideHits(doc, hits, replacement) {
  const find = (objs,id) => { for(const o of objs) { if(o.id===id)return o; const child=find(o.kids||[],id); if(child)return child; } };
  const buckets = new Map();
  for (const hit of hits) { const key=[hit.slide,hit.id,hit.r,hit.c].join(':'); if(!buckets.has(key))buckets.set(key,[]); buckets.get(key).push(hit); }
  for (const list of buckets.values()) {
    list.sort((a,b)=>b.index-a.index); const hit=list[0],o=find(doc.slides[hit.slide].objs,hit.id); if(!o)continue;
    if(hit.r!=null) { let text=String(o.rows[hit.r][hit.c]); for(const h of list)text=text.slice(0,h.index)+replacement+text.slice(h.index+h.length); o.rows[hit.r][hit.c]=text; continue; }
    const {root,runs}=runsOf(o.html);
    for (const h of list) {
      let inserted=false;
      for(const run of runs) { const a=Math.max(run.start,h.index),b=Math.min(run.end,h.index+h.length); if(a>=b)continue;
        run.node.nodeValue=run.node.nodeValue.slice(0,a-run.start)+(inserted?'':replacement)+run.node.nodeValue.slice(b-run.start); inserted=true;
      }
    }
    o.html=root.innerHTML;
  }
  return hits.length;
}
