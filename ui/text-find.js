import { wildcardMatches } from './word-wildcards.js';
// Find across formatting runs while preserving document structure and atomic fields/objects.
export function editableText(root) {
  const runs = []; let text = '';
  const walk = n => {
    if (n.nodeType === 3) { runs.push({ node: n, start: text.length, end: text.length + n.nodeValue.length }); text += n.nodeValue; return; }
    if (n.nodeType !== 1) return;
    if (n.matches('script,style,del,[contenteditable="false"],[data-toc],[data-bib]')) { text += '\ufffc'; return; }
    if (n.tagName === 'BR') { text += '\n'; return; }
    for (const c of n.childNodes) walk(c);
    if (n !== root && /^(P|DIV|LI|H[1-6]|BLOCKQUOTE|PRE|TD|TH)$/.test(n.tagName)) text += '\n';
  };
  walk(root); return { runs, text };
}
export function textMatches(text, query, options = {}) {
  if (!query) return [];
  const re = options.wildcards ? null : new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.caseSensitive ? 'gu' : 'giu');
  const word = s => /[\p{L}\p{N}_]/u.test(s || '');
  const at = i => i < 0 || i >= text.length ? '' : String.fromCodePoint(text.codePointAt(i));
  const before = i => { const c = text.charCodeAt(i - 1); return at(c >= 0xdc00 && c <= 0xdfff ? i - 2 : i - 1); };
  if(options.wildcards)return wildcardMatches(text,query,options.caseSensitive).filter(m=>!options.wholeWord||(!word(before(m.index))&&!word(at(m.index+m.length))));
  return [...text.matchAll(re)].filter(m => m[0].length && !m[0].includes('\ufffc') && (!options.wholeWord || (!word(before(m.index)) && !word(at(m.index + m[0].length)))))
    .map(m => ({ index: m.index, length: m[0].length, ...(options.wildcards ? { groups:m.slice(1), match:m[0] } : {}) }));
}
export function matchesTextFormat(node,format={}) {
  const el=node.parentElement||node.parentNode,view=node.ownerDocument?.defaultView,css=view?.getComputedStyle?.(el)||el?.style||{};
  const has=selector=>!!el?.closest?.(selector);
  const values={bold:css.fontWeight?css.fontWeight==='bold'||+css.fontWeight>=600:has('b,strong'),italic:css.fontStyle?css.fontStyle==='italic':has('i,em'),underline:String(css.textDecorationLine||css.textDecoration||'').includes('underline')||has('u')};
  for(const k of ['bold','italic','underline'])if(typeof format[k]==='boolean'&&values[k]!==format[k])return false;
  if(format.font&&!String(css.fontFamily||'').toLowerCase().split(',').some(s=>s.trim().replace(/^['"]|['"]$/g,'')===format.font.toLowerCase()))return false;
  if(format.size&&(!Number.isFinite(parseFloat(css.fontSize))||Math.abs(parseFloat(css.fontSize)*.75-format.size)>.1))return false;
  return true;
}
export function findTextRanges(root, query, options = {}) {
  const { runs, text } = editableText(root);
  const hits=query?textMatches(text,query,options):options.format&&Object.keys(options.format).length?runs.filter(r=>r.end>r.start&&matchesTextFormat(r.node,options.format)).map(r=>({index:r.start,length:r.end-r.start})):[];
  let cursor=0;
  const formats=new Map(),formatted=r=>{if(!formats.has(r.node))formats.set(r.node,matchesTextFormat(r.node,options.format));return formats.get(r.node);};
  return hits.map(hit => {
    while(cursor<runs.length&&runs[cursor].end<=hit.index)cursor++;
    const touched=[];for(let j=cursor;j<runs.length&&runs[j].start<hit.index+hit.length;j++)touched.push(runs[j]);
    if (!touched.length || (options.format && touched.some(r=>!formatted(r)))) return null;
    const first = touched[0], last = touched.at(-1), range = root.ownerDocument.createRange();
    range.setStart(first.node, Math.max(0, hit.index - first.start)); range.setEnd(last.node, Math.min(last.end, hit.index + hit.length) - last.start);
    return { ...hit, range, runs: touched.map(r => ({ node: r.node, start: Math.max(0, hit.index - r.start), end: Math.min(r.end, hit.index + hit.length) - r.start })) };
  }).filter(Boolean);
}
export function replacementText(hit,replacement) {return hit.groups?String(replacement).replace(/\\([1-9])/g,(_,n)=>hit.groups[n-1]??''):replacement;}
function formatRun(run,format) {
  if(run.end<=run.start)return null;
  const node=run.node,doc=node.ownerDocument,middle=node.splitText(run.start);middle.splitText(run.end-run.start);
  const span=doc.createElement('span');middle.parentNode.insertBefore(span,middle);span.appendChild(middle);
  // A CSS underline propagates through descendants. Split its inline ancestors so
  // removing it from this match leaves the neighbouring text and link targets intact.
  if(format.underline===false){let child=span;while(child.parentElement&&!/^(P|DIV|LI|H[1-6]|TD|TH|BLOCKQUOTE|PRE)$/.test(child.parentElement.tagName)){
    const parent=child.parentElement;if(parent.getAttribute('contenteditable')==='true')break;
    const before=parent.cloneNode(false),after=parent.cloneNode(false);before.removeAttribute('id');after.removeAttribute('id');
    while(parent.firstChild!==child)before.appendChild(parent.firstChild);while(child.nextSibling)after.appendChild(child.nextSibling);
    if(before.childNodes.length)parent.before(before);if(after.childNodes.length)parent.after(after);
    let clean=parent;if(parent.tagName==='U'){clean=doc.createElement('span');for(const a of parent.attributes)clean.setAttribute(a.name,a.value);while(parent.firstChild)clean.appendChild(parent.firstChild);parent.replaceWith(clean);}
    if((clean.style.textDecoration||'').includes('underline')||(doc.defaultView?.getComputedStyle(clean).textDecorationLine||'').includes('underline'))clean.style.textDecoration=(clean.style.textDecoration||doc.defaultView?.getComputedStyle(clean).textDecorationLine||'').replace(/underline/g,'').trim()||'none';
    child=clean;
  }}
  if(typeof format.bold==='boolean')span.style.fontWeight=format.bold?'bold':'normal';
  if(typeof format.italic==='boolean')span.style.fontStyle=format.italic?'italic':'normal';
  if(typeof format.underline==='boolean')span.style.textDecoration=format.underline?'underline':'none';
  if(format.font)span.style.fontFamily=format.font;if(format.size>0)span.style.fontSize=format.size+'pt';
  if(format.color)span.style.color=format.color;if(format.highlight)span.style.backgroundColor=format.highlight;
  return middle;
}
export function replaceTextRanges(hits, replacement, options={}) {
  const format=options.format||{},styled=Object.keys(format).length>0;
  for(const hit of hits.slice().reverse()){
    if(options.keepText){if(styled)for(const run of hit.runs.slice().reverse()){const node=formatRun(run,format);if(node&&!hit.replacementEnd)hit.replacementEnd={node,offset:node.length};}continue;}
    const text=replacementText(hit,replacement);
    hit.runs.forEach((r,i)=>{r.node.nodeValue=r.node.nodeValue.slice(0,r.start)+(i===0?text:'')+r.node.nodeValue.slice(r.end);});
    const first=hit.runs[0],node=styled&&text.length?formatRun({...first,end:first.start+text.length},format):first.node;
    hit.replacementEnd={node,offset:node===first.node?first.start+text.length:node.length};
  }
  return hits.length;
}
