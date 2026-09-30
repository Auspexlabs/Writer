/** Measure rendered text, keeping numbering out of the editable document and its offsets. */
export function lineNumberPositions(ed,{columns=1,gap=32}={}) {
  const box=ed.getBoundingClientRect(),zoom=box.width/(ed.offsetWidth||box.width||1)||1,width=ed.offsetWidth||box.width/zoom,colWidth=(width-gap*(columns-1))/columns,step=colWidth+gap;
  const out=[];let number=0;
  for(const block of ed.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre')) {
    if(block.closest('table,[data-toc],[data-bib],[data-office-object]')||block.querySelector('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre'))continue;
    const walker=ed.ownerDocument.createTreeWalker(block,4),rects=[];let node;
    while((node=walker.nextNode())){if(!node.nodeValue||node.parentElement.closest('[data-office-object],[data-eq],[data-shape],[data-cite],sup[data-fn],del'))continue;const range=ed.ownerDocument.createRange();range.selectNodeContents(node);for(const r of range.getClientRects())if(r.width>.1&&r.height>.1)rects.push({top:(r.top-box.top)/zoom,bottom:(r.bottom-box.top)/zoom,col:Math.max(0,Math.floor(((r.left-box.left)/zoom+.5)/step))});}
    if(!rects.length&&!block.textContent.trim()){const r=block.getBoundingClientRect();rects.push({top:(r.top-box.top)/zoom,bottom:(r.bottom-box.top)/zoom,col:Math.max(0,Math.floor(((r.left-box.left)/zoom+.5)/step))});}
    rects.sort((a,b)=>a.col-b.col||a.top-b.top);const lines=[];
    for(const r of rects){const last=lines.at(-1);if(last&&last.col===r.col&&Math.min(last.bottom,r.bottom)-Math.max(last.top,r.top)>Math.min(last.bottom-last.top,r.bottom-r.top)/2){last.top=Math.min(last.top,r.top);last.bottom=Math.max(last.bottom,r.bottom);}else lines.push({...r});}
    for(const line of lines)out.push({number:++number,x:line.col*step-42,y:(line.top+line.bottom)/2});
  }
  return out;
}
export function lineNumbersHtml(lines) {return lines.map(line=>`<span data-line-number="${line.number}" style="position:absolute;left:${line.x.toFixed(2)}px;top:${line.y.toFixed(2)}px;transform:translateY(-50%);width:30px;text-align:right;font:10px Arial,sans-serif;color:#6e6e73">${line.number}</span>`).join('');}
