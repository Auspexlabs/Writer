// Lay out explicit paragraph tab stops. The save bridge unwraps wd-tab spans back to tab characters.
const pixels=value=>{const m=/^(-?[\d.]+)(cm|mm|in|pt|px)?$/.exec(value);return m?+m[1]*({cm:96/2.54,mm:96/25.4,in:96,pt:96/72,px:1}[m[2]||'cm']):NaN;};
export function tabAdvance(current,stops,afterWidth,decimalWidth=afterWidth){
 for(const stop of stops){if(stop.at<=current+.1)continue;const offset=stop.kind==='right'?afterWidth:stop.kind==='center'?afterWidth/2:stop.kind==='decimal'?decimalWidth:0,width=stop.at-current-offset;if(width>.1)return width;}
 return (Math.floor((current+.1)/48)+1)*48-current;
}
export function layoutTabs(root){
 const tabs=Array.from(root.querySelectorAll('.wd-tab'));if(!tabs.length)return;
 const canvas=root.ownerDocument.createElement('canvas'),ctx=canvas.getContext('2d');if(!ctx)return;
 const groups=new Map();for(const tab of tabs){const p=tab.closest('p,h1,h2,h3,h4,h5,h6,li,td,th');if(!p)continue;if(!groups.has(p))groups.set(p,[]);groups.get(p).push(tab);}
 for(const [p,items] of groups){const raw=p.getAttribute('data-w-tabs')||'',stops=raw.split(/[,;]/).map(x=>x.trim().split(/\s+/)).map(([kind,pos])=>({kind,at:pixels(pos||'')})).filter(x=>Number.isFinite(x.at)&&x.kind!=='clear').sort((a,b)=>a.at-b.at);
  if(!stops.length){for(const tab of items)tab.removeAttribute('style');continue;}
  for(const tab of items){Object.assign(tab.style,{display:'inline-block',width:'0px',overflow:'hidden',verticalAlign:'baseline',whiteSpace:'pre'});}
  const rect=p.getBoundingClientRect(),scale=rect.width/p.offsetWidth||1,style=getComputedStyle(p),origin=rect.left+(parseFloat(style.paddingLeft)||0)*scale;
  for(let i=0;i<items.length;i++){const tab=items[i],next=items[i+1];let after='',width=0,decimal=null,started=false;
   const walker=root.ownerDocument.createTreeWalker(p,4);let n;
   while((n=walker.nextNode())){if(tab.contains(n)){started=true;continue;}if(!started)continue;if(next?.contains(n))break;const text=n.nodeValue.split('\n')[0];ctx.font=getComputedStyle(n.parentElement).font;const dot=text.indexOf('.');if(decimal===null&&dot>=0)decimal=width+ctx.measureText(text.slice(0,dot)).width;width+=ctx.measureText(text).width;after+=text;if(n.nodeValue.includes('\n'))break;}
   const current=(tab.getBoundingClientRect().left-origin)/scale,advance=tabAdvance(current,stops,width,decimal??width);tab.style.width=Math.max(0,advance)+'px';
  }
 }
}
