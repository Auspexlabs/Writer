// Text runs stay separate through wrapping and Canvas painting, preserving mixed formatting.
const decode=s=>s.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|nbsp|apos);/gi,(m,k)=>k[0]==='#'?String.fromCodePoint(Math.min(0x10ffff,parseInt(k.slice(k[1]?.toLowerCase()==='x'?2:1),k[1]?.toLowerCase()==='x'?16:10)||0)):({amp:'&',lt:'<',gt:'>',quot:'"',nbsp:'\u00a0',apos:"'"}[k.toLowerCase()]||m));
export function richParagraphs(html,base={}){
 const out=[], stack=[{tag:'',style:base}], lists=[];let p=null;
 const paragraph=()=>p||(p={runs:[],lvl:0,bullet:'',align:base.align});
 const finish=()=>{if(p){out.push(p);p=null;}};
 for(const token of String(html||'').match(/<[^>]*>|[^<]+/g)||[]){
  if(token[0]!=='<'){const text=decode(token);if(text)paragraph().runs.push({text,...stack.at(-1).style});continue;}
  const m=/^<\s*(\/?)\s*([\w]+)([^>]*)>/.exec(token);if(!m)continue;const [,close,name,attrs]=m, tag=name.toLowerCase();
  if(close){if(/^(p|div|li|h[1-6])$/.test(tag))finish();if(tag==='ul'||tag==='ol')lists.pop();const i=stack.map(x=>x.tag).lastIndexOf(tag);if(i>0)stack.splice(i);continue;}
  if(tag==='br'){paragraph().runs.push({text:'\n',...stack.at(-1).style});continue;}
  if(tag==='ul'||tag==='ol'){finish();lists.push({tag,n:Number(/\bstart=["'](-?\d+)["']/.exec(attrs)?.[1]||1)-1});}
  const style={...stack.at(-1).style}, css={};const raw=/\bstyle\s*=\s*(["'])(.*?)\1/i.exec(attrs)?.[2]||'';
  for(const field of raw.split(';')){const i=field.indexOf(':');if(i>0)css[field.slice(0,i).trim().toLowerCase()]=field.slice(i+1).trim();}
  if(tag==='b'||tag==='strong')style.bold=true;if(tag==='i'||tag==='em')style.italic=true;if(tag==='u')style.underline=true;if(['s','del','strike'].includes(tag))style.strike=true;
  if(css['font-weight'])style.bold=css['font-weight']==='bold'||parseInt(css['font-weight'])>=600;
  if(css['font-style'])style.italic=css['font-style']==='italic';if(css.color)style.color=css.color;
  if(css['font-family'])style.font=css['font-family'];
  if(css['font-size']){const size=css['font-size'],n=parseFloat(size);if(Number.isFinite(n))style.fs=/pt$/.test(size)?n*4/3:/em$/.test(size)?n*(style.fs||32):/%$/.test(size)?n*(style.fs||32)/100:n;}
  if(css['text-decoration']){style.underline=css['text-decoration'].includes('underline');style.strike=css['text-decoration'].includes('line-through');}
  if(tag==='sup'||tag==='sub'){style.shift=tag==='sup'?-.35:.2;style.fs=(style.fs||32)*.7;}
  if(/^(p|div|li|h[1-6])$/.test(tag)){finish();paragraph().align=css['text-align']||base.align;p.lvl=+( /data-lvl=["'](\d+)["']/.exec(attrs)?.[1]||0);if(tag==='li'){const l=lists.at(-1),start=/\b(?:data-start|value)=["'](-?\d+)["']/.exec(attrs);if(l&&start)l.n=Number(start[1])-1;const fallback=l?.tag==='ol'?++l.n+'.':'•',marker=/\bdata-marker=(["'])(.*?)\1/.exec(attrs);p.bullet=marker?decode(marker[2]):fallback;}}
  if(!/\/$/.test(attrs)&&!['img','hr','input'].includes(tag))stack.push({tag,style});
 }
 finish();return out;
}
export function paintText(g,o,paras,box,base){
 const font=s=>`${s.italic?'italic ':''}${s.bold?700:400} ${s.fs||32}px ${s.font||'sans-serif'}`;
 const width=(text,s)=>{g.font=font(s);return g.measureText(text).width;}, lines=[];
 for(const p of paras){let first=true,line;const newLine=()=>({runs:[],w:0,h:(base.fs||32)*(o.lh||1.35),ind:(p.lvl||0)*(base.fs||32)*1.6+(p.bullet?(base.fs||32)*1.1:0),bullet:first?p.bullet:'',align:p.align||base.align});line=newLine();
  const flush=()=>{lines.push(line);first=false;line=newLine();};
  for(const run of p.runs||[{text:p.text,...base}]){const s={...base,...run};for(const part of run.text.split(/(\n)/)){
   if(part==='\n'){flush();continue;}
   for(const token of part.match(/[\u3000-\u9fff\uff00-\uffef]|[^\s\u3000-\u9fff\uff00-\uffef]+|\s+/g)||[]){
    const max=Math.max(1,box.w-line.ind),pieces=width(token,s)>max?Array.from(token):[token];
    for(const t of pieces){const w=width(t,s);if(line.runs.length&&line.w+w>max)flush();if(!line.runs.length&&/^ +$/.test(t))continue;const previous=line.runs.at(-1);if(previous&&JSON.stringify(previous.s)===JSON.stringify(s)){previous.text+=t;previous.w+=w;}else line.runs.push({text:t,s,w});line.w+=w;line.h=Math.max(line.h,(s.fs||32)*(o.lh||1.35));}
   }
  }}flush();
 }
 const total=lines.reduce((n,l)=>n+l.h,0);let y=box.y+(o.va==='middle'?(box.h-total)/2:o.va==='bottom'?box.h-total:0);
 g.textBaseline='middle';g.textAlign='left';
 for(const l of lines){const baseline=y+l.h/2;let x=box.x+l.ind+(l.align==='center'?(box.w-l.ind-l.w)/2:l.align==='right'?box.w-l.ind-l.w:0);
  if(l.bullet){g.font=font(base);g.fillStyle=base.color;g.fillText(l.bullet,box.x+l.ind-(base.fs||32)*1.1,baseline);}
  for(const run of l.runs){const {s,w,text}=run;g.font=font(s);g.fillStyle=s.color;const yy=baseline+(s.shift||0)*(base.fs||32);g.fillText(text,x,yy);
   for(const offset of [s.underline?(s.fs||32)*.45:null,s.strike?0:null])if(offset!==null){g.beginPath();g.strokeStyle=s.color;g.lineWidth=Math.max(1,(s.fs||32)/18);g.moveTo(x,yy+offset);g.lineTo(x+w,yy+offset);g.stroke();}x+=w;
  }y+=l.h;
 }
}
