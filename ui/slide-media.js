import { bindCaptions } from './slide-captions.js';
/** PPT trimEnd is seconds removed from the end, not the final playback timestamp. */
export function mediaBounds(options={},duration=Infinity) {
 const start=Math.max(0,Number(options.trimStart)||0),end=Number.isFinite(duration)?Math.max(start,duration-Math.max(0,Number(options.trimEnd)||0)):Infinity;
 return {start,end};
}
export function mediaVolume(options={},time=0,duration=Infinity) {
 const {start,end}=mediaBounds(options,duration),fi=Number(options.fadeIn)||0,fo=Number(options.fadeOut)||0;
 return Math.min(1,Math.max(0,(Number(options.volume??100)/100)))*Math.min(fi>0?Math.max(0,Math.min(1,(time-start)/fi)):1,fo>0?Math.max(0,Math.min(1,(end-time)/fo)):1);
}
const bindings=new WeakMap();
export function bindMediaElement(el,options={},autoplay=false,captionRoot=el.parentElement) {
 const signature=JSON.stringify([el.getAttribute('src'),options,autoplay]),prev=bindings.get(el);if(prev?.signature===signature)return prev.dispose;
 prev?.dispose();let frame=0,ended=false,disposed=false;const disposeCaptions=bindCaptions(el,options.captions,captionRoot);
 const bounds=()=>mediaBounds(options,el.duration),play=()=>{const p=el.play();if(p?.catch)p.catch(()=>{el.controls=true;});};
 const update=()=>{const {start,end}=bounds();if(!(end>start)){el.pause();return;}if(!el.seeking&&el.currentTime<start-.02)el.currentTime=start;
  if(el.currentTime>=end-.01&&!el.paused){if(options.loop){el.currentTime=start;play();}else{ended=true;el.pause();el.currentTime=end;}}
  el.volume=mediaVolume(options,el.currentTime,el.duration);el.muted=!!options.muted;
  if(autoplay&&options.hideStopped)el.style.visibility=el.paused?'hidden':'visible';
 };
 const pump=()=>{update();if(!el.paused&&!disposed)frame=requestAnimationFrame(pump);};
 const start=()=>{if(ended||el.currentTime>=bounds().end-.01){ended=false;el.currentTime=bounds().start;}cancelAnimationFrame(frame);pump();};
 const ready=()=>{const {start,end}=bounds();if(start>0&&end>start)el.currentTime=start;update();if(autoplay&&options.autoplay&&end>start)play();};
 const finish=()=>{if(options.loop){el.currentTime=bounds().start;play();}else update();};
 const events={loadedmetadata:ready,play:start,pause:update,ended:finish,timeupdate:update,seeking:update};for(const [event,fn]of Object.entries(events))el.addEventListener(event,fn);
 const dispose=()=>{disposed=true;cancelAnimationFrame(frame);disposeCaptions();for(const [event,fn]of Object.entries(events))el.removeEventListener(event,fn);el.pause();bindings.delete(el);};
 bindings.set(el,{signature,dispose});if(el.readyState>=1)ready();return dispose;
}
export function bindSlideMedia(root,objects,autoplay=false,token='') {
 if(!root)return;const map=new Map();const visit=objs=>{for(const o of objs||[]){if(o.mediaType)map.set(o.id,o);visit(o.kids);}};visit(objects);
 for(const el of root.querySelectorAll('audio[data-media-id],video[data-media-id]')){const o=map.get(el.dataset.mediaId);if(o)bindMediaElement(el,{...o.playback,captions:o.captions,_epoch:token},autoplay,root);}
}
export function disposeSlideMedia(root){for(const el of root?.querySelectorAll('audio[data-media-id],video[data-media-id]')||[])bindings.get(el)?.dispose();}

export function editMediaPlayback(object,duration=Infinity) {return new Promise(resolve=>{
 const dialog=document.createElement('dialog');dialog.dataset.mediaEditor='1';dialog.style.cssText='width:610px;max-width:94vw;max-height:90vh;overflow:auto;border:1px solid #bbb;border-radius:16px;padding:24px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
 const title=document.createElement('h3');title.textContent='媒体播放与剪辑';dialog.append(title);const p=object.playback||{},controls=document.createElement('div'),inputs={};controls.style.cssText='display:grid;grid-template-columns:1fr 1fr;gap:12px';dialog.append(controls);
 for(const [key,label,defaultValue,type]of [['trimStart','裁去开头（秒）',0],['trimEnd','裁去结尾（秒）',0],['fadeIn','淡入（秒）',0],['fadeOut','淡出（秒）',0],['volume','音量（0–100）',100],['slideCount','连续播放的幻灯片数',1],['autoplay','进入幻灯片自动播放',false,'checkbox'],['loop','循环播放',false,'checkbox'],['muted','静音',false,'checkbox'],['hideStopped','停止时隐藏',false,'checkbox']]){const labelEl=document.createElement('label');labelEl.textContent=label;labelEl.style.cssText='display:grid;gap:6px';const input=document.createElement('input');input.name=key;input.type=type||'number';input.min='0';input.step=key==='slideCount'?'1':'.1';input.value=p[key]??defaultValue;input.checked=!!p[key];labelEl.append(input);controls.append(labelEl);inputs[key]=input;}
 const label=document.createElement('p');label.textContent='书签：每行“秒数 | 名称”';const marks=document.createElement('textarea');marks.name='bookmarks';marks.rows=4;marks.style.width='100%';marks.value=(p.bookmarks||[]).map(b=>b.time+' | '+b.name).join('\n');dialog.append(label,marks);
 const error=document.createElement('p');error.setAttribute('role','alert');error.style.color='#b3261e';dialog.append(error);const actions=document.createElement('div');actions.style.cssText='display:flex;justify-content:flex-end;gap:12px';dialog.append(actions);let ended=false;
 const done=v=>{if(ended)return;ended=true;dialog.close();dialog.remove();resolve(v);};
 for(const [text,fn]of [['取消',()=>done(null)],['应用播放设置',()=>{try{const out=Object.fromEntries(Object.entries(inputs).map(([k,input])=>[k,input.type==='checkbox'?input.checked:Number(input.value)]));if(Object.values(out).some(v=>typeof v==='number'&&(!Number.isFinite(v)||v<0))||out.volume>100||!Number.isInteger(out.slideCount)||out.slideCount<1||out.slideCount>999)throw Error('请输入有效的时间、音量和幻灯片数');
 const {start,end}=mediaBounds(out,duration);if(end<=start)throw Error('剪辑后必须保留至少一段媒体');if(Number.isFinite(end)&&out.fadeIn+out.fadeOut>end-start)throw Error('淡入与淡出的总时长不能超过剪辑后的媒体长度');
 out.bookmarks=marks.value.split('\n').filter(l=>l.trim()).map(line=>{const pos=line.indexOf('|'),time=Number(line.slice(0,pos).trim()),name=line.slice(pos+1).trim();if(pos<0||!name||!Number.isFinite(time)||time<0||time>duration)throw Error('书签格式应为“秒数 | 名称”，位置须在媒体内');return {name,time};});if(new Set(out.bookmarks.map(b=>b.name)).size!==out.bookmarks.length)throw Error('书签名称不能重复');done(out);
 }catch(e){error.textContent=e.message;}}]]){const b=document.createElement('button');b.textContent=text;b.onclick=fn;actions.append(b);}dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});document.body.append(dialog);dialog.showModal();
});}
const carries=new WeakMap();
/** Audio can continue across forward slide changes; leaving its span or ending the show releases it. */
export function carrySlideMedia(root,objects,from,to) {
 if(!root)return;let active=carries.get(root)||[];active=active.filter(item=>{if(to<=from||to>=item.until){item.dispose();item.el.remove();return false;}return true;});
 const map=new Map();const visit=objs=>{for(const o of objs||[]){map.set(o.id,o);visit(o.kids);}};visit(objects);
 for(const player of root.querySelectorAll('audio[data-media-id]')){const o=map.get(player.dataset.mediaId),count=Number(o?.playback?.slideCount)||1;if(player.paused||player.ended||count<=1||to<=from||to>=from+count)continue;
  const el=document.createElement('audio');el.src=player.src;el.preload='auto';el.style.display='none';root.append(el);const time=player.currentTime;player.pause();
  const dispose=bindMediaElement(el,{...o.playback,captions:o.captions,autoplay:false},false,root);el.addEventListener('loadedmetadata',()=>{if(!el.isConnected)return;el.currentTime=time;el.play()?.catch(()=>{});},{once:true});active.push({el,dispose,until:from+count});
 }
 carries.set(root,active);
}
export function stopCarriedMedia(root){for(const item of carries.get(root)||[]){item.dispose();item.el.remove();}carries.delete(root);}
