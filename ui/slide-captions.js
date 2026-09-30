/** Import SubRip alongside native WebVTT. Caption text stays local to the deck. */
export function captionText(source) {
  let text=String(source).replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n');
  if(text.length>2_000_000)throw Error('单个字幕文件不能超过 2 MB');
  if(!/^WEBVTT(?:[ \t][^\n]*)?(?:\n|$)/.test(text)) {
    if(!/^\d*\s*\n?\d{2,}:\d{2}:\d{2},\d{3}\s*-->/m.test(text))throw Error('请选择 WebVTT 或 SRT 字幕文件');
    text='WEBVTT\n\n'+text.replace(/(\d{2,}:\d{2}:\d{2}),(\d{3})/g,'$1.$2');
  }
  const time=s=>{const parts=s.split(':').map(Number);if(parts.at(-1)>=60||parts.at(-2)>=60)return NaN;return parts.reduce((n,x)=>n*60+x,0);};
  const cues=text.split(/\n\s*\n/).slice(1).filter(block=>!/^NOTE(?:\s|$)|^STYLE(?:\s|$)|^REGION(?:\s|$)/.test(block));
  let count=0;for(const block of cues){if(!block.trim())continue;const line=block.split('\n').find(l=>l.includes('-->')),match=line?.match(/^((?:\d{2,}:)?\d{2}:\d{2}\.\d{3})\s+-->\s+((?:\d{2,}:)?\d{2}:\d{2}\.\d{3})(?:\s.*)?$/);if(!match||!(time(match[2])>time(match[1])))throw Error('字幕时间格式无效，结束时间必须晚于开始时间');count++;}
  if(!count)throw Error('字幕文件中没有有效的字幕段');return text;
}
export function bindCaptions(player,captions={},root=player.parentElement) {
  const tracks=[],urls=[],listeners=[];let overlay,select,box;
  const first=(captions.tracks||[]).findIndex(t=>typeof t.text==='string');
  const custom=player.tagName==='AUDIO'||captions.display==='slide';
  const update=()=>{if(!box)return;const selected=tracks[Number(select.value)]?.track;box.textContent=[...(selected?.activeCues||[])].map(c=>c.getCueAsHTML?.().textContent??c.text).join('\n');box.hidden=!box.textContent;};
  if(custom&&(captions.tracks||[]).length&&root){overlay=document.createElement('div');overlay.dataset.captionOverlay=player.dataset.mediaId||'audio';overlay.style.cssText='position:absolute;inset:auto 5% 4%;z-index:100;display:grid;justify-items:center;gap:8px;pointer-events:none';box=document.createElement('div');box.style.cssText='white-space:pre-wrap;text-align:center;color:#fff;background:#000c;padding:8px 18px;border-radius:4px;font:24px/1.4 sans-serif;max-width:90%';box.hidden=true;select=document.createElement('select');select.setAttribute('aria-label','字幕语言');select.style.cssText='pointer-events:auto;justify-self:end;font:12px system-ui;color:#fff;background:#222b;border:0;border-radius:4px;padding:3px';select.add(new Option('关闭字幕','-1'));(captions.tracks||[]).forEach((t,i)=>{const option=new Option(t.label||t.lang||String(i+1),String(i));option.disabled=typeof t.text!=='string';select.add(option);});select.value=String(first);select.onpointerdown=e=>e.stopPropagation();select.onclick=e=>e.stopPropagation();select.onchange=()=>{tracks.forEach((t,i)=>t.track.mode=i===Number(select.value)?'hidden':'disabled');update();};overlay.append(box,select);root.append(overlay);}
  (captions.tracks||[]).forEach((item,i)=>{
    // Imported external tracks remain in the file; loading them is an explicit
    // import in the editor rather than an automatic network dependency in a show.
    if(typeof item.text!=='string')return;
    const track=document.createElement('track');track.kind='captions';track.label=item.label||'';track.srclang=item.lang||'';track.default=i===first;const url=URL.createObjectURL(new Blob([item.text],{type:'text/vtt'}));urls.push(url);track.src=url;player.append(track);tracks[i]=track;
    const loaded=()=>{track.track.mode=i===(select?Number(select.value):first)?custom?'hidden':'showing':'disabled';update();};track.addEventListener('load',loaded);track.track.addEventListener('cuechange',update);listeners.push(()=>{track.removeEventListener('load',loaded);track.track.removeEventListener('cuechange',update);});loaded();
  });
  player.addEventListener('timeupdate',update);player.addEventListener('seeked',update);
  return ()=>{player.removeEventListener('timeupdate',update);player.removeEventListener('seeked',update);listeners.forEach(fn=>fn());tracks.forEach(t=>t.remove());urls.forEach(url=>URL.revokeObjectURL(url));overlay?.remove();};
}
export function editCaptions(current={display:'media',tracks:[]}) {return new Promise(resolve=>{
  const data=structuredClone(current),dialog=document.createElement('dialog');data.tracks||=[];dialog.dataset.captionsEditor='1';dialog.style.cssText='width:660px;max-width:90vw;max-height:90vh;overflow:auto;border:1px solid #bbb;border-radius:16px;padding:24px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
  const heading=document.createElement('h3');heading.textContent='媒体字幕';dialog.append(heading);
  const field=(label,node)=>{const row=document.createElement('label');row.style.cssText='display:grid;gap:6px;margin:10px 0';row.append(document.createTextNode(label),node);dialog.append(row);return node;};
  const display=field('字幕位置',document.createElement('select'));display.name='display';display.add(new Option('媒体画面','media'));display.add(new Option('幻灯片底部','slide'));display.value=data.display||'media';
  const picker=field('字幕轨',document.createElement('select'));picker.name='track';const label=field('名称',document.createElement('input'));label.name='label';const lang=field('语言代码（如 zh-CN、en）',document.createElement('input'));lang.name='lang';const text=field('WebVTT 字幕',document.createElement('textarea'));text.name='text';text.rows=10;text.style.fontFamily='monospace';
  const file=document.createElement('input');file.type='file';file.accept='.vtt,.srt,text/vtt';field('导入字幕文件',file);
  const actions=document.createElement('div');actions.style.cssText='display:flex;gap:10px;flex-wrap:wrap';const error=document.createElement('p');error.setAttribute('role','alert');error.style.color='#b3261e';dialog.append(error,actions);
  let selected=0,ended=false;const capture=()=>{const t=data.tracks[selected];if(t){t.label=label.value;t.lang=lang.value;if(text.value||!t.link){t.text=text.value;delete t.link;}}};
  const render=()=>{picker.replaceChildren();data.tracks.forEach((t,i)=>picker.add(new Option(t.label||'字幕 '+(i+1),String(i))));picker.value=String(selected);const t=data.tracks[selected];label.value=t?.label||'';lang.value=t?.lang||'';text.value=t?.text||'';for(const el of [picker,label,lang,text])el.disabled=!t;};
  picker.onchange=()=>{const next=Number(picker.value);capture();selected=next;render();};
  file.onchange=async()=>{const f=file.files[0];if(!f)return;try{const value=captionText(await f.text());capture();if(data.tracks.length>=32)throw Error('最多可保存 32 条字幕轨');data.tracks.push({label:f.name.replace(/\.(vtt|srt)$/i,''),lang:'',text:value});selected=data.tracks.length-1;render();error.textContent='';}catch(e){error.textContent=e.message;}file.value='';};
  const done=v=>{if(ended)return;ended=true;dialog.close();dialog.remove();resolve(v);};
  for(const [title,fn]of [['新增字幕轨',()=>{capture();if(data.tracks.length>=32)return;data.tracks.push({label:'字幕 '+(data.tracks.length+1),lang:'zh-CN',text:'WEBVTT\n\n00:00:00.000 --> 00:00:03.000\n'});selected=data.tracks.length-1;render();}],['删除字幕轨',()=>{data.tracks.splice(selected,1);selected=Math.max(0,selected-1);render();}],['取消',()=>done(null)],['保存字幕',()=>{try{capture();for(const t of data.tracks){if(!t.label.trim())throw Error('请填写字幕名称');if(t.text!=null)t.text=captionText(t.text);}done({...data,display:display.value});}catch(e){error.textContent=e.message;}}]]){const b=document.createElement('button');b.textContent=title;b.onclick=fn;actions.append(b);}
  dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});render();document.body.append(dialog);dialog.showModal();
});}
