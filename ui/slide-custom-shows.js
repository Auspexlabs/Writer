export function editCustomShow(doc,current) {
 return new Promise(resolve=>{
  const dialog=document.createElement('dialog');dialog.dataset.customShowEditor='1';dialog.style.cssText='width:540px;max-width:85vw;padding:24px;border:1px solid #bbb;border-radius:16px;background:var(--fp-bg,#fff);color:var(--fp-ink,#222);font:14px system-ui';
  const title=document.createElement('h3');title.textContent='自定义放映';dialog.append(title);
  const input=(name,label,value)=>{const l=document.createElement('label'),i=document.createElement('input');l.textContent=label;l.style.cssText='display:grid;gap:6px;margin:16px 0';i.name=name;i.value=value;i.style.width='100%';l.append(i);dialog.append(l);return i;};
  const name=input('name','名称',current?.name||'自定义放映'),pages=input('slides','页码顺序（例如 3,1,2,3；可重复同一页）',current?current.slides.map(id=>doc.slides.findIndex(s=>s.id===id)+1).filter(n=>n>0).join(','):doc.slides.map((_,i)=>i+1).join(','));
  const error=document.createElement('p');error.setAttribute('role','alert');error.style.color='#b3261e';dialog.append(error);const actions=document.createElement('div');actions.style.cssText='display:flex;gap:12px;justify-content:flex-end';dialog.append(actions);let ended=false;
  const done=v=>{if(ended)return;ended=true;dialog.close();dialog.remove();resolve(v);},button=(text,action)=>{const b=document.createElement('button');b.textContent=text;b.onclick=action;actions.append(b);};
  if(current)button('删除放映',()=>done({remove:true}));button('取消',()=>done(null));button('保存放映',()=>{
   const n=name.value.trim(),numbers=pages.value.trim().split(/[,，\s]+/).filter(Boolean).map(Number);
   if(!n||n.length>255||(doc.customShows||[]).some(s=>s.name!==current?.name&&s.name.toLowerCase()===n.toLowerCase())){error.textContent='请输入不重复的名称（最多 255 字）';return;}
   if(!numbers.length||numbers.some(n=>!Number.isInteger(n)||n<1||n>doc.slides.length)){error.textContent='请输入有效的页码';return;}
   done({name:n,slides:numbers.map(n=>doc.slides[n-1].id)});
  });
  dialog.addEventListener('cancel',e=>{e.preventDefault();done(null);});document.body.append(dialog);dialog.showModal();
 });
}
