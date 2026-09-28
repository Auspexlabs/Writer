import { marginsCm } from './office-io.js';
const esc = v => String(v ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
const cssText = style => Array.from(style).map(k => `${k}:${style.getPropertyValue(k)}`).join(';');
/** Print the laid-out pages. Only blocks intersecting a page are copied; split paragraphs retain their line spacers. */
export function wordPrint(editor) {
  const ed = editor.edRef.current, pg = editor.page, box = ed.getBoundingClientRect(), zoom = box.width / ed.offsetWidth || 1;
  const sheet = ed.parentElement, rect = sheet.getBoundingClientRect(), W = rect.width / zoom;
  const margins = marginsCm(pg.margin).map(v => v * 96 / 2.54), [mt, mr, mb, ml] = margins;
  const sizes = { A4: [793.7,1122.52], A3: [1122.52,1587.4], A5: [559.37,793.7], B5: [665.2,944.88], Letter: [816,1056], Legal: [816,1344] };
  const sz = sizes[pg.size] || sizes.A4, H = pg.orient === 'landscape' ? sz[0] : sz[1];
  const count = editor.state.info.pages, gap = 20;
  // Use the editor's real physical spacing, rather than assuming rounded paper dimensions.
  const paper = Array.from(sheet.children).filter(e => e.style.pointerEvents === 'none' && parseFloat(e.style.height) > 500);
  const actualH = paper[0] ? parseFloat(paper[0].style.height) : H, actualStep = paper[1] ? parseFloat(paper[1].style.top) : actualH + gap;
  const styleRules = [];
  for (const style of document.styleSheets) { try { for (const rule of style.cssRules) if (/wd-|data-pg|data-fnpg|katex|@font-face|@counter-style|wfont/.test(rule.cssText)) styleRules.push(rule.cssText); } catch {} }
  // The print body contains a subset of the original children. Keep page spacer selectors tied to their original indices.
  const prefix = `[data-pg="${editor.pgId}"]`;
  let css = styleRules.join('\n').replaceAll(prefix + '>:nth-child(', prefix + '>[data-print-index="').replace(/(\[data-print-index="\d+)\)/g, '$1"]');
  const bodyStyle = cssText(getComputedStyle(ed));
  const kids = Array.from(ed.children).map((el, index) => { const r = el.getBoundingClientRect(); return { el, index, top: (r.top - box.top) / zoom, bottom: (r.bottom - box.top) / zoom, left: (r.left - box.left) / zoom, width: r.width / zoom }; });
  const parts = [], meta = editor.sectionPages();
  for (let i = 0; i < count; i++) {
    const y = i * actualStep, p = meta[i] || editor.pageMeta(i);
    const root = ed.cloneNode(false); root.removeAttribute('contenteditable'); root.style.cssText = bodyStyle;
    Object.assign(root.style, { position: 'absolute', top: mt + 'px', left: ml + 'px', width: ed.offsetWidth + 'px', height: actualH - mt - mb + 'px', minHeight: '0', overflow: 'hidden', margin: '0', padding: '0', opacity: '1', zoom: '1', display: 'block' });
    for (const k of kids) {
      if (k.bottom <= y || k.top >= y + actualH - mt - mb) continue;
      const c = k.el.cloneNode(true); c.setAttribute('data-print-index', k.index + 1);
      Object.assign(c.style, { position: 'absolute', top: k.top - y + 'px', left: k.left + 'px', width: k.width + 'px', boxSizing: 'border-box' }); c.style.setProperty('margin', '0', 'important');
      c.removeAttribute('contenteditable'); c.querySelectorAll('[contenteditable]').forEach(e => e.removeAttribute('contenteditable')); root.append(c);
    }
    const hf = kind => { const key = editor.hfKey(kind, i), html = String(p.section[key] || '').replace(/\{page\}/g, p.label).replace(/\{pages\}/g, count); return `<div class="wd-hf print-${kind}" style="position:absolute;left:${ml}px;right:${mr}px;${kind === 'header' ? 'top:' + Math.max(0, mt / 2 - 10) : 'bottom:' + Math.max(0, mb / 2 - 10)}px">${html}</div>`; };
    const notes = Array.from(editor.fnRef.current?.querySelectorAll(':scope > .wd-fn-area') || []).filter(n => Math.abs(parseFloat(n.style.top) - (y + actualH - mb)) < 2).map(n => { const c = n.cloneNode(true); c.style.top = actualH - mb + 'px'; c.querySelectorAll('[contenteditable]').forEach(e => e.removeAttribute('contenteditable')); return c.outerHTML; }).join('');
    const watermark = pg.wm ? `<div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none"><span style="font:64px serif;color:rgba(128,128,128,.18);transform:rotate(-35deg)">${esc(pg.wm)}</span></div>` : '';
    parts.push(`<section class="writer-print-page" style="width:${W}px;height:${actualH}px;background:${/^#[0-9a-f]{6}$/i.test(pg.color) ? pg.color : '#fff'}">${paper[i]?.querySelector('.wd-page-border')?.outerHTML || ''}${watermark}${root.outerHTML}${hf('header')}${hf('footer')}<div class="wd-fnotes" data-fnpg="${editor.pgId}" style="--fn-w:${ed.offsetWidth}px">${notes}</div></section>`);
  }
  css += `\n@page{size:${W}px ${actualH}px;margin:0}html,body{margin:0;padding:0;background:white;color:#1d1d1f}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}.writer-print-page{position:relative;overflow:hidden;break-after:page;box-sizing:border-box}.writer-print-page:last-child{break-after:auto}.wd-ed [data-w-sectionbreak]::after,.wd-ed hr[data-pb]::after{display:none!important}.wd-ed a{color:inherit}.wd-hf{font-size:12px}.wd-fnotes{position:absolute;inset:0;pointer-events:none}.wd-fn-bin{display:none!important}`;
  return { css, body: parts.join('') };
}
