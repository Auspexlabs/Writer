import { autoRowHeights } from './sheet-layout.js';
// Paginated spreadsheet print/export, using the same visible values and dimensions as the editor.
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const px = mm => mm * 96 / 25.4;
const papers = { 1: [215.9, 279.4], 5: [215.9, 355.6], 8: [297, 420], 9: [210, 297], 11: [148, 210] };
const list = text => String(text || '').split(/,(?=(?:[^']*'[^']*')*[^']*$)/).map(s => s.slice(s.lastIndexOf('!') + 1).replace(/\$/g, '').trim());
function partition(items, size, available, repeats) {
  const prefix = items.filter(i => repeats.has(i)), rest = items.filter(i => !repeats.has(i));
  const space = Math.max(1, available - prefix.reduce((n, i) => n + size(i), 0)), groups = []; let group = [], used = 0;
  for (const i of rest) { const n = size(i); if (group.length && used + n > space) { groups.push(prefix.concat(group)); group = []; used = 0; } group.push(i); used += n; }
  if (group.length || !groups.length) groups.push(prefix.concat(group)); return groups;
}
function hf(text, page, total, doc, sh) {
  const textOf = s => esc(s.replace(/&(?:"[^"]*"|\d+|[BIUS])/g, '').replace(/&[PNADF&T]/g, k => ({ '&P': page, '&N': total, '&A': sh.name, '&F': doc.title || '', '&D': new Date().toLocaleDateString(), '&T': new Date().toLocaleTimeString(), '&&': '&' }[k] ?? k)));
  const zones = { L: '', C: '', R: '' }; let side = 'C';
  for (const chunk of String(text || '').split(/(&[LCR])/)) { if (/^&[LCR]$/.test(chunk)) side = chunk[1]; else zones[side] += chunk; }
  return ['L', 'C', 'R'].map(z => `<span style="text-align:${({ L: 'left', C: 'center', R: 'right' })[z]}">${textOf(zones[z])}</span>`).join('');
}
function cellStyle(s, grid) {
  const css = [`font-family:${esc(s.font || 'Arial')}`, `font-size:${s.fs || 11}pt`, `font-weight:${s.b ? 700 : 400}`, `font-style:${s.i ? 'italic' : 'normal'}`, `text-decoration:${[s.u && 'underline', s.strike && 'line-through'].filter(Boolean).join(' ') || 'none'}`, `text-align:${s.align === 'general' ? 'left' : s.align || 'left'}`, `vertical-align:${s.va || 'bottom'}`, `white-space:${s.wrap ? 'pre-wrap' : 'pre'}`, `color:${s.color || '#000'}`, `background:${s.fill || 'transparent'}`];
  for (const side of ['top', 'right', 'bottom', 'left']) {
    const b = typeof s.bd === 'object' ? s.bd[side] : s.bd === 'all' ? 'thin' : '';
    css.push(`border-${side}:${b && b !== 'none' ? (b === 'thick' ? '3px' : b === 'medium' || b === 'double' ? '2px' : '1px') + ' ' + (/dash/i.test(b) ? 'dashed' : /dot/i.test(b) ? 'dotted' : b === 'double' ? 'double' : 'solid') + ' ' + (s.bdc || '#000') : grid ? '1px solid #bbb' : 'none'}`);
  }
  return css.join(';');
}
export function sheetPrint(doc, E) {
  const calc = new E.Calc(doc), pages = [];
  for (const [si, sh] of doc.sheets.entries()) {
    if (sh.visibility && sh.visibility !== 'visible') continue;
    const p = sh.print || {}, paper = papers[p.paper] || papers[9], [w, h] = p.orientation === 'landscape' ? [paper[1], paper[0]] : paper;
    const margin = [p.top ?? .75, p.right ?? .7, p.bottom ?? .75, p.left ?? .7].map(n => n * 25.4);
    const usableW = px(w - margin[1] - margin[3]), usableH = px(h - margin[0] - margin[2]) - 48;
    const hiddenR = new Set([...(sh.hiddenRows || []), ...(sh.frows || [])]), hiddenC = new Set(sh.hiddenCols || []);
    const autoH = autoRowHeights(sh, E.parseA, E.colName, (r, c) => calc.value(si, r, c), doc.fs || 11, doc);
    const cw = c => sh.colW?.[E.colName(c)] || 100, rh = r => sh.rowH?.[r + 1] || autoH[r + 1] || 26;
    const repeatR = new Set(), repeatC = new Set();
    for (const t of list(p.titles)) {
      let m = /^(\d+):(\d+)$/.exec(t); if (m) for (let r = +m[1] - 1; r < +m[2] && r < 1048576; r++) repeatR.add(r);
      m = /^([A-Z]+):([A-Z]+)$/i.exec(t); if (m) for (let c = E.parseA(m[1] + '1')?.c; c <= E.parseA(m[2] + '1')?.c && c < 16384; c++) repeatC.add(c);
    }
    const ranges = list(p.area).filter(Boolean).map(s => { const [a, b] = s.split(':').map(E.parseA); return a ? { r1: a.r, c1: a.c, r2: (b || a).r, c2: (b || a).c } : null; }).filter(Boolean);
    if (!ranges.length) { const u = E.usedRange(sh); if (u) ranges.push(u); }
    for (const u of ranges) {
      const rows = Array.from({ length: u.r2 - u.r1 + 1 }, (_, i) => u.r1 + i).filter(r => !hiddenR.has(r)), cols = Array.from({ length: u.c2 - u.c1 + 1 }, (_, i) => u.c1 + i).filter(c => !hiddenC.has(c));
      // Titles may lie outside the print area; include them on every printed page.
      for (const r of repeatR) if (!hiddenR.has(r) && !rows.includes(r)) rows.unshift(r);
      for (const c of repeatC) if (!hiddenC.has(c) && !cols.includes(c)) cols.unshift(c);
      rows.sort((a, b) => a - b); cols.sort((a, b) => a - b);
      if (!rows.length || !cols.length) continue;
      let scale = Math.max(.1, Math.min(4, (p.scale || 100) / 100));
      if (p.fitWidth > 0 || p.fitHeight > 0) scale = Math.min(1, p.fitWidth > 0 ? usableW * p.fitWidth / cols.reduce((n, c) => n + cw(c), 0) : Infinity, p.fitHeight > 0 ? usableH * p.fitHeight / rows.reduce((n, r) => n + rh(r), 0) : Infinity);
      for (const rr of partition(rows, rh, usableH / scale, repeatR)) for (const cc of partition(cols, cw, usableW / scale, repeatC)) {
        let table = '<table><colgroup>' + cc.map(c => `<col style="width:${cw(c)}px">`).join('') + '</colgroup><tbody>';
        for (const r of rr) {
          table += `<tr style="height:${rh(r)}px">`;
          for (const c of cc) {
            const merge = sh.merges?.find(m => r >= m.r && r < m.r + m.rs && c >= m.c && c < m.c + m.cs);
            let rs = 1, cs = 1, ar = r, ac = c;
            if (merge) { const mr = rr.filter(i => i >= merge.r && i < merge.r + merge.rs), mc = cc.filter(i => i >= merge.c && i < merge.c + merge.cs); if (r !== mr[0] || c !== mc[0]) continue; rs = mr.length; cs = mc.length; ar = merge.r; ac = merge.c; }
            const cell = sh.cells[E.A(ar, ac)], s = cell?.s || {}, v = calc.value(si, ar, ac), styled = { ...s, align: s.align || (typeof v === 'number' ? 'right' : 'left') };
            table += `<td rowspan="${rs}" colspan="${cs}" style="${cellStyle(styled, p.gridlines)}">${esc(E.fmt(v, s))}</td>`;
          }
          table += '</tr>';
        }
        table += '</tbody></table>';
        const images = (sh.images || []).filter(im => im.x >= cc[0] * 100 && im.x < (cc.at(-1) + 1) * 100 && im.y >= rr[0] * 26 && im.y < (rr.at(-1) + 1) * 26).map(im => `<img src="${esc(im.src)}" style="position:absolute;left:${im.x - cols.filter(c => c < cc[0]).reduce((n, c) => n + cw(c), 0)}px;top:${im.y - rows.filter(r => r < rr[0]).reduce((n, r) => n + rh(r), 0)}px;width:${im.w}px;height:${im.h}px">`).join('');
        pages.push({ sh, p, w, h, margin, scale, content: table + images });
      }
    }
  }
  const css = `body{margin:0;color:#000}.sheet-page{box-sizing:border-box;break-after:page;position:relative;overflow:hidden}.sheet-page:last-child{break-after:auto}.sheet-page table{border-collapse:collapse;table-layout:fixed}.sheet-page td{box-sizing:border-box;padding:2px 5px;overflow:hidden;overflow-wrap:anywhere}.sheet-head,.sheet-foot{height:24px;display:flex;font:9pt Arial}.sheet-head span,.sheet-foot span{width:33.333%}.sheet-foot{position:absolute;left:0;right:0;bottom:0}.sheet-content{position:relative;transform-origin:top left}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}` + pages.map((p, i) => `@page sheet${i}{size:${p.w}mm ${p.h}mm;margin:0}`).join('');
  const body = pages.map((p, i) => `<section class="sheet-page" style="page:sheet${i};width:${p.w}mm;height:${p.h}mm;padding:${p.margin.map(n => n + 'mm').join(' ')}"><div class="sheet-head">${hf(p.p.header, i + 1, pages.length, doc, p.sh)}</div><div class="sheet-content" style="transform:scale(${p.scale});width:${px(p.w - p.margin[1] - p.margin[3]) / p.scale}px">${p.content}</div><div class="sheet-foot" style="bottom:${p.margin[2] / 2}mm;left:${p.margin[3]}mm;right:${p.margin[1]}mm">${hf(p.p.footer, i + 1, pages.length, doc, p.sh)}</div></section>`).join('');
  return { css, body, pages };
}
