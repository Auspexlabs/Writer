// Shared by the editing grid and print/export, so rules have one visual result.
export function rangeBoxes(range, E) { return String(range || '').split(/\s+/).map(one => { const [a,b=a] = one.split(':'), p = E.parseA(a), q = E.parseA(b); return p && q ? { r1: Math.min(p.r,q.r), c1: Math.min(p.c,q.c), r2: Math.max(p.r,q.r), c2: Math.max(p.c,q.c) } : null; }).filter(Boolean); }
const mix = (a,b,t) => { const p = h => [1,3,5].map(i=>parseInt(h.slice(i,i+2),16)), x=p(a), y=p(b); return '#' + x.map((v,i)=>Math.round(v+(y[i]-v)*t).toString(16).padStart(2,'0')).join(''); };
export function conditionalFormats(sh, calc, si, E) {
    // conditional formats: each rule's boxes and what it needs of its range (numbers for bars, scales, top/bottom and averages; counts for duplicates)
    const cfs = (sh.cf || []).map((rule,i)=>({...rule,priority:rule.priority??i+1})).sort((a,b)=>a.priority-b.priority).map(rule => {
      const o = { rule, boxes: rangeBoxes(rule.range, E).map(b => ({ r1: b.r1, c1: b.c1, r2: Math.min(b.r2, 1048575), c2: Math.min(b.c2, 16383) })) }, t = rule.type;
      const each = fn => { const seen = new Set(); for (const box of o.boxes) for (const {r,c,value} of calc.rangeEntries({...box,si})) { const key=r+':'+c; if (!seen.has(key)) { seen.add(key); fn(value); } } };
      if (t === 'iconSet' || t === 'dataBar' || t === 'colorScale' || t === 'top10' || t === 'aboveAverage') { const ns = []; each(v => { if (typeof v === 'number') ns.push(v); }); o.mn = ns.reduce((a, b) => Math.min(a, b), Infinity); o.mx = ns.reduce((a, b) => Math.max(a, b), -Infinity); o.n = ns.length; o.values = ns;
        if (t === 'top10') { const k = Math.max(1, Math.round(rule.percent ? ns.length * (rule.rank || 10) / 100 : (rule.rank || 10))); const sorted = ns.slice().sort((a, b) => rule.bottom ? a - b : b - a); o.edge = sorted.length ? sorted[Math.min(k, sorted.length) - 1] : NaN; }
        if (t === 'aboveAverage') { const avg = ns.reduce((a, b) => a + b, 0) / (ns.length || 1), sd = Math.sqrt(ns.reduce((a, b) => a + (b - avg) ** 2, 0) / (ns.length || 1)); o.avg = avg + (rule.stdDev || 0) * sd * (rule.above === false ? -1 : 1); } }
      if (t === 'duplicateValues' || t === 'uniqueValues') { const cnt = {}; each(v => { const k = String(v); if (k !== '') cnt[k] = (cnt[k] || 0) + 1; }); o.cnt = cnt; }
      return o;
    }).reverse(); // applied last to first, so the first rule that holds wins a clash, as in Excel
    const cfHolds = (o, v, row, col) => {
      if (o.rule.type === 'expression') { const first = o.boxes[0], f = E.shiftF(o.rule.value || '', row - first.r1, col - first.c1), answer = calc.evaluate(f, si, row, col); return !E.isErr(answer) && !!answer; }
      const r = o.rule, t = r.type, num = typeof v === 'number', str = String(v), lo = str.toLowerCase(), txt = String(r.text || '').toLowerCase(), raw = x => String(x == null ? '' : x).replace(/^"|"$/g, '').toLowerCase();
      const threshold = x => { const first=o.boxes[0]; if(x==null)return ''; const source=String(x), value=calc.evaluate(E.shiftF(source,row-first.r1,col-first.c1),si,row,col);
        // Older UI rules store literal words without Excel's surrounding quotes.
        return !num && E.isErr(value) && value.err==='#NAME?' && !/[=+*/^&()<>":'!$]/.test(source) && !E.parseA(source) ? source : value; };
      const a = t === 'cellIs' ? threshold(r.value) : r.value, b = t === 'cellIs' ? threshold(r.value2) : r.value2;
      if (E.isErr(a) || E.isErr(b)) return false;
      if (t === 'cellIs') { switch (r.operator) { case 'greaterThan': return num && v > a; case 'lessThan': return num && v < a; case 'greaterThanOrEqual': return num && v >= a; case 'lessThanOrEqual': return num && v <= a; case 'between': return num && v >= Math.min(a, b) && v <= Math.max(a, b); case 'notBetween': return num && (v < Math.min(a, b) || v > Math.max(a, b)); case 'notEqual': return num ? v !== a : lo !== raw(a); default: return num ? v === a : lo === raw(a); } }
      if (t === 'containsText') return lo.includes(txt); if (t === 'notContainsText') return !lo.includes(txt); if (t === 'beginsWith') return lo.startsWith(txt); if (t === 'endsWith') return lo.endsWith(txt);
      if (t === 'duplicateValues') return o.cnt[str] > 1; if (t === 'uniqueValues') return o.cnt[str] === 1;
      if (t === 'top10') return num && (r.bottom ? v <= o.edge : v >= o.edge); if (t === 'aboveAverage') return num && (r.above === false ? v < o.avg : v > o.avg);
      return false; // bars and scales colour by value; icon sets, formulas and time periods are kept for the file, not drawn
    };

  return (r,c,v,s={}) => {
    let bg = s.fill || '', color = E.fmtColor(v,s.code) || s.color || '', bold = null, italic = null, icon = '', iconColor = '', hideValue = false;
        const stop = cfs.slice().reverse().find(o => o.rule.stopIfTrue && o.boxes.some(b => r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2) && cfHolds(o,v,r,c));
        cfs.forEach(o => {
          if (stop && o.rule.priority > stop.rule.priority) return;
          if (v === '' && o.rule.type !== 'expression' || !o.boxes.some(b => r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2)) return;
          const rule = o.rule, t = rule.type;
          if (t === 'iconSet' && typeof v === 'number' && o.n) {
            const n = +String(rule.iconSet || '3Arrows')[0] || 3, limits = rule.thresholds || Array.from({ length: n }, (_, i) => ({ type: 'percent', value: i * 100 / n }));
            let rank = 0;
            for (let i = 1; i < limits.length; i++) { const t = limits[i], threshold = t.type === 'num' ? +t.value : t.type === 'formula' ? calc.evaluate(t.value, si, r, c) : t.type === 'min' ? o.mn : t.type === 'max' ? o.mx : t.type === 'percentile' ? o.values.slice().sort((a, b) => a - b)[Math.min(o.n - 1, Math.floor((o.n - 1) * +t.value / 100))] : o.mn + (o.mx - o.mn) * +t.value / 100; if (t.gte === false ? v > threshold : v >= threshold) rank = i; }
            if (rule.reverse) rank = n - rank - 1;
            icon = /Traffic|Light/.test(rule.iconSet) ? '●' : /Rating/.test(rule.iconSet) ? ['▁', '▃', '▅', '▆', '█'][Math.round(rank * 4 / (n - 1))] : ['↓', '↘', '→', '↗', '↑'][Math.round(rank * 4 / (n - 1))]; iconColor = rank === 0 ? '#C43D3D' : rank === n - 1 ? '#368251' : '#B88A16'; if (rule.showValue === false) hideValue = true; return;
          }
          if (t === 'dataBar' && typeof v === 'number' && o.n) { const lo = Math.min(0, o.mn), k = o.mx === lo ? 1 : Math.max(0, (v - lo) / (o.mx - lo)); bg = `linear-gradient(90deg,${rule.color || '#638EC6'}66 ${Math.round(k * 100)}%,${s.fill || 'transparent'} ${Math.round(k * 100)}%)`; return; }
          if (t === 'colorScale' && typeof v === 'number' && o.n) { const cols = rule.colors && rule.colors.length >= 2 ? rule.colors : ['#F8696B', '#FFEB84', '#63BE7B'], k = o.mx === o.mn ? 0.5 : (v - o.mn) / (o.mx - o.mn), seg = (cols.length - 1) * k, i = Math.min(cols.length - 2, Math.floor(seg)); bg = mix(cols[i], cols[i + 1], seg - i); return; }
          if (!cfHolds(o, v, r, c)) return;
          // Excel's default look when a rule names none: 浅红填充色深红色文本
          const noLook = !rule.fill && !rule.color && !rule.bold && !rule.italic;
          if (rule.fill || noLook) bg = rule.fill || '#FFC7CE'; if (rule.color || noLook) color = rule.color || '#9C0006'; if (rule.bold) bold = true; if (rule.italic) italic = true;
        });
    return { bg, color, bold, italic, icon, iconColor, hideValue };
  };
}
