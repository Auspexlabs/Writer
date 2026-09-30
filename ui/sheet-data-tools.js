import { A, parseA, Calc, shiftF, isErr, conditionalFormats, tableStyle, fmtColor } from './sheet-engine.js';

const blank = v => v === '' || v == null;
const lower = v => String(v ?? '').toLocaleLowerCase();
const escape = v => v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
function wildcard(text, prefix) {
  let pattern = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '~' && /[?*~]/.test(text[i + 1] || '')) pattern += escape(text[++i]);
    else pattern += ch === '*' ? '.*' : ch === '?' ? '.' : escape(ch);
  }
  return new RegExp('^' + pattern + (prefix ? '' : '$'), 'iu');
}
/** Excel advanced criteria: repeated headers are AND, rows are OR, bare text is a prefix. */
export function advancedCriterion(criterion) {
  if (typeof criterion !== 'string') return v => v === criterion;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/s.exec(criterion), op = match?.[1] || '=', text = match?.[2] ?? criterion;
  if (text === '' && match) return v => op === '<>' ? !blank(v) : op === '=' && blank(v);
  const numeric = text.trim() !== '' && Number.isFinite(Number(text));
  const logical = /^(TRUE|FALSE)$/i.test(text);
  const pattern = !numeric && !logical && (op === '=' || op === '<>') ? wildcard(text, !match) : null;
  return value => {
    if (isErr(value)) return false;
    if (pattern) return op === '<>' ? !pattern.test(String(value ?? '')) : pattern.test(String(value ?? ''));
    const right = numeric ? Number(text) : logical ? /^true$/i.test(text) : lower(text);
    const left = numeric ? typeof value === 'number' ? value : blank(value) ? 0 : NaN : logical ? value : lower(value);
    if (Number.isNaN(left)) return op === '<>';
    return op === '=' ? left === right : op === '<>' ? left !== right : op === '<' ? left < right : op === '<=' ? left <= right : op === '>' ? left > right : left >= right;
  };
}
export function dataRange(doc, si, text) {
  const calc = new Calc(doc), b = calc.rangeBounds(text, si);
  if (!b || b.si !== si) throw Error('请选择当前工作表中的区域');
  return b;
}
export function advancedFilter(doc, si, spec) {
  const calc = new Calc(doc), sh = doc.sheets[si], source = dataRange(doc, si, spec.range);
  if (source.r1 === source.r2) throw Error('列表区域需要标题行和至少一行数据');
  const headers = Array.from({ length: source.c2 - source.c1 + 1 }, (_, i) => lower(calc.value(si, source.r1, source.c1 + i)));
  if (headers.some(blank)) throw Error('列表的每一列都需要标题');
  const groups = []; let criteriaRange = '';
  if (String(spec.criteria || '').trim()) {
    const criteria = dataRange(doc, si, spec.criteria); criteriaRange = A(criteria.r1,criteria.c1)+':'+A(criteria.r2,criteria.c2);
    if (criteria.r1 === criteria.r2) throw Error('条件区域需要标题行和至少一行条件');
    if ((criteria.r2 - criteria.r1) * (criteria.c2 - criteria.c1 + 1) > 10000) throw Error('条件区域过大');
    const fields = Array.from({ length: criteria.c2 - criteria.c1 + 1 }, (_, i) => headers.indexOf(lower(calc.value(si, criteria.r1, criteria.c1 + i))));
    for (let r = criteria.r1 + 1; r <= criteria.r2; r++) {
      const terms = [];
      for (let c = criteria.c1; c <= criteria.c2; c++) {
        const cell = sh.cells[A(r, c)], field = fields[c - criteria.c1];
        if (blank(cell?.v)) continue;
        if (field < 0) {
          if (cell.literal || !String(cell.v).startsWith('=')) throw Error('条件标题必须对应列表标题，或在该列填写布尔公式');
          terms.push(row => { const value = calc.evaluate('=' + shiftF(cell.v.slice(1), row - source.r1 - 1, 0), si, row, source.c1); return !isErr(value) && (value === true || typeof value === 'number' && value !== 0); });
        } else {
          const criterion = calc.value(si, r, c); if (isErr(criterion)) throw Error('条件包含公式错误：' + A(r, c));
          const accepts = advancedCriterion(criterion); terms.push(row => accepts(calc.value(si, row, source.c1 + field)));
        }
      }
      groups.push(terms);
    }
  }
  const rows = [], hidden = [], seen = new Set();
  for (let r = source.r1 + 1; r <= source.r2; r++) {
    let pass = !groups.length || groups.some(terms => terms.every(test => test(r)));
    if (pass && spec.unique) { const key = JSON.stringify(headers.map((_, i) => { const value = calc.value(si, r, source.c1 + i); return typeof value === 'string' ? lower(value) : value; })); pass = !seen.has(key); seen.add(key); }
    (pass ? rows : hidden).push(r);
  }
  return { source, rows, hidden, calc, criteriaRange };
}
/** Copy calculated values, preserving display formats. Reject any destructive overlap before changing cells. */
export function copyAdvancedFilter(doc, si, spec, result = advancedFilter(doc, si, spec)) {
  const sh = doc.sheets[si], start = parseA(spec.target || ''); if (!start) throw Error('请填写复制到的单元格地址');
  const { source, rows, calc } = result, width = source.c2 - source.c1 + 1;
  calc.ensureSpills(si);
  const target = { r1: start.r, c1: start.c, r2: start.r + rows.length, c2: start.c + width - 1 };
  if (target.r2 >= 1048576 || target.c2 >= 16384) throw Error('复制结果超出工作表上限');
  const overlap = b => b.r1 <= target.r2 && b.r2 >= target.r1 && b.c1 <= target.c2 && b.c2 >= target.c1;
  if (overlap(source) || spec.criteria && overlap(dataRange(doc, si, spec.criteria))) throw Error('复制结果不能覆盖列表或条件区域');
  if ((sh.merges || []).some(m => overlap({ r1: m.r, r2: m.r + m.rs - 1, c1: m.c, c2: m.c + m.cs - 1 }))) throw Error('复制结果不能覆盖合并单元格');
  for (let r = target.r1; r <= target.r2; r++) for (let c = target.c1; c <= target.c2; c++) {
    const ref = A(r,c); if (!blank(sh.cells[ref]?.v) || calc.spills.has(`${si}:${r}:${c}`)) throw Error('复制结果会覆盖已有单元格：' + ref);
  }
  return { target, cells: [source.r1, ...rows].flatMap((r, i) => Array.from({ length: width }, (_, j) => {
    const value = calc.value(si, r, source.c1 + j), style = sh.cells[A(r, source.c1 + j)]?.s;
    return [A(start.r + i, start.c + j), { v: isErr(value) ? value.err : value, ...(style ? { s: { ...style } } : {}), ...(typeof value === 'string' && value[0] === '=' ? { literal: true } : {}) }];
  })) };
}
export function sortAppearance(doc, si, calc = new Calc(doc)) {
  const sh = doc.sheets[si], conditional = conditionalFormats(sh, calc, si, { parseA, fmtColor, isErr, shiftF });
  return (r,c) => { const value = calc.value(si,r,c), style = { ...tableStyle(sh,r,c), ...sh.cells[A(r,c)]?.s }, look = conditional(r,c,value,style); return { fill: lower(look.bg || '#ffffff'), color: lower(look.color || '#000000'), icon: look.icon ? look.icon + ' ' + look.iconColor : '' }; };
}
export function sortRowOrder(doc,si,box,keys,{header=true,caseSensitive=false}={}) {
  const calc = new Calc(doc), appearance = keys.some(k => ['fill','color','icon'].includes(k.by)) ? sortAppearance(doc,si,calc) : null;
  const rows = [];
  for (let r = box.r1 + (header ? 1 : 0); r <= box.r2; r++) rows.push({ r, values: keys.map(k => {
    const value = calc.value(si,r,k.c);
    if (['fill','color','icon'].includes(k.by)) return lower(appearance(r,k.c)[k.by]) === lower(k.match) ? 0 : 1;
    if (k.by === 'list') { if(blank(value))return null;const i = (k.list || []).findIndex(v => caseSensitive ? String(v) === String(value) : lower(v) === lower(value)); return i < 0 ? (k.list||[]).length : i; }
    return value;
  }) });
  const rank = value => isErr(value) ? 3 : typeof value === 'boolean' ? 2 : typeof value === 'number' ? 0 : 1;
  rows.sort((a,b) => { for (let i = 0; i < keys.length; i++) { const x = a.values[i], y = b.values[i]; if (blank(x) || blank(y)) { if (blank(x) !== blank(y)) return blank(x) ? 1 : -1; continue; } if(keys[i].by==='list'){const n=keys[i].list?.length||0;if((x===n)!==(y===n))return x===n?1:-1;} const d = rank(x) - rank(y) || (typeof x === 'number' && typeof y === 'number' ? x-y : String(isErr(x)?x.err:x).localeCompare(String(isErr(y)?y.err:y),'zh',{sensitivity:caseSensitive?'variant':'accent'})); if (d) return keys[i].asc === false ? -d : d; } return a.r-b.r; });
  return rows.map(row=>row.r);
}
