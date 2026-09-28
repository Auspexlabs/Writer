/** Page labels are separate from physical page indices (odd/even headers use the latter). */
export function pageNumber(n, format = 'decimal') {
  if (/Roman$/.test(format) && n > 0 && n < 4000) {
    let out = ''; for (const [v, s] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']]) while (n >= v) { out += s; n -= v; }
    return format === 'lowerRoman' ? out.toLowerCase() : out;
  }
  if (/Letter$/.test(format) && n > 0) { let out = ''; for (; n; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out; return format === 'lowerLetter' ? out.toLowerCase() : out; }
  if (format === 'decimalEnclosedCircleChinese' && n > 0 && n <= 20) return String.fromCharCode(0x245f + n);
  if (format === 'chineseCounting' && n >= 0 && n < 10000) { const d = '零一二三四五六七八九'; let out = '', zero = false; for (let p = 3; p >= 0; p--) { const k = Math.floor(n / 10 ** p) % 10; if (k) { if (zero) out += d[0]; out += (k === 1 && p === 1 && !out ? '' : d[k]) + ['', '十', '百', '千'][p]; zero = false; } else if (out) zero = true; } return out || d[0]; }
  return String(n);
}
export function pageSections(count, sections) {
  let previous = 0;
  return Array.from({ length: count }, (_, i) => {
    let j = sections.findIndex(s => i <= s.end); if (j < 0) j = sections.length - 1;
    const section = sections[j] || {}, first = j ? sections[j - 1].end + 1 : 0;
    if (i === first && section.pageNumberStart !== '' && section.pageNumberStart != null && section.pageNumberStart !== 'none') previous = +section.pageNumberStart - 1;
    const number = ++previous;
    return { section, index: j, first: i === first, number, label: pageNumber(number, section.pageNumberFormat) };
  });
}
export function headerKey(kind, page, evenAndOdd) {
  return page?.first && page.section.titlePg ? (kind === 'header' ? 'firstHeader' : 'firstFooter') : evenAndOdd && page?.physical % 2 === 0 ? (kind === 'header' ? 'evenHeader' : 'evenFooter') : kind;
}
