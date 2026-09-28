import { readTSV } from './sheet-clipboard.js';

export function mergeRecords(text, name = '') {
  text = text.replace(/^\uFEFF/, '');
  let records;
  if (/\.json$/i.test(name) || /^\s*\[/.test(text)) records = JSON.parse(text);
  else {
    const rows = readTSV(text, /\.tsv$/i.test(name) || text.split(/\r?\n/, 1)[0].includes('\t') ? '\t' : ',');
    const headers = rows.shift().map(x => x.trim());
    if (new Set(headers).size !== headers.length || headers.some(h => !/^[\p{L}_][\p{L}\p{N}_]{0,63}$/u.test(h))) throw new Error('第一行需要唯一的字段名，只能包含字母、汉字、数字和下划线。');
    records = rows.filter(r => r.some(v => v !== '')).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ''])));
  }
  if (!Array.isArray(records) || !records.length || records.some(r => !r || typeof r !== 'object' || Array.isArray(r) || Object.values(r).some(v => v != null && typeof v === 'object'))) throw new Error('请选择包含数据记录的 CSV、TSV 或 JSON 数组。');
  return records;
}

/** Each recipient starts from the same saved template; generated files are separate documents. */
export async function mergeDocuments(EN, template, records, progress) {
  const result = [];
  for (let i = 0; i < records.length; i++) {
    const doc = await EN.copy({ ...template, title: template.title + '-' + (i + 1) });
    try { await EN.run(['set', doc.path, '/', '--prop', 'mergeData=' + JSON.stringify(records[i])]); }
    catch (e) { await EN.discard(doc); e.message += `（已完成 ${result.length} 份）`; throw e; }
    result.push(doc); progress?.(i + 1, records.length);
  }
  return result;
}
