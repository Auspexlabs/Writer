import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as E from '../sheet-engine.js';
import { sheetPrint } from '../sheet-print.js';
test('print preserves merges, width, borders and excludes filtered rows and hidden sheets', () => {
 const doc = { sheets: [{ name: 'Report', cells: { A1: { v: 'Merged', s: { bd: { bottom: 'double' }, b: true } }, A2: { v: 'hidden' }, A3: { v: 'shown' } }, colW: { A: 180 }, merges: [{ r: 0, c: 0, rs: 1, cs: 2 }], frows: [1], print: { area: 'A1:B3', header: '&LReport&C&P/&N' } }, { name: 'Hidden', visibility: 'hidden', cells: { A1: { v: 'secret' } } }] };
 const out = sheetPrint(doc, E); assert.equal(out.pages.length, 1); assert.match(out.body, /colspan="2"/); assert.match(out.body, /width:180px/); assert.match(out.body, /border-bottom:2px double/); assert.match(out.body, /shown/); assert.doesNotMatch(out.body, /hidden|secret/); assert.match(out.body, /1\/1/);
});
test('print repeats rows and columns and fits to one page when requested', () => {
 const cells = {}; for (let r = 0; r < 90; r++) for (let c = 0; c < 8; c++) cells[E.A(r, c)] = { v: r === 0 ? 'Title' : c === 0 ? 'Label' : `${r},${c}` };
 const sh = { name: 'Rows', cells, print: { titles: '$1:$1,$A:$A' } }, doc = { sheets: [sh] }, out = sheetPrint(doc, E);
 assert.ok(out.pages.length > 3); for (const p of out.pages) { assert.match(p.content, /Title/); assert.match(p.content, /Label/); }
 sh.print.fitWidth = sh.print.fitHeight = 1; assert.equal(sheetPrint(doc, E).pages.length, 1);
});
