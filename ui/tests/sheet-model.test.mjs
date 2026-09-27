import { test } from 'node:test';
import assert from 'node:assert/strict';
import { editWorkbook, shareCells, changedCellKeys, cloneWorkbook } from '../sheet-model.js';
const base = () => ({ active: 0, sheets: [{ name: 'A', cells: shareCells({ A1: { v: 1, s: { b: true } }, B2: { v: '=A1' } }), cf: [{ range: 'A1:B2' }], charts: [] }, { name: 'B', cells: shareCells({ A1: { v: 3 } }) }] });
test('single-cell and nested style edits share all untouched data and preserve undo snapshots', () => {
  const a = base(), b = editWorkbook(a, sh => { sh.cells.A1.s.b = false; sh.cells.C3 = { v: 4 }; delete sh.cells.B2; });
  assert.equal(a.sheets[0].cells.A1.s.b, true);
  assert.equal(b.sheets[0].cells.A1.s.b, false);
  assert.equal(b.sheets[1], a.sheets[1]);
  assert.equal(b.sheets[0].cf, a.sheets[0].cf);
  assert.deepEqual(changedCellKeys(a.sheets[0].cells, b.sheets[0].cells).sort(), ['A1', 'B2', 'C3']);
  assert.deepEqual(Object.keys(b.sheets[0].cells).sort(), ['A1', 'C3']);
  assert.equal(JSON.parse(JSON.stringify(b)).sheets[0].cells.C3.v, 4);
  assert.equal(editWorkbook(a, sh => sh.cells.A1.v), a);
});
test('arrays, cross-sheet changes, deleted children and a replaced cell map finalize without retained drafts', () => {
  const a = base(), b = editWorkbook(a, (sh, d) => { sh.cf.push({ range: 'C1' }); sh.cf.splice(0, 1); d.sheets[1].cells.A1.v = 7; sh.cells = { D4: sh.cells.A1 }; });
  assert.equal(a.sheets[1].cells.A1.v, 3); assert.equal(b.sheets[1].cells.A1.v, 7);
  assert.equal(b.sheets[0].cf[0].range, 'C1');
  assert.equal(b.sheets[0].cells.D4.v, 1);
  const copy = cloneWorkbook(b); assert.equal(copy.sheets[0].cells.D4.v, 1); assert.equal(copy.sheets[1].cells, b.sheets[1].cells);
});
test('one edit in a million-cell sheet reports one dirty key and keeps the saved snapshot intact', () => {
  const cells = {}; for (let r = 1; r <= 100000; r++) for (let c = 0; c < 10; c++) cells[String.fromCharCode(65 + c) + r] = { v: r + c };
  const doc = { sheets: [{ cells: shareCells(cells) }] }, start = performance.now();
  const next = editWorkbook(doc, sh => { sh.cells.E50000.v = 99; });
  const elapsed = performance.now() - start;
  assert.equal(doc.sheets[0].cells.E50000.v, 50004);
  assert.deepEqual(changedCellKeys(doc.sheets[0].cells, next.sheets[0].cells), ['E50000']);
  assert.ok(elapsed < 500, `single edit took ${elapsed.toFixed(1)}ms`);
  console.log(`million-cell edit: ${elapsed.toFixed(1)}ms`);
});
