import { test } from 'node:test';
import assert from 'node:assert/strict';
import { gridWindow, axisCoordinates, MAX_ROWS, MAX_COLS } from '../sheet-grid.js';
import { readTSV, writeTSV, parseA, Calc, shiftF } from '../sheet-engine.js';
test('virtual tracks retain full geometry, freeze visible panes and include offscreen merge anchors', () => {
  const xs = axisCoordinates(MAX_COLS, () => 80), ys = axisCoordinates(MAX_ROWS, i => i === 1 ? 0 : 24);
  const v = gridWindow(xs, ys, { x: 8000, y: 240000, width: 800, height: 600, frC: 2, frR: 3 }, [{ r: 9990, c: 95, rs: 30, cs: 30 }]);
  assert.ok(v.rows.indices.includes(9990)); assert.ok(v.cols.indices.includes(95));
  assert.ok(v.rows.indices.includes(0)); assert.ok(!v.rows.indices.includes(1)); assert.ok(v.cols.indices.includes(1));
  assert.ok(v.rows.indices.length < 50 && v.cols.indices.length < 25);
  assert.equal(v.rows.css.split(' ').reduce((n, x) => n + parseFloat(x), 0), ys[MAX_ROWS]);
  assert.ok(v.rows.line(10020) > v.rows.line(9990));
});
test('quoted TSV keeps linebreaks, tabs, quotes, empty trailing cells and CRLF boundaries', () => {
  const rows = [['first\nsecond', 'a\tb', 'he said "yes"', ''], ['', 'tail', '', '']];
  assert.deepEqual(readTSV(writeTSV(rows)), rows);
  assert.deepEqual(readTSV('"line\r\nbreak"\t""\r\nnext\t\r\n'), [['line\r\nbreak', ''], ['next', '']]);
});
test('full Excel addresses parse, calculate and shift with valid upper boundaries', () => {
  assert.deepEqual(parseA('XFD1048576'), { r: 1048575, c: 16383 }); assert.equal(parseA('XFE1'), null); assert.equal(parseA('A0'), null);
  const calc = new Calc({ sheets: [{ cells: { AAA100: { v: 7 }, A1: { v: '=AAA100*2' } } }] });
  assert.equal(calc.value(0, 0, 0), 14); assert.equal(shiftF('AAA100+$XFD$1', 1, 1), 'AAB101+$XFD$1');
});
