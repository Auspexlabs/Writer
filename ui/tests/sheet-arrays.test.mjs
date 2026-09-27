import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Calc, shareCells } from '../sheet-engine.js';
const calc = cells => new Calc({ sheets: [{ name: 'Data', cells: shareCells(cells) }] });
const formula = text => calc({ A1: { v: '=' + text } }).value(0, 0, 0);
test('array spills calculate from a child first, provide # references and reject blocked or merged destinations', () => {
  const c = calc({ A1: { v: '=SEQUENCE(3,2)' }, D1: { v: '=SUM(A1#)' } });
  assert.equal(c.value(0, 2, 1), 6); assert.equal(c.value(0, 1, 0), 3); assert.equal(c.value(0, 0, 3), 21);
  assert.deepEqual(calc({ A1: { v: '=SEQUENCE(3)' }, A2: { v: 'occupied' } }).value(0, 0, 0), { err: '#SPILL!' });
  const merge = new Calc({ sheets: [{ cells: { A1: { v: '=SEQUENCE(2)' } }, merges: [{ r: 1, c: 0, rs: 1, cs: 2 }] }] });
  assert.deepEqual(merge.value(0, 0, 0), { err: '#SPILL!' });
});
test('loaded spill caches are recomputed rather than blocking the array', () => {
  const c = calc({ A1: { v: '=SEQUENCE(2,2,10)' }, B1: { v: 2, spill: 'A1' }, A2: { v: 3, spill: 'A1' }, B2: { v: 4, spill: 'A1' } });
  assert.equal(c.value(0, 1, 1), 13);
});
test('additional functions calculate representative values and errors', () => {
  assert.equal(formula('SUMSQ(3,4)'), 25); assert.ok(Math.abs(formula('SIN(PI()/2)') - 1) < 1e-12);
  for (const [i, roman] of ['CDXCIX','LDVLIV','XDIX','VDIV','ID'].entries()) assert.equal(formula(`ROMAN(499,${i})`), roman);
  assert.equal(formula('ARABIC("MCMXCIX")'), 1999);
  assert.equal(formula('HYPERLINK("https://example.com","open")'), 'open');
  assert.equal(formula('SUM(HSTACK({1;2},{3;4}))'), 10); assert.equal(formula('SUM(VSTACK({1,2},{3,4}))'), 10);
  assert.equal(formula('SUM(TAKE(SEQUENCE(3,3),-1,2))'), 15); assert.equal(formula('SUM(DROP(SEQUENCE(3,3),1,-1))'), 24);
  assert.equal(formula('REGEXTEST("ABC123","[a-z]+",1)'), true); assert.equal(formula('REGEXEXTRACT("ID:42","[0-9]+")'), '42'); assert.equal(formula('REGEXREPLACE("a1 b2","[0-9]","x",2)'), 'a1 bx');
  assert.equal(formula('DSUM({"name","amount";"a",2;"b",4;"a",3},"amount",{"name";"a"})'), 5);
});
