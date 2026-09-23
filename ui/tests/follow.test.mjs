// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { planScroll } from '../follow.js';

// Scroller client box 1000 × 700 at (0, 0); default insets (56 top, 48 bottom) → visible box y 56…652 (596 tall).
// Comfort box: x 120…880, y 56 + 119.2 = 175.2 … 652 − 119.2 = 532.8.
const view = { left: 0, top: 0, width: 1000, height: 700 };
const caret = (left, top, h = 20) => ({ left, top, right: left, bottom: top + h, width: 0, height: h });
const cur = { x: 300, y: 400 }, max = { x: 2000, y: 3000 };
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 1e-6, `${msg}: ${a} ≠ ${b}`);

test('inside the comfort box: unchanged', () => {
  assert.deepEqual(planScroll(caret(500, 300), view, cur, max), cur);
});

test('past the right edge: caret lands on the right box edge', () => {
  const to = planScroll(caret(950, 300), view, cur, max);
  near(to.x, 300 + (950 - 880), 'x'); assert.equal(to.y, 400);
});

test('clamped at max', () => {
  assert.deepEqual(planScroll(caret(1990, 300), view, cur, { x: 320, y: 3000 }), { x: 320, y: 400 });
  assert.deepEqual(planScroll(caret(5, 300), view, { x: 20, y: 400 }, max), { x: 0, y: 400 });
});

test('line wrap: caret x moved left by more than half the view → left box edge', () => {
  // previous caret at content x 2160 (client 860 + scroll 1300); now at client 200 (content 1500): moved 660 left
  const to = planScroll(caret(200, 300), view, { x: 1300, y: 400 }, max, { prevX: 2160 });
  near(to.x, 1300 + (200 - 120), 'x');
  // the same caret without wrap memory sits inside the box: nothing happens
  assert.equal(planScroll(caret(200, 300), view, { x: 1300, y: 400 }, max).x, 1300);
});

test('below the box: caret lands on the bottom box edge', () => {
  const to = planScroll(caret(500, 640), view, cur, max);
  near(to.y, 400 + (660 - 532.8), 'y'); assert.equal(to.x, 300);
});

test('above the box: caret lands on the top box edge', () => {
  near(planScroll(caret(500, 60), view, cur, max).y, 400 + (60 - 175.2), 'y');
});

test('content fits on an axis: that axis is untouched', () => {
  assert.deepEqual(planScroll(caret(990, 690), view, cur, { x: 0, y: 0 }), cur);
  assert.equal(planScroll(caret(990, 690), view, cur, { x: 0, y: 3000 }).x, 300);
});

test('insets shrink the visible box; padX / padY are honoured', () => {
  const to = planScroll(caret(500, 640), view, cur, max, { insets: { top: 0, bottom: 0 } });
  near(to.y, 400 + (660 - 560), 'no insets → box edge at 80% of 700');
  const t2 = planScroll(caret(500, 640), view, cur, max, { insets: { top: 100, bottom: 200 }, padY: 0.1 });
  near(t2.y, 400 + (660 - (100 + 400 * 0.9)), 'custom insets and padY');
  near(planScroll(caret(950, 300), view, cur, max, { padX: 0.2 }).x, 300 + (950 - 800), 'padX 0.2');
});

test('a rect wider than the box keeps its start visible', () => {
  const to = planScroll({ left: 200, top: 300, right: 1800, bottom: 320 }, view, cur, max);
  near(to.x, 300 + (200 - 120), 'start at the left box edge, not pushed off by the right overflow');
});
