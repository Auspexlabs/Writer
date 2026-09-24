// node --test ui/tests/
//
// desktop/src-tauri/src/native.js draws the settings/about/gestures panels' traffic-light close button
// at 12x12px inside a much taller header strip, and treats any mousedown on that header's own background
// as the start of a window drag (invoke('plugin:window|start_dragging')). Without a hit box wider than the
// drawn circle, a click a few pixels off the button lands on the header instead of the button: the drag
// starts (with no visible movement, so it looks like nothing happened) and the button's onClick — which
// calls props.onClose — never fires. Confirmed against the real page in a browser: a click just below the
// lights hit the header div and fired start_dragging; with the ::before hit box below it hits the button
// and fires aux_close instead.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const native = readFileSync(new URL('../../desktop/src-tauri/src/native.js', import.meta.url), 'utf8');

test('the close button gets a hit box wider than its 12px drawing', () => {
  const rule = native.match(/button\[title="关闭"\]::before\s*\{[^}]*inset:\s*(-?\d+)px/);
  assert.ok(rule, 'native.js should widen the close button\'s clickable area past its drawn 12px circle');
  const inset = Number(rule[1]);
  assert.ok(inset <= -8, `inset (${inset}px) should reach well past the button, not just pad it by a pixel or two`);
});

test('the drag mousedown listener still targets only the header background, not the button', () => {
  const dragHandler = native.match(/document\.addEventListener\('mousedown'[\s\S]*?\}\);/);
  assert.ok(dragHandler, 'native.js should still install the header-drag mousedown listener');
  assert.match(dragHandler[0], /#dc-root>\.sc-host>div>div:first-child/, 'the drag region must stay scoped to the header div, not the whole panel');
});
