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
  const rule = native.match(/button\[data-close\]::before\s*\{[^}]*inset:\s*(-?\d+)px/);
  assert.ok(rule, 'native.js should widen the close button\'s clickable area past its drawn 12px circle');
  const inset = Number(rule[1]);
  assert.ok(inset <= -8, `inset (${inset}px) should reach well past the button, not just pad it by a pixel or two`);
  // found by an attribute rather than its title, which the English UI translates
  for (const page of ['MacSettings', 'MacAbout', 'MacGestures'])
    assert.match(readFileSync(new URL(`../${page}.dc.html`, import.meta.url), 'utf8'), /<button data-close="1" onClick="\{\{ close \}\}"/, page);
});

test('the lights show their symbols on hover and darken while pressed, like macOS', () => {
  // panels: the × sits in ::after (::before stays the hit box), the button darkens on :active and greys out with the window
  assert.match(native, /button\[data-close\]::after\{[^}]*opacity:0[^}]*svg/);
  assert.match(native, /button\[data-close\]:hover::after\{opacity:1\}/);
  assert.match(native, /button\[data-close\]:active\{filter:brightness\(0\.8\)\}/);
  assert.match(native, /html\.w-dim button\[data-close\]\{background:#D6D6D6!important/);
  assert.match(native, /classList\.toggle\('w-dim', !document\.hasFocus\(\)\)/);
  // document windows: one svg symbol per light, all three following the hover flag, the green one's arrows turning in for full screen
  const mac = readFileSync(new URL('../mac.dc.html', import.meta.url), 'utf8');
  const lights = mac.match(/<button data-light="1" onClick="\{\{ do(Close|Min|Zoom) \}\}"[^>]*style-active="filter:brightness\(0\.8\)"><svg[^>]*opacity:\{\{ gOp \}\}"/g) || [];
  assert.equal(lights.length, 3, 'close, minimize and zoom each carry a symbol bound to gOp, a pressed style and the wider hit box');
  assert.match(mac, /\[data-mactest\] \[data-light\]::before\{content:'';position:absolute;inset:-4px\}/);
  assert.match(mac, /gOp: st\.lh \? 1 : 0/);
  assert.match(mac, /style="display:\{\{ fsOut \}\}"/);
  assert.match(mac, /style="display:\{\{ fsIn \}\}"/);
  assert.match(mac, /fsOut: st\.full \? 'none' : 'block', fsIn: st\.full \? 'block' : 'none'/);
  assert.doesNotMatch(mac, /\{\{ g[123] \}\}/, 'the old text symbols are gone');
});

test('the drag mousedown listener still targets only the header background, not the button', () => {
  const dragHandler = native.match(/document\.addEventListener\('mousedown'[\s\S]*?\}\);/);
  assert.ok(dragHandler, 'native.js should still install the header-drag mousedown listener');
  assert.match(dragHandler[0], /#dc-root>\.sc-host>div>div:first-child/, 'the drag region must stay scoped to the header div, not the whole panel');
});
