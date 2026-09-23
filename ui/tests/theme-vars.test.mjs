// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { extract, apply, strip, pageCss, SKIP_LITERAL } from '../tools/theme-vars.mjs';

/** A tiny designer export: the shell wraps #1D1D1F as text (--k1) twice and as a border (--k7) once, leaves the accent tint bare. */
function fakeExport() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'theme-'));
  fs.writeFileSync(path.join(dir, 'Writer.dc.html'), [
    '<helmet><link rel="stylesheet" href="./theme-dark.css"><style>',
    'html,body{background:var(--k0, #FFFFFF);color:var(--k1, #1D1D1F)}',
    '</style></helmet>',
    '<div style="color:var(--k1, #1D1D1F);border:1px solid var(--k7, #1D1D1F);box-shadow:0 1px 3px var(--k19, rgba(0,0,0,0.1)),inset 0 1px 0 var(--k18, rgba(255,255,255,0.95))">x</div>',
    '<div style="background:var(--k43, rgba(29,29,31,0.72));color:#fff;border:2px dashed rgba(63,125,92,0.6)">toast</div>',
    '<div style="background:linear-gradient(180deg,var(--k15, rgba(255,255,255,0.62)),var(--k16, rgba(255,255,255,0.34)));backdrop-filter:blur(10px)">bar</div>',
    '<script type="text/x-dc" data-dc-script>',
    "const a = { color: on ? 'var(--k1, #1D1D1F)' : 'var(--k11, #6E6E73)', sh: '0 1px 4px var(--k36, rgba(0,0,0,0.12)), inset 0 1px 0 var(--k51, #fff)' };",
    '</script>'
  ].join('\n'));
  fs.writeFileSync(path.join(dir, 'theme-dark.css'), 'html[data-theme="dark"]{--k1:#F5F5F7;--k15:rgba(28,28,30,0.8);--k16:rgba(28,28,30,0.52)}');
  return dir;
}

test('extract: most frequent variable per literal, ambiguities reported, designer lines and segments remembered', () => {
  const map = extract(fakeExport());
  assert.equal(map.vars['#1D1D1F'], '--k1');
  assert.deepEqual(map.ambiguous['#1D1D1F'], { '--k1': 3, '--k7': 1 });
  assert.equal(map.vars['#6E6E73'], '--k11');
  assert.equal(map.vars['rgba(63,125,92,0.6)'], undefined, 'a literal the designer never wrapped has no variable');
  assert.ok(Object.keys(map.lines).length >= 4 && Object.keys(map.segs).length >= 5);
  assert.deepEqual(map.segs[Object.keys(map.segs).find(k => map.segs[k].length === 3)], ['--k43', '', '']);
});

test('apply: designer lines get their exact variables, new lines the role/frequency mapping, document colours stay', () => {
  const map = extract(fakeExport());
  const src = [
    '<helmet>',
    '<style>',
    'html,body{background:#FFFFFF;color:#1D1D1F}',
    '.wd-ed [data-ai]{background:#DCEBDF;box-shadow:0 0 0 3px #DCEBDF}',
    '</style></helmet>',
    '<div style="color:#1D1D1F;border:1px solid #1D1D1F;box-shadow:0 1px 3px rgba(0,0,0,0.1),inset 0 1px 0 rgba(255,255,255,0.95)">x</div>',
    '<div style="background:rgba(29,29,31,0.72);color:#fff;border:2px dashed rgba(63,125,92,0.6)">toast</div>',
    '<span style="width:19px;border:1.7px solid #1D1D1F;color:#1D1D1F">icon</span>',
    '<div style="border-top:1px solid rgba(0,0,0,0.1);box-shadow:0 12px 32px rgba(0,0,0,0.1)">line vs shadow</div>',
    '<button style="background:#1D1D1F;color:#fff">go</button>',
    "<div data-props='{\"doc\":{\"html\":\"<p style=\\\"color:#1D1D1F\\\">x</p>\"}}'>",
    '<script type="text/x-dc" data-dc-script>',
    "const a = { color: on ? '#1D1D1F' : '#6E6E73', sh: '0 1px 4px rgba(0,0,0,0.12), inset 0 1px 0 #fff' };",
    "const b = { bg: sel ? '#1D1D1F' : 'transparent', fill: '#EFEFF4', t: `<b style=\"color:#6E6E73\">` };",
    "const FILLS = ['#FFE8A3', '#D9EAD3']; const PAL = ['#3F7D5C', '#E3B25A'];",
    "g.fillStyle = '#FFFFFF'; g.fillRect(0, 0, 1, 1);",
    "const keep = { fg: fill ? '#1D1D1F' : 'x' }; // theme: keep",
    "const DEF = { theme: 'light', accent: '#3F7D5C' };",
    "M('bd', '边框', [['#1D1D1F', '黑色'], ['#6E6E73', '灰色']].map(([c, n]) => I(n, () => this.style({ bdc: c }))));",
    '</script>'
  ].join('\n');
  const { text, summary } = apply(src, map);
  const L = text.split('\n');
  assert.equal(L[1], '<link rel="stylesheet" href="./theme-dark.css">', 'the dark theme is linked from the helmet');
  assert.equal(L[3], 'html,body{background:var(--k0, #FFFFFF);color:var(--k1, #1D1D1F)}');
  assert.equal(L[4], '.wd-ed [data-ai]{background:#DCEBDF;box-shadow:0 0 0 3px #DCEBDF}', 'a literal the designer never wrapped stays bare');
  assert.equal(L[6], '<div style="color:var(--k1, #1D1D1F);border:1px solid var(--k7, #1D1D1F);box-shadow:0 1px 3px var(--k19, rgba(0,0,0,0.1)),inset 0 1px 0 var(--k18, rgba(255,255,255,0.95))">x</div>', 'a designer line: exact variables per occurrence');
  assert.equal(L[7], '<div style="background:var(--k43, rgba(29,29,31,0.72));color:#fff;border:2px dashed rgba(63,125,92,0.6)">toast</div>', 'what the designer left bare stays bare');
  assert.equal(L[8], '<span style="width:19px;border:1.7px solid var(--k7, #1D1D1F);color:var(--k1, #1D1D1F)">icon</span>', 'role: border → --k7, text → --k1');
  assert.equal(L[9], '<div style="border-top:1px solid var(--k29, rgba(0,0,0,0.1));box-shadow:0 12px 32px var(--k19, rgba(0,0,0,0.1))">line vs shadow</div>');
  assert.equal(L[10], '<button style="background:var(--k7, #1D1D1F);color:var(--kinv, #fff)">go</button>', 'white text on the inverted button');
  assert.equal(L[11], src.split('\n')[10], 'data-props (default documents) untouched');
  assert.equal(L[13], "const a = { color: on ? 'var(--k1, #1D1D1F)' : 'var(--k11, #6E6E73)', sh: '0 1px 4px var(--k36, rgba(0,0,0,0.12)), inset 0 1px 0 var(--k51, #fff)' };", 'JS strings: designer segment reused');
  assert.equal(L[14], "const b = { bg: sel ? 'var(--k7, #1D1D1F)' : 'transparent', fill: '#EFEFF4', t: `<b style=\"color:var(--k11, #6E6E73)\">` };", 'JS key decides the role; sheet fill literal skipped; template literals handled');
  assert.equal(L[15], src.split('\n')[14], 'palettes stay literal');
  assert.equal(L[16], src.split('\n')[15], 'canvas exports stay literal');
  assert.equal(L[17], src.split('\n')[16], 'theme: keep opts a line out');
  assert.equal(L[18], src.split('\n')[17], 'settings defaults are stored values, not styles');
  assert.equal(L[19], src.split('\n')[18], 'an inline swatch list is a palette too: its colours go into the document');
  assert.ok(summary.skipped >= 4 && summary.changed >= 8);
  assert.ok(SKIP_LITERAL.has('#FFE066'), 'the PDF highlight is document content');
});

test('apply is idempotent and never double-wraps a partly converted file', () => {
  const map = extract(fakeExport());
  const src = '<div style="color:var(--k1, #1D1D1F);background:#FFFFFF;box-shadow:0 1px 3px var(--k19, rgba(0,0,0,0.1))">x</div>\n<script type="text/x-dc" data-dc-script>\nconst c = on ? \'var(--k7, #1D1D1F)\' : \'#1D1D1F\';\n</script>';
  const once = apply(src, map).text, twice = apply(once, map).text;
  assert.equal(twice, once);
  assert.equal(once.split('\n')[0], '<div style="color:var(--k1, #1D1D1F);background:var(--k0, #FFFFFF);box-shadow:0 1px 3px var(--k19, rgba(0,0,0,0.1))">x</div>');
  assert.equal(once.split('\n')[2], "const c = on ? 'var(--k7, #1D1D1F)' : 'var(--k1, #1D1D1F)';", 'an already wrapped value keeps its variable');
  assert.doesNotMatch(once, /var\(--k\d+, var\(/);
  assert.equal(strip(once), strip(src));
});

test('pageCss: light palette for document pages, alpha rescaling for the frosted surfaces only', () => {
  const map = extract(fakeExport()), css = pageCss(map);
  assert.deepEqual(map.glass, ['--k15', '--k16'], 'only variables painting a backdrop-filter surface');
  assert.equal(map.light['--k1'], '#1D1D1F');
  assert.equal(map.dark['--k15'], 'rgba(28,28,30,0.8)');
  assert.match(css, /\[data-light\]\{color-scheme:light;--k0:#FFFFFF;--k1:#1D1D1F;/);
  assert.match(css, /html\[data-glass\]\{--k15:rgba\(255,255,255,clamp\(0\.03,calc\(1 - 0\.38 \* var\(--gt, 1\)\),1\)\);--k16:/);
  assert.match(css, /html\[data-glass\]\[data-theme="dark"\]\{--k15:rgba\(28,28,30,clamp\(0\.03,calc\(1 - 0\.2 \* var\(--gt, 1\)\),1\)\)/);
  assert.doesNotMatch(css, /--k43:rgba\(29,29,31,clamp/, 'the toast has no backdrop-filter in this export');
});

test('a designer segment only fills bare literals; scrims stay black', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'theme-'));
  fs.writeFileSync(path.join(dir, 'Mac.dc.html'), ['<script type="text/x-dc" data-dc-script>', "const a = 'var(--goff, rgba(0,0,0,0.03))';", '</script>'].join('\n'));
  const map = extract(dir);
  const src = ['<div style="position:absolute;inset:0;z-index:60;background:rgba(0,0,0,0.12);display:flex">', '<script type="text/x-dc" data-dc-script>', "const b = on ? 'var(--k36, rgba(0,0,0,0.03))' : 'rgba(0,0,0,0.03)';", '</script>'].join('\n');
  const L = apply(src, map).text.split('\n');
  assert.equal(L[0], src.split('\n')[0], 'full-cover backdrop keeps its black');
  assert.equal(L[2], "const b = on ? 'var(--k36, rgba(0,0,0,0.03))' : 'var(--goff, rgba(0,0,0,0.03))';", 'existing wrapper kept, bare one filled');
});
