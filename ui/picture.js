// picture.js — the 图片 tab shared by the Word, slide and sheet editors. A picture's look is the engine's picture props as
// `get` prints them ({ crop: '10,0,10,0', brightness: '20', grayscale: 'true', … }); the engine writes them as Office markup,
// and here they are drawn with CSS, edited by the tab's tools and cropped with handles.

export const SHAPES = [['rect', '矩形'], ['roundRect', '圆角矩形'], ['ellipse', '椭圆'], ['triangle', '三角形'], ['diamond', '菱形'], ['hexagon', '六边形'], ['star5', '五角星']];
export const RATIOS = [['1:1', '1:1 方形'], ['4:3', '4:3'], ['3:2', '3:2'], ['16:9', '16:9 宽屏'], ['3:4', '3:4 竖版'], ['9:16', '9:16 竖版']];
export const WIDTHS = [['none', '无边框'], ['0.75pt', '0.75 磅'], ['1.5pt', '1.5 磅'], ['3pt', '3 磅'], ['4.5pt', '4.5 磅'], ['6pt', '6 磅']];
export const COMPRESS = [['print', '打印（220 ppi）'], ['web', '网页（150 ppi）'], ['email', '电子邮件（96 ppi）']];

const num = (v, d = 0) => { const n = parseFloat(v); return isFinite(n) ? n : d; };
const pct = v => +(v * 100).toFixed(3) + '%';
// $t under node (this module is node-tested): falls back to the Chinese, vars filled the same way.
const T = (s, v) => globalThis.$t ? globalThis.$t(s, v) : v ? String(s).replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : s;

/** The crop as left, top, right, bottom fractions of the whole picture. */
export function cropOf(look) { const c = String((look && look.crop) || '').split(',').map(v => num(v) / 100); return c.length === 4 ? c : [0, 0, 0, 0]; }
/** Fractions → the engine's crop prop, in percent. */
export const cropText = c => c.map(v => +(Math.max(0, v) * 100).toFixed(3)).join(',');

/** The crop that shows the largest centred part of the whole picture at ratio "w:h", for a frame now w × h (any unit). */
export function cropToRatio(look, w, h, ratio) {
  const [l, t, r, b] = cropOf(look), [rw, rh] = String(ratio).split(':').map(Number);
  const fw = w / Math.max(0.01, 1 - l - r), fh = h / Math.max(0.01, 1 - t - b), want = rw / rh;
  if (fw / fh > want) { const side = (1 - want * fh / fw) / 2; return cropText([side, 0, side, 0]); }
  const side = (1 - fw / want / fh) / 2;
  return cropText([0, side, 0, side]);
}

/** The look the props leave: every picture prop the engine reports, and none with its neutral value. */
export function lookFrom(props) {
  const look = {};
  for (const k of ['crop', 'rotation', 'flipH', 'flipV', 'brightness', 'contrast', 'grayscale', 'transparency', 'line', 'lineWidth', 'shadow', 'geometry']) {
    const v = props && props[k];
    if (v == null || v === '' || v === 'false' || v === 'none' || (k === 'geometry' && v === 'rect')) continue;
    if ((['rotation', 'brightness', 'contrast', 'transparency'].includes(k) && num(v) === 0) || (k === 'crop' && cropOf({ crop: v }).every(x => !x))) continue;
    look[k] = String(v);
  }
  if (!look.line) delete look.lineWidth;
  return look;
}

// Office's shapes as polygons in a unit box (x, y from the top left), with the default adjust values.
const POLYGONS = {
  triangle: [[0.5, 0], [1, 1], [0, 1]],
  diamond: [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]],
  star5: Array.from({ length: 10 }, (_, i) => { const a = -Math.PI / 2 + i * Math.PI / 5, rr = i % 2 ? 0.382 : 1; return [0.5 + 0.5 * rr * Math.cos(a) / 0.951, (rr * Math.sin(a) + 1) / 1.809]; }),
};

const lumIds = new Set();
/** Office's a:lum as an SVG filter in this page: contrast scales about mid-grey, then brightness moves every colour toward
 *  white (above 0) or black (below). Returns the CSS filter that uses it. */
export function lumFilter(bright, contrast) {
  const id = 'wlum' + Math.round(bright * 100) + '_' + Math.round(contrast * 100);
  if (typeof document !== 'undefined' && !lumIds.has(id)) {
    lumIds.add(id);
    let defs = document.getElementById('wlum-defs');
    if (!defs) { defs = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); defs.id = 'wlum-defs'; defs.setAttribute('aria-hidden', 'true'); defs.style.cssText = 'position:absolute;width:0;height:0;overflow:hidden'; document.body.appendChild(defs); }
    const k = contrast >= 0 ? 1 / Math.max(0.01, 1 - contrast) : 1 + contrast, [s, i] = bright >= 0 ? [1 - bright, bright] : [1 + bright, 0];
    const f = (slope, icpt) => '<feComponentTransfer>' + ['R', 'G', 'B'].map(c => `<feFunc${c} type="linear" slope="${slope}" intercept="${icpt}"/>`).join('') + '</feComponentTransfer>';
    defs.insertAdjacentHTML('beforeend', `<filter id="${id}" color-interpolation-filters="sRGB">${f(k, 0.5 - 0.5 * k)}${f(s, i)}</filter>`);
  }
  return `url(#${id})`;
}

/** How a picture with this look is drawn in a w × h frame (w, h in the frame's CSS px; ptPx = those px per point). The frame
 *  element clips (`frame`: shape, border and shadow outside it); the image inside it is placed and sized to show the crop,
 *  mirrored about the frame's centre and coloured (`image`). Rotation belongs to whatever places the frame. */
export function pictureView(look, w, h, ptPx) {
  look = look || {};
  const [l, t, r, b] = cropOf(look), vw = Math.max(0.01, 1 - l - r), vh = Math.max(0.01, 1 - t - b);
  const fx = look.flipH === 'true' ? -1 : 1, fy = look.flipV === 'true' ? -1 : 1;
  const filters = [], bright = num(look.brightness) / 100, contrast = num(look.contrast) / 100;
  if (bright || contrast) filters.push(lumFilter(bright, contrast));
  if (look.grayscale === 'true') filters.push('grayscale(1)');
  const g = look.geometry || 'rect', m = Math.min(w, h);
  let radius = '0', clip = 'none';
  if (g === 'roundRect') radius = +(m * 0.16667).toFixed(2) + 'px';
  else if (g === 'ellipse') radius = '50%';
  else if (g === 'hexagon') { const i = +(m * 0.25).toFixed(2); clip = `polygon(${i}px 0,calc(100% - ${i}px) 0,100% 50%,calc(100% - ${i}px) 100%,${i}px 100%,0 50%)`; }
  else if (POLYGONS[g]) clip = 'polygon(' + POLYGONS[g].map(([x, y]) => pct(fx < 0 ? 1 - x : x) + ' ' + pct(fy < 0 ? 1 - y : y)).join(',') + ')';
  const shadows = [], lw = look.line && look.line !== 'none' ? +(num(look.lineWidth, 0.75) * ptPx).toFixed(2) : 0;
  if (lw) shadows.push(`0 0 0 ${lw}px #${look.line}`);
  // Office's shadow falls from the picture with its border: spread by the border, so the border does not hide it
  if (look.shadow === 'true') { const d = +(3 * Math.SQRT1_2 * ptPx).toFixed(2); shadows.push(`${d}px ${d}px ${+(4 * ptPx).toFixed(2)}px ${lw}px rgba(0,0,0,0.4)`); }
  const v = {
    left: pct(-l / vw), top: pct(-t / vh), width: pct(1 / vw), height: pct(1 / vh),
    flip: fx < 0 || fy < 0 ? `scale(${fx},${fy})` : 'none', origin: pct((1 + l - r) / 2) + ' ' + pct((1 + t - b) / 2),
    filter: filters.join(' ') || 'none', opacity: String(1 - num(look.transparency) / 100),
    radius, clip, shadow: shadows.join(',') || 'none',
  };
  v.frame = `overflow:hidden;border-radius:${v.radius};clip-path:${v.clip};box-shadow:${v.shadow}`;
  v.image = `position:absolute;left:${v.left};top:${v.top};width:${v.width};height:${v.height};max-width:none;max-height:none;margin:0;transform:${v.flip};transform-origin:${v.origin};filter:${v.filter};opacity:${v.opacity}`;
  return v;
}

// A picture's bytes change under the same URL after 抠图, 压缩 and 替换: every place that draws it asks for picSrc(url),
// which carries a version the browser has not cached yet.
const versions = new Map();
const bare = url => String(url || '').replace(/&v=\d+$/, '');
export const picSrc = url => url && versions.has(bare(url)) ? bare(url) + '&v=' + versions.get(bare(url)) : url;
export function bumpPic(url) { const u = bare(url); versions.set(u, (versions.get(u) || 0) + 1); return picSrc(u); }

const size = n => n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
/** What a compress did, in a sentence. */
export const savedText = (before, after) => after < before ? T('已压缩：{before} → {after}，节省 {saved}', { before: size(before), after: size(after), saved: size(before - after) }) : T('这张图片已经不大于显示所需，没有再压缩');

/** A picture file the user picks, as a data URL (null when they cancel). */
export function pickImage() {
  return new Promise(resolve => {
    const input = document.createElement('input');
    input.type = 'file'; input.accept = 'image/png,image/jpeg,image/gif,image/bmp';
    input.onchange = () => { const f = input.files[0]; if (!f) return resolve(null); const rd = new FileReader(); rd.onload = () => resolve(rd.result); rd.onerror = () => resolve(null); rd.readAsDataURL(f); };
    input.click();
  });
}

/**
 * The tools of one selected picture. `host` is the editor's side:
 *   look()          the picture's look in the model now
 *   show(look)      draws a look at once (the live preview)
 *   send(props)     runs one `set <picture> --prop …` through engine.js; resolves with the picture's props afterwards
 *   commit(props, binary)  takes that answer into the model and its saved snapshot (binary: the image itself changed)
 *   busy(op)        a slow tool started (op) or ended (null)
 *   frame()         the frame as { cx, cy, w, h } on screen (px, unrotated size), for cropping with handles
 *   src()           the picture's URL; toast(message)
 * A slider's value is sent once it rests; everything else at once. A failed command puts the model's look back.
 */
export function pictureTools(host) {
  let timer = null, pending = null;
  const fail = e => { host.show(host.look()); host.toast && host.toast(e && e.message ? e.message + (e.hint ? '（' + e.hint + '）' : '') : T('图片没有改成')); }; // i18n-ok — parens wrap a dynamic engine message, not translated text
  const send = async props => { try { host.commit(await host.send(props), false); } catch (e) { fail(e); } };
  const tools = {
    set(patch, rest) {
      host.show(Object.assign({}, host.look(), pending, patch));
      if (!rest) { clearTimeout(timer); const p = Object.assign({}, pending, patch); pending = null; return send(p); }
      pending = Object.assign({}, pending, patch); clearTimeout(timer);
      timer = setTimeout(() => { const p = pending; pending = null; send(p); }, 350);
    },
    async run(op, props) {
      host.busy(op);
      try {
        const r = await host.send(props);
        host.commit(r, true);
        if (r.before != null && host.toast) host.toast(savedText(r.before, +r.props.bytes || 0));
      } catch (e) { fail(e); } finally { host.busy(null); }
    },
    cutout: () => tools.run('cutout', { background: 'remove' }),
    compress: level => tools.run('compress', { compress: level || 'print' }),
    reset: () => tools.run('reset', { reset: 'true' }),
    async replace() { const src = await pickImage(); if (src) await tools.run('replace', { src }); },
    ratio(r) { const f = host.frame(); tools.set({ crop: cropToRatio(host.look(), f.w, f.h, r) }); },
    async crop() {
      const look = host.look(), f = host.frame();
      const next = await cropBox({ frame: f, crop: cropOf(look), rotation: num(look.rotation), flipH: look.flipH === 'true', flipV: look.flipV === 'true', src: host.src() });
      if (next) tools.set({ crop: cropText(next) });
    },
  };
  return tools;
}

/** The ribbon once the selected picture changed (`key` is the editor's 图片 tab): a ribbon open on another tab turns to 图片 when a
 *  picture is selected; open on 图片, it closes once none is. Returns the state to set, or null. */
export function tabAfterSelect(st, key, selected) {
  if (selected) return st.bubble && st.tab !== key ? { tab: key } : null;
  return st.tab === key ? { tab: 'home', bubble: false } : null;
}

/** The 图片 tab's items, made with the editor's ribbon makers k = { B, M, I, C, SEP } (and sliders made here). */
export function pictureRibbon(k, tools, look, busy) {
  look = look || {};
  const { B, M, I, C, SEP } = k, on = key => look[key] === 'true', rot = num(look.rotation);
  const R = (label, value, min, max, key) => ({ isRange: true, label: T(label), title: T(label), min, max, value, text: String(value), onChange: e => tools.set({ [key]: String(Math.round(+e.target.value)) }, true) });
  const line = look.line && look.line !== 'none' ? '#' + look.line : '';
  return [
    B(busy === 'cutout' ? '抠图中…' : '抠图', () => { if (!busy) tools.cutout(); }, { title: '移除背景，只留下主体（Apple Vision，macOS 14 以上）', dis: !!busy }),
    M('pic-crop', '裁剪', [I('拖动裁剪…', () => tools.crop(), { hint: '拖边角' })].concat(RATIOS.map(([r, l]) => I(l, () => tools.ratio(r))), [I('取消裁剪', () => tools.set({ crop: '0,0,0,0' }), { on: !look.crop })]), '裁剪到比例，或拖动边角'),
    M('pic-shape', '形状', SHAPES.map(([g, l]) => I(l, () => tools.set({ geometry: g }), { on: (look.geometry || 'rect') === g })), '裁剪为形状'), SEP,
    B('旋转 90°', () => tools.set({ rotation: String((rot + 90) % 360) }), { title: '向右旋转 90°' }),
    B('水平翻转', () => tools.set({ flipH: on('flipH') ? 'false' : 'true' }), { on: on('flipH') }),
    B('垂直翻转', () => tools.set({ flipV: on('flipV') ? 'false' : 'true' }), { on: on('flipV') }), SEP,
    R('亮度', num(look.brightness), -100, 100, 'brightness'), R('对比度', num(look.contrast), -100, 100, 'contrast'),
    B('灰度', () => tools.set({ grayscale: on('grayscale') ? 'false' : 'true' }), { on: on('grayscale') }),
    R('透明度', num(look.transparency), 0, 100, 'transparency'), SEP,
    C('边框', line || 'transparent', e => tools.set({ line: e.target.value.slice(1).toUpperCase(), lineWidth: look.lineWidth || '1.5pt' }), '边框颜色'),
    M('pic-width', '粗细', WIDTHS.map(([w, l]) => I(l, () => tools.set(w === 'none' ? { line: 'none' } : { lineWidth: w, line: look.line || '1D1D1F' }), { on: w === 'none' ? !line : !!line && num(look.lineWidth, 0.75) === num(w) })), '边框粗细'),
    B('阴影', () => tools.set({ shadow: on('shadow') ? 'false' : 'true' }), { on: on('shadow') }),
    B('圆角', () => tools.set({ geometry: look.geometry === 'roundRect' ? 'rect' : 'roundRect' }), { on: look.geometry === 'roundRect' }), SEP,
    M('pic-compress', busy === 'compress' ? '压缩中…' : '压缩图片', COMPRESS.map(([v, l]) => I(l, () => { if (!busy) tools.compress(v); })), '按显示尺寸重新编码，并删除裁掉的部分'),
    B(busy === 'reset' ? '重置中…' : '重置图片', () => { if (!busy) tools.reset(); }, { title: '去掉所有调整，恢复原图', dis: !!busy }),
    B(busy === 'replace' ? '替换中…' : '替换图片', () => { if (!busy) tools.replace(); }, { title: '换一张图片，保留宽度与位置', dis: !!busy }),
  ];
}

// ---- where a Word picture sits: in the line of text, or floating in its paragraph (the engine's wrap / x / y / xFrom / yFrom /
// xAlign / yAlign, as `get` prints them, lengths in cm) ----

export const PLACE = ['wrap', 'x', 'y', 'xFrom', 'yFrom', 'xAlign', 'yAlign'];
const UNIT_PX = { cm: 96 / 2.54, mm: 9.6 / 2.54, in: 96, pt: 96 / 72, px: 1, emu: 96 / 914400 };
/** A length as the engine prints it ("2.5cm", "12pt", "96px"; a bare number is cm) in px. */
export function lengthPx(s) { const m = /^(-?[\d.]+)\s*([a-z]+)?$/i.exec(String(s == null ? '' : s).trim()); return m ? +m[1] * (UNIT_PX[(m[2] || 'cm').toLowerCase()] || UNIT_PX.cm) : 0; }
export const cmText = px => +(px / UNIT_PX.cm).toFixed(3) + 'cm';
export const floating = place => !!place && !!place.wrap && place.wrap !== 'inline';

/** How a picture placed by `place` sits in its paragraph, as CSS properties for its frame (w × h px): square, tight and through
 *  float at their offset with the text beside them (on the side they leave free), topBottom breaks the text, front and behind lie
 *  over and under it. geo, once the editor knows it: { colW, pageW, pageH, mL, mT, top } — the text column's width, the page, its
 *  margins and the paragraph's top on its page, px; without it, offsets count from the column and the paragraph. Every property is
 *  set, so applying it again undoes the last placement. ponytail: a float wraps text on one side only and a page-relative picture
 *  sits absolutely without wrapping — real Word wrapping needs a layout engine. */
export function placeStyle(place, w, h, geo) {
  const s = { position: '', left: '', top: '', float: '', clear: '', margin: '', zIndex: '' };
  if (!floating(place)) return s;
  const p = place, g = geo || {}, colW = g.colW || 0, W = p.xFrom === 'page' ? g.pageW || colW : colW, H = p.yFrom === 'page' ? g.pageH || 0 : p.yFrom === 'margin' ? (g.pageH || 0) - 2 * (g.mT || 0) : 0;
  let x = p.xAlign ? ({ center: (W - w) / 2, right: W - w, outside: W - w }[p.xAlign] || 0) : lengthPx(p.x);
  if (p.xFrom === 'page') x -= g.mL || 0;
  let y = p.yAlign ? ({ center: (H - h) / 2, bottom: H - h, outside: H - h }[p.yAlign] || 0) : lengthPx(p.y);
  if (p.yFrom === 'page') y -= g.top || 0; else if (p.yFrom === 'margin') y += (g.mT || 0) - (g.top || 0);
  const r = v => Math.round(v * 100) / 100, flows = (!p.yFrom || p.yFrom === 'paragraph' || p.yFrom === 'line') && !p.yAlign;
  if ((p.wrap === 'square' || p.wrap === 'tight' || p.wrap === 'through') && flows) {
    const right = colW > 0 && x + w / 2 > colW / 2;
    return Object.assign(s, { position: 'relative', float: right ? 'right' : 'left', margin: right ? `${r(y)}px ${r(Math.max(0, colW - x - w))}px 8px 12px` : `${r(y)}px 12px 8px ${r(x)}px` });
  }
  if (p.wrap === 'topBottom' && flows) return Object.assign(s, { position: 'relative', float: 'left', clear: 'both', margin: `${r(y)}px calc(100% - ${r(x + w)}px) 8px ${r(x)}px` }); // a margin box as wide as the column: the text goes on below
  return Object.assign(s, { position: 'absolute', left: r(x) + 'px', top: r(y) + 'px', margin: '0', zIndex: p.wrap === 'behind' ? '-1' : '1' });
}

/** A frame resized by a handle: `edge` is l, r, t, b or a corner (lt, rt, lb, rb); dx, dy the pointer's move along the frame's own
 *  axes. A corner keeps the aspect (the axis moved more decides) unless `free`; an edge stretches one side. Never under 16 px. */
export function resizeMath(w0, h0, edge, dx, dy, free) {
  const sx = edge.includes('l') ? -1 : edge.includes('r') ? 1 : 0, sy = edge.includes('t') ? -1 : edge.includes('b') ? 1 : 0;
  let w = Math.max(16, w0 + sx * dx), h = Math.max(16, h0 + sy * dy);
  if (sx && sy && !free) { const k = Math.abs(dx) >= Math.abs(dy) ? w / w0 : h / h0; w = Math.max(16, w0 * k); h = Math.max(16, h0 * k); }
  return { w: Math.round(w), h: Math.round(h) };
}

/** A floating picture nudged by an arrow key: `step` px from where it stands now in its paragraph (`at`: { x, y } px), as an offset from
 *  the column and the paragraph (an alignment becomes the offset it stood for). Null for an inline picture, whose arrows move the caret. */
export function nudgePlace(place, key, step, at) {
  const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[key];
  if (!d || !floating(place)) return null;
  const { xAlign, yAlign, ...p } = place;
  return Object.assign(p, { xFrom: 'column', yFrom: 'paragraph', x: cmText(at.x + d[0]), y: cmText(at.y + d[1]) });
}

/** Resize handles on a selected picture: eight squares on its frame's edges and corners, over everything (a fixed layer), turned with
 *  the picture. host: frame() → { cx, cy, w, h } on screen (px, unrotated) or null once it is gone; rotation() in degrees; resize(w, h)
 *  while a handle drags; resized(w, h) on release (screen px). Returns { layout, remove }; layout() again after a scroll, zoom or edit. */
export function handleBox(host) {
  const layer = document.createElement('div');
  layer.setAttribute('data-handles', '1');
  layer.style.cssText = 'position:fixed;left:0;top:0;width:0;height:0;z-index:60;pointer-events:none';
  const HANDLES = [['lt', 0, 0, 'nwse'], ['t', 0.5, 0, 'ns'], ['rt', 1, 0, 'nesw'], ['r', 1, 0.5, 'ew'], ['rb', 1, 1, 'nwse'], ['b', 0.5, 1, 'ns'], ['lb', 0, 1, 'nesw'], ['l', 0, 0.5, 'ew']];
  const handles = HANDLES.map(([edge, hx, hy, cur]) => {
    const h = document.createElement('span');
    h.setAttribute('data-edge', edge);
    h.style.cssText = `position:absolute;width:10px;height:10px;margin:-5px 0 0 -5px;border-radius:2px;background:var(--k0, #FFFFFF);border:1.5px solid var(--k59, #3F7D5C);box-sizing:border-box;cursor:${cur}-resize;pointer-events:auto`;
    layer.appendChild(h);
    return [h, hx, hy];
  });
  let f = null;
  const rot = () => (host.rotation ? host.rotation() : 0) || 0;
  const layout = () => {
    f = host.frame();
    layer.style.display = f ? '' : 'none';
    if (!f) return;
    layer.style.transform = `translate(${f.cx}px,${f.cy}px) rotate(${rot()}deg)`;
    handles.forEach(([el, hx, hy]) => { el.style.left = (hx - 0.5) * f.w + 'px'; el.style.top = (hy - 0.5) * f.h + 'px'; });
  };
  layer.addEventListener('pointerdown', e => {
    const edge = e.target.getAttribute && e.target.getAttribute('data-edge');
    if (!edge || !f) return;
    e.preventDefault(); e.stopPropagation();
    const d = { x: e.clientX, y: e.clientY, w: f.w, h: f.h, a: -rot() * Math.PI / 180, r: null };
    const move = ev => {
      const dx = ev.clientX - d.x, dy = ev.clientY - d.y; // as a move along the unturned frame's axes
      d.r = resizeMath(d.w, d.h, edge, dx * Math.cos(d.a) - dy * Math.sin(d.a), dx * Math.sin(d.a) + dy * Math.cos(d.a), ev.shiftKey);
      host.resize(d.r.w, d.r.h); layout();
    };
    const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); if (d.r) host.resized(d.r.w, d.r.h); layout(); };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  });
  document.body.appendChild(layer);
  layout();
  return { layout, remove: () => layer.remove() };
}

/** Drags a picture with the pointer, from the pointerdown `e` on it (not the browser's own drag): nothing under 4 px; then a translucent
 *  ghost of `frame` follows the pointer and move(x, y, ev) gets the ghost's top left on screen, drop(x, y, ev) on release, cancel() on Esc
 *  or a press that never moved. */
export function dragPicture(e, frame, { move, drop, cancel }) {
  const r = frame.getBoundingClientRect(), gx = e.clientX - r.left, gy = e.clientY - r.top, x0 = e.clientX, y0 = e.clientY;
  const img = frame.matches('img') ? frame : frame.querySelector('img');
  let ghost = null;
  const at = ev => [ev.clientX - gx, ev.clientY - gy];
  const onMove = ev => {
    if (!ghost) {
      if (Math.hypot(ev.clientX - x0, ev.clientY - y0) < 4) return;
      ghost = document.createElement('div');
      ghost.style.cssText = `position:fixed;left:0;top:0;width:${r.width}px;height:${r.height}px;z-index:70;pointer-events:none;opacity:.55;background:url("${img ? img.src : ''}") center/100% 100% no-repeat;outline:1.5px dashed var(--k59, #3F7D5C)`;
      document.body.appendChild(ghost);
    }
    const [x, y] = at(ev);
    ghost.style.transform = `translate(${x}px,${y}px)`;
    move && move(x, y, ev);
  };
  const end = () => { window.removeEventListener('pointermove', onMove); window.removeEventListener('pointerup', onUp); document.removeEventListener('keydown', onKey, true); ghost && ghost.remove(); };
  const onUp = ev => { const moved = !!ghost; end(); if (moved) drop(...at(ev), ev); else cancel && cancel(); };
  const onKey = ev => { if (ev.key === 'Escape') { ev.preventDefault(); end(); cancel && cancel(); } };
  window.addEventListener('pointermove', onMove); window.addEventListener('pointerup', onUp); document.addEventListener('keydown', onKey, true);
}

/** The caret under a point on screen, as a collapsed Range (null where there is none). */
export function caretAt(x, y) {
  if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
  const p = document.caretPositionFromPoint && document.caretPositionFromPoint(x, y);
  if (!p) return null;
  const r = document.createRange(); r.setStart(p.offsetNode, p.offset); r.collapse(true);
  return r;
}
/** The insertion mark a drag shows: show(range) draws a bar at that caret, hide() takes it away. */
export function caretMark() {
  let bar = null;
  return {
    show(range) {
      let rc = range.getBoundingClientRect();
      if (!rc.height) { const c = range.startContainer, el = c.nodeType === 1 ? c : c.parentElement; if (el) rc = el.getBoundingClientRect(); } // a caret in an empty block: the block's own line
      if (!bar) { bar = document.createElement('div'); bar.style.cssText = 'position:fixed;width:2px;border-radius:1px;background:var(--k59, #3F7D5C);z-index:70;pointer-events:none'; document.body.appendChild(bar); }
      Object.assign(bar.style, { left: rc.left - 1 + 'px', top: rc.top + 'px', height: (rc.height || 20) + 'px' });
    },
    hide() { bar && bar.remove(); bar = null; }
  };
}

/**
 * Crop with handles: the whole picture shows dimmed around the part kept, whose edges and corners drag (inside it, drag to
 * move it). `frame` is the picture's frame on screen { cx, cy, w, h } (unrotated size); the box turns and mirrors with the
 * picture. Resolves with the new crop as fractions, or null when cancelled (Esc, 取消).
 */
export function cropBox({ frame, crop, rotation, flipH, flipV, src }) {
  return new Promise(resolve => {
    let [l, t, r, b] = crop;
    const FW = frame.w / Math.max(0.01, 1 - l - r), FH = frame.h / Math.max(0.01, 1 - t - b); // the whole picture on screen
    const cx0 = (l + (1 - l - r) / 2) * FW, cy0 = (t + (1 - t - b) / 2) * FH; // the frame's centre on the whole picture
    const layer = document.createElement('div');
    layer.setAttribute('data-crop', '1');
    layer.style.cssText = 'position:fixed;inset:0;z-index:120;cursor:default;touch-action:none';
    const stage = document.createElement('div');
    stage.style.cssText = `position:absolute;left:${frame.cx - cx0}px;top:${frame.cy - cy0}px;width:${FW}px;height:${FH}px;transform-origin:${cx0}px ${cy0}px;transform:rotate(${rotation || 0}deg) scale(${flipH ? -1 : 1},${flipV ? -1 : 1})`;
    const picture = () => { const i = document.createElement('img'); i.src = src; i.draggable = false; i.style.cssText = `position:absolute;left:0;top:0;width:${FW}px;height:${FH}px;max-width:none;pointer-events:none`; return i; };
    const dim = document.createElement('div'), win = document.createElement('div'), inner = picture();
    dim.style.cssText = 'position:absolute;inset:0;opacity:0.38';
    dim.appendChild(picture());
    win.setAttribute('data-win', '1');
    win.style.cssText = 'position:absolute;overflow:hidden;cursor:move;outline:1.5px solid var(--k59, #3F7D5C);box-shadow:0 0 0 1px rgba(255,255,255,0.6)';
    win.appendChild(inner);
    stage.append(dim, win);
    const HANDLES = [['l', 0, 0.5, 'ew'], ['r', 1, 0.5, 'ew'], ['t', 0.5, 0, 'ns'], ['b', 0.5, 1, 'ns'], ['lt', 0, 0, 'nwse'], ['rt', 1, 0, 'nesw'], ['lb', 0, 1, 'nesw'], ['rb', 1, 1, 'nwse']];
    const handles = HANDLES.map(([edge, hx, hy, cur]) => {
      const h = document.createElement('span');
      h.setAttribute('data-edge', edge);
      h.style.cssText = `position:absolute;width:12px;height:12px;margin:-6px 0 0 -6px;border-radius:3px;background:var(--k0, #FFFFFF);border:1.5px solid var(--k59, #3F7D5C);box-sizing:border-box;cursor:${cur}-resize`;
      stage.appendChild(h);
      return [h, hx, hy];
    });
    const layout = () => {
      const x = l * FW, y = t * FH, w = (1 - l - r) * FW, h = (1 - t - b) * FH;
      Object.assign(win.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
      Object.assign(inner.style, { left: -x + 'px', top: -y + 'px' });
      handles.forEach(([el, hx, hy]) => Object.assign(el.style, { left: x + hx * w + 'px', top: y + hy * h + 'px' }));
    };
    const bar = document.createElement('div');
    bar.style.cssText = 'position:fixed;left:50%;bottom:64px;transform:translateX(-50%);display:flex;gap:6px;padding:5px;border-radius:999px;background:linear-gradient(180deg,var(--k15, rgba(255,255,255,0.62)),var(--k16, rgba(255,255,255,0.34)));backdrop-filter:blur(10px) saturate(190%);-webkit-backdrop-filter:blur(10px) saturate(190%);border:1px solid var(--k17, rgba(0,0,0,0.07));box-shadow:inset 0 1px 0 var(--k18, rgba(255,255,255,0.95)),0 12px 32px var(--k22, rgba(0,0,0,0.12));font-size:13px';
    const btn = (label, strong) => { const x = document.createElement('button'); x.textContent = label; x.style.cssText = `height:30px;padding:0 16px;border:none;border-radius:999px;cursor:pointer;font:inherit;${strong ? 'background:var(--k7, #1D1D1F);color:var(--kinv, #fff)' : 'background:transparent;color:var(--k1, #1D1D1F)'}`; bar.appendChild(x); return x; };
    const cancel = btn(T('取消')), ok = btn(T('完成'), true);
    layer.append(stage, bar);
    document.body.appendChild(layer);
    layout();
    // a pointer move on screen, as a move on the unturned, unmirrored picture
    const a = -(rotation || 0) * Math.PI / 180;
    const local = (dx, dy) => { const x = dx * Math.cos(a) - dy * Math.sin(a), y = dx * Math.sin(a) + dy * Math.cos(a); return [x * (flipH ? -1 : 1), y * (flipV ? -1 : 1)]; };
    let drag = null;
    const MIN = 0.02;
    const down = e => {
      const edge = e.target.getAttribute && (e.target.getAttribute('data-edge') || (e.target.closest && e.target.closest('[data-win]') && 'move'));
      if (!edge) return;
      e.preventDefault(); e.stopPropagation();
      drag = { edge, x: e.clientX, y: e.clientY, c: [l, t, r, b] };
    };
    const move = e => {
      if (!drag) return;
      const [dx, dy] = local(e.clientX - drag.x, e.clientY - drag.y), fx = dx / FW, fy = dy / FH, [l0, t0, r0, b0] = drag.c;
      if (drag.edge === 'move') {
        const mx = Math.max(-l0, Math.min(r0, fx)), my = Math.max(-t0, Math.min(b0, fy));
        [l, t, r, b] = [l0 + mx, t0 + my, r0 - mx, b0 - my];
      } else {
        if (drag.edge.includes('l')) l = Math.max(0, Math.min(1 - r0 - MIN, l0 + fx));
        if (drag.edge.includes('r')) r = Math.max(0, Math.min(1 - l0 - MIN, r0 - fx));
        if (drag.edge.includes('t')) t = Math.max(0, Math.min(1 - b0 - MIN, t0 + fy));
        if (drag.edge.includes('b')) b = Math.max(0, Math.min(1 - t0 - MIN, b0 - fy));
      }
      layout();
    };
    const done = value => { layer.remove(); document.removeEventListener('keydown', key, true); window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); resolve(value); };
    let beside = false;
    const up = () => { drag = null; if (beside) done([l, t, r, b]); };
    const key = e => { if (e.key === 'Escape') { e.preventDefault(); done(null); } else if (e.key === 'Enter' && !(e.isComposing || e.keyCode === 229)) { e.preventDefault(); done([l, t, r, b]); } };
    stage.addEventListener('pointerdown', down);
    // a click beside the picture finishes, as in Office — on release, so the press cannot start dragging the picture underneath
    layer.addEventListener('pointerdown', e => { if (e.target === layer) { e.preventDefault(); beside = true; } });
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    document.addEventListener('keydown', key, true);
    cancel.onclick = () => done(null); ok.onclick = () => done([l, t, r, b]);
  });
}
