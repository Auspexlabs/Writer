// panel.js — the building blocks of the right-hand format panel (FormatPanel.dc.html draws it). An editor describes each of its
// tabs as groups of rows of plain items; kit(editor) makes those items, and its popup buttons open the editor's own menus
// (editor.menus[id], editor.openPop(id, event)), the same menus the pure-mode toolbar opens.

// ui/i18n.js's $t in the page; under node the Chinese itself, without its @@ context
const T = (s, v) => { if (globalThis.$t) return globalThis.$t(s, v); s = String(s).replace(/@@.*$/, ''); return v ? s.replace(/\{(\w+)\}/g, (m, k) => (k in v ? v[k] : m)) : s; };

/** The item makers for one editor. Static Chinese labels go through $t here; a label that is a value (a font name, 1.5 倍) comes
 *  translated already. */
export function kit(ed) {
  const openId = () => ed.state.pop && ed.state.pop.id;
  const menu = (id, items) => { if (items) ed.menus[id] = items; return e => ed.openPop(id, e); };
  const k = {
    /** A group: its title, then its rows (falsy rows are left out). */
    G: (title, ...rows) => ({ title: title ? T(title) : '', rows: rows.flat().filter(Boolean) }),
    /** A row of items side by side; grid(n, …) shares n equal columns, wrap(…) lets them flow onto more lines. */
    R: (...items) => ({ items: items.flat().filter(Boolean) }),
    grid: (n, ...items) => ({ grid: n, items: items.flat().filter(Boolean) }),
    wrap: (...items) => ({ wrap: true, items: items.flat().filter(Boolean) }),
    /** A popup button opening the editor's menu `id` (items registered here), its label the current value. */
    sel: (id, label, items, o = {}) => ({ t: 'sel', label, title: o.title ? T(o.title) : label, w: o.w, lcss: o.lcss, dis: o.dis, open: openId() === id, onClick: menu(id, items) }),
    /** A popup button over a native menu: options [[value, label]], onChange(value). */
    pick: (value, options, onChange, o = {}) => ({ t: 'sel', options: options.map(([v, l, dis]) => [v, l, dis]), value, onChange, w: o.w, title: o.title ? T(o.title) : '',
      label: o.label || ((options.find(x => String(x[0]) === String(value)) || [])[1] || '') }),
    /** A colour square, a label and the chevron: onPick(colour) through the system colour picker, or a menu. */
    swatch: (label, color, onPick, o = {}) => o.id ? { t: 'swatch', label: T(label), color, w: o.w, title: T(o.title || label), open: openId() === o.id, onClick: menu(o.id, o.items) }
      : { t: 'swatch', label: T(label), color, value: o.value || color, w: o.w, title: T(o.title || label), onChange: onPick },
    btn: (label, icon, onClick, o = {}) => ({ t: 'btn', label: T(label), icon, onClick, on: o.on, dis: o.dis, w: o.w, title: o.title ? T(o.title) : T(label) }),
    /** A button that opens a menu. */
    menuBtn: (id, label, icon, items, o = {}) => ({ t: 'btn', label: T(label), icon, on: o.on, dis: o.dis, w: o.w, title: o.title ? T(o.title) : T(label), open: openId() === id, onClick: menu(id, items) }),
    card: (label, icon, onClick, o = {}) => ({ t: 'card', label: T(label), icon, onClick, dis: o.dis, title: o.title ? T(o.title) : T(label) }),
    menuCard: (id, label, icon, items, o = {}) => ({ t: 'card', label: T(label), icon, title: o.title ? T(o.title) : T(label), open: openId() === id, onClick: menu(id, items) }),
    /** A segmented control: opts [{ label | icon | rich, title, on, onClick, css }]; icon-only segments keep their title for the tooltip. */
    seg: (opts, o = {}) => ({ t: 'seg', w: o.w, title: o.title ? T(o.title) : '', opts: opts.filter(Boolean).map(x => ({ ...x, title: x.title ? T(x.title) : x.label ? T(x.label) : '', label: x.label ? T(x.label) : x.label })) }),
    /** A number field and its stepper: value as shown (with its unit), onStep(±1), onSet(text typed); lab puts a label before it. */
    num: (value, onStep, onSet, o = {}) => ({ t: 'num', value: String(value), onStep, onSet, w: o.w, lab: o.lab ? T(o.lab) : '', labW: o.labW ? o.labW + 'px' : 'auto', title: T(o.title || o.lab || ''),
      upTip: T(o.upTip || '增加'), downTip: T(o.downTip || '减少') }),
    chk: (label, on, onChange, o = {}) => ({ t: 'chk', label: T(label), on: !!on, onChange, dis: o.dis, title: o.title ? T(o.title) : T(label) }),
    sw: (on, onChange, title) => ({ t: 'sw', on: !!on, onChange, title: T(title || '') }),
    lab: (label, w) => ({ t: 'lab', label: T(label), w }),
    sp: () => ({ t: 'sp' }),
    /** A− value ⌄ A+: onDec, onInc, and the value's menu. */
    size: (value, onDec, onInc, id, items, o = {}) => ({ t: 'size', value: String(value), onDec, onInc, w: o.w, h: o.h, minW: o.minW, title: T(o.title || '字号'),
      decTitle: T(o.decTitle || '减小字号'), incTitle: T(o.incTitle || '增大字号'), open: openId() === id, onMenu: menu(id, items) }),
    /** A slider with its label and value. */
    range: (label, value, min, max, onChange, o = {}) => ({ t: 'range', label: T(label), value, min, max, onChange, text: o.text != null ? String(o.text) : String(value), title: T(o.title || label) }),
    link: (label, onClick, o = {}) => ({ t: 'link', label: T(label), onClick, dis: o.dis }),
    stat: (n, label) => ({ t: 'stat', n: String(n), label: T(label) }),
    input: (label, value, onChange, onKey, ref, title) => ({ t: 'input', label: T(label), value, onChange, onKey, ref, title: T(title || label) }),
    empty: label => ({ t: 'empty', label }),
    tsty: (look, on, onClick, title) => ({ t: 'tsty', ...look, on, onClick, title: T(title) }),
    /** A row of colour dots, a palette: [{ color, on, title, onClick }]; no color draws "none". */
    dots: list => ({ t: 'dots', dots: list.filter(Boolean).map(x => ({ ...x, title: x.title ? T(x.title) : x.color || T('无') })) }),
    /** A row of small marker buttons, each its glyph in its own colour: [{ label, color, fs, on, dis, title, onClick }]. */
    marks: list => ({ t: 'marks', marks: list.filter(Boolean).map(x => ({ ...x, title: T(x.title || '') })) }),
    /** A slide theme as a tile: its background with its text and accent colours, the name under it. */
    theme: (label, look, on, onClick) => ({ t: 'theme', label: T(label), title: T(label), bg: look.bg, fg: look.fg, acc: look.acc, on: !!on, onClick })
  };
  return k;
}

/** The tabs row: [[key, label, context]] with the current key; a context tab (表格, 图片, 形状) is drawn in blue. */
export function tabs(defs, cur, onPick) {
  return defs.map(([key, label, ctx]) => ({ key, label: T(label), on: key === cur, ctx: !!ctx, onClick: () => onPick(key) }));
}

// ----- numbers with their units, as the fields show and take them -----
const CM = { cm: 1, 厘米: 1, 公分: 1, mm: 0.1, 毫米: 0.1, in: 2.54, 英寸: 2.54, '"': 2.54, pt: 2.54 / 72, 磅: 2.54 / 72 };
/** A number from what was typed ("2.5", "2.5 厘米", "25mm", "1in"), in the unit the field shows; NaN when there is none. */
export function parseCm(text) {
  const m = /^\s*(-?\d+(?:\.\d+)?|-?\.\d+)\s*([a-z"]+|厘米|公分|毫米|英寸|磅)?\s*$/i.exec(String(text || ''));
  return m ? +m[1] * (CM[(m[2] || 'cm').toLowerCase()] || 1) : NaN;
}
/** A length in cm as the fields show it, to the hundredth and halves up as Word shows them (1800 twips, 3.175 cm, is 3.18 厘米). */
export const cmLabel = cm => T('{n} 厘米', { n: Math.round(+cm * 100 + 1e-9) / 100 });
/** Rounds to the step, so a stepper lands on round values. */
export const snap = (v, step) => Math.round(Math.round(v / step) * step * 1000) / 1000;
