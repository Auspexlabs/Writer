// A small DOM for node tests: enough of Element for engine.js's html → blocks walk (blocksFromHtml, inlineHtml, runsOf) on
// well-formed markup. Selectors: tag names, [attr], [attr="v"], .class, > and descendant combinators, comma lists.
const VOID = /^(br|hr|img|input|col|wbr)$/i;
const camel = k => k.replace(/-([a-z])/g, (m, c) => c.toUpperCase());
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export class Text { constructor(t) { this.nodeType = 3; this.nodeValue = t; this.parentNode = null; } get textContent() { return this.nodeValue; } cloneNode() { return new Text(this.nodeValue); } remove() { const p = this.parentNode; if (p) p.childNodes.splice(p.childNodes.indexOf(this), 1); this.parentNode = null; } }

export class Element {
  constructor(tag, attrs = {}) { this.nodeType = 1; this.tagName = tag.toUpperCase(); this.attrs = {}; this.childNodes = []; this.parentNode = null; this.style = styleProxy(this); for (const [k, v] of Object.entries(attrs)) this.attrs[k] = v; }
  get parentElement() { return this.parentNode && this.parentNode.nodeType === 1 ? this.parentNode : null; }
  get children() { return this.childNodes.filter(n => n.nodeType === 1); }
  get firstChild() { return this.childNodes[0] || null; }
  get lastChild() { return this.childNodes[this.childNodes.length - 1] || null; }
  get firstElementChild() { return this.children[0] || null; }
  get previousElementSibling() { const s = this.parentNode ? this.parentNode.children : []; return s[s.indexOf(this) - 1] || null; }
  get nextElementSibling() { const s = this.parentNode ? this.parentNode.children : []; return s[s.indexOf(this) + 1] || null; }
  get nextSibling() { const s = this.parentNode ? this.parentNode.childNodes : []; return s[s.indexOf(this) + 1] || null; }
  get colSpan() { return +this.attrs.colspan || 1; } set colSpan(v) { this.attrs.colspan = String(v); }
  get rowSpan() { return +this.attrs.rowspan || 1; } set rowSpan(v) { this.attrs.rowspan = String(v); }
  get id() { return this.attrs.id || ''; } set id(v) { this.attrs.id = v; }
  get className() { return this.attrs.class || ''; }
  get classList() { const el = this; return { contains: c => el.className.split(/\s+/).includes(c), add: c => { if (!this.contains(c)) el.attrs.class = (el.className + ' ' + c).trim(); }, remove: c => { el.attrs.class = el.className.split(/\s+/).filter(x => x !== c).join(' '); }, toggle: (c, on) => (on ?? !this.contains(c)) ? this.add(c) : this.remove(c) }; }
  getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; }
  hasAttribute(k) { return k in this.attrs; }
  setAttribute(k, v) { this.attrs[k] = String(v); }
  removeAttribute(k) { delete this.attrs[k]; }
  appendChild(n) { n.remove && n.remove(); n.parentNode = this; this.childNodes.push(n); return n; }
  append(...ns) { ns.forEach(n => this.appendChild(typeof n === 'string' ? new Text(n) : n)); }
  prepend(...ns) { ns.forEach(n => this.insertBefore(n, this.firstChild)); }
  insertBefore(n, ref) { n.remove && n.remove(); n.parentNode = this; const i = ref ? this.childNodes.indexOf(ref) : -1; if (i < 0) this.childNodes.push(n); else this.childNodes.splice(i, 0, n); return n; }
  after(...ns) { const p = this.parentNode; let ref = this.nextSibling; ns.forEach(n => p.insertBefore(n, ref)); }
  before(...ns) { const p = this.parentNode; ns.forEach(n => p.insertBefore(n, this)); }
  replaceWith(n) { const p = this.parentNode; p.insertBefore(n, this); this.remove(); }
  replaceChildren(...ns) { this.childNodes.slice().forEach(n => n.remove()); this.append(...ns); }
  remove() { const p = this.parentNode; if (p) p.childNodes.splice(p.childNodes.indexOf(this), 1); this.parentNode = null; }
  cloneNode(deep) { const c = new Element(this.tagName, this.attrs); if (deep) this.childNodes.forEach(n => c.appendChild(n.cloneNode(true))); return c; }
  get textContent() { return this.childNodes.map(n => n.textContent).join(''); }
  set textContent(t) { this.replaceChildren(new Text(String(t))); }
  get innerText() { return this.textContent; }
  get innerHTML() { return this.childNodes.map(n => n.nodeType === 3 ? esc(n.nodeValue) : n.outerHTML).join(''); }
  set innerHTML(h) { this.replaceChildren(...parseNodes(h)); }
  get outerHTML() { const a = Object.entries(this.attrs).map(([k, v]) => ` ${k}="${String(v).replace(/"/g, '&quot;')}"`).join(''), t = this.tagName.toLowerCase(); return VOID.test(t) ? `<${t}${a}>` : `<${t}${a}>${this.innerHTML}</${t}>`; }
  matches(sel) { return sel.split(',').some(s => matchesOne(this, s.trim())); }
  closest(sel) { for (let e = this; e; e = e.parentElement) if (e.matches(sel)) return e; return null; }
  contains(n) { for (let e = n; e; e = e.parentNode) if (e === this) return true; return false; }
  querySelectorAll(sel) { const out = []; const walk = e => { for (const c of e.children) { if (c.matches(sel)) out.push(c); walk(c); } }; walk(this); return out; }
  querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
  getBoundingClientRect() { return { left: 0, top: 0, width: 0, height: 0, right: 0, bottom: 0 }; }
}
function styleProxy(el) {
  const read = () => Object.fromEntries((el.attrs.style || '').split(';').map(d => d.split(':')).filter(x => x.length > 1).map(([k, ...v]) => [camel(k.trim()), v.join(':').trim()]));
  const write = o => { const s = Object.entries(o).filter(([, v]) => v !== '' && v != null).map(([k, v]) => `${k.replace(/[A-Z]/g, c => '-' + c.toLowerCase())}: ${v}`).join('; '); if (s) el.attrs.style = s + ';'; else delete el.attrs.style; };
  const setProp = (k, v) => { const o = read(); o[camel(k)] = v; write(o); };
  return new Proxy({}, { get: (_, k) => k === 'cssText' ? el.attrs.style || '' : k === 'setProperty' ? setProp : k === 'removeProperty' ? (n => setProp(n, '')) : k === 'getPropertyValue' ? (n => read()[camel(n)] || '') : read()[k] || '', set: (_, k, v) => { if (k === 'cssText') { if (v) el.attrs.style = v; else delete el.attrs.style; } else setProp(k, v); return true; } });
}
function matchesOne(el, sel) {
  const parts = sel.split(/\s*(>)\s*|\s+/).filter(Boolean); // ['a', '>', 'b'] or ['a', 'b']
  let e = el, i = parts.length - 1;
  if (!simple(e, parts[i--])) return false;
  while (i >= 0) {
    if (parts[i] === '>') { e = e.parentElement; if (!e || !simple(e, parts[i - 1])) return false; i -= 2; continue; }
    e = e.parentElement; while (e && !simple(e, parts[i])) e = e.parentElement; if (!e) return false; i--;
  }
  return true;
}
function simple(el, s) {
  if (s.startsWith(':scope')) return true;
  const m = /^([a-z0-9]*)((?:[.#\[][^\[\]]*\]?)*)$/i.exec(s); if (!m) return false;
  if (m[1] && m[1] !== '*' && el.tagName !== m[1].toUpperCase()) return false;
  for (const c of m[2].match(/[.#][^.#\[]+|\[[^\]]+\]/g) || []) {
    if (c[0] === '.') { if (!el.classList.contains(c.slice(1))) return false; }
    else if (c[0] === '#') { if (el.id !== c.slice(1)) return false; }
    else { const a = /^\[([^=~^$*|\]]+)(?:([~^$*|]?)=["']?([^"'\]]*)["']?)?\]$/.exec(c); if (!a) return false; const v = el.getAttribute(a[1]); if (v === null) return false; if (a[3] !== undefined) { if (a[2] === '^' ? !v.startsWith(a[3]) : a[2] === '*' ? !v.includes(a[3]) : v !== a[3]) return false; } }
  }
  return true;
}
const decode = s => s.replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
/** Nodes of a well-formed html fragment. */
export function parseNodes(html) {
  const root = new Element('body'); let cur = root, i = 0; html = String(html || '');
  while (i < html.length) {
    const lt = html.indexOf('<', i);
    if (lt < 0) { cur.appendChild(new Text(decode(html.slice(i)))); break; }
    if (lt > i) cur.appendChild(new Text(decode(html.slice(i, lt))));
    if (html.startsWith('<!--', lt)) { i = html.indexOf('-->', lt); i = i < 0 ? html.length : i + 3; continue; }
    let gt = lt + 1, q = null; for (; gt < html.length; gt++) { const ch = html[gt]; if (q) { if (ch === q) q = null; } else if (ch === '"' || ch === "'") q = ch; else if (ch === '>') break; }
    const raw = html.slice(lt + 1, gt); i = gt + 1;
    if (raw[0] === '/') { const name = raw.slice(1).trim().toUpperCase(); for (let e = cur; e && e !== root; e = e.parentElement) if (e.tagName === name) { cur = e.parentElement; break; } continue; }
    const name = /^[a-zA-Z0-9-]+/.exec(raw)?.[0]; if (!name) continue;
    const el = new Element(name); const re = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g; let m; const rest = raw.slice(name.length);
    while ((m = re.exec(rest))) el.attrs[m[1].toLowerCase()] = decode(m[2] ?? m[3] ?? m[4] ?? '');
    cur.appendChild(el);
    if (!VOID.test(name) && !raw.endsWith('/')) cur = el;
  }
  return root.childNodes.slice();
}
export class DOMParser { parseFromString(s) { const body = new Element('body'); parseNodes(s).forEach(n => body.appendChild(n)); return { body, createElement: t => new Element(t), createTextNode: t => new Text(t) }; } }
export const document = { createElement: t => new Element(t), createTextNode: t => new Text(t), createTreeWalker() { throw new Error('no tree walker'); } };
/** Installs the stub globally (DOMParser, document, Node); returns the root of html parsed as a body. */
export function install() { globalThis.DOMParser = DOMParser; globalThis.document ??= document; globalThis.Node ??= { ELEMENT_NODE: 1, TEXT_NODE: 3 }; return html => new DOMParser().parseFromString(html).body; }
