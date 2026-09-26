// node --test ui/tests/ — the title bar: ⌘N opens the 新建 page (it does not make a blank file of some type), the 新建 page
// shows none of the controls that act on a document, and the ⋯ menu is gone (沉浸 has its own button in the pill).
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const read = f => readFileSync(new URL('../' + f, import.meta.url), 'utf8');
const script = f => read(f).match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];

function shell() {
  const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, getComputedStyle: () => ({}), window: { innerWidth: 1200 },
    document: { querySelector: () => null, querySelectorAll: () => [], documentElement: { dataset: {} }, activeElement: null }, localStorage: { getItem: () => null, setItem() { } },
    $t: s => String(s).replace(/@@.*$/, ''), React: { createRef: () => ({ current: null }) },
    DCLogic: class { setState(u, cb) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(script('index.dc.html') + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell(); c.props = { mac: true }; c.prefs = {};
  Object.assign(c.state, { ready: true, view: 'doc', cur: 'a', docs: [{ id: 'a', type: 'docx', title: '报告', path: '报告.docx', loaded: true, html: '' }] });
  return c;
}
const key = (c, k, o = {}) => { const e = Object.assign({ key: k, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, preventDefault() { this.prevented = true; } }, o); c.onKey(e); return !!e.prevented; };

test('⌘N opens the 新建 page instead of making a blank document, and leaves 沉浸书写 on the way', () => {
  const c = shell(); const made = []; c.createBlank = t => made.push(t);
  assert.equal(key(c, 'n', { metaKey: true }), true);
  assert.deepEqual([c.state.view, made.length], ['create', 0], 'the page with the templates and blanks, no new file');
  c.setState({ view: 'doc' }); c.setPure(true);
  key(c, 'n', { metaKey: true });
  assert.deepEqual([c.state.view, c.state.pure], ['create', false]);
  c.setState({ view: 'doc' }); key(c, 't', { metaKey: true }); assert.equal(c.state.view, 'create', '⌘T goes to the same page');
});

for (const page of ['mac.dc.html', 'win.dc.html']) {
  test(`${page}: no ⋯ menu, and the 新建 page shows no sidebar, undo / redo or Aa / 沉浸 / AI`, () => {
    const src = read(page);
    assert.doesNotMatch(src, /tb-more|moreOpen|moreToggle/, 'the ⋯ button and its menu are gone');
    const side = src.indexOf('class="tb-btn tb-side"'), undo = src.indexOf('onClick="{{ goBack }}"'), pill = src.indexOf('data-on="{{ barOn }}"');
    for (const [name, at] of [['sidebar', side], ['undo', undo], ['Aa pill', pill]]) {
      const open = src.lastIndexOf('<sc-if value="{{ docBar }}"', at), close = src.lastIndexOf('</sc-if>', at);
      assert.ok(open > -1 && open > close, `${name} sits inside <sc-if value="{{ docBar }}">`);
    }
    assert.match(script(page), /docBar: !!\(sh && sh\.isDoc\)/);
  });
}
assert.doesNotMatch(read('theme-dark.css'), /\.tb-more/);
