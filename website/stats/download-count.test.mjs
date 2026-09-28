import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const source = readFileSync(new URL('../dist/assets/download-count.js', import.meta.url), 'utf8');
function browser(beacon = true, failFetch = false) {
  const handlers = {}, sent = [], fetched = [];
  vm.runInNewContext(source, {
    URL,
    location: { href: 'https://thewriter.cn/', origin: 'https://thewriter.cn' },
    navigator: { sendBeacon(url, body) { if (beacon === 'throws') throw Error('blocked'); sent.push({ url, body }); return beacon; } },
    fetch(url, options) { fetched.push({ url, ...options }); if (failFetch) throw Error('offline'); return Promise.resolve(); },
    document: { addEventListener(type, handler) { handlers[type] = handler; } },
  });
  return { sent, fetched, click(href, type = 'click', button = 0, defaultPrevented = false) {
    const link = { href };
    handlers[type]({ type, button, defaultPrevented, target: { closest: () => link }, preventDefault() { throw Error('download blocked'); } });
  } };
}

test('each Mac/Windows click counts exactly once; repeats, keyboard and middle click count', () => {
  const b = browser();
  b.click('/download/mac');
  b.click('/download/mac');
  b.click('/download/windows');
  b.click('/download/windows', 'auxclick', 1);
  assert.deepEqual(b.sent.map(s => JSON.parse(s.body).platform), ['mac', 'mac', 'windows', 'windows']);
  assert.equal(b.fetched.length, 0);
});
test('page navigation, installer files, external links and right-click do not count', () => {
  const b = browser();
  b.click('/#download');
  b.click('/download/Writer-0.1.8-mac.dmg');
  b.click('https://elsewhere.example/download/mac');
  b.click('/download/mac', 'auxclick', 2);
  b.click('/download/mac', 'click', 0, true);
  assert.equal(b.sent.length, 0);
});
test('beacon rejection uses one keepalive request without interfering with download', () => {
  for (const beacon of [false, 'throws']) {
    const b = browser(beacon);
    b.click('/download/mac');
    assert.equal(b.fetched.length, 1);
    assert.equal(b.fetched[0].keepalive, true);
    assert.equal(b.fetched[0].body, '{"platform":"mac"}');
  }
  assert.doesNotThrow(() => browser('throws', true).click('/download/mac'));
});
