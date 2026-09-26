// node --test ui/tests/ — the whole app embedded (mode app, ui/embed/host.js): every document the site gives becomes a tab, in
// the site's order, the one to open comes to the front, and the page shows 正在打开… until then.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { page } from './embed-page.mjs';

const { E, say, until, find, embed, api, shown } = await page({ id: 'app1', mode: 'app' });

test('the whole app: the site\'s documents open as tabs in its order, then the one to show; hidden until then', async () => {
  assert.equal(embed.starting, true, 'the page waits for the site\'s documents');
  embed.shell({ path: '', isDoc: false }, api); // the 新建 page is up before the documents are
  const bytes = () => new Uint8Array([1]).buffer;
  say({ type: 'init', files: [{ name: 'a.docx', data: bytes() }, { name: 'b.xlsx', data: bytes() }, { name: 'c.pptx', data: bytes() }], open: 'b.xlsx' });
  const ready = await until(() => find('ready'));
  assert.deepEqual([...E.files.keys()], ['a.docx', 'b.xlsx', 'c.pptx']);
  assert.deepEqual(shown, ['a.docx', 'b.xlsx', 'c.pptx', 'b.xlsx'], 'each opened in turn (a tab each), then b.xlsx in front');
  assert.equal(ready.msg.file, 'b.xlsx');
  await until(() => embed.starting === false);
  assert.equal(embed.empty, undefined, 'the whole app is never "nothing to show"');
});
