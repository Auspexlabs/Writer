// node --test ui/tests/ — one editor embedded with no document (ui/embed/host.js): a blank one of its kind is made and shown,
// named as the site asked; the command that makes it does not wait for the documents it is part of.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { page } from './embed-page.mjs';

const { E, say, until, find, embed, api, shown } = await page({ id: 'x1', mode: 'xlsx', blankName: '预算' });

test('one editor given nothing: a blank document of its kind, named blankName, is created and shown', async () => {
  embed.shell({ path: '', isDoc: false }, api);
  say({ type: 'init', files: [] });
  const ready = await until(() => find('ready'));
  assert.deepEqual(E.calls[0], ['create', '预算.xlsx']);
  assert.deepEqual([ready.msg.file, shown], ['预算.xlsx', ['预算.xlsx']]);
  assert.equal(find('event', m => m.name === 'change'), undefined, 'making it is no change to tell the site about');
  await until(() => embed.starting === false);
  assert.equal(embed.empty, undefined);
});
