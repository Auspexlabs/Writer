// node --test ui/tests/   — drafts and user-chosen locations in the engine bridge (desktop apps)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const EN = await import('../engine.js');

/** fetch stub: answers /files and /files?drafts=1, records every request */
function stub({ workspace, drafts, files = [], draftFiles = [] }) {
  const calls = [];
  globalThis.fetch = async (url, opts = {}) => {
    calls.push({ url, method: opts.method || 'GET', body: opts.body });
    const json = x => new Response(JSON.stringify(x), { status: 200, headers: { 'Content-Type': 'application/json' } });
    if (url === '/files') return json({ workspace, drafts, files: files.map(path => ({ path, format: path.split('.').pop() })) });
    if (url === '/files?drafts=1') return json({ workspace, drafts, files: draftFiles.map(path => ({ path, format: path.split('.').pop() })) });
    if (url === '/run') return json({ code: 0, output: '' });
    return json({ path: 'x', size: 0 });
  };
  return calls;
}

test('isDraft: paths inside the drafts folder, relative ones in the drafts window, any separators and case', async () => {
  stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  assert.equal(EN.isDraft('/U/app/Drafts/未命名.docx'), true);
  assert.equal(EN.isDraft('未命名.docx'), false);
  assert.equal(EN.isDraft('/U/app/Drafts2/a.docx'), false);
  stub({ workspace: '/U/app/Drafts', drafts: '/U/app/Drafts' }); await EN.files();
  assert.equal(EN.isDraft('未命名.docx'), true);
  stub({ workspace: 'C:/Users/me/Documents', drafts: 'C:/Users/me/AppData/Roaming/cn.thewriter.app/Drafts' }); await EN.files();
  assert.equal(EN.isDraft('C:\\Users\\me\\AppData\\Roaming\\cn.thewriter.app\\Drafts\\a.docx'), true);
  assert.equal(EN.isDraft('c:/users/me/appdata/roaming/cn.thewriter.app/drafts/a.docx'), true);
  stub({ workspace: '/U/Docs' }); await EN.files();
  assert.equal(EN.isDraft('/U/app/Drafts/a.docx'), false);   // browser version: no drafts
});

test('new documents go to the drafts folder with a free name; in the drafts window as relative paths', async () => {
  let calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts', draftFiles: ['/U/app/Drafts/未命名.docx'] }); await EN.files();
  const d = await EN.create('docx');
  assert.equal(d.path, '/U/app/Drafts/未命名 2.docx');
  assert.deepEqual(JSON.parse(calls.find(c => c.url === '/run').body).argv, ['create', '/U/app/Drafts/未命名 2.docx']);
  calls = stub({ workspace: '/U/app/Drafts', drafts: '/U/app/Drafts', files: ['未命名.docx'] }); await EN.files();
  assert.equal((await EN.create('docx')).path, '未命名 2.docx');
  calls = stub({ workspace: '/U/Docs', files: ['未命名.docx'] }); await EN.files();
  assert.equal((await EN.create('docx')).path, '未命名 2.docx');   // browser version: the workspace, as before
});

test('saveAs moves or copies through PUT /file; discard deletes; relOf keeps workspace paths relative', async () => {
  const calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  const doc = { path: '/U/app/Drafts/未命名.docx' };
  assert.equal(await EN.saveAs(doc, 'C:\\Users\\me\\报告.docx'), 'C:/Users/me/报告.docx');
  assert.equal(calls.at(-1).url, '/file?file=' + encodeURIComponent('C:/Users/me/报告.docx') + '&from=' + encodeURIComponent('/U/app/Drafts/未命名.docx'));
  assert.equal(calls.at(-1).method, 'PUT');
  await EN.saveAs({ path: '/U/Docs/a.docx' }, '/U/Other/b.docx', true);
  assert.match(calls.at(-1).url, /&keep=1$/);
  await EN.discard(doc);
  assert.deepEqual([calls.at(-1).method, calls.at(-1).url], ['DELETE', '/file?file=' + encodeURIComponent('/U/app/Drafts/未命名.docx')]);
  assert.equal(EN.relOf('/U/Docs/sub/x.docx'), 'sub/x.docx');
  assert.equal(EN.relOf('/U/Other/x.docx'), '/U/Other/x.docx');
  assert.equal(EN.lazy('C:\\x\\y\\报告.docx').title, '报告');
});

// ---- the shell (index.dc.html): 存储 / 另存为 through the host's Save dialog, the close prompt for drafts ----
/** The shell's logic in a vm over the real bridge; only saving the model and watching the file are faked. */
function shell(props, docs, cur) {
  const code = readFileSync(new URL('../index.dc.html', import.meta.url), 'utf8').match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { location: { search: '' }, URLSearchParams, structuredClone, setTimeout: () => 0, clearTimeout: () => { }, document: { querySelector: () => null },
    React: { createRef: () => ({ current: null }) }, DCLogic: class { setState(u) { Object.assign(this.state, typeof u === 'function' ? u(this.state) : u); } } };
  vm.runInNewContext(code + '\nglobalThis.Shell = Component;', ctx);
  const c = new ctx.Shell();
  Object.assign(c, { props, EN: { ...EN, save: async () => 0, watch: () => () => { } } });
  Object.assign(c.state, { docs, cur });
  return c;
}
const draft = n => ({ id: 'd' + n, title: '未命名 ' + n, type: 'md', path: `/U/app/Drafts/未命名 ${n}.md`, loaded: true, text: '' });

test('the shell: 存储 moves a draft to where the Save dialog says, 另存为 copies a folder document; undo keeps the new place', async () => {
  const calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  const asked = [];
  const c = shell({ onSaveDialog: async (name, ext) => { asked.push(name + '.' + ext); return '/U/Desktop/报告.md'; } }, [draft(1), { id: 'a', title: 'a', type: 'md', path: 'a.md', loaded: true, text: '' }], 'd1');
  assert.equal(await c.save(), true);
  assert.deepEqual(asked, ['未命名 1.md']);
  assert.deepEqual([calls.at(-1).method, calls.at(-1).url], ['PUT', '/file?file=' + encodeURIComponent('/U/Desktop/报告.md') + '&from=' + encodeURIComponent('/U/app/Drafts/未命名 1.md')]);
  assert.deepEqual([c.state.docs[0].id, c.state.docs[0].path, c.state.docs[0].title], ['d1', '/U/Desktop/报告.md', '报告']);
  c.hist.d1 = { past: [draft(1)], future: [] }; c.undo();
  assert.equal(c.state.docs[0].path, '/U/Desktop/报告.md', 'undo brings back content, not the draft it was');
  assert.equal(await c.save(), true); assert.equal(asked.length, 1, 'a saved document is autosaved: 存储 does not ask again');
  c.state.cur = 'a'; assert.equal(await c.saveAs(), true);
  assert.match(calls.at(-1).url, /&keep=1$/); assert.equal(c.state.docs[1].path, '/U/Desktop/报告.md');
});

test('the shell: closing a draft asks 存储 / 不存储 / 取消, the window close goes through every draft; without the host nothing asks', async () => {
  const calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  const c = shell({ onSaveDialog: async () => '/U/Desktop/报告.md' }, [draft(1), draft(2), { id: 'a', title: 'a', type: 'md', path: 'a.md', loaded: true, text: '' }], 'd1');
  assert.equal(await c.closeDoc('a'), true); assert.equal(c.state.ask, null, 'a document of the folder closes at once');
  let closing = c.closeDoc('d1'); assert.equal(c.state.ask.title, '未命名 1');
  c.state.ask.res('cancel'); assert.equal(await closing, false); assert.ok(c.state.docs.some(x => x.id === 'd1'));
  closing = c.closeDoc('d1'); c.state.ask.res('discard'); assert.equal(await closing, true);
  assert.deepEqual([calls.at(-1).method, calls.at(-1).url], ['DELETE', '/file?file=' + encodeURIComponent('/U/app/Drafts/未命名 1.md')]);
  closing = c.confirmClose(); c.state.ask.res('save'); assert.equal(await closing, true);
  assert.equal(calls.at(-1).method, 'PUT'); assert.deepEqual(c.state.docs, []);
  const b = shell({}, [draft(3)], 'd3');
  assert.equal(await b.closeDoc('d3'), true); assert.equal(b.state.ask, null);
});

test('ids stay unique when a new draft takes the name of a document that was saved elsewhere', () => {
  assert.notEqual(EN.lazy('未命名.docx').id, EN.lazy('未命名.docx').id);
});

test('saveAs and discard wait their turn behind a write already under way on the same file', async () => {
  const order = [];
  let release; const gate = new Promise(r => { release = r; });
  globalThis.fetch = async (url, opts = {}) => {
    const m = opts.method || 'GET'; order.push('start ' + m);
    if (m === 'PUT') await gate;
    order.push('end ' + m);
    return new Response('{}', { status: 200, headers: { 'Content-Type': 'application/json' } });
  };
  const doc = { path: '/U/app/Drafts/lane.docx' };
  const put = EN.saveAs(doc, '/U/x/lane.docx', true);
  const del = EN.discard(doc);
  await new Promise(r => setTimeout(r, 20));
  assert.deepEqual(order, ['start PUT']);
  release(); await put; await del;
  assert.deepEqual(order, ['start PUT', 'end PUT', 'start DELETE', 'end DELETE']);
});
