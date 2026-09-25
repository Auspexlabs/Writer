// node --test ui/tests/   — compatibility formats in the engine bridge: the engine's readable-type list decides the editor,
// and a .doc/.xls/.ppt/.wps/.odt/.rtf/.csv opens as a draft of the editor's own type (Office's compatibility mode)
import { test } from 'node:test';
import assert from 'node:assert/strict';

globalThis.window = globalThis;
globalThis.DOMParser ??= class { parseFromString(s) { return { body: { innerHTML: s, childNodes: [], children: [], querySelectorAll: () => [] } }; } }; // enough for an empty-ish docx to open
const EN = await import('../engine.js');

const OPEN = { docx: 'docx', doc: 'docx', wps: 'docx', odt: 'docx', rtf: 'docx', xlsx: 'xlsx', xls: 'xlsx', et: 'xlsx', csv: 'xlsx', pptx: 'pptx', ppt: 'pptx', dps: 'pptx', md: 'md', txt: 'md', pdf: 'pdf', mm: 'mm' };

/** fetch stub: /files with the engine's open map, /run recording argv, a minimal docx tree for /json, /stat */
function stub({ workspace, drafts, files = [] }) {
  const calls = [];
  globalThis.fetch = async (url, opts = {}) => {
    calls.push({ url, method: opts.method || 'GET', body: opts.body });
    const json = x => new Response(JSON.stringify(x), { status: 200, headers: { 'Content-Type': 'application/json' } });
    if (url === '/files') return json({ workspace, drafts, open: OPEN, files: files.map(path => ({ path, format: path.split('.').pop() })) });
    if (url === '/files?drafts=1') return json({ workspace, drafts, open: OPEN, files: [] });
    if (url === '/run') return json({ code: 0, output: '{"file":"x","format":"docx","warnings":[]}' });
    if (url.startsWith('/json?')) return json({ kind: 'document', path: '/', props: { format: 'docx' }, children: [{ kind: 'body', path: '/body', props: {}, children: [{ kind: 'paragraph', path: '/body/paragraph[1]', props: { text: 'hi', html: 'hi' }, children: [] }] }] });
    if (url.startsWith('/stat?')) return json({ path: 'x', mtime: 1, size: 0 });
    return json({ path: 'x', size: 0 });
  };
  return calls;
}

test('the engine says which editor a file opens in; unknown types are left out of the listing', async () => {
  stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' });
  const files = await EN.files();
  assert.equal(EN.editorOf('报告.DOC'), 'docx');
  assert.equal(EN.editorOf('/U/Docs/表.et'), 'xlsx');
  assert.equal(EN.editorOf('a.txt'), 'md');
  assert.equal(EN.editorOf('a.zip'), null);
  assert.equal(EN.isCompat('a.doc'), true);
  assert.equal(EN.isCompat('a.docx'), false);
  assert.equal(EN.isCompat('a.txt'), false);   // text is saved in place as text
  assert.equal(EN.isCompat('a.pdf'), false);
  assert.ok(EN.accept().split(',').includes('.wps'));
  assert.equal(EN.lazy('/U/Docs/合同.wps').type, 'docx');
  assert.equal(EN.lazy('/U/Docs/新建.dps').type, 'pptx');
  assert.equal(files.length, 0);
  stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts', files: ['/U/Docs/a.doc', '/U/Docs/b.zip', '/U/Docs/c.csv'] });
  const listed = await EN.files();
  assert.deepEqual(listed.map(f => [f.path.split('/').pop(), f.type]), [['a.doc', 'docx'], ['c.csv', 'xlsx']]);
});

test('opening a .doc converts it into a draft of the editor type; the original path is remembered, never written', async () => {
  const calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  const doc = await EN.open(EN.lazy('/U/Docs/合同.doc'));
  const run = calls.find(c => c.url === '/run');
  assert.deepEqual(JSON.parse(run.body).argv, ['export', '/U/Docs/合同.doc', '--to', '/U/app/Drafts/合同.docx']);
  assert.equal(doc.path, '/U/app/Drafts/合同.docx');
  assert.equal(doc.type, 'docx');
  assert.equal(doc.from, '/U/Docs/合同.doc');
  assert.equal(EN.isDraft(doc.path), true);
  assert.ok(doc.html.includes('hi'));
  assert.ok(!calls.some(c => c.method === 'PUT' && c.url.includes('%E5%90%88%E5%90%8C.doc')), 'the .doc is never written');
});

test('a native document opens as before, without a conversion', async () => {
  const calls = stub({ workspace: '/U/Docs', drafts: '/U/app/Drafts' }); await EN.files();
  const doc = await EN.open(EN.lazy('/U/Docs/a.docx'));
  assert.equal(doc.path, '/U/Docs/a.docx');
  assert.equal(doc.from, undefined);
  assert.ok(!calls.some(c => c.url === '/run'));
});
