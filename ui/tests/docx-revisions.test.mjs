// node --test ui/tests/   (Node 20) — the docx bridge's pure parts for tracked changes and comments:
// canonical runs treat ins/del as revisions, comments come out of the tree and go back as add/set/remove.
import { test } from 'node:test';
import assert from 'node:assert/strict';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');

// A tiny DOM stand-in: runsOf accepts a parsed root, so no DOMParser is needed.
const text = t => ({ nodeType: 3, nodeValue: t });
const el = (tag, children, attrs = {}) => ({ nodeType: 1, tagName: tag, style: {}, childNodes: children, getAttribute: k => attrs[k] ?? null });

test('runsOf keeps inserted and deleted text apart from plain text and never reads ins as underline', () => {
  const plain = JSON.parse(EN.runsOf(el('BODY', [el('P', [text('Keep '), text('old')])])));
  const tracked = JSON.parse(EN.runsOf(el('BODY', [el('P', [text('Keep '), el('DEL', [text('old')], { 'data-t': '1' }), el('INS', [text('new\u200B')], { 'data-t': '1' })])])));
  assert.equal(plain.length, 1, 'plain text is one run');
  assert.equal(tracked.length, 3);
  assert.equal(tracked[1].t, 'old', 'deleted text is kept, not dropped');
  assert.equal(tracked[2].t, 'new', 'zero-width fillers are ignored');
  assert.notEqual(tracked[1].s, plain[0].s);
  assert.notEqual(tracked[2].s, plain[0].s);
  assert.notEqual(tracked[1].s, tracked[2].s);
  const underlined = JSON.parse(EN.runsOf(el('BODY', [el('P', [el('U', [text('new')])])])));
  assert.notEqual(underlined[0].s, tracked[2].s, 'ins is a revision, not an underline');
  assert.equal(EN.runsOf(el('BODY', [el('P', [el('INS', [text('a')], { 'data-author': 'Ann' })])])), EN.runsOf(el('BODY', [el('P', [el('INS', [text('a')], { 'data-t': '1' })])])), 'author and date do not make a difference');
});

test('trackHtml gives the engine ins/del the editor marker once', () => {
  assert.equal(EN.trackHtml('a <ins data-author="Ann" data-date="2026-01-01T00:00:00Z">b</ins><del>c</del> <ins data-t="1">d</ins>'),
    'a <ins data-t="1" data-author="Ann" data-date="2026-01-01T00:00:00Z">b</ins><del data-t="1">c</del> <ins data-t="1">d</ins>');
  assert.equal(EN.trackHtml('<p>insert</p>'), '<p>insert</p>');
});

test('commentsOf collects comments under paragraphs anywhere in the tree with their engine ids', () => {
  const list = EN.commentsOf([
    { kind: 'paragraph', path: '/body/paragraph[1]', props: { text: 'a' }, children: [
      { kind: 'run', path: '/body/paragraph[1]/run[1]', props: { text: 'a' } },
      { kind: 'comment', path: '/body/paragraph[1]/comment[1]', props: { id: '3', author: 'Ann', date: '2026-01-15T09:30:00Z', text: 'why', quote: 'a', resolved: 'true' } }
    ] },
    { kind: 'table', path: '/body/table[1]', props: {}, children: [{ kind: 'row', path: '/body/table[1]/row[1]', children: [{ kind: 'cell', path: '/body/table[1]/row[1]/cell[1]', children: [
      { kind: 'paragraph', path: '/body/table[1]/row[1]/cell[1]/paragraph[1]', props: {}, children: [{ kind: 'comment', path: '/body/table[1]/row[1]/cell[1]/paragraph[1]/comment[1]', props: { id: '0', text: 'in a cell' } }] }
    ] }] }] }
  ]);
  assert.equal(list.length, 2);
  assert.deepEqual({ cid: list[0].cid, id: list[0].id, path: list[0].path, author: list[0].author, text: list[0].text, quote: list[0].quote, resolved: list[0].resolved },
    { cid: '3', id: '3', path: '/body/paragraph[1]', author: 'Ann', text: 'why', quote: 'a', resolved: true });
  assert.match(list[0].time, /^Ann · 1月15日 \d\d:\d\d$/);
  assert.deepEqual([list[1].cid, list[1].path, list[1].resolved, list[1].time], ['0', '/body/table[1]/row[1]/cell[1]/paragraph[1]', false, '']);
});

test('commentsOf keeps each author\'s initials and tells my comments (written as the document\'s author) from the others', () => {
  const list = EN.commentsOf([{ kind: 'paragraph', path: '/body/paragraph[1]', props: {}, children: [
    { kind: 'comment', path: '/body/paragraph[1]/comment[1]', props: { id: '0', author: 'Ann Lee', initials: 'AL', text: 'a' } },
    { kind: 'comment', path: '/body/paragraph[1]/comment[2]', props: { id: '1', author: 'Writer', text: 'b' } }] }], 'Writer');
  assert.deepEqual(list.map(c => [c.author, c.initials, c.mine]), [['Ann Lee', 'AL', false], ['Writer', '', true]]);
});


const addExec = (calls, fail) => async argv => {
  calls.push(argv);
  if (fail && fail(argv)) throw new Error('refused');
  return argv[0] === 'add' ? { path: argv[2] + '/comment[1]', props: { id: '7', author: 'Writer', date: '2026-02-02T00:00:00Z' } } : {};
};

test('planComments adds, updates and removes by engine id and leaves unchanged or unanchored comments alone', async () => {
  const calls = [];
  const orig = [{ cid: '3', id: '3', path: '/body/paragraph[1]', text: 'why', resolved: false }, { cid: '4', id: '4', path: '/body/paragraph[2]', text: 'gone', resolved: false }, { cid: '5', id: '5', path: '/body/paragraph[3]', text: 'same', resolved: true }];
  const current = [
    { cid: '3', text: 'why not', resolved: true, parent: '/body/paragraph[1]', origin: '/body/paragraph[1]', quote: 'a' },
    { cid: '5', text: 'same', resolved: true, parent: '/body/paragraph[2]', origin: '/body/paragraph[3]', quote: '' },
    { cid: 'cx1', text: 'new one', resolved: false, parent: '/body/paragraph[3]', origin: null, quote: 'grew 25%' },
    { cid: 'cx2', text: 'homeless', resolved: false, parent: null, origin: null, quote: '' }
  ];
  const r = await EN.planComments('a.docx', orig, current, null, addExec(calls));
  assert.deepEqual(calls, [
    ['set', 'a.docx', '//comment[@id=3]', '--prop', 'text=why not', '--prop', 'resolved=true'],
    ['add', 'a.docx', '/body/paragraph[3]', '--type', 'comment', '--prop', 'text=new one', '--prop', 'quote=grew 25%'],
    ['remove', 'a.docx', '//comment[@id=4]']
  ]);
  assert.equal(r.count, 3);
  assert.deepEqual(r.list.map(c => [c.cid, c.id, c.path, c.text, c.resolved]),
    [['3', '3', '/body/paragraph[1]', 'why not', true], ['5', '5', '/body/paragraph[2]', 'same', true], ['cx1', '7', '/body/paragraph[3]', 'new one', false]],
    'paths follow the renumbered paragraphs');

  calls.length = 0;
  const again = await EN.planComments('a.docx', r.list, r.list.map(c => ({ cid: c.cid, text: c.text, resolved: c.resolved, parent: c.path, origin: c.path, quote: '' })), null, addExec(calls));
  assert.deepEqual(calls, [], 'nothing differs, nothing runs');
  assert.equal(again.count, 0);
});

test('planComments anchors a comment again when its span moved to another paragraph, keeping author and date', async () => {
  const calls = [];
  const orig = [{ cid: '2', id: '2', path: '/body/paragraph[2]', author: 'Ann', date: '2026-01-15T09:30:00Z', text: 'check', resolved: true }];
  const merged = [{ cid: '2', text: 'check', resolved: true, parent: '/body/paragraph[1]', origin: '/body/paragraph[1]', quote: 'fox' }];
  const r = await EN.planComments('a.docx', orig, merged, null, addExec(calls, argv => argv[0] === 'remove'));
  assert.deepEqual(calls, [
    ['remove', 'a.docx', '//comment[@id=2]'],
    ['add', 'a.docx', '/body/paragraph[1]', '--type', 'comment', '--prop', 'text=check', '--prop', 'author=Ann', '--prop', 'date=2026-01-15T09:30:00Z', '--prop', 'quote=fox', '--prop', 'resolved=true']
  ], 'the old comment went with its paragraph, so the failed remove is fine');
  assert.deepEqual(r.list.map(c => [c.cid, c.id, c.path]), [['2', '7', '/body/paragraph[1]']]);
  assert.equal(r.count, 1);

  calls.length = 0;
  await EN.planComments('a.docx', [Object.assign({}, orig[0], { initials: 'AL' })], merged, null, addExec(calls));
  assert.ok(calls[1].includes('initials=AL'), 'and the initials Word shows for its author');
});

test('planComments falls back to the whole paragraph when the quote is not found, and drops comments already gone', async () => {
  const calls = [];
  const exec = addExec(calls, argv => (argv[0] === 'add' && argv.includes('quote=across\nlines')) || argv[0] === 'set');
  const r = await EN.planComments('a.docx', [{ cid: '9', id: '9', path: '/body/paragraph[4]', text: 'old', resolved: false }], [
    { cid: 'c1', text: 't', resolved: false, parent: '/body/paragraph[1]', origin: '/body/paragraph[1]', quote: 'across\nlines' },
    { cid: '9', text: 'edited', resolved: false, parent: null, origin: null, quote: '' }
  ], null, exec);
  assert.deepEqual(calls.map(a => a.slice(0, 3).concat(a.filter(x => /^quote=/.test(x)))), [
    ['add', 'a.docx', '/body/paragraph[1]', 'quote=across\nlines'],
    ['add', 'a.docx', '/body/paragraph[1]'],
    ['set', 'a.docx', '//comment[@id=9]']
  ]);
  assert.deepEqual(r.list.map(c => c.cid), ['c1'], 'the comment whose paragraph is gone is not remembered');
  assert.equal(r.count, 1);
});
