// node --test ui/tests/
import { test } from 'node:test';
import assert from 'node:assert/strict';
globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const K = await import('../pdf-kit.js');

test('the text layer is laid out unrotated at scale 1 and rotated and zoomed with one transform, as pdf.js does', () => {
  assert.equal(K.layerTransform(0, 1.5), 'scale(1.5)');
  assert.equal(K.layerTransform(90, 1), 'scale(1) rotate(90deg) translateY(-100%)');
  assert.equal(K.layerTransform(-90, 2), 'scale(2) rotate(270deg) translateX(-100%)');
  assert.equal(K.layerTransform(180, 1), 'scale(1) rotate(180deg) translate(-100%,-100%)');
});

test('a page\'s text keeps the PDF\'s line ends and skips marked-content items', () => {
  const items = [{ str: 'Quarterly ', hasEOL: false }, { type: 'beginMarkedContent' }, { str: 'Report', hasEOL: true }, { str: '', hasEOL: true }, { str: 'Costs', hasEOL: false }];
  assert.equal(K.textOf(items), 'Quarterly Report\n\nCosts');
});

test('search: case-insensitive hits across pages, a space in the query also matching a line end, numbered in order', () => {
  const hits = K.findIn(['Quarterly Report\nRevenue grew', 'Second page: revenue\ngrew again'], 'revenue grew');
  assert.deepEqual(hits, [{ p: 0, start: 17, end: 29, i: 0 }, { p: 1, start: 13, end: 25, i: 1 }]);
  assert.deepEqual(K.findIn(['a.b a+b'], 'a.b'), [{ p: 0, start: 0, end: 3, i: 0 }], 'regex characters are literal');
  assert.deepEqual(K.findIn(['abc'], '  '), []);
});

test('selection rects merge into one box per line', () => {
  const lines = K.mergeRects([{ x: 50, y: 100, w: 40, h: 12 }, { x: 92, y: 100.5, w: 60, h: 12 }, { x: 0, y: 0, w: 0, h: 12 }, { x: 50, y: 116, w: 30, h: 12 }]);
  assert.deepEqual(lines, [{ x: 50, y: 100, w: 102, h: 12.5 }, { x: 50, y: 116, w: 30, h: 12 }]);
});

test('page → PDF space follows the page\'s rotation', () => {
  const box = [0, 0, 595, 842];
  assert.deepEqual(K.toPdf(10, 20, 0, box), [10, 822]);
  assert.deepEqual(K.toPdf(10, 20, 90, box), [20, 10]);
  assert.deepEqual(K.toPdf(10, 20, 180, box), [585, 20]);
  assert.deepEqual(K.toPdf(10, 20, 270, box), [575, 832]);
  assert.deepEqual(K.toPdf(10, 20, 0, [20, 30, 615, 872]), [30, 852], 'a media box away from the origin');
});

test('annotations become standard /Annot objects: quads flipped to PDF space, ink points, a note with its icon', () => {
  const o = { num: 30, pageRef: '4 0 R', rot: 0, box: [0, 0, 595, 842] };
  const hl = K.annotObjects({ id: 'h1', t: 'hl', quads: [[50, 50, 100, 20]], color: '#FFE066' }, o);
  assert.equal(hl.next, 31);
  assert.match(hl.objs[0].body, /\/Subtype \/Highlight \/Rect \[50 772 150 792\] \/P 4 0 R/);
  assert.match(hl.objs[0].body, /\/QuadPoints \[50 792 150 792 50 772 150 772\] \/C \[1 0\.88 0\.4\]/);
  const ul = K.annotObjects({ id: 'u', t: 'so', quads: [[0, 0, 10, 10]], color: '#000000' }, Object.assign({}, o, { rot: 90 }));
  assert.match(ul.objs[0].body, /\/Subtype \/StrikeOut \/Rect \[0 0 10 10\]/);
  const ink = K.annotObjects({ id: 'k', t: 'ink', pts: [[10, 10], [20, 30]], color: '#2F5D8A', sw: 2 }, o);
  assert.match(ink.objs[0].body, /\/Subtype \/Ink \/Rect \[8 810 22 834\]/);
  assert.match(ink.objs[0].body, /\/InkList \[\[10 832 20 812\]\] \/C \[0\.18 0\.36 0\.54\] \/BS << \/W 2 >>/);
  const note = K.annotObjects({ id: 'n', t: 'note', x: 100, y: 100, text: '中', color: '#FFD54A' }, o);
  assert.equal(note.objs.length, 2);
  assert.match(note.objs[0].body, /\/Subtype \/Text \/Rect \[100 722 120 742\] .* \/F 28 .*\/Contents <FEFF4E2D> \/Name \/Comment .*\/AP << \/N 31 0 R >>/);
  assert.match(note.objs[1].dict, /\/Subtype \/Form \/BBox \[0 0 20 20\]/);
  const sq = K.annotObjects({ id: 'w', t: 'white', x: 1, y: 1, w: 2, h: 2 }, o);
  assert.match(sq.objs[0].body, /\/Subtype \/Square .*\/C \[1 1 1\] \/BS << \/W 0 \/S \/S >> \/IC \[1 1 1\]/);
  assert.deepEqual(K.annotObjects({ id: 't', t: 'text', x: 0, y: 0, w: 10, h: 10, text: 'x' }, o).objs, [], 'a text annotation needs its appearance image');
});

test('an incremental update keeps the original bytes and chains a cross-reference table, or a stream when the file ends with one', () => {
  const pdf = '%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>\nendobj\n';
  const xref = pdf.length, tail = `xref\n0 4\n0000000000 65535 f \n0000000009 00000 n \n0000000058 00000 n \n0000000115 00000 n \ntrailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  const bytes = Uint8Array.from(pdf + tail, c => c.charCodeAt(0));
  assert.equal(K.firstFree(bytes), 4);
  const out = K.appendUpdate(bytes, [{ num: 4, body: '<< /Type /Annot >>' }, { num: 3, gen: 0, body: '<< /Type /Page /Annots [4 0 R] >>' }, { num: 5, dict: '<< /Length 3 >>', stream: Uint8Array.from([65, 66, 67]) }], { size: 6, root: '1 0 R', id: '[<AB> <CD>]' });
  const s = String.fromCharCode(...out);
  assert.equal(s.slice(0, bytes.length), pdf + tail, 'the original is untouched');
  const upd = s.slice(bytes.length);
  const m = /startxref\n(\d+)\n%%EOF\n$/.exec(upd); assert.ok(m);
  assert.equal(s.slice(+m[1], +m[1] + 4), 'xref', 'startxref points at the new table');
  assert.match(upd, /xref\n3 3\n(\d{10}) 00000 n \n(\d{10}) 00000 n \n(\d{10}) 00000 n \ntrailer\n<< \/Size 6 \/Root 1 0 R \/Prev \d+ \/ID \[<AB> <CD>\] >>/);
  const offs = /xref\n3 3\n(\d{10}) 00000 n \n(\d{10}) 00000 n \n(\d{10})/.exec(upd).slice(1).map(Number);
  assert.equal(s.slice(offs[0], offs[0] + 7), '3 0 obj'); assert.equal(s.slice(offs[1], offs[1] + 7), '4 0 obj'); assert.equal(s.slice(offs[2], offs[2] + 7), '5 0 obj');
  assert.match(upd, new RegExp('/Prev ' + xref + ' '), 'chained to the original table');
  assert.match(upd, /5 0 obj\n<< \/Length 3 >>\nstream\nABC\nendstream\nendobj/);
  // a file whose last section is a cross-reference stream gets a stream back
  const streamed = Uint8Array.from(pdf + `4 0 obj\n<< /Type /XRef /Size 5 /W [1 4 2] /Root 1 0 R /Length 0 >>\nstream\n\nendstream\nendobj\nstartxref\n${xref}\n%%EOF\n`, c => c.charCodeAt(0));
  const out2 = String.fromCharCode(...K.appendUpdate(streamed, [{ num: 5, body: '<< >>' }], { size: 6, root: '1 0 R' }));
  assert.match(out2, /6 0 obj\n<< \/Type \/XRef \/Size 7 \/W \[1 4 2\] \/Index \[5 2\] \/Root 1 0 R \/Prev \d+ \/Length 14 >>\nstream\n/);
  assert.throws(() => K.appendUpdate(Uint8Array.from('%PDF-1.4 garbage', c => c.charCodeAt(0)), [], { size: 1, root: '1 0 R' }), /xref/);
});

test('a page list: ranges, single pages, Chinese commas, out-of-range pages dropped, nonsense rejected', () => {
  assert.deepEqual(K.parseRange('1-3,5', 9), [0, 1, 2, 4]);
  assert.deepEqual(K.parseRange('5，3-1 12', 9), [0, 1, 2, 4]);
  assert.deepEqual(K.parseRange('', 9), []);
  assert.equal(K.parseRange('a-b', 9), null);
});

test('dirty: a document differs from its file when it has pending annotations, removals, form edits or a changed page list', () => {
  K.setGen('d', 1, new Uint8Array([37, 80, 68, 70]));
  const clean = { _gen: 1, _n: 2, pages: [{ id: 'a', src: 0, rot: 0, from: 'd@1' }, { id: 'b', src: 1, rot: 0, from: 'd@1' }], annots: [], removed: [], form: 0, _formSaved: 0 };
  assert.equal(K.dirty('d', clean), false);
  assert.equal(K.dirty('d', Object.assign({}, clean, { annots: [{ id: 'x' }] })), true);
  assert.equal(K.dirty('d', Object.assign({}, clean, { removed: ['12R'] })), true);
  assert.equal(K.dirty('d', Object.assign({}, clean, { form: 1 })), true);
  assert.equal(K.dirty('d', Object.assign({}, clean, { pages: clean.pages.slice().reverse() })), true);
  assert.equal(K.dirty('d', Object.assign({}, clean, { pages: [clean.pages[0], Object.assign({}, clean.pages[1], { rot: 90 })] })), true);
  assert.equal(K.dirty('d', Object.assign({}, clean, { _gen: 0 })), true, 'an undo to before the last save');
  assert.equal(K.baseKey('d', clean), 'd@1');
});
