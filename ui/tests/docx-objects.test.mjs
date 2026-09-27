import { test } from 'node:test';
import assert from 'node:assert/strict';
import { install } from './dom-stub.mjs';
globalThis.window = globalThis;
const parse = install();
const EN = await import('../engine.js');
const obj = (path, props = {}) => ({ kind: 'object', path, props: { type: 'ole', progId: 'Equation.DSMT4', width: '2.54cm', height: '1.27cm', ...props } });

test('standalone Office objects remain objects through render, unchanged save and explicit deletion', async () => {
  const before = EN.blocksOf([obj('/body/object[1]'), obj('/body/object[2]')], 'a.docx');
  const root = parse(EN.blocksToHtml(before));
  assert.equal(root.children.length, 2);
  assert.equal(root.children[0].getAttribute('contenteditable'), 'false');
  assert.match(root.innerHTML, /width:96px;height:48px/);
  const calls = [], run = async argv => calls.push(argv);
  assert.equal(await EN.planDocxBlocks('a.docx', before, EN.blocksFromHtml(root), run), 0);
  root.children[0].remove();
  await EN.planDocxBlocks('a.docx', before, EN.blocksFromHtml(root), run);
  assert.deepEqual(calls, [['remove', 'a.docx', '/body/object[1]']]);
});

test('inline preview labels and images never become paragraph text or loose pictures', async () => {
  const nodes = [{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'abcdef' }, children: [obj('/body/paragraph[1]/object[1]', { at: 3, src: '/word/media/preview.png' }), obj('/body/paragraph[1]/object[2]', { at: 4 })] }];
  const objects = EN.officeObjectsOf(nodes, 'a.docx');
  const root = parse(EN.anchorObjects(EN.blocksToHtml(EN.blocksOf(nodes, 'a.docx')), [], [], objects));
  const [b] = EN.blocksFromHtml(root);
  assert.equal(b.props.html, 'abcdef');
  assert.equal(b.pics, undefined);
  const marks = root.querySelectorAll('[data-office-object]');
  assert.equal(EN.offsetIn(root.children[0], marks[1]), 4);
  marks[0].remove();
  const calls = [];
  const result = await EN.planOfficeObjects('a.docx', objects, root, async argv => calls.push(argv));
  assert.deepEqual(calls, [['remove', 'a.docx', '/body/paragraph[1]/object[1]']]);
  assert.equal(result.list[0].path, '/body/paragraph[1]/object[1]');
  assert.equal(marks[1].getAttribute('data-path'), result.list[0].path);
});

test('a chart has a drawn SVG and its labels cannot leak into the saved text', () => {
  const [b] = EN.blocksOf([obj('/body/object[1]', { type: 'chart', chart: { kind: 'bar', cats: ['甲', '乙'], series: [{ name: '收入', values: [3, 5] }] } })], 'a.docx');
  assert.match(EN.blocksToHtml([b]), /<svg/);
  assert.equal(EN.blocksFromHtml(parse(EN.blocksToHtml([b])))[0].kind, 'object');
});
