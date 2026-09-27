// node --test ui/tests/ — citations in the Word editor: a citation (Word's citation control) sits in its paragraph at a character
// offset and shows what the engine drew; its text is not the paragraph's; the save adds, moves, changes and removes citations as
// commands; a works-cited list is a block drawn from the sources. The round trip runs the real engine when the CLI is built.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { install } from './dom-stub.mjs';

globalThis.window = globalThis;
const parse = install();
const EN = await import('../engine.js');

const PEGG = { type: 'article', authors: [{ last: 'Pegg', first: 'Ian L.' }], title: 'Behavior of technetium in nuclear waste vitrification processes', container: 'Journal of Radioanalytical and Nuclear Chemistry', year: '2015', volume: '305', issue: '1', pages: '287-292', doi: '10.1007/s10967-014-3900-9' };

test('a citation sits at its offset, shows the engine\'s text and is not part of its paragraph\'s', () => {
  const tree = [{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'Glass holds it .' }, children: [{ kind: 'citation', path: '/body/paragraph[1]/citation[1]', props: { id: '77', sources: 'Peg15', pages: '288', at: '15', text: '(Pegg 288)', html: '(Pegg 288)' } }] }];
  const cites = EN.citesOf(tree);
  assert.deepEqual(cites, [{ id: '77', path: '/body/paragraph[1]', at: 15, sources: 'Peg15', pages: '288', noAuthor: false, noYear: false, html: '(Pegg 288)' }]);
  const html = EN.anchorCites(EN.blocksToHtml(EN.blocksOf(tree, 'a.docx')), cites);
  assert.match(html, /Glass holds it <span data-cite="Peg15" data-citeid="77" data-pages="288" contenteditable="false">\(Pegg 288\)<\/span>\.<\/p>/);
  const el = parse(html), span = el.querySelector('span[data-cite]');
  assert.equal(EN.offsetIn(el.firstChild, span), 15);
  assert.equal(EN.blocksFromHtml(el)[0].props.html, 'Glass holds it .', 'the save writes the paragraph without it');
});

test('the save: a new citation is added where it sits and learns its id, a moved one again, a changed one is set, one gone is removed', async () => {
  const sent = [], exec = async argv => { sent.push(argv.slice(2).join(' ')); return argv[0] === 'add' ? { path: argv[2] + '/citation[1]', props: { id: '901', html: '(Pegg 288)' } } : {}; };
  const orig = [{ id: '77', path: '/body/paragraph[1]', at: 15, sources: 'Peg15', pages: '', noAuthor: false, noYear: false }, { id: '78', path: '/body/paragraph[2]', at: 3, sources: 'Kuh62', pages: '' }];
  const el = parse('<p data-path="/body/paragraph[1]">Glass holds it <span data-cite="Peg15" data-citeid="77" data-pages="288">(Pegg)</span>.</p><p data-path="/body/paragraph[2]">New <span data-cite="Kuh62">(Kuhn)</span> one.</p>');
  const blocks = EN.blocksFromHtml(el);
  const cur = Array.from(el.querySelectorAll('span[data-cite]')).map(s => { const b = blocks.find(x => x.el.contains(s)); return { el: s, id: s.getAttribute('data-citeid') || '', sources: s.getAttribute('data-cite'), pages: s.getAttribute('data-pages') || '', noAuthor: false, noYear: false, parent: b.path, at: EN.offsetIn(b.el, s) }; });
  const r = await EN.planCites('a.docx', orig, cur, null, exec);
  assert.deepEqual(sent, ['//citation[@id=77] --prop pages=288', '/body/paragraph[2] --type citation --prop sources=Kuh62 --prop at=4', '//citation[@id=78]'], 'pages set on the one kept; the one in paragraph 2 is new (Kuh62 had id 78 at 3: that one is gone)');
  assert.equal(r.count, 3);
  assert.equal(el.querySelectorAll('span[data-cite]')[1].getAttribute('data-citeid'), '901', 'the new one knows its id for the next save');
  assert.equal(el.querySelectorAll('span[data-cite]')[1].textContent, '(Pegg 288)', 'and shows what the engine drew');
  sent.length = 0;
  await EN.planCites('a.docx', r.list, cur.map(x => Object.assign({}, x, { id: x.el.getAttribute('data-citeid') })), null, exec);
  assert.deepEqual(sent, [], 'saved again unchanged: nothing to write');
});

test('a works-cited list is a block the editor does not type into, on a page of its own, saved as the engine\'s bibliography', () => {
  const tree = [{ kind: 'bibliography', path: '/body/bibliography[1]', props: { title: 'Works Cited', html: '<p>Pegg, Ian L. <i>X</i>.</p>' } }];
  const html = EN.blocksToHtml(EN.blocksOf(tree, 'a.docx'));
  assert.equal(html, '<nav data-bib="1" data-path="/body/bibliography[1]" data-title="Works Cited" data-pb="before" contenteditable="false"><div class="wd-bib-title">Works Cited</div><p>Pegg, Ian L. <i>X</i>.</p></nav>');
  assert.deepEqual(EN.blocksFromHtml(parse(html)).map(b => [b.kind, b.path, b.props.title]), [['bibliography', '/body/bibliography[1]', 'Works Cited']]);
  assert.equal(EN.listTitle('apa'), 'References');
  assert.match(EN.bibHtml('', 'References', ''), /data-title="References"[^>]*><div class="wd-bib-title">References<\/div><p class="wd-bib-empty">/, 'a new list waits for the save');
});

test('a note\'s citation as the engine drew it at the save reaches the editor: a new note by the id the save gave it, the doc edited since too', async () => {
  const exec = async argv => argv[2] === '//footnote' ? [{ kind: 'footnote', path: '/body/paragraph[2]/footnote[1]', props: { id: '1', cite: 'Peg15', citeHtml: 'Ian L. Pegg, “Behavior”' } }] : [];
  const doc = { path: 'a.docx', notes: [{ id: 'nabc', kind: 'footnote', text: '', cite: 'Peg15', pages: '', citeHtml: 'Pegg' }], _orig: { blocks: [], cites: [], notes: [{ nid: 'nabc', id: '1', cite: 'Peg15' }] } };
  await EN.refreshCites(doc, null, exec);
  assert.equal(doc.notes[0].citeHtml, 'Ian L. Pegg, “Behavior”', 'the note added in this session is the file\'s footnote 1');
  const typed = { type: 'docx', notes: [{ id: 'nabc', kind: 'footnote', text: 'See also', cite: 'Peg15', pages: '', citeHtml: 'Pegg' }, { id: 'nx', kind: 'footnote', text: '', cite: 'Kuh62', citeHtml: 'Kuhn' }] };
  const saved = { type: 'docx', notes: [doc.notes[0], { id: 'nx', kind: 'footnote', text: '', cite: 'Peg15', citeHtml: 'Pegg, “Behavior,”' }] };
  EN.adopt(typed, saved, false);
  assert.deepEqual(typed.notes.map(x => [x.text, x.citeHtml]), [['See also', 'Ian L. Pegg, “Behavior”'], ['', 'Kuhn']], 'what was typed since stays; a note citing another source now keeps its own');
});

test('a tab shows in a span of its own that keeps it (Word\'s half-inch stops); the save writes the tab alone', () => {
  const el = parse('<p data-path="/body/paragraph[1]">Name:\tValue\tand <b>more\there</b></p>');
  EN.wrapTabs(el);
  assert.equal(el.innerHTML, '<p data-path="/body/paragraph[1]">Name:<span class="wd-tab">\t</span>Value<span class="wd-tab">\t</span>and <b>more<span class="wd-tab">\t</span>here</b></p>');
  assert.equal(EN.blocksFromHtml(el)[0].props.html, 'Name:\tValue\tand <b>more\there</b>');
  EN.wrapTabs(el); assert.equal(el.querySelectorAll('.wd-tab').length, 3, 'once only');
});

// ----- round trips through the engine -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-cites-'));
  server = spawn('dotnet', [cli, 'serve', '--dir', dir, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  url = await new Promise((resolve, reject) => {
    let err = '';
    server.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1]); });
    server.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
  });
});
after(() => { server && server.kill(); dir && rmSync(dir, { recursive: true, force: true }); });
const run = async argv => {
  const r = await (await fetch(url + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) })).json();
  if (r.code !== 0) throw new Error(argv.join(' ') + ' → ' + JSON.stringify(r.error));
  return r.output && /^[{[]/.test(r.output.trim()) ? JSON.parse(r.output) : r.output;
};
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';

test('engine: a source, a citation and a works-cited list saved, then the style changed: all drawn again and shown', { skip: skip() }, async () => {
  const file = join(dir, 'essay.docx');
  await run(['create', file]);
  await run(['add', file, '/body', '--type', 'paragraph', '--prop', 'text=Glass holds most of it .']);
  const doc = { path: file, sources: [], _orig: { blocks: [], cites: [] } };
  const r = await EN.citeEdit(doc, PEGG, null, run);
  assert.equal(r.tag, 'Peg15');
  assert.equal(doc.citeStyle, 'mla', 'a document starts in MLA');
  const el = parse('<p data-path="/body/paragraph[1]">Glass holds most of it <span data-cite="Peg15" data-pages="288">(Pegg)</span>.</p>' + EN.bibHtml('', 'Works Cited', ''));
  const blocks = EN.blocksFromHtml(el);
  await EN.planDocxBlocks(file, [{ kind: 'paragraph', path: '/body/paragraph[1]', props: { html: 'Glass holds most of it .' } }], blocks, run);
  const cur = [{ el: el.querySelector('span'), id: '', sources: 'Peg15', pages: '288', parent: '/body/paragraph[1]', at: 23 }];
  const saved = await EN.planCites(file, [], cur, null, run);
  doc._orig.cites = saved.list;
  await EN.refreshCites(doc, el, run);
  assert.equal(el.querySelector('span[data-cite]').textContent, '(Pegg 288)');
  assert.match(el.querySelector('nav[data-bib]').textContent, /^Works CitedPegg, Ian L\. “Behavior of Technetium in Nuclear Waste Vitrification Processes\.”/);
  await EN.citeEdit(doc, { citationStyle: 'apa' }, el, run);
  assert.equal(el.querySelector('span[data-cite]').textContent, '(Pegg, 2015, p. 288)');
  assert.equal(el.querySelector('nav[data-bib]').getAttribute('data-title'), 'References');
  assert.equal((await run(['view', file, 'text'])).split('\n')[0], 'Glass holds most of it (Pegg, 2015, p. 288).');
  await assert.rejects(EN.citeEdit(doc, { tag: 'Peg15', remove: true }, el, run), /cited/, 'a cited source stays');
});
