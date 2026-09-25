// node --test ui/tests/ — the new-file gallery (ui/templates, built by ui/tools/templates/build.mjs from its sources): every
// template a source lists is in index.json, exists in Chinese and English with its thumbnail, has its English name in
// ui/i18n/en-shell.js, and opens through the engine; a deck has four to six slides. Needs the CLI built (dotnet build).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, readFileSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..'), TPL = join(root, 'ui/templates');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
const MIN = { docx: 12, xlsx: 8, pptx: 10, md: 3, mm: 3 }; // what the gallery promises, at least
const sources = {};
for (const type of Object.keys(MIN)) sources[type] = await import(`../tools/templates/${type}.mjs`);

let server, base;
before(async () => {
  if (!existsSync(cli)) return;
  server = spawn('dotnet', [cli, 'serve', '--dir', TPL, '--port', '0', '--no-token'], { stdio: ['ignore', 'ignore', 'pipe'] });
  base = await new Promise((resolve, reject) => {
    let err = '';
    server.stderr.on('data', d => { err += d; const m = /"url"\s*:\s*"([^"]+)"/.exec(err); if (m) resolve(m[1]); });
    server.on('exit', code => reject(new Error('writer serve exited ' + code + ': ' + err)));
  });
});
after(() => server && server.kill());
async function run(argv) {
  const r = await (await fetch(base + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ argv }) })).json();
  assert.equal(r.code, 0, argv.join(' ') + ' → ' + JSON.stringify(r.error));
  return JSON.parse(r.output);
}

test('every template of every source is built in both languages with a thumbnail, listed, named in English, and opens in the engine', { skip: !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx' }, async () => {
  const index = JSON.parse(readFileSync(join(TPL, 'index.json'), 'utf8'));
  const win = {}; new Function('window', readFileSync(join(root, 'ui/i18n/en-shell.js'), 'utf8'))(win);
  const there = f => assert.ok(existsSync(join(TPL, f)) && statSync(join(TPL, f)).size > 0, f + ' is built');
  let count = 0;
  for (const [type, src] of Object.entries(sources)) {
    assert.ok(src.default.length >= MIN[type], `${type}: ${src.default.length} templates, at least ${MIN[type]} promised`);
    const cats = new Map(src.cats);
    for (const tpl of src.default) {
      count++;
      const listed = index.find(x => x.type === type && x.id === tpl.id);
      assert.deepEqual(listed, { id: tpl.id, type, cat: tpl.cat, name: tpl.name[0] }, `${type}/${tpl.id} is in index.json as its source says`);
      assert.equal(win.I18N_EN[tpl.name[0]], tpl.name[1], `${type}/${tpl.id}: its English name is in ui/i18n/en-shell.js`);
      assert.equal(win.I18N_EN[tpl.cat], cats.get(tpl.cat), `${type}/${tpl.id}: its category is in ui/i18n/en-shell.js`);
      for (const lang of ['zh', 'en']) {
        const file = `${lang}/${type}/${tpl.id}.${type}`;
        there(file); there(`${lang}/${type}/${tpl.id}.webp`);
        const doc = await run(['get', file, '/', '--depth', '1']);
        assert.equal(doc.kind, 'document', file + ' opens');
        assert.ok(type === 'md' || (doc.children || []).length > 0, file + ' has content');
        if (type === 'pptx') { const n = doc.children.filter(c => c.kind === 'slide').length; assert.ok(n >= 4 && n <= 6, `${file}: ${n} slides, four to six expected`); }
      }
    }
  }
  assert.equal(index.length, count, 'index.json lists exactly the templates of the sources');
});
