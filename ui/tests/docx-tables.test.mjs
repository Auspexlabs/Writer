// node --test ui/tests/ — Word tables in the editor: the grid model behind the table ribbon, the commands a save sends for it,
// and (when the CLI is built) those commands run by the real engine, which must end with exactly the editor's table.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { existsSync, mkdtempSync, copyFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

globalThis.window = globalThis; // pdf-kit.js keeps its store on window at import time
const EN = await import('../engine.js');

// ----- helpers: tables as blocks, and the editor's side of a command without a DOM -----
const P = (r, c) => `/body/table[1]/row[${r}]/cell[${c}]`;
/** A table block from rows of cell specs: 'text' or { html, colspan, rowspan, fill }. */
function table(rows, props = {}) {
  return { kind: 'table', path: '/body/table[1]', props, ops: [], rows: rows.map((cells, r) => ({ kind: 'row', path: `/body/table[1]/row[${r + 1}]`, props: {},
    cells: cells.map((c, i) => ({ kind: 'cell', path: P(r + 1, i + 1), props: Object.assign({ html: typeof c === 'string' ? c : c.html || '' },
      typeof c === 'string' ? {} : Object.fromEntries(Object.entries(c).filter(([k]) => k !== 'html'))) })) })) };
}
const modelOf = b => EN.gridModel(b.rows.map(r => ({ id: r.path, ref: r, cells: r.cells.map(c => ({ id: c.path, ref: c, cs: c.props.colspan, rs: c.props.rowspan })) })));
const shape = m => m.rows.map(r => r.cells.map(x => `${x.c}:${x.cs}x${x.rs}`).join(' '));
let seq = 0; const newId = () => '~t' + (++seq);
/** What the editor does for a ribbon command, on blocks instead of elements (runTableOp does the same on the DOM). */
function edit(b, op, cell, flag) {
  const r = EN.tableCommand(modelOf(b), op, cell, flag, newId);
  const byId = new Map(modelOf(b).rows.flatMap(row => [[row.id, row.ref], ...row.cells.map(x => [x.id, x.ref])]));
  if (r.into) byId.get(r.into).props.html = EN.joinCells(byId.get(r.into).props.html, byId.get(r.from).props.html);
  const spans = x => Object.assign({}, x.cs > 1 ? { colspan: String(x.cs) } : {}, x.rs > 1 ? { rowspan: String(x.rs) } : {});
  b.rows = r.m.rows.map(row => {
    const cells = row.cells.map(x => {
      const own = x.ref || { kind: 'cell', path: x.id, props: Object.fromEntries(Object.entries((byId.get(x.from) || { props: {} }).props).filter(([k]) => ['fill', 'borders', 'valign', 'width', 'align'].includes(k))) };
      const props = Object.assign({}, own.props, spans(x)); if (!x.ref) props.html = ''; if (x.cs === 1) delete props.colspan; if (x.rs === 1) delete props.rowspan;
      return Object.assign({}, own, { props });
    });
    if (row.ref) return Object.assign({}, row.ref, { cells });
    const from = byId.get(row.from); const props = Object.assign({}, from && from.props); delete props.header;
    return { kind: 'row', path: row.id, props, cells };
  });
  b.ops = (b.ops || []).concat([r.record]);
  return b;
}
const plan = async (before, after) => { const calls = []; await EN.planDocxBlocks('t.docx', [before], [after], async argv => { calls.push(argv.slice(0, 1).concat(argv.slice(2))); return { path: argv[2] + '/x[1]' }; }); return calls; };
const clone = x => structuredClone(x);

// ----- the model -----
test('the grid model places cells like a browser, and the commands keep it rectangular', () => {
  const m = modelOf(table([[{ html: 'a', rowspan: '2' }, 'b', 'c'], ['e', 'f'], ['g', 'h', 'i']]));
  assert.deepEqual(shape(m), ['0:1x2 1:1x1 2:1x1', '1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1']);
  const ins = EN.tableOps.insertRow(m, P(1, 2), true, newId); // below row 1: inside the merge in column 0
  assert.deepEqual(shape(ins.m), ['0:1x3 1:1x1 2:1x1', '1:1x1 2:1x1', '1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1']);
  assert.equal(ins.at, 1);
  assert.deepEqual(shape(EN.tableOps.deleteRow(m, '/body/table[1]/row[1]').m), ['0:1x1 1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1'], 'the merged cell moves down');
  assert.deepEqual(shape(EN.tableOps.insertCol(m, P(1, 1), true, newId).m), ['0:1x2 1:1x1 2:1x1 3:1x1', '1:1x1 2:1x1 3:1x1', '0:1x1 1:1x1 2:1x1 3:1x1']);
  const wide = modelOf(table([[{ html: 'a', colspan: '2' }, 'c'], ['d', 'e', 'f']]));
  assert.deepEqual(shape(EN.tableOps.insertCol(wide, P(2, 1), true, newId).m), ['0:3x1 3:1x1', '0:1x1 1:1x1 2:1x1 3:1x1'], 'a merged cell across the new column grows');
  assert.deepEqual(shape(EN.tableOps.deleteCol(wide, 1).m), ['0:1x1 1:1x1', '0:1x1 1:1x1'], 'and shrinks when one of its columns goes');
  assert.deepEqual(shape(EN.tableOps.mergeRight(m, P(1, 2)).m), ['0:1x2 1:2x1', '1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1']);
  assert.deepEqual(shape(EN.tableOps.mergeDown(m, P(2, 1)).m), ['0:1x2 1:1x1 2:1x1', '1:1x2 2:1x1', '0:1x1 2:1x1']);
  assert.deepEqual(shape(EN.tableOps.split(m, P(1, 1), newId).m), ['0:1x1 1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1', '0:1x1 1:1x1 2:1x1']);
});

test('the ribbon refuses merges that would not be rectangular, as Word does', () => {
  const m = modelOf(table([[{ html: 'a', rowspan: '2' }, 'b'], ['d']]));
  assert.throws(() => EN.tableOps.mergeRight(m, P(1, 1)), /行数不同/, 'a covers two rows, b one');
  assert.throws(() => EN.tableOps.mergeDown(m, P(2, 1)), /下方没有/, 'd is in the last row');
  assert.throws(() => EN.tableOps.mergeRight(m, P(1, 2)), /右侧没有/, 'b is at the right edge');
  assert.throws(() => EN.tableOps.mergeDown(modelOf(table([['a', 'b'], [{ html: 'c', colspan: '2' }]])), P(1, 1)), /列数不同/, 'c is two columns wide');
  assert.throws(() => EN.tableOps.split(m, P(1, 2), newId), /没有合并/);
  assert.throws(() => EN.tableOps.deleteCol(modelOf(table([[{ html: 'a', colspan: '2' }], ['b']])), 0), /没有单元格/);
  assert.equal(EN.tableCommand(m, 'deleteRows', P(1, 1), false, newId).empty, true, 'deleting every row takes the table away');
  assert.equal(EN.joinCells('<b>A</b>', '<br>'), '<b>A</b>');
  assert.equal(EN.joinCells('A', 'B'), 'A<br>B');
});

// ----- the commands a save sends -----
test('an untouched table and contents send nothing', async () => {
  const b = table([[{ html: 'a', rowspan: '2' }, 'b'], ['c']], { borders: 'all' });
  assert.deepEqual(await plan(b, clone(b)), []);
});

test('a merge sends one colspan, before any cell is looked up by position', async () => {
  const b = table([['A', 'B', 'C']], { borders: 'all' }), a = edit(clone(b), 'mergeRight', P(1, 1));
  a.props.borders = 'none';
  assert.deepEqual(await plan(b, a), [['set', '/body/table[1]', '--prop', 'borders=none'], ['set', P(1, 1), '--prop', 'colspan=2']]);
  assert.equal(a.rows[0].cells[0].props.html, 'A<br>B');
});

test('rows go in and out where the engine puts them, merged cells following', async () => {
  const b = table([[{ html: 'a', rowspan: '2' }, 'b'], ['c'], ['d', 'e']]);
  assert.deepEqual(await plan(b, edit(clone(b), 'insertRow', P(1, 2), true)), [
    ['add', '/body/table[1]', '--type', 'row', '--index', '2'], ['set', '/body/table[1]/row[2]/cell[1]', '--prop', 'fill=none']]);
  assert.deepEqual(await plan(b, edit(clone(b), 'deleteRows', P(1, 1))), [
    ['remove', '/body/table[1]/row[2]'], ['remove', '/body/table[1]/row[1]']], 'deleting a merged cell deletes the rows it covers, as Word');
  assert.deepEqual(await plan(b, edit(clone(b), 'deleteRows', P(1, 2))), [['remove', '/body/table[1]/row[1]']], 'the merged cell moves into row 2 in the engine too');
});

test('a column next to a vertical merge is added in the right grid column', async () => {
  const b = table([[{ html: 'a', rowspan: '2' }, 'b'], ['c']]);
  const left = await plan(b, edit(clone(b), 'insertCol', P(1, 1), false));
  assert.deepEqual(left, [
    ['set', P(1, 1), '--prop', 'rowspan=1'],
    ['add', '/body/table[1]/row[1]', '--type', 'cell', '--index', '1'], ['add', '/body/table[1]/row[2]', '--type', 'cell', '--index', '1'],
    ['set', P(1, 2), '--prop', 'rowspan=2'],
    ['set', P(2, 1), '--prop', 'fill=none'], ['set', P(1, 1), '--prop', 'fill=none']], 'the merge is split for the moment, else row 2 would get its cell after it');
  const right = await plan(b, edit(clone(b), 'insertCol', P(1, 1), true));
  assert.deepEqual(right.filter(c => c[0] !== 'set'), [['add', '/body/table[1]/row[1]', '--type', 'cell', '--index', '2'], ['add', '/body/table[1]/row[2]', '--type', 'cell', '--index', '1']]);
});

test('tables whose commands were lost are written anew instead of being patched cell by cell', () => {
  const b = table([['A', 'B', 'C']]), a = clone(b);
  a.rows[0].cells[0].props.colspan = '2'; a.rows[0].cells.splice(1, 1); // merged, but no command recorded
  assert.deepEqual(EN.tablesToRewrite([b], [a]), [a]);
  assert.deepEqual(EN.tablesToRewrite([b], [edit(clone(b), 'mergeRight', P(1, 1))]), []);
});

test('a new table gets its merges from the grid: colspans row by row, then rowspans at their visible places', () => {
  const rows = table([[{ html: 'a', colspan: '2', rowspan: '2' }, 'b'], ['c'], ['d', { html: 'e', rowspan: '1' }, 'f']]).rows;
  assert.deepEqual(EN.newTableCommands('/body/table[1]', rows, 'f.docx').map(c => c.slice(2)), [
    ['/body/table[1]/row[1]/cell[1]', '--prop', 'colspan=2'], ['/body/table[1]/row[1]/cell[1]', '--prop', 'rowspan=2']]);
});

test('tables render their lines as the file draws them, and contents as the file lists them', () => {
  assert.equal(EN.tableLook({ style: 'ThreeLineTable' }), 'three');
  assert.equal(EN.tableLook({ style: 'TableGrid', borders: 'none' }), 'none');
  assert.equal(EN.tableLook({}), 'none', 'a table without style or borders has no lines in Word');
  const three = (r, R) => EN.cellLines('three', r, { c: 0, cs: 1, rs: 1 }, R, 2);
  assert.match(three(0, 3), /border-top:1\.5px solid[^;]*;.*border-bottom:1px solid #1D1D1F/);
  assert.match(three(2, 3), /border-bottom:1\.5px solid/);
  assert.match(three(1, 3), /border-top:1px dashed/);
  const html = EN.tocHtml({ path: '/body/toc[1]', levels: '2', title: '目录', entries: [{ text: 'A', level: 1, page: '3' }, { text: 'B', level: 2 }] });
  assert.match(html, /^<nav data-toc="1" data-path="\/body\/toc\[1\]" data-levels="2" data-title="目录" contenteditable="false"/);
  assert.match(html, /padding-left:18px"><span style="flex:1">B<\/span>/);
  assert.match(html, />3<\/span>/);
  assert.match(EN.tocHtml({ levels: '3', title: '', entries: [] }), /添加标题后/);
});

// ----- the real engine: the commands must end with the editor's table in the file -----
const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const cli = join(root, 'src/Writer.Cli/bin/Debug/net10.0/writer.dll');
let server, url, dir;
before(async () => {
  if (!existsSync(cli)) return;
  dir = mkdtempSync(join(tmpdir(), 'writer-tables-'));
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
/** The body as the editor opens it: every block, so paths are numbered as in a real save. */
const bodyBlocks = async file => EN.blocksOf((await run(['get', file, '/body', '--depth', '4'])).children, file);
const shapeOf = b => b.rows.map(r => r.cells.map(c => `${c.props.html.replace(/<br>/g, '/')} ${c.props.colspan || 1}x${c.props.rowspan || 1}${c.props.fill ? ' #' + c.props.fill : ''}`).join(' | '));
/** Opens the file, runs ribbon commands on its n-th table as the editor does, saves as the editor does, and checks that the file
 * now holds exactly the table the editor shows, and that a second save of the same view sends nothing. */
async function scenario(name, make, steps) {
  const file = join(dir, name + '.docx'), n = steps.table || 1;
  await make(file);
  const opened = await bodyBlocks(file), edited = clone(opened), t = edited.filter(b => b.kind === 'table')[n - 1];
  for (const [op, cell, flag] of steps.ops) edit(t, op, cell(t), flag);
  const calls = [];
  await EN.planDocxBlocks(file, opened, edited, async argv => { calls.push(argv); return run(argv); });
  const reopened = await bodyBlocks(file), saved = reopened.filter(b => b.kind === 'table')[n - 1];
  assert.deepEqual(shapeOf(saved), shapeOf(t), name + ': ' + JSON.stringify(calls.map(c => c.slice(2).join(' '))));
  assert.equal(await EN.planDocxBlocks(file, reopened, clone(reopened), async argv => { throw new Error('second save sent ' + argv.join(' ')); }), 0);
  return calls;
}
const grid = rows => async file => { await run(['create', file]); await run(['add', file, '/body', '--type', 'table', '--prop', 'data=' + JSON.stringify(rows)]); };
const withMerges = (rows, merges) => async file => { await grid(rows)(file); for (const [path, prop] of merges) await run(['set', file, path, '--prop', prop]); };
const at = (r, c) => b => b.rows[r - 1].cells[c - 1].path;
const skip = () => !existsSync(cli) && 'build the CLI first: dotnet build Writer.slnx';

test('engine: rows in and out of vertical merges', { skip: skip() }, async () => {
  const make = withMerges([['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']], [[P(1, 1), 'rowspan=3']]);
  await scenario('row-inside', make, { ops: [['insertRow', at(1, 2), true], ['insertRow', at(1, 2), false], ['insertRow', at(4, 1), true]] });
  await scenario('row-delete-top', make, { ops: [['deleteRows', at(1, 2)]] });
  const four = withMerges([['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I'], ['J', 'K', 'L']], [[P(1, 1), 'rowspan=3']]);
  await scenario('row-delete-merged', four, { ops: [['deleteRows', at(1, 1)]] });
  await scenario('row-below-merge', four, { ops: [['insertRow', at(1, 1), true], ['insertRow', at(4, 1), true]] });
});

test('engine: columns beside, inside and across merges', { skip: skip() }, async () => {
  const make = withMerges([['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']], [[P(1, 2), 'colspan=2'], [P(2, 1), 'rowspan=2']]);
  await scenario('col-left-of-vmerge', make, { ops: [['insertCol', at(2, 1), false]] });
  await scenario('col-right-of-vmerge', make, { ops: [['insertCol', at(2, 1), true]] });
  await scenario('col-inside-colspan', make, { ops: [['insertCol', at(2, 2), true]] });
  await scenario('col-delete-under-colspan', make, { ops: [['deleteCols', at(2, 3)]] });
  await scenario('col-delete-vmerge', make, { ops: [['deleteCols', at(2, 1)]] });
});

test('engine: merge, split and a burst of commands in one save', { skip: skip() }, async () => {
  const make = grid([['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']]);
  await scenario('merge-right-down', make, { ops: [['mergeRight', at(1, 1)], ['mergeRight', at(2, 1)], ['mergeDown', at(1, 1)]] });
  await scenario('merge-then-split', make, { ops: [['mergeRight', at(1, 2)], ['mergeRight', at(2, 2)], ['mergeDown', at(1, 2)], ['split', at(1, 2)]] });
  await scenario('everything', make, { ops: [['mergeDown', at(2, 1)], ['insertCol', at(1, 2), false], ['insertRow', at(2, 2), true], ['mergeRight', at(1, 2)], ['deleteCols', at(1, 3)], ['deleteRows', at(3, 1)]] });
});

test('engine: the merged financial table from tables.docx takes every ribbon command', { skip: skip() }, async () => {
  const make = async file => copyFileSync(join(root, 'tests/Writer.Tests/Fixtures/docx/tables.docx'), file);
  const t2 = { table: 2 };
  await scenario('fixture-row', make, Object.assign({ ops: [['insertRow', at(3, 1), true], ['deleteRows', at(1, 1)]] }, t2));
  await scenario('fixture-col', make, Object.assign({ ops: [['insertCol', at(2, 2), true], ['deleteCols', at(1, 4)]] }, t2));
  await scenario('fixture-merge', make, Object.assign({ ops: [['mergeDown', at(3, 4)], ['split', at(1, 3)]] }, t2));
});
