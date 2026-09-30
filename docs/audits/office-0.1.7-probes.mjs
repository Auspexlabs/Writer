// Read-only capability probes against the actual editor/formula implementations.
// Run from any directory: node docs/audits/office-0.1.7-probes.mjs
// These are Node logic probes, not Microsoft Office or desktop UI acceptance tests.
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import vm from 'node:vm';
import * as E from '../../ui/sheet-engine.js';
import * as K from '../../ui/office-io.js';
import { sheetPrint } from '../../ui/sheet-print.js';

const root = new URL('../../', import.meta.url);
const output = {
  sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim(),
  workingTreeModified: !!execFileSync('git', ['status', '--porcelain'], { cwd: root, encoding: 'utf8' }).trim(), generatedAt: new Date().toISOString(),
  runtime: process.version,
  scope: 'Node logic probes; no Microsoft Office or desktop GUI comparison',
  registeredFunctions: E.FUNCS.length,
  formulaSupport: {}, formulaResults: [], validation: [], wordPagination: [],
};
const doc = { active: 0, sheets: [1, 2, 3].map(n => ({ name: 'Sheet' + n, cells: { A1: { v: String(n) } } })) };
for (const name of ['LET', 'LAMBDA', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'FILTER', 'SORT', 'UNIQUE', 'HSTACK', 'VSTACK', 'TAKE', 'DROP', 'TOCOL', 'TOROW', 'WRAPROWS', 'CHOOSECOLS', 'TEXTSPLIT', 'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE', 'LINEST', 'LOGEST', 'T.TEST', 'CHISQ.TEST', 'FORECAST.ETS', 'PRICE', 'YIELD', 'GETPIVOTDATA', 'CUBEVALUE', 'STOCKHISTORY', 'IMAGE', 'PY']) output.formulaSupport[name] = E.FUNCS.includes(name);
for (const formula of ['=SUM(A1:A10)', '=SUM(A1:A1048576)', '=LET(x,2,x+1)', '=LAMBDA(x,x+1)(2)', '=MAP({1,2},LAMBDA(x,x*2))', '=SUM(Sheet1:Sheet3!A1)', "='[Other.xlsx]Sheet1'!A1", '=IF(TRUE,B1,9)', '=COUNTBLANK(A:A)', '=ROWS(A:A)']) output.formulaResults.push({ formula, result: new E.Calc(doc).evaluate(formula) });
output.lastCellAddress = E.parseA('XFD1048576');

function component(file) {
  const html = readFileSync(new URL('../../ui/' + file, import.meta.url), 'utf8');
  const script = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)?.[1];
  if (!script) throw new Error('Editor script not found: ' + file);
  const ctx = {
    React: { createRef: () => ({ current: null }), createElement: (...args) => args },
    DCLogic: class { setState(value, cb) { Object.assign(this.state, typeof value === 'function' ? value(this.state) : value); cb?.(); } forceUpdate() {} },
    document: { documentElement: { dataset: {} } }, structuredClone, setTimeout, clearTimeout,
    $t: (s, values) => String(s).replace(/@@.*$/, '').replace(/\{(\w+)\}/g, (m, k) => values?.[k] ?? m), $lang: () => 'zh',
  };
  vm.runInNewContext(script + '\nglobalThis.Editor = Component;', ctx);
  return new ctx.Editor();
}
const sheet = component('SheetEditor.dc.html'); sheet.E = E;
for (const [rule, value] of [
  [{ type: 'whole', operator: 'between', value: '0', value2: '10' }, '11'],
  [{ type: 'time', operator: 'between', value: '0.375', value2: '0.7083333333' }, '08:00'],
  [{ type: 'custom', value: 'A1>0' }, '-3'],
]) {
  sheet.props = { doc: { active: 0, sheets: [{ name: 'Data', cells: {}, dv: [{ range: 'A1', ...rule }] }] } };
  output.validation.push({ rule, value, rejection: sheet.dvBad(0, 0, value) });
}
const word = component('WordEditor.dc.html'); word.state.view = 'page';
for (const columns of [1, 2, 3]) { word.props = { doc: { page: { cols: columns } } }; output.wordPagination.push({ columns, paginated: word.paged }); }
const pivotDoc = { sheets: [{ name: 'Data', cells: { A1: { v: 'Region' }, B1: { v: 'Product' }, C1: { v: 'Revenue' }, A2: { v: 'East' }, B2: { v: 'Book' }, C2: { v: '10' } } }] };
try { E.pivotOutput(pivotDoc, { sourceSheet: 'Data', sourceRange: 'A1:C2', target: 'A1', rows: [0, 1], cols: [], values: [{ field: 2, fn: 'sum' }] }); output.multiFieldPivot = 'accepted'; }
catch (error) { output.multiFieldPivot = error.message; }
const cfDoc = { title: 'Conditional formatting probe', sheets: [{ name: 'Data', cells: { A1: { v: '20' } }, cf: [{ range: 'A1', type: 'cellIs', operator: 'greaterThan', value: '10', fill: '#123456' }] }] };
sheet.props = { doc: { active: 0, ...cfDoc } };
const rendered = sheet.renderVals();
output.conditionalFormatting = {
  editorContainsRuleColor: JSON.stringify(rendered.cells).includes('#123456'),
  printedHtmlContainsRuleColor: sheetPrint(cfDoc, E).body.includes('#123456'),
};
const object = { x: 0, y: 0, w: 100, h: 100 };
output.powerPoint = {
  authoredEffectCount: K.FX.reduce((n, group) => n + group[2].length, 0),
  transitionNames: K.TRANS.map(t => t[0]),
  unsupportedEntranceFallsBackToFade: JSON.stringify(K.fxFrames({ fx: 'unsupported-test-effect', cls: 'entr' }, object, 900)) === JSON.stringify(K.fxFrames({ fx: 'fade' }, object, 900)),
  dissolveUsesFadeFrames: JSON.stringify(K.transitionFrames('dissolve')) === JSON.stringify(K.transitionFrames('fade')),
  morphFramesForRotationOnly: K.morphFrames({ ...object, rot: 0 }, { ...object, rot: 90 }),
};
console.log(JSON.stringify(output, null, 2));
