// node --test ui/tests — the formula engine. Expected values come from Microsoft's documented examples for each function.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as E from '../sheet-engine.js';

const calc = (cells, other = {}) => new E.Calc({ sheets: [
  { name: 'Data', cells }, { name: 'Other', cells: other }
] });
const value = (formula, cells = {}, other = {}) => calc({ ...cells, T1: { v: '=' + formula } }, other).value(0, 0, 19);

test('editor API remains available and formulas traverse cells and sheets', () => {
  for (const name of ['Calc', 'shiftF', 'adjF', 'upperF', 'usedRange', 'FUNC_GROUPS', 'FUNCS', 'fmt', 'fmtCode', 'REF_SRC']) assert.ok(E[name], name);
  const c = calc({ A1: { v: '2' }, A2: { v: '3' }, B1: { v: '=SUM(A1:A2)' }, B2: { v: '=B1+Other!A1' } }, { A1: { v: '7' } });
  assert.equal(c.value(0, 1, 1), 12);
  assert.equal(E.shiftF('SUM($A1,B$2,"A1")', 2, 1), 'SUM($A3,C$2,"A1")');
  assert.equal(E.adjF('A1+Other!B3', 'Data', 'Data', 'r', 0, 1), 'A2+Other!B3');
});

test('expanded criteria, lookup and array formulas evaluate through Calc', () => {
  const cells = { A1: { v: 'East' }, A2: { v: 'West' }, A3: { v: 'East' }, B1: { v: '10' }, B2: { v: '20' }, B3: { v: '30' } };
  assert.equal(value('SUMIFS(B1:B3,A1:A3,"East")', cells), 40);
  assert.deepEqual(value('SUMIFS(B1:B2,A1:A3,"East")', cells), { err: '#VALUE!' });
  assert.equal(value('COUNTIFS(A1:A3,"East",B1:B3,">15")', cells), 1);
  assert.equal(value('XLOOKUP("West",A1:A3,B1:B3)', cells), 20);
  assert.equal(value('LET(x,3,x*4)'), 12);
  assert.equal(value('SUM(SEQUENCE(2,2))'), 10);
  assert.equal(value('INDEX(SORT(B1:B3,1,-1),1)', cells), 30);
});

test('whole-column references include imported cells beyond the visible grid', () => {
  const cells = { A1: { v: '1' }, A100: { v: '7' } };
  assert.equal(value('SUM(A:A)', cells), 8);
  assert.deepEqual(E.usedRange({ cells }), { r1: 0, c1: 0, r2: 99, c2: 0 });
});

test('date and legacy type functions remain usable', () => {
  assert.equal(value('DATE(2024,1,1)'), 45292);
  assert.equal(value('YEAR(DATE(2024,1,1))'), 2024);
  assert.equal(value('MONTH(DATE(2024,1,1))'), 1);
  assert.equal(value('ISBLANK(A1)'), true);
  assert.equal(value('ISNUMBER(5)'), true);
  assert.equal(value('ISERROR(1/0)'), true);
  assert.equal(E.fmt(45292, { fmt: 'date' }), '2024/1/1');
});

test('errors are displayed instead of throwing or silently changing cells', () => {
  assert.deepEqual(value('1/0'), { err: '#DIV/0!' });
  assert.deepEqual(value('MISSING(1)'), { err: '#NAME?' });
  const c = calc({ A1: { v: '=A1+1' } });
  assert.deepEqual(c.value(0, 0, 0), { err: '#CIRC!' });
});

// ---- the specified library: every function present, grouped for the picker with a Chinese description ----
const SPEC = 'SUM SUMIF SUMIFS SUMPRODUCT PRODUCT ABS ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD POWER SQRT EXP LN LOG LOG10 CEILING CEILING.MATH FLOOR FLOOR.MATH MROUND SIGN RAND RANDBETWEEN PI FACT COMBIN GCD LCM QUOTIENT EVEN ODD SUBTOTAL AGGREGATE AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MAXA MIN MINA MAXIFS MINIFS MEDIAN MODE MODE.SNGL STDEV STDEV.S STDEV.P STDEVP VAR VAR.S VAR.P VARP LARGE SMALL RANK RANK.EQ RANK.AVG PERCENTILE PERCENTILE.INC PERCENTILE.EXC QUARTILE QUARTILE.INC CORREL PEARSON SLOPE INTERCEPT FORECAST FORECAST.LINEAR TREND GROWTH FREQUENCY GEOMEAN HARMEAN NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV LEN LENB LEFT LEFTB RIGHT RIGHTB MID MIDB UPPER LOWER PROPER TRIM CLEAN CONCAT CONCATENATE TEXTJOIN TEXT VALUE NUMBERVALUE FIND FINDB SEARCH REPLACE SUBSTITUTE REPT EXACT CHAR CODE UNICHAR UNICODE T N FIXED DOLLAR RMB TEXTBEFORE TEXTAFTER TEXTSPLIT IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE SWITCH CHOOSE LET VLOOKUP HLOOKUP XLOOKUP LOOKUP INDEX MATCH XMATCH OFFSET INDIRECT ROW ROWS COLUMN COLUMNS ADDRESS AREAS TRANSPOSE UNIQUE FILTER SORT SORTBY SEQUENCE TODAY NOW DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY WEEKNUM ISOWEEKNUM DATEDIF DATEVALUE TIMEVALUE EDATE EOMONTH DAYS DAYS360 NETWORKDAYS NETWORKDAYS.INTL WORKDAY WORKDAY.INTL YEARFRAC ISBLANK ISERROR ISERR ISNA ISNUMBER ISTEXT ISNONTEXT ISLOGICAL ISEVEN ISODD ISFORMULA ISREF TYPE NA ERROR.TYPE CELL INFO PMT PV FV NPER RATE IPMT PPMT NPV IRR XNPV XIRR SLN DB DDB SYD EFFECT NOMINAL CUMIPMT CUMPRINC DEC2BIN BIN2DEC DEC2HEX HEX2DEC DEC2OCT OCT2DEC CONVERT DELTA'.split(' ');
test('the function library covers the specification (≥150 functions) and the picker groups describe every one in Chinese', () => {
  const have = new Set(E.FUNCS);
  assert.deepEqual(SPEC.filter(f => !have.has(f)), [], 'missing from FUNCS');
  assert.ok(E.FUNCS.length >= 150, `only ${E.FUNCS.length} functions`);
  const grouped = new Map(E.FUNC_GROUPS.flatMap(([, fs]) => fs));
  for (const f of SPEC) assert.match(grouped.get(f) || '', /[一-龥]/, f + ' needs a Chinese description in FUNC_GROUPS');
  for (const [g, fs] of E.FUNC_GROUPS) { assert.match(g, /[一-龥]/); for (const [f] of fs) assert.ok(have.has(f), g + ' lists unknown ' + f); }
});

// ---- table-driven semantics: one sheet of fixtures, [formula, expected] rows per group ----
const CELLS = {
  A1: { v: 'apple' }, A2: { v: 'Banana' }, A3: { v: 'apricot' }, A4: { v: '' }, A5: { v: 'x' },
  B1: { v: '10' }, B2: { v: '20' }, B3: { v: '30' }, B4: { v: '40' }, B5: { v: 'text' },
  C1: { v: '0.1' }, C2: { v: '0.5' }, C3: { v: '1' }, D1: { v: 'a' }, D2: { v: 'b' }, D3: { v: 'c' },
  F1: { v: '4.14' }, F2: { v: '4.19' }, F3: { v: '5.17' }, F4: { v: '5.77' }, F5: { v: '6.39' },
  G1: { v: 'red' }, G2: { v: 'orange' }, G3: { v: 'yellow' }, G4: { v: 'green' }, G5: { v: 'blue' },
  H1: { v: '25' }, H2: { v: '38' }, H3: { v: '40' }, H4: { v: '41' },
  I1: { v: '7' }, I2: { v: '3.5' }, I3: { v: '3.5' }, I4: { v: '1' }, I5: { v: '2' },
  J1: { v: '=1/0' }, J2: { v: '=SUM(1,2)' }, J3: { v: '2024-01-01', s: { fmt: 'date' } }, J4: { v: '=NA()' }, J5: { v: 'TRUE' }, J6: { v: '' }, J7: { v: '1899-12-30 12:30:00', s: { fmt: 'time', code: 'h:mm:ss' } }, J8: { v: '2024-03-05 12:00:00', s: { fmt: 'date', code: 'yyyy/m/d h:mm' } },
  K1: { v: '1345' }, K2: { v: '1301' }, K3: { v: '1368' }, K4: { v: '1322' }, K5: { v: '1310' }, K6: { v: '1370' }, K7: { v: '1318' }, K8: { v: '1350' }, K9: { v: '1303' }, K10: { v: '1299' },
  L1: { v: '3' }, L2: { v: '2' }, L3: { v: '4' }, L4: { v: '5' }, L5: { v: '6' }, M1: { v: '9' }, M2: { v: '7' }, M3: { v: '12' }, M4: { v: '15' }, M5: { v: '17' },
  N1: { v: '6' }, N2: { v: '7' }, N3: { v: '9' }, N4: { v: '15' }, N5: { v: '21' }, O1: { v: '20' }, O2: { v: '28' }, O3: { v: '31' }, O4: { v: '38' }, O5: { v: '40' },
  P1: { v: 'East' }, P2: { v: 'West' }, P3: { v: 'East' }, Q1: { v: '1' }, Q2: { v: '2' }, Q3: { v: '3' },
  R1: { v: '-70000' }, R2: { v: '12000' }, R3: { v: '15000' }, R4: { v: '18000' }, R5: { v: '21000' }, R6: { v: '26000' },
  S1: { v: '-10000' }, S2: { v: '2750' }, S3: { v: '4250' }, S4: { v: '3250' }, S5: { v: '2750' },
  U1: { v: '=DATE(2008,1,1)' }, U2: { v: '=DATE(2008,3,1)' }, U3: { v: '=DATE(2008,10,30)' }, U4: { v: '=DATE(2009,2,15)' }, U5: { v: '=DATE(2009,4,1)' },
  V1: { v: '=DATE(2012,11,22)' }, V2: { v: '=DATE(2012,12,4)' }, V3: { v: '=DATE(2013,1,21)' },
  W1: { v: '=DATE(2008,11,26)' }, W2: { v: '=DATE(2008,12,4)' }, W3: { v: '=DATE(2009,1,21)' },
  X1: { v: '-10000' }, X2: { v: '3000' }, X3: { v: '4200' }, X4: { v: '6800' }, Y1: { v: '8000' }, Y2: { v: '9200' }, Y3: { v: '10000' }, Y4: { v: '12000' }, Y5: { v: '14500' },
};
const DOC = { sheets: [{ name: 'Data', cells: CELLS }, { name: 'My Sheet', cells: { A1: { v: '99' } } }], names: { RATE: '=Data!C2', DATA: 'Data!B1:B4' } };
const at = f => { const c = new E.Calc({ ...DOC, sheets: [{ name: 'Data', cells: { ...CELLS, Z99: { v: '=' + f } } }, DOC.sheets[1]] }); const v = c.value(0, 98, 25); return E.isErr(v) ? v.err : v; };
const near = (got, exp) => typeof exp === 'number' && typeof got === 'number' ? Math.abs(got - exp) <= Math.max(5e-3, Math.abs(exp) * 1e-6) : got === exp;
const TABLE = {
  '数学': [
    ['ROUND(2.15,1)', 2.2], ['ROUND(-1.475,2)', -1.48], ['ROUND(21.5,-1)', 20], ['ROUND(626.3,-3)', 1000], ['ROUND(1.98,-1)', 0],
    ['ROUNDUP(3.2,0)', 4], ['ROUNDUP(-3.14159,1)', -3.2], ['ROUNDUP(31415.92654,-2)', 31500], ['ROUNDDOWN(-3.14159,1)', -3.1], ['ROUNDUP(76.9,0)', 77],
    ['INT(-8.9)', -9], ['TRUNC(-8.9)', -8], ['MOD(-3,2)', 1], ['MOD(3,-2)', -1], ['MOD(3,0)', '#DIV/0!'],
    ['CEILING(2.5,1)', 3], ['CEILING(-2.5,-2)', -4], ['CEILING(1.5,0.1)', 1.5], ['CEILING(0.234,0.01)', 0.24], ['CEILING(-2.5,2)', '#NUM!'],
    ['CEILING.MATH(-5.5,2,-1)', -6], ['CEILING.MATH(24.3,5)', 25], ['CEILING.MATH(-8.1,2)', -8], ['CEILING.MATH(6.7)', 7],
    ['FLOOR(3.7,2)', 2], ['FLOOR(-2.5,-2)', -2], ['FLOOR(0.234,0.01)', 0.23], ['FLOOR.MATH(-5.5,2,-1)', -4], ['FLOOR.MATH(-8.1,2)', -10], ['FLOOR(2.5,-2)', '#NUM!'],
    ['MROUND(10,3)', 9], ['MROUND(-10,-3)', -9], ['MROUND(1.3,0.2)', 1.4], ['MROUND(5,-2)', '#NUM!'],
    ['QUOTIENT(5,2)', 2], ['QUOTIENT(-10,3)', -3], ['EVEN(1.5)', 2], ['EVEN(-1)', -2], ['ODD(1.5)', 3], ['ODD(-2)', -3], ['ODD(-1)', -1], ['ODD(3)', 3],
    ['GCD(24,36)', 12], ['LCM(24,36)', 72], ['COMBIN(8,2)', 28], ['FACT(5)', 120], ['POWER(5,2)', 25], ['POWER(98.6,3.2)', 2401077.222], ['SQRT(-16)', '#NUM!'], ['LOG(8,2)', 3], ['LOG10(100)', 2], ['LN(EXP(2))', 2], ['SIGN(-3)', -1], ['ABS(-2)', 2], ['PI()', Math.PI],
    ['SUM(B1:B5)', 100], ['SUM("3",TRUE)', 4], ['SUM(B1:B4,"5")', 105], ['SUM(1,,2)', 3], ['PRODUCT(B1:B2)', 200],
    ['SUMIF(B1:B4,">15")', 90], ['SUMIF(A1:A3,"a*",B1:B3)', 40], ['SUMIF(A1:A5,"<>",B1:B5)', 60], ['SUMIFS(B1:B4,A1:A4,"a*",B1:B4,">15")', 30],
    ['SUMPRODUCT(L1:L5,M1:M5)', 3 * 9 + 2 * 7 + 4 * 12 + 5 * 15 + 6 * 17], ['SUMPRODUCT((P1:P3="East")*Q1:Q3)', 4], ['SUMPRODUCT(--(P1:P3="East"))', 2],
    ['SUBTOTAL(9,B1:B5)', 100], ['SUBTOTAL(1,B1:B4)', 25], ['SUBTOTAL(109,B1:B4)', 100], ['AGGREGATE(9,6,J1:J2,B1:B2)', 33], ['AGGREGATE(14,6,B1:B4,2)', 30],
  ],
  '统计': [
    ['AVERAGE(B1:B5)', 25], ['AVERAGEA(B1:B5)', 20], ['COUNT(B1:B5)', 4], ['COUNTA(A1:A5)', 4], ['COUNTBLANK(A1:A5)', 1], ['MAX(B1:B5)', 40], ['MIN(B1:B5)', 10], ['MAXA(J5:J6)', 1], ['MINA(B1:B5)', 0],
    ['COUNTIF(A1:A5,"a*")', 2], ['COUNTIF(A1:A5,"?????")', 1], ['COUNTIF(A1:A5,"*an*")', 1], ['COUNTIF(B1:B4,B2)', 1], ['COUNTIF(B1:B4,">="&B2)', 3], ['COUNTIF(A1:A5,"")', 1], ['COUNTIF(A1:A5,"<>")', 4],
    ['COUNTIFS(P1:P3,"East",Q1:Q3,">1")', 1], ['AVERAGEIF(B1:B4,">15")', 30], ['AVERAGEIFS(Q1:Q3,P1:P3,"East")', 2], ['MAXIFS(Q1:Q3,P1:P3,"East")', 3], ['MINIFS(Q1:Q3,P1:P3,"East")', 1],
    ['MEDIAN(1,2,3,4,5,6)', 3.5], ['MODE(5.6,4,4,3,2,4)', 4], ['MODE.SNGL(1,2,2)', 2], ['STDEV(K1:K10)', 27.46391572], ['STDEV.S(K1:K10)', 27.46391572], ['STDEV.P(K1:K10)', 26.05455814], ['STDEVP(K1:K10)', 26.05455814], ['VAR(K1:K10)', 754.2666667], ['VAR.S(K1:K10)', 754.2666667], ['VARP(K1:K10)', 678.84], ['VAR.P(K1:K10)', 678.84],
    ['LARGE(L1:L5,2)', 5], ['SMALL(L1:L5,2)', 3], ['RANK(3.5,I1:I5)', 2], ['RANK.EQ(1,I1:I5,1)', 1], ['RANK.AVG(3.5,I1:I5)', 2.5], ['RANK(9,I1:I5)', '#N/A'],
    ['PERCENTILE(Q1:Q3,0.5)', 2], ['PERCENTILE.INC(L1:L5,0.3)', 3.2], ['PERCENTILE.EXC(L1:L5,0.5)', 4], ['QUARTILE(L1:L5,1)', 3], ['QUARTILE.INC(L1:L5,3)', 5], ['QUARTILE.EXC(L1:L5,1)', 2.5],
    ['CORREL(L1:L5,M1:M5)', 0.997054486], ['PEARSON(L1:L5,M1:M5)', 0.997054486], ['SLOPE(M1:M5,L1:L5)', 2.6], ['INTERCEPT(M1:M5,L1:L5)', 1.6], ['FORECAST(30,N1:N5,O1:O5)', 10.60725309], ['FORECAST.LINEAR(30,N1:N5,O1:O5)', 10.60725309],
    ['INDEX(TREND(M1:M5,L1:L5,{7}),1)', 19.8], ['INDEX(GROWTH({2,4,8},{1,2,3},{4}),1)', 16], ['INDEX(FREQUENCY(L1:L5,{3,5}),2)', 2],
    ['GEOMEAN(4,5,8,7,11,4,3)', 5.476986969], ['HARMEAN(4,5,8,7,11,4,3)', 5.028375962],
    ['NORM.DIST(42,40,1.5,TRUE)', 0.9087888], ['NORM.DIST(42,40,1.5,FALSE)', 0.10934005], ['NORM.INV(0.908789,40,1.5)', 42.000002], ['NORM.S.DIST(1.333333,TRUE)', 0.908788726], ['NORM.S.INV(0.908789)', 1.3333347],
  ],
  '文本': [
    ['LEFT("Sale Price",4)', 'Sale'], ['RIGHT("Sale Price",5)', 'Price'], ['MID("Fluid Flow",7,20)', 'Flow'], ['MID("Fluid Flow",20,5)', ''], ['LEFT("abc")', 'a'], ['RIGHT("abc",0)', ''],
    ['FIND("M","Miriam McGovern")', 1], ['FIND("m","Miriam McGovern")', 6], ['FIND("M","Miriam McGovern",3)', 8], ['FIND("x","abc")', '#VALUE!'],
    ['SEARCH("e","Statements",6)', 7], ['SEARCH("margin","Profit Margin")', 8], ['SEARCH("~?","a?b")', 2], ['SEARCH("a*c","xxabbc")', 3],
    ['SUBSTITUTE("Sales Data","Sales","Cost")', 'Cost Data'], ['SUBSTITUTE("Quarter 1, 2008","1","2",1)', 'Quarter 2, 2008'], ['SUBSTITUTE("Quarter 1, 2011","1","2",3)', 'Quarter 1, 2012'],
    ['REPLACE("abcdefghijk",6,5,"*")', 'abcde*k'], ['REPLACE("2009",3,2,"10")', '2010'], ['REPLACE("123456",1,3,"@")', '@456'],
    ['TEXT(1234.567,"$#,##0.00")', '$1,234.57'], ['TEXT(0.285,"0.0%")', '28.5%'], ['TEXT(12200000,"0.00E+00")', '1.22E+07'], ['TEXT(DATE(2012,3,14),"MM/DD/YY")', '03/14/12'], ['TEXT(DATE(2012,3,14),"dddd")', 'Wednesday'],
    ['TEXT(0.5,"h:mm AM/PM")', '12:00 PM'], ['TEXT(-1234,"#,##0;(#,##0)")', '(1,234)'], ['TEXT(1234567.89,"#,##0.00")', '1,234,567.89'], ['TEXT(0.25,"0%")', '25%'], ['TEXT(12,"0000")', '0012'], ['TEXT(45292,"yyyy-mm-dd")', '2024-01-01'], ['TEXT(TIME(13,5,9),"hh:mm:ss")', '13:05:09'], ['TEXT(1.5,"[h]:mm")', '36:00'], ['TEXT(0.75,"[mm]")', '1080'], ['TEXT(DATE(2024,1,1),"yyyy年m月d日")', '2024年1月1日'], ['TEXT(DATE(2024,1,1),"aaaa")', '星期一'], ['TEXT(1234.5,"0.0")', '1234.5'], ['TEXT(0.1234,"0.00%")', '12.34%'], ['TEXT(1234,"#,##0.00 ""元""")', '1,234.00 元'], ['TEXT(DATE(2024,3,5),"mmm d, yyyy")', 'Mar 5, 2024'], ['TEXT(DATE(2024,3,5),"mmmm")', 'March'], ['TEXT(1,"General")', '1'], ['TEXT(1234,"")', '1234'],
    ['VALUE("$1,000")', 1000], ['VALUE("16:48:00")-VALUE("12:00:00")', 0.2], ['VALUE("12%")', 0.12], ['VALUE("abc")', '#VALUE!'],
    ['PROPER("this is a TITLE")', 'This Is A Title'], ['PROPER("2-way street")', '2-Way Street'], ['PROPER("76BudGet")', '76Budget'],
    ['TRIM(" First Quarter   Earnings ")', 'First Quarter Earnings'], ['CONCAT("The"," ","sun")', 'The sun'], ['CONCATENATE("a",1)', 'a1'], ['TEXTJOIN(", ",TRUE,"a","","b")', 'a, b'], ['TEXTJOIN("-",FALSE,A1:A3)', 'apple-Banana-apricot'],
    ['EXACT("word","word")', true], ['EXACT("Word","word")', false], ['REPT("*-",3)', '*-*-*-'], ['CHAR(65)', 'A'], ['CODE("A")', 65], ['UNICHAR(20320)', '你'], ['UNICODE("你")', 20320],
    ['FIXED(1234.567,1)', '1,234.6'], ['FIXED(1234.567,-1)', '1,230'], ['FIXED(-1234.567,-1,TRUE)', '-1230'], ['FIXED(44.332)', '44.33'], ['FIXED(1234.5678,4)', '1,234.5678'],
    ['DOLLAR(1234.567,2)', '$1,234.57'], ['DOLLAR(-1234.567,-2)', '($1,200)'], ['DOLLAR(-0.123,4)', '($0.1230)'], ['DOLLAR(99.888)', '$99.89'], ['RMB(1234.567,2)', '¥1,234.57'], ['RMB(12345.678,1)', '¥12,345.7'],
    ['TEXTBEFORE("Red riding hood\'s, red hood","hood")', 'Red riding '], ['TEXTAFTER("Red riding hood\'s, red hood","hood")', "'s, red hood"], ['TEXTAFTER("a-b-c","-",2)', 'c'], ['TEXTBEFORE("Red riding hood\'s, red hood","HOOD",1,1)', 'Red riding '], ['TEXTBEFORE("abc","x",1,0,0,"nf")', 'nf'],
    ['INDEX(TEXTSPLIT("a,b;c,d",",",";"),2,1)', 'c'], ['NUMBERVALUE("2.500,27",",",".")', 2500.27], ['NUMBERVALUE("3.5%")', 0.035], ['LEN("Phoenix, AZ")', 11], ['LENB("中文a")', 5], ['LEFTB("中文a",3)', '中 '], ['RIGHTB("中文a",3)', '文a'], ['MIDB("中文a",3,2)', '文'], ['FINDB("a","中文a")', 5],
    ['T("Rainfall")', 'Rainfall'], ['T(19)', ''], ['N("7")', 0], ['N(TRUE)', 1], ['N(DATE(2011,4,17))', 40650], ['UPPER("abc")', 'ABC'], ['LOWER("ABC")', 'abc'], ['CLEAN("a"&CHAR(7)&"b")', 'ab'],
  ],
  '逻辑': [
    ['IF(B1>5,"big","small")', 'big'], ['IF(B1>50,"big")', false], ['IF(1,,3)', 0], ['IFS(B1>50,"a",B1>5,"b")', 'b'], ['IFS(B1>50,"a")', '#N/A'], ['IFERROR(1/0,"x")', 'x'], ['IFERROR(5,"x")', 5], ['IFNA(NA(),"x")', 'x'], ['IFNA(1,"x")', 1], ['IFNA(1/0,"x")', '#DIV/0!'],
    ['SWITCH(3,1,"a",2,"b","c")', 'c'], ['SWITCH(2,1,"a",2,"b","c")', 'b'], ['CHOOSE(2,"a","b")', 'b'], ['XOR(TRUE,TRUE)', false], ['XOR(TRUE,FALSE,FALSE)', true], ['AND(TRUE,1)', true], ['AND(TRUE,0)', false], ['OR(FALSE,B1>5)', true], ['NOT(TRUE)', false], ['LET(x,2,y,3,x*y)', 6], ['TRUE()', true], ['FALSE()', false],
  ],
  '查找引用': [
    ['VLOOKUP(0.7,C1:D3,2)', 'b'], ['VLOOKUP(0.05,C1:D3,2)', '#N/A'], ['VLOOKUP(2,C1:D3,2)', 'c'], ['VLOOKUP("BANANA",A1:B3,2,FALSE)', 20], ['VLOOKUP("b*",A1:B3,2,FALSE)', 20], ['VLOOKUP(0.5,C1:D3,3)', '#REF!'], ['VLOOKUP(0.5,C1:D3,0)', '#VALUE!'],
    ['HLOOKUP("Axles",A1:C1,1)', 'apple'], ['HLOOKUP("apple",A1:C1,1,FALSE)', 'apple'], ['LOOKUP(4.19,F1:F5,G1:G5)', 'orange'], ['LOOKUP(5,F1:F5,G1:G5)', 'orange'], ['LOOKUP(7.66,F1:F5,G1:G5)', 'blue'], ['LOOKUP(0,F1:F5,G1:G5)', '#N/A'],
    ['MATCH(39,H1:H4,1)', 2], ['MATCH(41,H1:H4,0)', 4], ['MATCH(40,H1:H4,-1)', '#N/A'], ['MATCH(2,{1,2,2,2,3},1)', 4], ['MATCH(3,{9,7,5,3,1},-1)', 4], ['VLOOKUP(5,{1,"a";3,"b";"x","c";7,"d"},2)', 'b'], ['MATCH("ban*",A1:A3,0)', 2],
    ['XLOOKUP("West",P1:P3,Q1:Q3)', 2], ['XLOOKUP("North",P1:P3,Q1:Q3,"none")', 'none'], ['XLOOKUP(39,H1:H4,H1:H4,,-1)', 38], ['XLOOKUP(39,H1:H4,H1:H4,,1)', 40], ['XLOOKUP("East",P1:P3,Q1:Q3,,0,-1)', 3], ['XLOOKUP("b*",A1:A3,B1:B3,,2)', 20], ['XMATCH(39,H1:H4,-1)', 2], ['XMATCH("East",P1:P3,0,-1)', 3],
    ['INDEX(A1:B3,2,2)', 20], ['INDEX(B1:B3,3)', 30], ['SUM(INDEX(A1:B3,0,2))', 60], ['INDEX(A1:B3,4,1)', '#REF!'], ['INDEX(A1:B3,MATCH("apricot",A1:A3,0),2)', 30],
    ['SUM(OFFSET(B1,1,0,2,1))', 50], ['OFFSET(B1,-1,0)', '#REF!'], ['INDIRECT("B"&2)', 20], ["INDIRECT(\"'My Sheet'!A1\")", 99], ['INDIRECT("nope")', '#REF!'],
    ['ROW()', 99], ['COLUMN()', 26], ['ROW(B3)', 3], ['COLUMN(B3)', 2], ['ROWS(A1:B3)', 3], ['COLUMNS(A1:B3)', 2], ['SUM(ROW(A1:A3))', 6], ['AREAS(A1:B2)', 1],
    ['ADDRESS(2,3)', '$C$2'], ['ADDRESS(2,3,2)', 'C$2'], ['ADDRESS(2,3,2,FALSE)', 'R2C[3]'], ['ADDRESS(2,3,1,TRUE,"EXCEL SHEET")', "'EXCEL SHEET'!$C$2"], ['ADDRESS(2,3,4)', 'C2'], ['ADDRESS(2,3,1,FALSE)', 'R2C3'],
    ['INDEX(TRANSPOSE(A1:B2),2,1)', 10], ['ROWS(UNIQUE(P1:P3))', 2], ['INDEX(UNIQUE(P1:P3),2)', 'West'], ['SUM(FILTER(Q1:Q3,P1:P3="East"))', 4], ['FILTER(Q1:Q3,P1:P3="X","none")', 'none'], ['INDEX(SORT(A1:A3),1)', 'apple'], ['INDEX(SORT(A1:B3,2,-1),1,1)', 'apricot'], ['INDEX(SORTBY(A1:A3,B1:B3,-1),1)', 'apricot'], ['SUM(SEQUENCE(3,2,1,2))', 36], ['INDEX(SEQUENCE(2,2),2,2)', 4],
  ],
  '日期时间': [
    ['DATE(2008,1,35)', 39482], ['DATE(2008,14,2)', 39846], ['DATE(2008,-3,2)', 39327], ['DATE(1900,2,29)', 60], ['DATE(1900,3,1)', 61], ['DATE(1900,1,1)', 1], ['DATE(9999,12,31)', 2958465], ['DATE(24,1,1)', 8767],
    ['DATEVALUE("8/22/2011")', 40777], ['DATEVALUE("2011/02/23")', 40597], ['DATEVALUE("2024-01-01")', 45292], ['DATEVALUE("2024年1月1日")', 45292], ['DATEVALUE("22-AUG-2011")', 40777], ['DATEVALUE("Aug 22, 2011")', 40777], ['DATEVALUE("22 August 2011")', 40777], ['DATEVALUE("13/13/2011")', '#VALUE!'],
    ['TIMEVALUE("2:24 AM")', 0.1], ['TIMEVALUE("22-Aug-2011 6:35 AM")', 0.274305556], ['TIMEVALUE("18:00")', 0.75],
    ['TIME(12,0,0)', 0.5], ['TIME(16,48,10)', 0.7001157], ['HOUR(0.75)', 18], ['MINUTE(TIME(1,2,3))', 2], ['SECOND(TIME(1,2,3))', 3], ['YEAR(45292)', 2024], ['MONTH(45292)', 1], ['DAY(45292)', 1], ['YEAR(60)', 1900], ['MONTH(60)', 2], ['DAY(60)', 29], ['DAY(59)', 28], ['DAY(61)', 1],
    ['WEEKDAY(DATE(2008,2,14))', 5], ['WEEKDAY(DATE(2008,2,14),2)', 4], ['WEEKDAY(DATE(2008,2,14),3)', 3], ['WEEKDAY(DATE(2008,2,14),11)', 4], ['WEEKDAY(DATE(2008,2,14),15)', 7], ['WEEKDAY(DATE(2008,2,14),16)', 6], ['WEEKDAY(DATE(2008,2,14),17)', 5], ['WEEKDAY(1)', 1], ['WEEKDAY(DATE(2024,1,1))', 2],
    ['WEEKNUM(DATE(2012,3,9))', 10], ['WEEKNUM(DATE(2012,3,9),2)', 11], ['WEEKNUM(DATE(2012,3,9),21)', 10], ['ISOWEEKNUM(DATE(2012,3,9))', 10], ['ISOWEEKNUM(DATE(2021,1,1))', 53], ['ISOWEEKNUM(DATE(2024,12,30))', 1], ['WEEKNUM(DATE(2024,1,1))', 1], ['WEEKNUM(DATE(2023,1,1),2)', 1], ['WEEKNUM(DATE(2023,12,31))', 53],
    ['DATEDIF("1/1/2001","1/1/2003","Y")', 2], ['DATEDIF("6/1/2001","8/15/2002","D")', 440], ['DATEDIF("6/1/2001","8/15/2002","YD")', 75], ['DATEDIF("6/1/2001","8/15/2002","MD")', 14], ['DATEDIF("6/1/2001","8/15/2002","YM")', 2], ['DATEDIF("6/1/2001","8/15/2002","M")', 14], ['DATEDIF("6/1/2001","8/15/2002","Y")', 1], ['DATEDIF(DATE(2024,3,31),DATE(2024,4,30),"M")', 0], ['DATEDIF(DATE(2024,3,1),DATE(2024,2,1),"D")', '#NUM!'], ['DATEDIF(DATE(2020,1,15),DATE(2024,1,14),"Y")', 3], ['DATEDIF(DATE(2020,1,15),DATE(2024,1,15),"Y")', 4],
    ['EDATE(DATE(2011,1,15),1)', 40589], ['EDATE(DATE(2011,1,15),-1)', 40527], ['EDATE(DATE(2011,1,15),2)', 40617], ['EDATE(DATE(2024,1,31),1)', 45351], ['EDATE(DATE(2024,3,31),-1)', 45351], ['EOMONTH(DATE(2011,1,1),1)', 40602], ['EOMONTH(DATE(2011,1,1),-3)', 40482], ['EOMONTH(DATE(2024,2,10),0)', 45351], ['EOMONTH(DATE(1900,2,1),0)', 60],
    ['DAYS("3/15/11","2/1/11")', 42], ['DAYS(DATE(2021,12,31),DATE(2021,1,1))', 364], ['DAYS360("1/30/2011","2/1/2011")', 1], ['DAYS360("1/1/2011","12/31/2011")', 360], ['DAYS360("1/1/2011","2/1/2011")', 30], ['DAYS360(DATE(2011,1,31),DATE(2011,3,31))', 60], ['DAYS360(DATE(2011,1,31),DATE(2011,3,31),TRUE)', 60], ['DAYS360(DATE(2012,2,29),DATE(2012,3,31))', 30], ['DAYS360(DATE(2011,1,15),DATE(2011,2,28))', 43],
    ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))', 110], ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))', 109], ['NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),V1:V3)', 107], ['NETWORKDAYS(DATE(2013,3,1),DATE(2012,10,1))', -110],
    ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))', 22], ['NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))', -21], ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"2006/1/2","2006/1/16"})', 22], ['NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"2006/1/2","2006/1/16"})', 20], ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)', 6], ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),"1111111")', '#VALUE!'], ['NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),8)', '#NUM!'],
    ['WORKDAY(DATE(2008,10,1),151)', 39933], ['WORKDAY(DATE(2008,10,1),151,W1:W3)', 39938], ['WORKDAY(DATE(2024,1,5),1)', 45299], ['WORKDAY(DATE(2024,1,8),-1)', 45296], ['WORKDAY(DATE(2024,1,1),0)', 45292],
    ['WORKDAY.INTL(DATE(2012,1,1),30,0)', '#NUM!'], ['WORKDAY.INTL(DATE(2012,1,1),90,11)', 41013], ['WORKDAY.INTL(DATE(2012,1,1),30,17)', 40944], ['TEXT(WORKDAY.INTL(DATE(2012,1,1),30,17),"m/dd/yyyy")', '2/05/2012'],
    ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))', 0.58055556], ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)', 0.57650273], ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)', 0.58611111], ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)', 0.57534247], ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)', 0.58055556], ['YEARFRAC(DATE(2012,7,30),DATE(2012,1,1))', 0.58055556], ['YEARFRAC(DATE(2020,1,1),DATE(2022,1,1),1)', 1.99908592], ['YEARFRAC(DATE(2012,2,29),DATE(2013,2,28),0)', 1], ['YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),5)', '#NUM!'],
    ['J3+1', 45293], ['YEAR(J3)', 2024], ['SUM(J3)', 45292], ['HOUR(J7)', 12], ['J7*24', 12.5], ['TEXT(J7,"h:mm")', '12:30'], ['J8', 45356.5], ['DAY(J8)', 5],
  ],
  '信息': [
    ['ISBLANK(A4)', true], ['ISBLANK(A1)', false], ['ISNUMBER(B1)', true], ['ISNUMBER("1")', false], ['ISERROR(J1)', true], ['ISERROR(J2)', false],
    ['ISERR(J1)', true], ['ISERR(J4)', false], ['ISNA(J4)', true], ['ISNA(J1)', false], ['ISERROR(J4)', true], ['ISTEXT(A1)', true], ['ISTEXT(B1)', false], ['ISTEXT(J6)', false], ['ISNONTEXT(B1)', true], ['ISNONTEXT(J6)', true], ['ISNONTEXT(A1)', false], ['ISLOGICAL(J5)', true], ['ISLOGICAL(1)', false], ['ISEVEN(2.5)', true], ['ISEVEN(-1)', false], ['ISODD(3)', true], ['ISODD("x")', '#VALUE!'],
    ['ISFORMULA(J2)', true], ['ISFORMULA(B1)', false], ['ISFORMULA("x")', '#VALUE!'], ['ISREF(J1)', true], ['ISREF(A1:B2)', true], ['ISREF("a")', false], ['ISREF(1)', false],
    ['TYPE(1)', 1], ['TYPE("a")', 2], ['TYPE(TRUE)', 4], ['TYPE(J1)', 16], ['TYPE({1,2})', 64], ['TYPE(J6)', 1], ['ISNA(NA())', true],
    ['ERROR.TYPE(J4)', 7], ['ERROR.TYPE(J1)', 2], ['ERROR.TYPE(#VALUE!)', 3], ['ERROR.TYPE(#REF!)', 4], ['ERROR.TYPE(#NAME?)', 5], ['ERROR.TYPE(#NUM!)', 6], ['ERROR.TYPE(#NULL!)', 1], ['ERROR.TYPE(1)', '#N/A'],
    ['CELL("address",B3)', '$B$3'], ['CELL("row",B3)', 3], ['CELL("col",B3)', 2], ['CELL("contents",B1)', 10], ['CELL("type",J6)', 'b'], ['CELL("type",A1)', 'l'], ['CELL("type",B1)', 'v'], ['CELL("address")', '$Z$99'], ['CELL("nope",A1)', '#VALUE!'], ['INFO("numfile")', 2], ['INFO("recalc")', 'Automatic'], ['INFO("nope")', '#VALUE!'],
  ],
  '财务': [
    ['PMT(0.08/12,10,10000)', -1037.03], ['PMT(0.08/12,10,10000,0,1)', -1030.16], ['PMT(0.06/12,18*12,0,50000)', -129.08], ['PMT(0,10,1000)', -100],
    ['PV(0.08/12,20*12,500)', -59777.15], ['FV(0.06/12,10,-200,-500,1)', 2581.40], ['FV(0.12/12,12,-1000)', 12682.50], ['FV(0.11/12,35,-2000,,1)', 82846.25], ['FV(0.06/12,12,-100,-1000,1)', 2301.40], ['FV(0,12,-100)', 1200],
    ['NPER(0.12/12,-100,-1000,10000,1)', 59.6738657], ['NPER(0.12/12,-100,-1000,10000)', 60.0821229], ['NPER(0.12/12,-100,-1000)', -9.57859404], ['NPER(0,-100,1000)', 10],
    ['RATE(4*12,-200,8000)', 0.007701472], ['RATE(4*12,-200,8000)*12', 0.09241767], ['IPMT(0.1/12,1,3*12,8000)', -66.67], ['IPMT(0.1,3,3,8000)', -292.45], ['IPMT(0.1/12,1,36,8000,0,1)', 0], ['IPMT(0.1,0,3,8000)', '#NUM!'],
    ['PPMT(0.1/12,1,2*12,2000)', -75.62], ['PPMT(0.08,10,10,200000)', -27598.05],
    ['NPV(0.1,-10000,3000,4200,6800)', 1188.44], ['NPV(0.1,X1:X4)', 1188.44], ['NPV(0.08,Y1:Y5)+(-40000)', 1922.06],
    ['IRR(R1:R5)', -0.02124485], ['IRR(R1:R6)', 0.0866309], ['IRR(R1:R3,-0.1)', -0.44350694], ['IRR(R2:R6)', '#NUM!'],
    ['XNPV(0.09,S1:S5,U1:U5)', 2086.647602], ['XIRR(S1:S5,U1:U5)', 0.373362535], ['XIRR(S1:S5,U1:U5,0.1)', 0.373362535],
    ['SLN(30000,7500,10)', 2250], ['DB(1000000,100000,6,1,7)', 186083.33], ['DB(1000000,100000,6,2,7)', 259639.42], ['DB(1000000,100000,6,3,7)', 176814.44], ['DB(1000000,100000,6,7,7)', 15845.10], ['DB(1000000,100000,6,8,7)', '#NUM!'],
    ['DDB(2400,300,10*365,1)', 1.32], ['DDB(2400,300,10*12,1)', 40], ['DDB(2400,300,10,1)', 480], ['DDB(2400,300,10,2,1.5)', 306], ['DDB(2400,300,10,10)', 22.12],
    ['SYD(30000,7500,10,1)', 4090.91], ['SYD(30000,7500,10,10)', 409.09], ['EFFECT(0.0525,4)', 0.0535427], ['NOMINAL(0.053543,4)', 0.05250032], ['EFFECT(0,4)', '#NUM!'],
    ['CUMIPMT(0.09/12,30*12,125000,13,24,0)', -11135.23], ['CUMIPMT(0.09/12,30*12,125000,1,1,0)', -937.50], ['CUMPRINC(0.09/12,30*12,125000,13,24,0)', -934.1071234], ['CUMPRINC(0.09/12,30*12,125000,1,1,0)', -68.27827118], ['CUMIPMT(0.09/12,30*12,125000,0,1,0)', '#NUM!'],
  ],
  '工程': [
    ['DEC2BIN(9)', '1001'], ['DEC2BIN(9,4)', '1001'], ['DEC2BIN(-100)', '1110011100'], ['DEC2BIN(9,3)', '#NUM!'], ['DEC2BIN(512)', '#NUM!'], ['BIN2DEC(1100100)', 100], ['BIN2DEC(1111111111)', -1], ['BIN2DEC("102")', '#NUM!'],
    ['DEC2HEX(100,4)', '0064'], ['DEC2HEX(-54)', 'FFFFFFFFCA'], ['DEC2HEX(28)', '1C'], ['HEX2DEC("A5")', 165], ['HEX2DEC("FFFFFFFF5B")', -165], ['HEX2DEC("3DA408B9")', 1034160313], ['HEX2DEC("G")', '#NUM!'],
    ['DEC2OCT(58,3)', '072'], ['DEC2OCT(-100)', '7777777634'], ['OCT2DEC(54)', 44], ['OCT2DEC(7777777533)', -165],
    ['CONVERT(1,"lbm","kg")', 0.45359237], ['CONVERT(68,"F","C")', 20], ['CONVERT(2.5,"ft","sec")', '#N/A'], ['CONVERT(CONVERT(100,"ft","m"),"ft","m")', 9.290304], ['CONVERT(1,"km","mi")', 0.621371192], ['CONVERT(100,"C","K")', 373.15], ['CONVERT(1,"kibyte","bit")', 8192], ['CONVERT(1,"hr","min")', 60], ['CONVERT(1,"gal","l")', 3.785411784], ['CONVERT(1,"xyz","m")', '#N/A'],
    ['DELTA(5,4)', 0], ['DELTA(5,5)', 1], ['DELTA(0.5,0)', 0], ['DELTA(0.5)', 0], ['DELTA(0)', 1],
  ],
  '运算符、引用、数组常量': [
    ['1=1', true], ['"a"<"b"', true], ['"apple"="APPLE"', true], ['-2^2', 4], ['2^3^2', 64], ['50%', 0.5], ['A4+1', 1], ['A4&"x"', 'x'], ['A4=0', true], ['A4=""', true], ['B1&B2', '1020'], ['"x"+1', '#VALUE!'], ['1+TRUE', 2], ['0.1+0.2=0.3', true], ['1.1*3', 3.3], ['J6', 0], ['J6&""', ''], ['""', ''],
    ['SUM(B:B)', 100], ['COUNTA(1:1)', 23], ['SUM(Data!B1:B2)', 30], ["SUM('My Sheet'!A1)", 99], ['SUM(DATA)', 100], ['RATE*2', 1], ['SUM(Nope!A1)', '#REF!'], ['NOPE', '#NAME?'], ['SUM(', '#VALUE!'],
    ['SUM({1,2;3,4})', 10], ['INDEX({1,2;3,4},2,1)', 3], ['SUMPRODUCT({1,2},{3,4})', 11], ['LARGE({3,5,3,5,4},3)', 4], ['VLOOKUP("b",{"a",1;"b",2},2,FALSE)', 2], ['{1,2', '#VALUE!'], ['{1,2;3}', '#VALUE!'], ['COUNT({1,"a",TRUE})', 1], ['MAX({1,#N/A})', '#N/A'], ['IFERROR(MAX({1,#N/A}),9)', 9], ['ROWS({1;2;3})', 3], ['LOOKUP(4.19,{4.14,4.19,5.17,5.77,6.39},{"red","orange","yellow","green","blue"})', 'orange'],
  ],
};
let n = 0;
for (const [group, rows] of Object.entries(TABLE)) for (const [f, exp] of rows) { n++; test(`${group}: =${f} → ${JSON.stringify(exp)}`, () => { const got = at(f); assert.ok(near(got, exp), `got ${JSON.stringify(got)}`); }); }
test('the table holds at least 300 cases', () => assert.ok(n >= 300, String(n)));

test('whole-column and whole-row references shift and adjust like cells, sheet prefixes and strings untouched', () => {
  assert.equal(E.shiftF('SUM(A:A,$B:$B,3:3,Sheet2!C:D,"A:A")+A1', 2, 1), 'SUM(B:B,$B:$B,5:5,Sheet2!D:E,"A:A")+B3');
  assert.equal(E.adjF('SUM(A:C)+B1+D:D', 'S', 'S', 'c', 1, 1), 'SUM(A:D)+C1+E:E');
  assert.equal(E.adjF('SUM(A:C)+B1+B:B+D:D', 'S', 'S', 'c', 1, -1), 'SUM(A:B)+#REF!+#REF!+C:C');
  assert.equal(E.adjF('SUM(2:3)+A5', 'S', 'S', 'r', 0, 2), 'SUM(4:5)+A7');
  assert.equal(E.adjF('SUM(2:3)+A5', 'S', 'Other', 'r', 0, 2), 'SUM(2:3)+A5', 'another sheet is left alone');
  assert.equal(E.shiftF('A:A', 0, -1), '#REF!');
});

test('typed formulas keep sheet names as written and sheets resolve case-insensitively (cross-sheet refs gave #REF!)', () => {
  assert.equal(E.upperF("sum(Other!a1, 'My Sheet'!b2, \"Keep\", x)"), "SUM(Other!A1, 'My Sheet'!B2, \"Keep\", X)");
  const c = new E.Calc({ sheets: [{ name: 'S', cells: { A1: { v: '=' + E.upperF("sum(Other!a1,'My Sheet'!b2)") }, A2: { v: '=OTHER!A1' } } }, { name: 'Other', cells: { A1: { v: '5' } } }, { name: 'My Sheet', cells: { B2: { v: '7' } } }] });
  assert.equal(c.value(0, 0, 0), 12); assert.equal(c.value(0, 1, 0), 5);
  assert.equal(E.adjF('other!A1+A1', 'S', 'Other', 'r', 0, 1), 'other!A2+A1', 'row insert on Other reaches other!A1');
});

test('named ranges resolve through the names hook, including named formulas', () => {
  const c = new E.Calc({ sheets: [{ name: 'S', cells: { A1: { v: '5' }, A2: { v: '7' }, B1: { v: '=SUM(Nums)' }, B2: { v: '=Twice' }, B3: { v: '=SUM(nums)' } } }], names: { NUMS: 'S!A1:A2', TWICE: '=A1*2' } });
  assert.equal(c.value(0, 0, 1), 12); assert.equal(c.value(0, 1, 1), 10); assert.equal(c.value(0, 2, 1), 12, 'names are case-insensitive');
});

test('number formats: codes drive display, the editor kinds keep their shortcuts', () => {
  assert.equal(E.fmt(45292, { fmt: 'date', code: 'yyyy-mm-dd' }), '2024-01-01');
  assert.equal(E.fmt(0.75, { fmt: 'time', code: 'hh:mm' }), '18:00');
  assert.equal(E.fmt(1234.5, { fmt: 'custom', code: '[Red]0.0' }), '1234.5');
  assert.equal(E.fmt(-5, { fmt: 'custom', code: '0;(0);"zero"' }), '(5)');
  assert.equal(E.fmt(0, { fmt: 'custom', code: '0;(0);"zero"' }), 'zero');
  assert.equal(E.fmt(1234.567, { fmt: 'money', dec: 1 }), '¥1,234.6');
  assert.equal(E.fmt(0.256, { fmt: 'pct', dec: 1 }), '25.6%');
  assert.equal(E.fmt(1e-7, {}), '1e-7');
  assert.equal(E.fmtCode(0.5, 'h:mm:ss AM/PM'), '12:00:00 PM');
});

test('performance: a 200×20 grid of chained formulas evaluates in well under a second', () => {
  const cells = {}; for (let r = 0; r < 200; r++) for (let c = 0; c < 20; c++) cells[E.A(r, c)] = { v: r === 0 ? String(c + 1) : c === 0 ? `=${E.A(r - 1, 0)}+1` : `=SUM(${E.A(r - 1, c - 1)}:${E.A(r, c - 1)})+${E.A(r, c - 1)}*0` };
  cells.T200 = { v: '=SUM(A:S)+SUMIF(A:A,">100")+COUNTIF(B:B,"<>")' };
  const t0 = performance.now(); const c = new E.Calc({ sheets: [{ name: 'S', cells }] }); let bad = 0;
  for (let r = 0; r < 200; r++) for (let cc = 0; cc < 20; cc++) if (E.isErr(c.value(0, r, cc))) bad++;
  const v = c.value(0, 199, 19), ms = performance.now() - t0;
  assert.equal(bad, 0); assert.equal(typeof v, 'number'); assert.ok(ms < 1000, `${ms.toFixed(0)} ms`);
});
