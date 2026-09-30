import {test} from 'node:test';
import assert from 'node:assert/strict';
import {statLinest,statTTest,statChiTest} from '../sheet-statistics.js';
import {Calc} from '../sheet-engine.js';
const near=(a,b,tol=1e-8)=>assert.ok(Math.abs(a-b)<tol*Math.max(1,Math.abs(b)),`${a} != ${b}`);
test('LINEST matches Microsoft multiple regression coefficient and diagnostic example',()=>{
 const x=[[2310,2,2,20],[2333,2,2,12],[2356,3,1.5,33],[2379,3,2,43],[2402,2,3,53],[2425,4,2,23],[2448,2,1.5,99],[2471,2,2,34],[2494,3,3,23],[2517,4,4,55],[2540,2,3,22]],y=[142000,144000,151000,150000,139000,169000,126000,142900,163000,169000,149000].map(v=>[v]),r=statLinest(y,x,true,true);
 near(r[0][0],-234.2371645);near(r[1][0],13.26801148);near(r[2][0],.996747993);near(r[3][0],459.7536742);assert.equal(r[3][1],6);near(r[4][0],1732393319);assert.equal(r[2][2].err,'#N/A');
});
test('regression handles rows, missing X, no intercept, exponential and dependent predictors',()=>{
 assert.deepEqual(statLinest([[1,9,5,7]],[[0,4,2,3]]).map(r=>r.map(v=>Math.round(v))),[[2,1]]);
 const r=statLinest([[2],[4],[6]],null,false,true);near(r[0][0],2);assert.equal(r[0][1],0);assert.equal(r[1][1].err,'#N/A');
 const ln=statLinest([[6],[18],[54]],null,true,false,true);near(ln[0][0],3);near(ln[0][1],2);
 const dependent=statLinest([[5],[9],[13],[17]],[[1,2],[2,4],[3,6],[4,8]],true,true);near(dependent[0][0]*2+dependent[0][1],4);near(dependent[0][2],1);assert.equal(dependent[3][1],2);
});
test('paired and Welch t-tests plus chi-square use distribution tails',()=>{
 near(statTTest([[3,4,5,8,9,1,2,4,5]],[[6,19,3,2,14,4,5,17,1]],2,1),.196016,5e-7);
 near(statTTest([[1,2,3]],[[2,3,4]],2,3),.28786413472669,1e-9);
 near(statChiTest([[10,20],[20,10]],[[15,15],[15,15]]),.009823274507519,1e-9);
});
test('formula calls return native dynamic arrays and preserve input errors',()=>{
 const calc=new Calc({sheets:[{name:'S',cells:{A1:{v:'=LINEST({1;9;5;7},{0;4;2;3},TRUE,TRUE)'}}}]});near(calc.value(0,0,0),2);near(calc.value(0,0,1),1);assert.equal(calc.arrays.get('0:0:0').length,5);
 near(calc.evaluate('T.TEST({1,2,3},{2,3,4},2,3)'),.28786413472669,1e-9);
 assert.equal(calc.evaluate('T.TEST({1,2},{2,3},3,1)').err,'#NUM!');assert.equal(calc.evaluate('LINEST({1;2},{1;2;3})').err,'#REF!');assert.equal(calc.evaluate('LOGEST({0;2})').err,'#NUM!');
});
test('distribution probabilities and inverses agree with analytic quantiles',()=>{
 const c=new Calc({sheets:[{name:'S',cells:{}}]}),v=f=>c.evaluate(f);
 near(v('T.INV.2T(0.05,6)'),2.44691185114497,1e-9);near(v('T.DIST.RT(1,1)'),.25);near(v('T.INV(0.25,1)'),-1);
 near(v('CHISQ.INV(0.95,2)'),5.991464547108,1e-9);near(v('CHISQ.DIST.RT(2,2)'),Math.exp(-1));near(v('F.DIST(1,5,5,TRUE)'),.5);near(v('F.INV(0.5,5,5)'),1);
});
test('bond prices match the documented example and yield inverts every day-count basis',()=>{
 const c=new Calc({sheets:[{name:'S',cells:{}}]}),v=f=>c.evaluate(f);
 assert.equal(v('PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)').toFixed(2),'94.63');
 for(let b=0;b<=4;b++)for(const f of [1,2,4]){const args=`DATE(2024,2,29),DATE(2028,8,31),0.04`,p=v(`PRICE(${args},0.06,100,${f},${b})`);near(v(`YIELD(${args},${p},100,${f},${b})`),.06,1e-9);}
 const args='DATE(2024,5,1),DATE(2024,8,31),0.03';const p=v(`PRICE(${args},0.05,100,2,1)`);near(v(`YIELD(${args},${p},100,2,1)`),.05,1e-9);
 assert.equal(v('PRICE(DATE(2024,1,1),DATE(2025,1,1),0.05,-0.01,100,2)').err,'#NUM!');
});
