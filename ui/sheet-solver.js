import {Calc,parseA,A} from './sheet-engine.js';
const EPS=1e-9;
// Two-phase simplex. Variables are nonnegative; callers split unrestricted variables.
export function linearProgram(cost,constraints,{limit=10000}={}) {
 const n=cost.length,rows=constraints.map(c=>({a:c.a.slice(),b:c.b,op:c.op}));
 if(!n||cost.some(x=>!Number.isFinite(x))||rows.some(c=>c.a.length!==n||!Number.isFinite(c.b)||c.a.some(x=>!Number.isFinite(x))||!['<=','>=','='].includes(c.op)))throw Error('无效的线性模型');
 let width=n;const basic=[],art=new Set();
 for(const r of rows){if(r.b<0){r.a=r.a.map(x=>-x);r.b=-r.b;r.op=r.op==='='?'=':r.op==='<='?'>=':'<=';}if(r.op!=='='){r.slack=width++;}if(r.op==='<=')basic.push(r.slack);else{r.art=width++;art.add(r.art);basic.push(r.art);}}
 const table=rows.map(r=>{const a=Array(width+1).fill(0);r.a.forEach((v,i)=>a[i]=v);if(r.slack!=null)a[r.slack]=r.op==='<='?1:-1;if(r.art!=null)a[r.art]=1;a[width]=r.b;return a;});let z=[],iterations=0;
 const pivot=(r,c)=>{const row=table[r],v=row[c];for(let j=0;j<=width;j++)row[j]/=v;for(let i=0;i<table.length;i++)if(i!==r){const k=table[i][c];if(Math.abs(k)>EPS)for(let j=0;j<=width;j++)table[i][j]-=k*row[j];}const k=z[c];for(let j=0;j<=width;j++)z[j]-=k*row[j];basic[r]=c;};
 const objective=co=>{z=Array(width+1).fill(0);co.forEach((v,i)=>z[i]=-v);for(let r=0;r<table.length;r++){const k=co[basic[r]]||0;for(let j=0;j<=width;j++)z[j]+=k*table[r][j];}};
 const optimize=()=>{while(++iterations<=limit){let c=-1;for(let j=0;j<width;j++)if(z[j]<-EPS){c=j;break;}if(c<0)return 'optimal';let r=-1,ratio=Infinity;for(let i=0;i<table.length;i++)if(table[i][c]>EPS){const t=table[i][width]/table[i][c];if(t<ratio-EPS||Math.abs(t-ratio)<=EPS&&(r<0||basic[i]<basic[r])){r=i;ratio=t;}}if(r<0)return 'unbounded';pivot(r,c);}return 'limit';};
 if(art.size){objective(Array.from({length:width},(_,i)=>art.has(i)?-1:0));const status=optimize();if(status==='limit')return {status,iterations};if(z[width]<-EPS)return {status:'infeasible',iterations};
  for(let i=table.length-1;i>=0;i--)if(art.has(basic[i])){const c=table[i].findIndex((v,j)=>j<width&&!art.has(j)&&Math.abs(v)>EPS);if(c>=0)pivot(i,c);else{table.splice(i,1);basic.splice(i,1);}}
  for(const row of table)for(const j of art)row[j]=0;
 }
 objective(Array.from({length:width},(_,i)=>art.has(i)?0:cost[i]||0));const status=optimize();if(status!=='optimal')return {status,iterations};
 const values=Array(n).fill(0);basic.forEach((b,i)=>{if(b<n)values[b]=Math.abs(table[i][width])<EPS?0:table[i][width];});return {status,values,value:values.reduce((s,x,i)=>s+x*cost[i],0),iterations};
}
export function variableRefs(text){const refs=[];for(const raw of String(text).split(/[,;\s]+/).filter(Boolean)){const [a,b=a]=raw.replace(/\$/g,'').split(':').map(parseA);if(!a||!b||b.r<a.r||b.c<a.c)throw Error('可变单元格地址无效');if((b.r-a.r+1)*(b.c-a.c+1)>200)throw Error('一次最多求解 200 个可变单元格');for(let r=a.r;r<=b.r;r++)for(let c=a.c;c<=b.c;c++){const ref=A(r,c);if(!refs.includes(ref))refs.push(ref);}}if(!refs.length||refs.length>200)throw Error('请选择 1 至 200 个可变单元格');return refs;}
function nonlinearSolve(evaluate,start,constraints,spec) {
 const n=start.length;if(n>40)throw Error('非线性求解一次最多支持 40 个可变单元格');if(spec.integer)throw Error('整数约束请使用线性求解');
 let evaluations=0,best=null,x=start.slice(),scale=1;const project=xs=>xs.map(v=>spec.nonnegative===false?v:Math.max(0,v));
 const sample=(xs,weight)=>{xs=project(xs);evaluations++;try{const values=evaluate(xs),objective=spec.mode==='value'?(values[0]-Number(spec.value))**2:values[0]*(spec.mode==='min'?1:-1),errors=constraints.map((c,i)=>c.op==='='?Math.abs(values[i+1]):Math.max(0,values[i+1]*(c.op==='<='?1:-1))),violation=errors.reduce((s,e)=>s+e*e,0),feasible=errors.every(e=>e<=1e-6)&&(spec.mode!=='value'||Math.abs(values[0]-Number(spec.value))<=1e-6);if(feasible&&(!best||objective<best.objective))best={x:xs.slice(),value:values[0],objective};return {x:xs,score:objective/scale+weight*violation};}catch{return {x:xs,score:1e100};}};
 try{scale=Math.max(1,Math.abs(evaluate(project(x))[0]));}catch{}
 let converged=false;
 for(const weight of [1,100,1e4,1e6,1e8]) {
  converged=false;
  let simplex=[sample(x,weight),...x.map((v,i)=>sample(x.map((a,j)=>i===j?a+Math.max(.1,Math.abs(v)*.05):a),weight))];
  for(let iteration=0;iteration<300*n+300;iteration++) {
   simplex.sort((a,b)=>a.score-b.score);const low=simplex[0],high=simplex[n],spread=Math.max(...simplex.map(p=>Math.max(...p.x.map((v,i)=>Math.abs(v-low.x[i])))));
   if(spread<1e-8*Math.max(1,...low.x.map(Math.abs))){converged=true;break;}if(evaluations>=20000)break;
   const centroid=x.map((_,i)=>simplex.slice(0,n).reduce((s,p)=>s+p.x[i],0)/n),point=f=>sample(centroid.map((v,i)=>v+f*(v-high.x[i])),weight),reflected=point(1);
   if(reflected.score<low.score){const expanded=point(2);simplex[n]=expanded.score<reflected.score?expanded:reflected;}
   else if(reflected.score<simplex[n-1].score)simplex[n]=reflected;
   else{const contracted=point(reflected.score<high.score?.5:-.5);if(contracted.score<Math.min(reflected.score,high.score))simplex[n]=contracted;else simplex=simplex.map((p,i)=>i?sample(p.x.map((v,j)=>(v+low.x[j])/2),weight):p);}
  }
  simplex.sort((a,b)=>a.score-b.score);x=simplex[0].x;if(evaluations>=20000)break;
 }
 return best&&converged?{ok:true,status:'local',x:best.x,result:best.value,iterations:evaluations}:{ok:false,status:evaluations>=20000?'limit':'infeasible',iterations:evaluations};
}
export function solveWorkbook(doc,si,spec) {
 const sh=doc.sheets[si],refs=variableRefs(spec.variables),target=parseA(spec.target),integer=new Set(spec.integer?variableRefs(spec.integer):[]);
 if([...integer].some(r=>!refs.includes(r)))throw Error('整数约束只能指定可变单元格');
 if(!target||!sh||!String(sh.cells[A(target.r,target.c)]?.v||'').startsWith('=')||sh.cells[A(target.r,target.c)]?.literal)throw Error('目标单元格必须包含公式');
 if(refs.some(r=>sh.cells[r]?.spill||!sh.cells[r]?.literal&&String(sh.cells[r]?.v||'').startsWith('=')))throw Error('可变单元格不能包含公式或动态数组结果');
 const constraints=(spec.constraints||[]).map(c=>{if(!['<=','>=','='].includes(c.op)||!c.left||c.right==null)throw Error('约束条件无效');return c;}),sign=spec.mode==='min'?-1:1;
 let values=Array(refs.length).fill(0);const map=new Map(refs.map((r,i)=>[r,i])),cells=new Proxy(sh.cells,{get:(base,k)=>map.has(k)?{v:values[map.get(k)]}:base[k],ownKeys:base=>[...new Set([...Reflect.ownKeys(base),...refs])],getOwnPropertyDescriptor:(base,k)=>map.has(k)?{enumerable:true,configurable:true}:Object.getOwnPropertyDescriptor(base,k)});const work={...doc,sheets:doc.sheets.map((s,i)=>i===si?{...s,cells}:s)};
 const evaluate=xs=>{values=xs;const calc=new Calc(work),value=expression=>{const v=calc.evaluate(String(expression).replace(/^=/,''),si);if(typeof v!=='number'||!Number.isFinite(v))throw Error('求解公式或约束没有返回有限数值');return v;};return [value(spec.target),...constraints.map(c=>value(c.left)-value(c.right))];};
 if(spec.method==='nonlinear'){const r=nonlinearSolve(evaluate,refs.map(ref=>Number(sh.cells[ref]?.v)||0),constraints,spec);return r.ok?{...r,values:Object.fromEntries(refs.map((ref,i)=>[ref,r.x[i]])),x:undefined}:r;}
 const base=evaluate(values),coeff=Array.from({length:base.length},()=>[]);for(let j=0;j<refs.length;j++){const xs=Array(refs.length).fill(0);xs[j]=1;const result=evaluate(xs);result.forEach((v,i)=>coeff[i].push(v-base[i]));}
 // Validate the affine model at independent points before allowing a linear optimum to change cells.
 for(let t=0;t<3;t++){const xs=refs.map((_,j)=>(t+1)*(j%3+1)*.713+(t===2?-.5:0)),result=evaluate(xs);for(let i=0;i<base.length;i++){const expected=base[i]+coeff[i].reduce((s,a,j)=>s+a*xs[j],0);if(Math.abs(result[i]-expected)>1e-7*Math.max(1,Math.abs(expected),Math.abs(result[i])))throw Error('当前模型包含非线性公式，请使用非线性求解');}}
 const free=spec.nonnegative===false,N=refs.length,expand=a=>free?a.concat(a.map(x=>-x)):a.slice(),cost=expand(coeff[0].map(a=>a*sign));let rows=constraints.map((c,i)=>({a:expand(coeff[i+1]),op:c.op,b:-base[i+1]}));
 if(spec.mode==='value'){const desired=Number(spec.value);if(!Number.isFinite(desired))throw Error('目标值无效');rows.push({a:expand(coeff[0]),op:'=',b:desired-base[0]});cost.fill(0);}
 let best=null,count=0,iterations=0;const pending=[rows],maxNodes=spec.maxNodes||512;
 while(pending.length&&count++<maxNodes){const branch=pending.pop(),r=linearProgram(cost,branch);iterations+=r.iterations;if(r.status==='infeasible')continue;if(r.status!=='optimal')return {ok:false,status:r.status,iterations};if(best&&r.value<=best.score+EPS)continue;
  const xs=r.values.slice(0,N).map((v,i)=>v-(free?r.values[i+N]:0)),fraction=refs.findIndex((ref,i)=>integer.has(ref)&&Math.abs(xs[i]-Math.round(xs[i]))>1e-7);
  if(fraction<0){best={values:xs.map((v,i)=>integer.has(refs[i])?Math.round(v):v),score:r.value};continue;}
  const a=Array(N).fill(0);a[fraction]=1;pending.push(branch.concat({a:expand(a),op:'>=',b:Math.ceil(xs[fraction])}),branch.concat({a:expand(a),op:'<=',b:Math.floor(xs[fraction])}));
 }
 if(pending.length)return {ok:false,status:'limit',iterations};if(!best)return {ok:false,status:'infeasible',iterations};
 const checked=evaluate(best.values);if(constraints.some((c,i)=>c.op==='='?Math.abs(checked[i+1])>1e-6:c.op==='<='?checked[i+1]>1e-6:checked[i+1]<-1e-6))return {ok:false,status:'precision',iterations};
 return {ok:true,status:'optimal',values:Object.fromEntries(refs.map((r,i)=>[r,best.values[i]])),result:checked[0],iterations,nodes:count};
}
