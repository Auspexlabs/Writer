import {Calc,parseA,A} from './sheet-engine.js';
export function whatIfBox(table){const [a,b=a]=String(table.range||'').split(':').map(parseA);if(!a||!b||b.r<=a.r||b.c<=a.c)throw Error('模拟运算表需要包含标题行、标题列和结果区域');return {r1:a.r,c1:a.c,r2:b.r,c2:b.c};}
export function validateDataTable(doc,si,table) {
 const sh=doc.sheets[si],b=whatIfBox(table),row=table.rowInput&&parseA(table.rowInput),col=table.colInput&&parseA(table.colInput);if(!row&&!col||table.rowInput&&!row||table.colInput&&!col)throw Error('请输入行输入单元格或列输入单元格');if(row&&col&&row.r===col.r&&row.c===col.c)throw Error('两个输入单元格不能相同');if((b.r2-b.r1)*(b.c2-b.c1)>100000)throw Error('模拟运算表一次最多计算 100000 个结果');
 for(const p of [row,col].filter(Boolean)){if(p.r>=b.r1&&p.r<=b.r2&&p.c>=b.c1&&p.c<=b.c2)throw Error('输入单元格不能位于模拟运算表内');const cell=sh.cells[A(p.r,p.c)];if(cell?.spill||!cell?.literal&&String(cell?.v||'').startsWith('='))throw Error('输入单元格必须是可修改的常量');}
 const formula=(r,c)=>{const x=sh.cells[A(r,c)];if(x?.literal||!String(x?.v||'').startsWith('='))throw Error('结果标题单元格必须包含公式');};
 if(row&&col)formula(b.r1,b.c1);else if(col)for(let c=b.c1+1;c<=b.c2;c++)formula(b.r1,c);else for(let r=b.r1+1;r<=b.r2;r++)formula(r,b.c1);
 for(const other of sh.dataTables||[]){if(other===table||other.range===table.range)continue;const o=whatIfBox(other);if(b.r1<=o.r2&&b.r2>=o.r1&&b.c1<=o.c2&&b.c2>=o.c1)throw Error('模拟运算表不能重叠');}if((sh.merges||[]).some(m=>m.r<=b.r2&&m.r+m.rs>b.r1&&m.c<=b.c2&&m.c+m.cs>b.c1))throw Error('模拟运算表不能包含合并单元格');return b;
}
export function dataTableValue(calc,si,r,c) {
 const sh=calc.doc.sheets[si];if(!sh?.dataTables?.length)return undefined;
 calc._whatIfRanges??=new Map();if(!calc._whatIfRanges.has(si))calc._whatIfRanges.set(si,sh.dataTables.map(table=>{try{return {table,box:whatIfBox(table)};}catch{return null;}}).filter(Boolean));
 const match=calc._whatIfRanges.get(si).find(({box:b})=>r>b.r1&&r<=b.r2&&c>b.c1&&c<=b.c2);if(!match)return undefined;
 const {table,box:b}=match,token=si+':'+table.range,stack=calc.doc._dataTableStack||[];if(stack.includes(token))return {err:'#CIRC!'};if(!table.rowInput&&!table.colInput||table.rowInput&&!parseA(table.rowInput)||table.colInput&&!parseA(table.colInput))return {err:'#REF!'};
 const inputs={};if(table.rowInput)inputs[table.rowInput.replace(/\$/g,'').toUpperCase()]=calc.value(si,b.r1,c);if(table.colInput)inputs[table.colInput.replace(/\$/g,'').toUpperCase()]=calc.value(si,r,b.c1);for(const value of Object.values(inputs))if(value?.err)return value;
 const cells=new Proxy(sh.cells,{get:(base,key)=>Object.hasOwn(inputs,key)?{v:inputs[key],literal:typeof inputs[key]==='string'&&inputs[key].startsWith('=')}:base[key],ownKeys:base=>[...new Set([...Reflect.ownKeys(base),...Object.keys(inputs)])],getOwnPropertyDescriptor:(base,k)=>Object.hasOwn(inputs,k)?{configurable:true,enumerable:true}:Object.getOwnPropertyDescriptor(base,k)});
 const doc={...calc.doc,_dataTableStack:[...stack,token],sheets:calc.doc.sheets.map((s,i)=>i===si?{...s,cells}:s)},trial=new Calc(doc),fr=table.colInput?b.r1:r,fc=table.rowInput?b.c1:c;const result=trial.value(si,fr,fc);return result;
}
