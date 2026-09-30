export const pivotDisplays=[['normal','原始汇总'],['percentOfRow','行汇总百分比'],['percentOfCol','列汇总百分比'],['percentOfTotal','总计百分比'],['index','指数'],['runTotal','累计值'],['difference','与上一项差值'],['percent','与上一项比值'],['percentDiff','与上一项百分比差异']];
/** Display calculations use the aggregate for each group, not the raw input rows. */
export function pivotDisplayed(value,row,col,spec,rows,cols,rowFields,colFields) {
 const mode=spec.showAs||'normal',v=value(row,col);
 if(mode==='normal'||typeof v==='string'&&v.startsWith('#'))return v;
 const number=x=>typeof x==='number'?x:x===''?0:Number(x),divide=(a,b)=>typeof b==='string'&&b.startsWith('#')?b:number(b)===0?'#DIV/0!':number(a)/number(b);
 if(mode==='percentOfRow')return divide(v,value(row,-1));
 if(mode==='percentOfCol')return divide(v,value(-1,col));
 if(mode==='percentOfTotal')return divide(v,value(-1,-1));
 if(mode==='index')return divide(number(v)*number(value(-1,-1)),number(value(row,-1))*number(value(-1,col)));
 const axis=rowFields.includes(spec.baseField)?rows:cols,fields=rowFields.includes(spec.baseField)?rowFields:colFields,index=axis===rows?row:col,pos=fields.indexOf(spec.baseField);
 if(index<0||pos<0)return mode==='runTotal'?v:'';
 const key=axis[index].value,items=axis.filter(item=>item.value.every((x,k)=>k===pos||JSON.stringify(x)===JSON.stringify(key[k])));
 const here=items.findIndex(item=>item.index===index),get=item=>axis===rows?value(item.index,col):value(row,item.index);
 if(mode==='runTotal'){let sum=0;for(const item of items.slice(0,here+1)){const x=get(item);if(typeof x==='string'&&x.startsWith('#'))return x;sum+=number(x);}return sum;}
 const other=items[here+(spec.baseItem==='next'?1:-1)];if(!other)return '';const base=get(other);if(typeof base==='string'&&base.startsWith('#'))return base;
 return mode==='difference'?number(v)-number(base):mode==='percent'?divide(v,base):divide(number(v)-number(base),base);
}
export function pivotItemMatches(value,items){return !Array.isArray(items)||items.some(item=>typeof item===typeof value&&JSON.stringify(item)===JSON.stringify(value));}
