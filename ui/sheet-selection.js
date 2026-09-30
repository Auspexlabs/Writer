// Disjoint selections are rectangles, never a set containing a million cell addresses.
export const inArea = (b,r,c) => r>=b.r1&&r<=b.r2&&c>=b.c1&&c<=b.c2;
export const areaBounds = areas => areas.reduce((b,a)=>({r1:Math.min(b.r1,a.r1),c1:Math.min(b.c1,a.c1),r2:Math.max(b.r2,a.r2),c2:Math.max(b.c2,a.c2)}),{r1:Infinity,c1:Infinity,r2:-1,c2:-1});
export function subtractArea(p,cut) {
  const r1=Math.max(p.r1,cut.r1),r2=Math.min(p.r2,cut.r2),c1=Math.max(p.c1,cut.c1),c2=Math.min(p.c2,cut.c2);
  return r1>r2||c1>c2?[p]:[{...p,r2:r1-1},{...p,r1:r2+1},{r1,r2,c1:p.c1,c2:c1-1},{r1,r2,c1:c2+1,c2:p.c2}].filter(x=>x.r1<=x.r2&&x.c1<=x.c2);
}
export function unionAreas(areas) {
  const result=[];
  for(const area of areas){let pieces=[area];for(const old of result)pieces=pieces.flatMap(p=>subtractArea(p,old));result.push(...pieces);}
  return result.sort((a,b)=>a.r1-b.r1||a.c1-b.c1);
}
export function areaContains(areas,box) { let pieces=[box];for(const a of areas)pieces=pieces.flatMap(p=>subtractArea(p,a));return !pieces.length; }
