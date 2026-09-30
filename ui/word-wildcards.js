// A bounded matcher for Word's wildcard syntax. Patterns never run as an
// unbounded backtracking RegExp on the editor's document text.
export function wildcardMatches(text, query, caseSensitive = false) {
  const tokens=[],groups=[],flags=caseSensitive?'u':'iu';let groupCount=0;
  const atom=(source)=>({test:new RegExp('^(?:'+source+')$',flags),min:1,max:1});
  const escape=c=>c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  if(query.length>256)throw new Error('Search pattern is too long');
  for(let i=0;i<query.length;i++){
    const c=String.fromCodePoint(query.codePointAt(i));i+=c.length-1;
    if(c==='\\'){if(++i>=query.length)throw new Error('Incomplete escape');const next=String.fromCodePoint(query.codePointAt(i));tokens.push(atom(escape(next)));i+=next.length-1;}
    else if(c==='^'&&/^(13|p|t|l)/.test(query.slice(i+1))){const m=/^(13|p|t|l)/.exec(query.slice(i+1))[0];tokens.push(atom(m==='t'?'\\t':'\\n'));i+=m.length;}
    else if(c==='?')tokens.push(atom('[^\\n\\ufffc]'));
    else if(c==='*')tokens.push({...atom('[^\\n\\ufffc]'),min:0,max:Infinity,lazy:true,quantified:true});
    else if(c==='['){const end=query.indexOf(']',i+1);if(end<0)throw new Error('Unclosed character set');let set=query.slice(i+1,end);if(!set||/[\\[\n]/.test(set))throw new Error('Invalid character set');if(set[0]==='!')set='^'+set.slice(1)+'\\n\\ufffc';tokens.push(atom('['+set+']'));i=end;}
    else if(c==='('){groups.push(groupCount);tokens.push({open:groupCount++});}
    else if(c===')'){if(!groups.length)throw new Error('Unbalanced group');tokens.push({close:groups.pop()});}
    else if(c==='<')tokens.push({boundary:'start'});
    else if(c==='>')tokens.push({boundary:'end'});
    else if(c==='@'||c==='{'){
      const t=tokens.at(-1);if(!t?.test||t.quantified)throw new Error('Invalid repetition');
      if(c==='@'){t.min=1;t.max=Infinity;}else{const m=/^\{(\d+)(?:,(\d*))?\}/.exec(query.slice(i));if(!m||+m[1]>10000||(m[2]&&(+m[2]<+m[1]||+m[2]>10000)))throw new Error('Invalid repetition');t.min=+m[1];t.max=m[2]===undefined?t.min:m[2]===''?Infinity:+m[2];i+=m[0].length-1;}t.quantified=true;
    }
    else tokens.push(atom(escape(c)));
  }
  if(groups.length)throw new Error('Unclosed group');
  const chars=Array.from(text),offsets=[0];for(const c of chars)offsets.push(offsets.at(-1)+c.length);
  const word=c=>/[\p{L}\p{N}_]/u.test(c||'');let work=0;
  const step=()=>{if(++work>2000000)throw new Error('Search pattern is too complex. Narrow the search.');};
  const matchAt=start=>{
    const stack=[{t:0,p:start,caps:[]}],seen=new Set();
    while(stack.length){step();const state=stack.pop(),{t,p,caps}=state,key=t+':'+p;if(seen.has(key))continue;seen.add(key);
      if(t===tokens.length){if(p>start)return {index:offsets[start],length:offsets[p]-offsets[start],match:text.slice(offsets[start],offsets[p]),groups:Array.from({length:groupCount},(_,i)=>caps[i]?text.slice(offsets[caps[i][0]],offsets[caps[i][1]]):''),end:p};continue;}
      const token=tokens[t];
      if(token.open!==undefined||token.close!==undefined){const next=caps.slice();if(token.open!==undefined)next[token.open]=[p,p];else next[token.close]=[next[token.close][0],p];stack.push({t:t+1,p,caps:next});}
      else if(token.boundary){if(token.boundary==='start'?!word(chars[p-1])&&word(chars[p]):word(chars[p-1])&&!word(chars[p]))stack.push({t:t+1,p,caps});}
      else{let end=p;while(end<chars.length&&end-p<token.max&&chars[end]!=='\ufffc'&&token.test.test(chars[end])){step();end++;}
        if(end-p<token.min)continue;
        if(token.lazy){for(let i=end;i>=p+token.min;i--){step();stack.push({t:t+1,p:i,caps});}}
        else for(let i=p+token.min;i<=end;i++){step();stack.push({t:t+1,p:i,caps});}
      }
    }
    return null;
  };
  const found=[];
  for(let start=0;start<chars.length;){step();const hit=matchAt(start);if(hit){found.push(hit);start=hit.end;}else start++;}
  return found.map(({end,...hit})=>hit);
}
