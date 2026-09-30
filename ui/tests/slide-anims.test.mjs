// node --test ui/tests/ — animations and transitions (PowerPoint's click steps, what shows when, morph pairs, the 动画 tab) and the
// show: its clicks, the presenter view, jumping, blanking, the pen and looping; sections and 幻灯片浏览.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

globalThis.window = globalThis;
const K = await import('../office-io.js');
const plain = x => JSON.parse(JSON.stringify(x)); // the editor's objects come from its vm

test('fxSteps and fxHidden: a click starts a step, with shares its start, after waits for the step before; entrances hide until they play', () => {
  const an = [{ id: 'a', fx: 'fade', start: 'click', dur: 500 }, { id: 'b', fx: 'fly', start: 'with', dur: 500, delay: 200 }, { id: 'a', fx: 'spin', start: 'after', dur: 1000 }, { id: 'b', fx: 'fadeOut', start: 'click', dur: 500 }];
  const st = K.fxSteps(an);
  assert.deepEqual(st.map(s => [s.auto, s.fx.map(f => f.at)]), [[false, [0, 200, 700]], [false, [0]]], 'the spin starts when the fly ends: 200 + 500');
  assert.deepEqual([K.fxHidden(an, 0), K.fxHidden(an, 1, true), K.fxHidden(an, 2, true), K.fxHidden(an, 2, false)], [['a', 'b'], [], [], ['b']]);
  assert.equal(K.fxSteps([{ id: 'a', fx: 'fade', start: 'after' }])[0].auto, true, 'before the first click: plays as the slide appears');
  assert.deepEqual(K.fxFrames({ fx: 'fly' }, { y: 300 }, 900), [{ translate: '0 600px' }, { translate: '0 0' }]);
  assert.equal(K.fxClass({ fx: 'other', cls: 'exit' }), 'exit');
});

test('morphPairs: the same picture, the same text, else the same kind of shape in the same fill', () => {
  const prev = [{ id: 'p1', t: 'text', html: '<p>Q3</p>' }, { id: 'p2', t: 'image', src: 'x.png' }, { id: 'p3', t: 'shape', shape: 'ellipse', fill: 'acc', html: '' }];
  const next = [{ id: 'n1', t: 'shape', shape: 'ellipse', fill: 'acc', html: '' }, { id: 'n2', t: 'text', html: '<p>Q3</p>' }, { id: 'n3', t: 'image', src: 'y.png' }];
  assert.deepEqual(Object.fromEntries(Object.entries(K.morphPairs(prev, next)).map(([k, v]) => [k, v.id])), { n1: 'p3', n2: 'p1' });
  assert.deepEqual(K.morphFrames({ x: 0, y: 0, w: 200, h: 100 }, { x: 100, y: 100, w: 100, h: 100 }), [{ translate: '-50px -100px', scale: '2 1' }, { translate: '0 0', scale: '1 1' }]);
  assert.equal(K.morphFrames({x:0,y:0,w:100,h:100,rot:350},{x:0,y:0,w:100,h:100,rot:10})[0].rotate,'-20deg','pure rotation uses the shortest path across zero');
});

function editor(get, set) {
  const html = readFileSync(new URL('../SlideEditor.dc.html', import.meta.url), 'utf8'), code = html.match(/<script type="text\/x-dc" data-dc-script[^>]*>([\s\S]*?)<\/script>/)[1];
  const ctx = { window: { innerWidth: 1360, innerHeight: 860 }, document: {}, React: { createRef: () => ({ current: null }) }, structuredClone, setTimeout: () => 0, setInterval: () => 0, clearInterval: () => { },
    $t: (s, v) => { const i = String(s).indexOf('@@'), bare = i < 0 ? String(s) : String(s).slice(0, i); return v ? bare.replace(/\{(\w+)\}/g, (m, k) => k in v ? v[k] : m) : bare; }, $lang: () => 'zh',
    DCLogic: class { setState(u, cb) { Object.assign(this.state, u); if (cb) cb(); } forceUpdate() { } } };
  vm.runInNewContext(code + '\nglobalThis.SlideEditor = Component;', ctx);
  const c = new ctx.SlideEditor(); c.props = { get doc() { return get(); }, onChange: set, toast() { } }; c.K = K;
  return c;
}
const slide = (id, extra) => Object.assign({ id, layout: 'blank', decor: [], objs: [], notes: '', trans: 'none', hidden: false, bg: null }, extra);

test('outline edits preserve rich text and nested object identity through the normal history path',()=>{
  const a=K.txt({id:'title',html:'<p><b>Title</b></p>'}),b=K.txt({id:'nested',html:'<p>Before</p>'});
  let doc={id:'d',type:'pptx',ratio:'16:9',slides:[slide('s1',{objs:[a]}),slide('s2',{objs:[{id:'group',t:'group',kids:[b]}]})]},changes=0;
  const c=editor(()=>doc,d=>{doc=d;changes++;});
  const boxes=c.renderVals().outlineItems[1].boxes;assert.equal(boxes.length,1);assert.equal(boxes[0].html.__html,'<p>Before</p>');
  boxes[0].onBlur({currentTarget:{innerHTML:'<p><b>After</b></p>'}});assert.equal(changes,1);
  assert.equal(doc.slides[1].objs[0].kids[0].html,'<p><b>After</b></p>');assert.equal(doc.slides[0].objs[0].html,'<p><b>Title</b></p>');
  c.editOutline(1,'nested','<p><b>After</b></p>');assert.equal(changes,1,'unchanged focus does not create a history entry');
});

test('the show: digits and Enter jump, B blanks until the next key, Ctrl+P draws, 循环放映 starts over, 演示者视图 covers the show without a second window', () => {
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [slide('s1', { notes: 'Say hello' }), slide('s2'), slide('s3')] };
  const c = editor(() => doc, d => { doc = d; }), key = (k, extra) => c.showKey(Object.assign({ key: k, preventDefault() { } }, extra));
  c.startShow(0, true);
  let v = c.renderVals();
  assert.equal(v.presenterIn, true, 'no window.open here: drawn over the show');
  assert.match(v.presenterHtml.__html, /Say hello/); assert.match(v.presenterHtml.__html, /data-act="next"/); assert.match(v.presenterHtml.__html, /data-clock/);
  key('3'); assert.equal(c.renderVals().jumpText, '跳转到第 3 张'); key('Enter'); assert.equal(c.state.show.i, 2);
  key('b'); assert.equal(c.state.show.blank, '#000'); key('x'); assert.deepEqual([c.state.show.blank, c.state.show.i], [null, 2], 'any key brings the slide back, and only that');
  key('p', { ctrlKey: true }); assert.equal(c.state.show.tool, 'pen');
  c.showSlideRef.current = { getBoundingClientRect: () => ({ left: 0, top: 0, width: 800, height: 450 }) };
  c.renderVals().onShowDown({ button: 0, clientX: 100, clientY: 100, preventDefault() { } }); c.renderVals().onShowMove({ clientX: 200, clientY: 100 }); c.renderVals().onShowUp();
  assert.deepEqual(plain(c.state.ink[2]), [[[200, 200], [400, 200]]]);
  assert.match(c.renderVals().inkSvg.__html, /points="200,200 400,200"/);
  c.persistInk = () => { c.state.ink = {}; }; // vector persistence is exercised by verify-slide-017 in Chrome
  key('Escape'); assert.equal(c.state.show.tool, null, 'Esc puts the pen away first');
  c.setState({ loop: true }); c.showNext(); assert.equal(c.state.show.i, 0, 'after the last slide, the first again');
  c.persistInk = () => { c.state.ink = {}; }; // vector persistence is exercised by verify-slide-017 in Chrome
  key('Escape'); assert.equal(c.state.show, null);
});

test('the editor: 动画 adds effects to the selection, the pane reorders them, a deleted object takes its effects along; the show steps through them', () => {
  const a = K.txt({ id: 'a', html: '<p>A</p>' }), b = K.shape({ id: 'b' });
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [slide('s1', { objs: [a, b] }), slide('s2', { trans: 'morph' })] };
  const c = editor(() => doc, d => { doc = d; });
  c.setState({ sel: 'a', sels: ['a', 'b'] }); c.addFx('fade');
  assert.deepEqual(plain(doc.slides[0].anims.map(x => [x.id, x.fx, x.dur])), [['a', 'fade', 500], ['b', 'fade', 500]]);
  c.setState({ tab: 'anim', fxPane: true });
  let v = c.renderVals();
  v.ribbon.find(i => i.isSel).onChange({ target: { value: 'after' } });
  c.menus.fxexit.find(i => i.label === '淡出').onClick(); // 退出 › 淡出 on both again
  v = c.renderVals();
  assert.deepEqual(plain(v.fxRows.map(r => [r.n, r.label])), [[1, '淡入 · A'], ['', '淡入 · 形状'], [2, '淡出 · A'], [3, '淡出 · 形状']]);
  assert.deepEqual(plain(v.objs.map(o => o.badge)), [1, 1]);
  v.fxRows[3].up({ stopPropagation() { } });
  assert.deepEqual(plain(doc.slides[0].anims.map(x => x.id + ':' + x.fx)), ['a:fade', 'b:fade', 'b:fadeOut', 'a:fadeOut']);
  // the show: both hidden until the first click, then B goes on the second, A on the third, and the next click moves on
  c.startShow(0);
  const hidden = () => plain(c.renderVals().showHidden);
  assert.deepEqual([c.state.show.step, hidden()], [0, ['a', 'b']]);
  c.showNext(); assert.deepEqual(hidden(), []);
  c.showNext(); c.showPrev(); c.showNext(); c.showNext(); assert.deepEqual([c.state.show.step, hidden()], [3, ['b']]);
  c.showNext(); assert.deepEqual([c.state.show.i, c.pendShow.trans.kind, c.pendShow.trans.from], [1, 'morph', 0]);
  c.showPrev(); assert.deepEqual([c.state.show.i, c.state.show.step, hidden()], [0, 3, ['a', 'b']]);
  c.exitShow();
  // deleting B takes its effects along; a duplicated slide's effects point at the copies
  c.setState({ show: null, cur: 0, sel: 'b', sels: ['b'] }); c.deleteSels();
  assert.deepEqual(plain(doc.slides[0].anims.map(x => x.id)), ['a', 'a']);
  c.dupSlide(0);
  assert.deepEqual(doc.slides[1].anims.map(x => x.id), [doc.slides[1].objs[0].id, doc.slides[1].objs[0].id]);
});

test('sections follow their first slide: a deleted first slide hands its section on, the first section starts at slide 1', () => {
  const S = (id, sec) => ({ id, sec }), ids = l => l.map(s => s.id + (s.sec ? ':' + s.sec : ''));
  assert.deepEqual(ids(K.keepSections([S('a', 'X'), S('b'), S('c', 'Y'), S('d')], [S('a', 'X'), S('b'), S('d')])), ['a:X', 'b', 'd:Y'], 'Y goes on from its next slide');
  assert.deepEqual(ids(K.keepSections([S('a', 'X'), S('b'), S('c', 'Y')], [S('a', 'X'), S('b')])), ['a:X', 'b'], "Y's only slide went: so did Y");
  assert.deepEqual(ids(K.keepSections([S('a', 'X'), S('b', 'Y'), S('c')], [S('a', 'X'), S('c')])), ['a:X', 'c:Y'], 'Y lives on in c');
  assert.deepEqual(ids(K.keepSections([S('a', 'X'), S('b'), S('c', 'Y')], [S('b'), S('c', 'Y')])), ['b:X', 'c:Y']);
  assert.deepEqual(ids(K.keepSections([S('a'), S('b', 'Y')], [S('b', 'Y'), S('a')].reverse())), ['a:Y', 'b'], 'moved behind: the first section still starts at slide 1');
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', slides: [slide('s1'), slide('s2'), slide('s3')] };
  const c = editor(() => doc, d => { doc = d; });
  c.addSection(2);
  assert.deepEqual(doc.slides.map(s => s.sec || ''), ['默认节', '', '无标题节']);
  c.setState({ cur: 2 }); let v = c.renderVals();
  assert.deepEqual(plain(v.thumbs.map(t => [t.hasSec, t.secCount])), [[true, '2 张'], [false, ''], [true, '1 张']]);
  v.thumbs[2].onSec({ target: { value: '结果' } });
  c.deleteSlide(2);
  assert.deepEqual(doc.slides.map(s => s.sec || ''), ['默认节', ''], 'its last slide gone, the section goes');
  v = c.renderVals(); v.toggleSorter();
  v = c.renderVals();
  assert.deepEqual([v.sorter, v.showThumbs], [true, false]);
  v.thumbs[1].onOpen();
  assert.deepEqual([c.state.sorter, c.state.cur], [false, 1]);
});

test('export: the outline as Markdown (titles as headings, text as bullets, notes quoted) and a slide painted on a canvas', () => {
  const t = K.txt({ ph: 'title', html: '<p>Q3 &amp; Q4</p>' }), body = K.txt({ html: '<ul><li>Revenue up</li><li data-lvl="1">Asia first</li></ul>' });
  const doc = { title: 'Deck', ratio: '16:9', theme: 'paper', slides: [{ id: 'a', objs: [t, body], notes: 'Smile' }, { id: 'b', objs: [], hidden: true }, { id: 'c', objs: [K.txt({ html: '<p>Thanks</p>' })] }] };
  assert.equal(K.slidesMarkdown(doc), '# Deck\n\n## Q3 & Q4\n\n- Revenue up\n  - Asia first\n\n> Smile\n\n## Thanks\n');
  const calls = [], g = new Proxy({}, { get: (o, k) => k in o ? o[k] : k === 'measureText' ? s => ({ width: s.length * 10 }) : (...a) => calls.push([k, ...a]), set: (o, k, v) => { o[k] = v; return true; } });
  globalThis.Path2D ??= class { ellipse() { } lineTo() { } moveTo() { } roundRect() { } closePath() { } };
  K.paintSlide(g, doc.slides[0], K.THEMES.paper, 900, new Map());
  assert.deepEqual(calls.filter(c => c[0] === 'fillText').map(c => c[1]), ['Q3 & Q4', '•', 'Revenue up', '•', 'Asia first']);
  assert.deepEqual(calls[0], ['fillRect', 0, 0, 1600, 900], 'the background first');
});

test('the format panel: the Word 定稿 panel with the deck\'s own tabs; a selected text box brings its text groups and the 形状 context tab', async () => {
  const FP = await import('../panel.js');
  const t = K.txt({ id: 't1', html: '<p>Hi</p>' });
  let doc = { id: 'd', type: 'pptx', ratio: '16:9', theme: 'paper', slides: [slide('s1', { objs: [t] })] };
  const c = editor(() => doc, d => { doc = d; }); c.FP = FP; c.props.formatOpen = true;
  const titles = v => Array.from(v.panelGroups, g => g.title);
  let v = c.renderVals();
  assert.deepEqual([v.showBar, v.bubbleOpen, v.formatOpen, v.panelPad], [false, false, true, '312px']);
  assert.deepEqual(Array.from(v.panelTabs, x => x.label), ['开始', '插入', '设计', '切换', '动画', '审阅', '视图', '放映']);
  assert.deepEqual(titles(v), ['幻灯片', '文字'], 'nothing selected: the slide, and where the text tools are');
  c.setState({ sel: 't1', sels: ['t1'] }); v = c.renderVals();
  assert.deepEqual(plain(v.panelTabs.slice(-1).map(x => [x.label, x.ctx])), [['形状', true]]);
  assert.deepEqual(titles(v), ['幻灯片', '字体', '段落', '文本框']);
  c.state.tab = 'format'; v = c.renderVals(); assert.deepEqual(titles(v), ['填充与轮廓', '大小', '排列', '大小与位置']);
  for (const [tab, want] of [['insert', ['常用', '页面']], ['design', ['主题', '背景', '幻灯片大小', '母版与版式']], ['trans', ['效果选项', '切换效果', '计时']], ['anim', ['添加动画', '效果选项', '计时', '重复与反向', '动画窗格']], ['show', ['开始放映', '自定义放映', '视图']]]) {
    c.state.tab = tab; v = c.renderVals(); assert.deepEqual(titles(v), want, tab);
  }
  c.state.tab = 'design'; v = c.renderVals();
  const themes = v.panelGroups[0].rows[0].items; assert.equal(themes[0].t, 'theme'); themes.find(x => !x.on).onClick();
  assert.notEqual(doc.theme, 'paper', 'a theme tile sets the deck\'s theme');
});

test('selection pane exposes nested objects; hidden items disappear and locked items reject edits',()=>{
 const a=K.txt({id:'a',name:'Title',html:'<p>Hello</p>'}),b=K.shape({id:'b',name:'Child',fill:'#ff0000',x:10,y:10,w:20,h:20}),g=K.group([b]);
 let doc={id:'pane',type:'pptx',ratio:'16:9',slides:[slide('s',{objs:[a,g]})]};const history=[],c=editor(()=>doc,d=>{history.push(doc);doc=d;});
 assert.deepEqual(plain(c.objectRows().map(x=>[x.o.id,x.depth])),[[g.id,0],['b',1],['a',0]]);
 c.objectState('b',{name:'Nested',locked:true});assert.equal(c.objectRows()[1].o.name,'Nested');
 c.patchObj(g.id,{x:999});assert.notEqual(doc.slides[0].objs[1].x,999);c.select(g.id,false);assert.equal(c.obj,undefined);
 c.state.sels=[g.id];c.deleteSels();assert.equal(doc.slides[0].objs.length,2);
 c.objectState('a',{hidden:true});assert.ok(!c.renderVals().objs.some(x=>x.id==='a'));
 c.objectState('b',{hidden:true});const group=doc.slides[0].objs[1];assert.equal(K.objView(group,K.THEMES.paper).groupInner.__html,'');
 c.objectState('b',{locked:false});c.objectState('a',{hidden:false});c.select('a',false);c.patchSel({x:123});assert.equal(doc.slides[0].objs[0].x,123);
 doc=history.at(-1);assert.notEqual(doc.slides[0].objs[0].x,123,'changes follow normal undo history');
});

test('nested group isolation edits members and preserves ancestor transforms and locked siblings',()=>{
 const a=K.txt({id:'a',x:100,y:200,html:'<p>Before</p>'}),b=K.shape({id:'b',x:400,locked:true}),inner=K.group([a,b]),outer=K.group([inner]);inner.rot=30;outer.rot=90;
 let doc={id:'groups',type:'pptx',ratio:'16:9',slides:[slide('s',{objs:[outer,K.txt({id:'outside'})]})]};const history=[],c=editor(()=>doc,d=>{history.push(doc);doc=d;});
 c.enterGroup(outer.id);c.enterGroup(inner.id);assert.deepEqual(plain(c.objects().map(o=>o.id)),['a','b']);c.select('a',false);c.patchSel({x:123,html:'<p>Changed</p>'});
 assert.equal(doc.slides[0].objs[0].kids[0].kids[0].x,123);assert.equal(doc.slides[0].objs[0].rot,90);assert.equal(doc.slides[0].objs[0].kids[0].rot,30);assert.equal(doc.slides[0].objs[0].kids[0].kids[1].x,400);
 c.insertObj(K.txt({id:'new'}));assert.ok(c.objects().some(o=>o.id==='new'));assert.equal(doc.slides[0].objs.length,2);c.deleteSels();assert.equal(c.objects().length,2);doc=history.at(-1);assert.ok(c.objects().some(o=>o.id==='new'));
 c.leaveGroup();assert.equal(c.objects()[0].id,inner.id);c.leaveGroup();assert.equal(c.objects().length,2);
});
test('handouts paginate slides, omit hidden content, escape notes and fit each column',()=>{
 const doc={type:'pptx',theme:'paper',ratio:'16:9',slides:Array.from({length:10},(_,i)=>slide('s'+i,{hidden:i===2,notes:'<script>note</script>',objs:[K.txt({html:'<p>Slide '+i+'</p>'})]}))};
 for(const count of [1,2,3,4,6,9]){const out=K.slidesPrint(doc,{mode:'handout',count});assert.equal((out.body.match(/class="handout"/g)||[]).length,Math.ceil(9/count));assert.doesNotMatch(out.body,/Slide 2</);assert.match(out.body,/Slide 9</);}
 const notes=K.slidesPrint(doc,{mode:'notes',hidden:true});assert.equal((notes.body.match(/class="slide-notes"/g)||[]).length,10);assert.match(notes.body,/&lt;script&gt;note/);assert.match(notes.css,/height:auto/);
});

test('repeated reversing effects extend the next step and return to their starting visibility',()=>{
 const fx={id:'a',fx:'fade',dur:500,repeat:3,autoReverse:true};
 assert.equal(K.fxSpan(fx),3000);assert.deepEqual(K.fxTiming(fx),{duration:500,iterations:6,direction:'alternate'});
 assert.equal(K.fxSteps([fx,{id:'b',fx:'spin',start:'after',dur:100}])[0].fx[1].at,3000);
 assert.deepEqual(K.fxHidden([fx],1,false),['a']);assert.deepEqual(K.fxHidden([fx],1,true),[]);
 assert.deepEqual(K.fxHidden([{...fx,fx:'fadeOut'}],1,false),[]);
 assert.deepEqual(K.fxHidden([{...fx,repeat:1.5}],1,false),[]);
 assert.deepEqual(K.fxHidden([{...fx,repeat:1.5,fx:'fadeOut'}],1,false),['a']);
 assert.deepEqual(K.fxHidden([{...fx,repeat:1.25,fx:'fadeOut'}],1,false),[]);
});

 test('custom shows preserve repeated and hidden slides, route navigation, jumps and looping',()=>{
 let doc={id:'custom',type:'pptx',ratio:'16:9',slides:[slide('a'),slide('b',{hidden:true}),slide('c'),slide('d')],customShows:[{name:'Demo',slides:['c','a','c','b','missing']}]};
 const c=editor(()=>doc,d=>doc=d),key=k=>c.showKey({key:k,preventDefault(){}});
 assert.deepEqual(K.customShowRoute(doc,'Demo'),[2,0,2,1]);
 c.startShow(0,false,'Demo');assert.equal(c.state.show.i,2);assert.equal(c.state.show.routePos,0);
 c.showNext();assert.equal(c.state.show.i,0);c.showNext();assert.equal(c.state.show.i,2);assert.equal(c.state.show.routePos,2);
 c.showPrev();assert.equal(c.state.show.i,0);key('End');assert.equal(c.state.show.i,1);assert.equal(c.state.show.routePos,3);
 c.state.loop=true;c.showNext();assert.equal(c.state.show.i,2);assert.equal(c.state.show.routePos,0);
 key('4');key('Enter');assert.equal(c.state.show.i,2);key('1');key('Enter');assert.equal(c.state.show.routePos,1);
 key('Home');assert.equal(c.state.show.routePos,0);key('End');c.state.loop=false;c.showNext();assert.equal(c.state.show,null);
 });
