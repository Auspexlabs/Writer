import {test} from 'node:test';
import assert from 'node:assert/strict';
import {mergeRecords,workbookRecords,mergeFields,mailLayout,addressText} from '../word-mailmerge.js';
import {customPaperCm} from '../word-pages.js';

test('mail merge reads Excel calculated and formatted values including alternate date systems',()=>{
 const doc={date1904:true,sheets:[{name:'Recipients',cells:{A1:{v:'Full Name'},B1:{v:'Date'},C1:{v:'Total'},A2:{v:'张三'},B2:{v:0,s:{code:'yyyy-mm-dd'}},C2:{v:'=10*2',s:{code:'0.00'}}}}]};
 assert.deepEqual(workbookRecords(doc),[{'Full Name':'张三',Date:'1904-01-01',Total:'20.00'}]);
 assert.deepEqual(mergeRecords('Full Name,Address\n"A, B","Line 1\nLine 2"','a.csv'),[{'Full Name':'A, B',Address:'Line 1\nLine 2'}]);
 assert.throws(()=>mergeRecords('Name,Name\nA,B','a.csv'),/不重复/);
});
test('label and envelope pages retain all recipients and escape data as ordinary text',()=>{
 const recipients=Array.from({length:23},(_,i)=>({Name:'<Person '+i+'>',Address:'Road '+i}));
 const labels=mailLayout(recipients,{mode:'labels',address:'{Name}\n{Address}',cols:3,rows:7,labelWidth:6.35,labelHeight:3.81,margin:.7});
 assert.equal((labels.html.match(/<table /g)||[]).length,2);assert.equal((labels.html.match(/&lt;Person/g)||[]).length,23);
 assert.equal((labels.html.match(/<tr /g)||[]).length,14);assert.match(labels.html,/data-w-height="3.81cm"/);
 assert.throws(()=>mailLayout(recipients,{mode:'labels',cols:4,labelWidth:6}),/超出/);
 const env=mailLayout(recipients.slice(0,2),{mode:'envelopes',address:'{Name}\n{Address}',width:22,height:11,sender:'Sender'});
 assert.equal(env.page.size,'22cm x 11cm');assert.equal((env.html.match(/data-pb/g)||[]).length,1);
 assert.equal(addressText('{Full Name}',{'Full Name':'Keep =SUM(A1)'}),'Keep =SUM(A1)');
 assert.deepEqual(mergeFields('<span data-field="MERGEFIELD Name">Name</span><span data-field="MERGEFIELD 地址">地址</span>'),['Name','地址']);
});
test('custom paper dimensions understand native units and reject invalid geometry',()=>{
 assert.deepEqual(customPaperCm('22cm x 110mm'),[22,11]);assert.deepEqual(customPaperCm('8.5in × 11in'),[21.59,27.94]);
 assert.equal(customPaperCm('0cm x 10cm'),null);assert.equal(customPaperCm('90cm x 2cm'),null);
});
