import {test} from 'node:test';
import assert from 'node:assert/strict';
import {captionText} from '../slide-captions.js';
test('captions import SRT, UTF-8 BOM, overlapping cues and native WebVTT settings without changing text',()=>{
 const vtt=captionText('\uFEFF1\r\n00:00:01,200 --> 00:00:02,300\r\n你好 & hello\r\n\r\n2\r\n00:00:01,500 --> 00:00:03,000\r\nSecond');
 assert.ok(vtt.startsWith('WEBVTT\n\n'));assert.ok(vtt.includes('00:00:01.200 --> 00:00:02.300'));assert.ok(vtt.includes('你好 & hello'));
 const native='WEBVTT\n\nNOTE import\nkeep\n\ncue\n00:01.200 --> 00:02.300 align:start\n<b>Words</b>';assert.equal(captionText(native),native);
});
test('invalid captions fail before they replace an existing track',()=>{
 for(const text of ['not captions','WEBVTT\n\n','WEBVTT\n\n00:00:02.000 --> 00:00:01.000\nWrong','WEBVTT\n\n00:00:02 --> 00:00:03\nBad','WEBVTT\n\n00:00:61.000 --> 00:02:03.000\nBad','WEBVTT\n\n00:00:02.000 --> 00:60:03.000\nBad','WEBVTT\n\n'+'a'.repeat(2_000_000)])assert.throws(()=>captionText(text));
});
