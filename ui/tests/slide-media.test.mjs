import {test} from 'node:test';
import assert from 'node:assert/strict';
import {mediaBounds,mediaVolume} from '../slide-media.js';
test('media trims remove time from both ends and fades apply inside the trimmed interval',()=>{
 const o={trimStart:2,trimEnd:3,fadeIn:2,fadeOut:1,volume:60};
 assert.deepEqual(mediaBounds(o,10),{start:2,end:7});
 assert.equal(mediaVolume(o,2,10),0);assert.equal(mediaVolume(o,3,10),.3);assert.equal(mediaVolume(o,4,10),.6);assert.equal(mediaVolume(o,6.5,10),.3);assert.equal(mediaVolume(o,7,10),0);
 assert.equal(mediaVolume({volume:0},5,10),0);assert.equal(mediaVolume({},5,10),1);
 assert.deepEqual(mediaBounds({trimStart:20},10),{start:20,end:20});
});
