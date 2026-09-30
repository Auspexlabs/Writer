import {inferFlashFill} from './sheet-flash-fill.js';
self.onmessage=({data})=>{try{self.postMessage({results:inferFlashFill(data.rows,data.examples)});}catch(e){self.postMessage({error:e.message});}};
