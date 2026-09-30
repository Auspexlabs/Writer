import {solveWorkbook} from './sheet-solver.js';
self.onmessage=e=>{try{self.postMessage({result:solveWorkbook(e.data.doc,e.data.si,e.data.spec)});}catch(error){self.postMessage({error:error.message});}};
