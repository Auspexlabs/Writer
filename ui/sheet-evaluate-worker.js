import {evaluateFormula} from './sheet-evaluate.js';
self.onmessage=e=>{try{const {doc,si,ref}=e.data;self.postMessage({result:evaluateFormula(doc,si,ref)});}catch(error){self.postMessage({error:error.message});}};
