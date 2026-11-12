import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {importPptx} from '../packages/pptx-parser/dist/index.js';
const paths=process.argv.slice(2);if(paths.length!==3)throw Error('Expected edited, round2 and round3 files.');
for(const path of paths){
 const bytes=await readFile(path),s=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9'}),slides=s.document.slides;
 const shapes=[slides[0].elements[0],slides[1].elements[0],slides[2].elements[0].children[0]];
 for(const [i,shape] of shapes.entries())assert.deepEqual(shape.style.fill,{type:'linearGradient',angle:i===0?120:45,scaled:i!==0,stops:[{position:0,color:'FF0000',opacity:[.5,.7,.35][i]},{position:1,color:'0000FF',opacity:i===2?.5:1}]});
 assert.equal(shapes[1].transform.rotation,30);assert.equal(slides[2].elements[0].transform.rotation,30);
 const ready=await s.prepareSave({mode:'no-op'});assert.equal(ready.status,'ready');assert.deepEqual((await s.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');s.dispose();
}console.log('Three native shape gradient imports and byte-identical no-op saves verified.');
