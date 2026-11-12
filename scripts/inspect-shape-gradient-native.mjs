import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const path=process.argv[2];if(!path)throw Error('Expected PowerPoint-saved sample');
const bytes=await readFile(path),session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9'}),slides=session.document.slides;
const shapes=[slides[0].elements[0],slides[1].elements[0],slides[2].elements[0].children[0]];
for(const [i,shape] of shapes.entries()){
 assert.equal(shape.type,'shape');assert.deepEqual(shape.style.fill,{type:'linearGradient',angle:45,scaled:true,stops:[{position:0,color:'FF0000',opacity:i===2?.35:.7},{position:1,color:'0000FF',opacity:i===2?.5:1}]});
}
assert.equal(shapes[1].transform.rotation,30);assert.equal(slides[2].elements[0].transform.rotation,30);
const ready=await session.prepareSave({mode:'no-op'});assert.equal(ready.status,'ready');assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');session.dispose();console.log('PowerPoint-saved gradients, transforms and no-op import verified.');
