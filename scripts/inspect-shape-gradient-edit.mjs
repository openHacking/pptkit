import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {importPptx} from '../packages/pptx-parser/dist/index.js';
const path=process.argv[2];if(!path)throw Error('Expected PowerPoint-edited file');
const bytes=await readFile(path),session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9'});
assert.deepEqual(session.document.slides[0].elements[0].style.fill,{type:'linearGradient',angle:120,scaled:false,stops:[{position:0,color:'FF0000',opacity:.5},{position:1,color:'0000FF',opacity:1}]});
const ready=await session.prepareSave({mode:'no-op'});assert.equal(ready.status,'ready');assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');session.dispose();console.log('Native edited shape import and no-op passed.');
