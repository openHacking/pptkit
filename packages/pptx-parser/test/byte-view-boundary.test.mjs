import test from 'node:test';
import assert from 'node:assert/strict';
import { inspectZipArchive, inspectXmlPart, compareOpcPackages, importPptx, inspectPptxMasterPictures } from '../dist/index.js';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';
import { structureFixture, encodeStructure } from './structure-fixture.mjs';

class Hostile extends Uint8Array { get byteLength(){return 0}get length(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('unexpected iterator')} }
const document=()=>encodeStructure(structureFixture()),xml=()=>Buffer.from('<r><t>hello</t></r>');

test('inspection and import use true view length and snapshot without invoking custom iterator',async()=>{
 const source=document(),hostile=new Hostile(source),sourceXml=xml(),hostileXml=new Hostile(sourceXml);
 assert.deepEqual(await inspectZipArchive(hostile),await inspectZipArchive(source));assert.deepEqual(await inspectXmlPart(hostileXml),await inspectXmlPart(sourceXml));
 const session=await importPptx(hostile,{projection:'structure'});try{const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(source));}finally{session.dispose();}
 const compared=await compareOpcPackages(hostile,source);assert.deepEqual(compared.differences,[]);assert.equal((await inspectPptxMasterPictures(hostile)).masters.length,1);
 await assert.rejects(inspectZipArchive(hostile,{limits:{maxInputBytes:1}}),{code:'archive-limit-exceeded'});await assert.rejects(inspectXmlPart(hostileXml,{limits:{maxInputBytes:1}}),{code:'xml-limit-exceeded'});
 await assert.rejects(importPptx(hostile,{projection:'structure',zipLimits:{maxInputBytes:1}}),{code:'import-limit-exceeded'});await assert.rejects(compareOpcPackages(hostile,source,{zipLimits:{maxInputBytes:1}}),{code:'archive-limit-exceeded'});
 await assert.rejects(captureXmlTextSource(hostileXml,{limits:{maxInputBytes:1}}),{code:'xml-limit-exceeded'});
});

test('inspection, import and preservation reject shared buffers across realms',async()=>{
 const vm=await import('node:vm'),source=document(),foreign=vm.runInNewContext(`new SharedArrayBuffer(${source.length})`),shared=new Hostile(foreign);shared.set(source);assert.equal(foreign instanceof SharedArrayBuffer,false);
 for(const run of [()=>inspectZipArchive(shared),()=>importPptx(shared,{projection:'structure'}),()=>compareOpcPackages(source,shared),()=>inspectPptxMasterPictures(shared),()=>replaceArchiveEntries(shared,[],{maxOutputBytes:1000000})])await assert.rejects(run(),/SharedArrayBuffer|ordinary byte buffer/);
 const sharedXml=new Hostile(vm.runInNewContext('new SharedArrayBuffer(32)'));sharedXml.set(xml());await assert.rejects(inspectXmlPart(sharedXml),/SharedArrayBuffer/);await assert.rejects(captureXmlTextSource(sharedXml),/ordinary byte buffer/);
});

test('ZIP replacement enforces actual patch size and captures each replacement once',async()=>{
 const source=document(),patch=new Hostile(Buffer.from('<replacement/>'));
 await assert.rejects(replaceArchiveEntries(source,[{name:'deck/a.xml',bytes:patch}],{maxOutputBytes:1000000,limits:{maxEntryBytes:1}}),{code:'archive-limit-exceeded'});
 let reads=0;const replacement={name:'deck/a.xml',get bytes(){reads++;return patch}};const output=await replaceArchiveEntries(source,[replacement],{maxOutputBytes:1000000});assert.equal(reads,1);assert.ok((await inspectZipArchive(output)).entries.some(x=>x.name==='deck/a.xml'));
});
