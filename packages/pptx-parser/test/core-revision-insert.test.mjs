import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture,encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',metadataSaveProfile:'core-revision-insert-v9'};
function fixture(strict=false,body='<d:title>Original</d:title>'){const f=metadataFixture(strict);f.add('core',body);return encodeMetadata(f);}

test('missing revision supports explicit normalized default and later edits over three rounds',async()=>{
 for(const strict of [false,true]){
  let bytes=fixture(strict);
  for(let round=0;round<3;round++){
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,options);
   try{if(round===0)assert.equal(s.report.coreMetadataSources.find(x=>x.field==='revision').status,'insertable');if(round>0)s.document.setMetadata({...s.document.metadata,revision:round===1?7:0});const p=await s.prepareSave(round===0?{mode:'edit',coreMetadataInsertions:['revision']}:{mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,value] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),value,name);if(round===0)assert.match(after.get('inventory/core.data').toString(),/<cp:revision xmlns:cp="http:\/\/schemas.openxmlformats.org\/package\/2006\/metadata\/core-properties">1<\/cp:revision>/);}finally{s.dispose();}
   const reopened=await importPptx(bytes,options);try{assert.equal(reopened.document.metadata.revision,[1,7,0][round]);assert.equal(reopened.report.coreMetadataSources.find(x=>x.field==='revision').status,'bound');}finally{reopened.dispose();}
  }
 }
});

test('revision insertion combines with core and custom part creation',async()=>{
 const source=encodeMetadata(metadataFixture()),s=await importPptx(source,{...options,customPropertySaveProfile:'create-and-append-v5'});
 try{s.document.setMetadata({...s.document.metadata,revision:8,title:'New',customProperties:[{name:'A',type:'integer',value:3}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,options);try{assert.equal(reopened.document.metadata.revision,8);assert.equal(reopened.document.metadata.title,'New');assert.equal(reopened.document.metadata.customProperties[0].value,3);}finally{reopened.dispose();}}finally{s.dispose();}
});

test('revision insertion rejects existing malformed or previously unsupported sources and unsafe values',async()=>{
 for(const body of ['<c:revision>01</c:revision>','<c:revision>9007199254740992</c:revision>']){const s=await importPptx(fixture(false,body),options);try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='revision').status,'unmapped');s.document.setMetadata({...s.document.metadata,revision:2});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}}
 const old=await importPptx(fixture(),{projection:'content',metadataSaveProfile:'core-create-v8'});try{old.document.setMetadata({...old.document.metadata,revision:2});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
 const s=await importPptx(fixture(),options);try{s.document.setMetadata({...s.document.metadata,revision:-0});const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['revision']});assert.equal(p.status,'blocked');}finally{s.dispose();}
});


test('revision insertion respects source guards, cancellation and shared limits',async()=>{
 const source=fixture(),limited=await importPptx(source,{...options,sessionLimits:{maxSaveChanges:1}});try{await assert.rejects(limited.prepareSave({mode:'edit',coreMetadataInsertions:['title','revision']}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
 const s=await importPptx(source,options);try{const no=await s.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await s.save(no.plan)).bytes),Buffer.from(source));const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',coreMetadataInsertions:['revision'],signal:ac.signal}),{code:'operation-aborted'});}finally{s.dispose();}
 const f=metadataFixture();f.add('core','');f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
 const guard=await importPptx(encodeMetadata(f),options);try{const p=await guard.prepareSave({mode:'edit',coreMetadataInsertions:['revision']});assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));}finally{guard.dispose();}
});
