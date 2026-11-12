import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture,encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'core-keywords-insert-v10'};
function fixture(strict=false,body='<d:title>Original</d:title>'){const f=metadataFixture(strict);f.add('core',body);return encodeMetadata(f);}

test('one missing keyword inserts and stays editable over three rounds in both dialects',async()=>{
 for(const strict of [false,true]){
  let bytes=fixture(strict);
  for(const [round,value] of ['', '中文 & <tag>', 'final, single'].entries()){
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{if(round===0)assert.equal(s.report.coreMetadataSources.find(x=>x.field==='keywords').status,'insertable');s.document.setMetadata({...s.document.metadata,keywords:[value]});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/core.data').toString(),/<cp:keywords xmlns:cp=/);}finally{s.dispose();}
   const reopened=await importPptx(bytes,config);try{assert.deepEqual(reopened.document.metadata.keywords,[value]);assert.equal(reopened.report.coreMetadataSources.find(x=>x.field==='keywords').status,'bound');}finally{reopened.dispose();}
  }
 }
});

test('keyword insertion combines with simultaneous core, revision and custom creation',async()=>{
 const s=await importPptx(encodeMetadata(metadataFixture()),{...config,customPropertySaveProfile:'create-and-append-v5'});
 try{s.document.setMetadata({...s.document.metadata,keywords:['value'],revision:1,customProperties:[{name:'A',type:'integer',value:2}]});const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['revision'],customPropertyAdditions:[0]});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,config);try{assert.deepEqual(reopened.document.metadata.keywords,['value']);assert.equal(reopened.document.metadata.revision,1);assert.equal(reopened.document.metadata.customProperties[0].value,2);}finally{reopened.dispose();}}finally{s.dispose();}
});

test('unsupported keyword structure and multiple values cannot be inserted',async()=>{
 for(const body of ['<c:keywords><c:value>one</c:value></c:keywords>','<c:keywords xml:lang="en">one</c:keywords>']){const s=await importPptx(fixture(false,body),config);try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='keywords').status,'unmapped');s.document.setMetadata({...s.document.metadata,keywords:['new']});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}}
 const s=await importPptx(fixture(),config);try{s.document.setMetadata({...s.document.metadata,keywords:['one','two']});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}
 const old=await importPptx(fixture(),{projection:'content',metadataSaveProfile:'core-revision-insert-v9'});try{old.document.setMetadata({...old.document.metadata,keywords:['one']});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
});


test('keyword insertion observes no-op, limits, invalid XML, guards and freshness',async()=>{
 const source=fixture(),s=await importPptx(source,config);try{const no=await s.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await s.save(no.plan)).bytes),Buffer.from(source));s.document.setMetadata({...s.document.metadata,keywords:['bad\u0000']});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');s.document.setMetadata({...s.document.metadata,keywords:['valid']});const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:ac.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,keywords:['changed']});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}
 const limited=await importPptx(source,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,keywords:['value']});await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
 const f=metadataFixture();f.add('core','');f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');const guard=await importPptx(encodeMetadata(f),config);try{guard.document.setMetadata({...guard.document.metadata,keywords:['value']});const p=await guard.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));}finally{guard.dispose();}
});
