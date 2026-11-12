import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata,MC} from './metadata-fixture.mjs';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'metadata-remove-v13'};

test('removes optional direct properties and retains siblings over three reimports',async()=>{
 for(const strict of [false,true]){
  const f=metadataFixture(strict);f.add('core','<d:subject>S</d:subject><!--sentinel--><d:description>D</d:description><c:revision>7</c:revision>');f.add('extended','<Manager>M</Manager><Company>C</Company><Application>A</Application>');let bytes=encodeMetadata(f);
  for(let round=0;round<3;round++){
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{
    if(round!==1){for(const field of ['subject','description','company'])assert.equal(s.report.coreMetadataSources.find(x=>x.field===field).removable,true);assert.equal(s.report.coreMetadataSources.find(x=>x.field==='revision').removable,false);}
    const metadata={...s.document.metadata,subject:round===1?'New':undefined,description:round===1?'New description':undefined,company:round===1?'New company':undefined};s.document.setMetadata(metadata);
    const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);
    for(const [name,payload] of before)if(!['inventory/core.data','inventory/extended.data'].includes(name))assert.deepEqual(after.get(name),payload,name);
    assert.match(after.get('inventory/core.data').toString(),/<!--sentinel-->/);assert.match(after.get('inventory/extended.data').toString(),/<Manager>M<\/Manager><Application>A<\/Application>/);
   }finally{s.dispose();}
   const reopened=await importPptx(bytes,config);try{assert.equal(reopened.document.metadata.subject,round===1?'New':undefined);assert.equal(reopened.document.metadata.description,round===1?'New description':undefined);assert.equal(reopened.document.metadata.company,round===1?'New company':undefined);assert.equal(reopened.document.metadata.revision,7);}finally{reopened.dispose();}
  }
 }
});

test('removal combines atomically with a neighboring insertion and value edit',async()=>{
 for(const strict of [false,true]){const f=metadataFixture(strict);f.add('core','<d:subject>S</d:subject><d:description>D</d:description>');f.add('extended','<Company>C</Company>');const s=await importPptx(encodeMetadata(f),config);try{
  s.document.setMetadata({...s.document.metadata,subject:undefined,description:'edited',company:undefined,createdAt:'2020-01-01T00:00:00.000Z'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const r=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(r.document.metadata.subject,undefined);assert.equal(r.document.metadata.description,'edited');assert.equal(r.document.metadata.company,undefined);assert.equal(r.document.metadata.createdAt,'2020-01-01T00:00:00.000Z');}finally{r.dispose();}
 }finally{s.dispose();}}
});

test('older profiles and sources with unknown attributes, mixed content or MCE block removal',async()=>{
 for(const [profile,body] of [['extended-create-v12','<Company>C</Company>'],['metadata-remove-v13','<Company>C<!--opaque--></Company>']]){const f=metadataFixture();f.add('extended',body);const s=await importPptx(encodeMetadata(f),{projection:'content',metadataSaveProfile:profile});try{s.document.setMetadata({...s.document.metadata,company:undefined});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}}
 const invalid=metadataFixture();invalid.add('extended','<Company custom="opaque">C</Company>');await assert.rejects(importPptx(encodeMetadata(invalid),config),{code:'import-invalid-package'});
 const f=metadataFixture(),part=f.add('extended','<Company>C</Company>');f.parts[part]=f.properties('extended','<Company>C</Company>',`xmlns:mc="${MC}"`);const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,company:undefined});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}
});

test('removal respects cancellation, stale plans and source budgets',async()=>{
 const f=metadataFixture();f.add('extended','<Company>C</Company>');const source=encodeMetadata(f),s=await importPptx(source,config);try{s.document.setMetadata({...s.document.metadata,company:undefined});const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:ac.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,company:'changed'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}
 const limited=await importPptx(source,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,company:undefined});await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});

test('removal preserves UTF-16 byte order and non-target source characters',async()=>{
 for(const endian of ['le','be']){const f=metadataFixture(),part=f.add('extended','<Manager>中文</Manager><Company>C</Company><Application>Keep</Application>');const text=f.parts[part];let encoded=Buffer.from(text,'utf16le');if(endian==='be')for(let i=0;i<encoded.length;i+=2)[encoded[i],encoded[i+1]]=[encoded[i+1],encoded[i]];f.parts[part]=Buffer.concat([Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]),encoded]);const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,company:undefined});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const bytes=(await s.save(p.plan)).bytes,partBytes=readGeneratedZip(bytes).get(part);assert.deepEqual(partBytes.subarray(0,2),Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]));const actual=Buffer.from(partBytes.subarray(2));if(endian==='be')for(let i=0;i<actual.length;i+=2)[actual[i],actual[i+1]]=[actual[i+1],actual[i]];assert.equal(actual.toString('utf16le'),text.replace('<Company>C</Company>',''));}finally{s.dispose();}}
});
