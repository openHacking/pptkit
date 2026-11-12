import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata} from './metadata-fixture.mjs';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'metadata-keywords-remove-v15'};

test('removes and reinserts a single keyword in both dialects across three reopens',async()=>{
 for(const strict of [false,true])for(const initial of ['','one, two']){const f=metadataFixture(strict);f.add('core',`<d:title>Title</d:title><!--sentinel--><c:keywords>${initial}</c:keywords><c:revision>7</c:revision>`);let bytes=encodeMetadata(f);
  for(let round=0;round<3;round++){const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{if(round!==1)assert.equal(s.report.coreMetadataSources.find(x=>x.field==='keywords').removable,true);s.document.setMetadata({...s.document.metadata,keywords:round===1?['中文 & one']:[]});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/core.data').toString(),/<!--sentinel-->/);}finally{s.dispose();}const r=await importPptx(bytes,config);try{assert.deepEqual(r.document.metadata.keywords??[],round===1?['中文 & one']:[]);assert.equal(r.document.metadata.revision,7);}finally{r.dispose();}}
 }
});

test('empty string is a value and keyword removal combines with other metadata edits',async()=>{
 const f=metadataFixture();f.add('core','<c:keywords>one</c:keywords><d:subject>Subject</d:subject><d:description>Old</d:description>');const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,keywords:['']});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');const r=await importPptx((await s.save(p.plan)).bytes,config);try{assert.deepEqual(r.document.metadata.keywords,['']);}finally{r.dispose();}s.document.setMetadata({...s.document.metadata,keywords:[],subject:undefined,description:'New',createdAt:'2020-01-01T00:00:00.000Z'});const mixed=await s.prepareSave({mode:'edit'});assert.equal(mixed.status,'ready',JSON.stringify(mixed.report));const again=await importPptx((await s.save(mixed.plan)).bytes,config);try{assert.deepEqual(again.document.metadata.keywords??[],[]);assert.equal(again.document.metadata.subject,undefined);assert.equal(again.document.metadata.description,'New');assert.equal(again.document.metadata.createdAt,'2020-01-01T00:00:00.000Z');}finally{again.dispose();}}finally{s.dispose();}
});

test('older profiles and language-tagged or segmented keyword sources cannot be removed',async()=>{
 const f=metadataFixture();f.add('core','<c:keywords>one</c:keywords>');const old=await importPptx(encodeMetadata(f),{projection:'content',metadataSaveProfile:'metadata-date-remove-v14'});try{assert.equal(old.report.coreMetadataSources.find(x=>x.field==='keywords').removable,false);old.document.setMetadata({...old.document.metadata,keywords:[]});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
 for(const xml of ['<c:keywords xml:lang="en">one</c:keywords>','<c:keywords><c:value>one</c:value></c:keywords>','<c:keywords>one<!--opaque-->two</c:keywords>']){const fixture=metadataFixture();fixture.add('core',xml);const s=await importPptx(encodeMetadata(fixture),config);try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='keywords').status,'unmapped');s.document.setMetadata({...s.document.metadata,keywords:[]});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}}
});

test('keyword removal enforces plan freshness, cancellation and work limits',async()=>{
 const f=metadataFixture();f.add('core','<c:keywords>one</c:keywords>');const source=encodeMetadata(f),s=await importPptx(source,config);try{s.document.setMetadata({...s.document.metadata,keywords:[]});const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:ac.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,keywords:['changed']});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}const limited=await importPptx(source,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,keywords:[]});await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});
