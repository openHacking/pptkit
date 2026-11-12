import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata,MC} from './metadata-fixture.mjs';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'metadata-date-remove-v14'};
const date=(name,value)=>`<t:${name} xsi:type="t:W3CDTF">${value}</t:${name}>`;

test('removes and reinserts both typed dates in two dialects through three reopens',async()=>{
 for(const strict of [false,true]){const f=metadataFixture(strict);f.add('core',date('created','2020-02-29T10:00:00.1200+08:00')+'<!--sentinel-->'+date('modified','2021-01-02T03:04:05Z')+'<d:subject>S</d:subject><c:revision>7</c:revision>');let bytes=encodeMetadata(f);
  for(let round=0;round<3;round++){const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{
   if(round!==1){assert.equal(s.report.coreMetadataSources.find(x=>x.field==='createdAt').removable,true);assert.equal(s.report.coreMetadataSources.find(x=>x.field==='modifiedAt').removable,true);}s.document.setMetadata({...s.document.metadata,createdAt:round===1?'2022-01-01T00:00:00.123Z':undefined,modifiedAt:round===1?'2023-01-01T00:00:00.000Z':undefined});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/core.data').toString(),/<!--sentinel--><d:subject>S<\/d:subject><c:revision>7<\/c:revision>/);
  }finally{s.dispose();}const r=await importPptx(bytes,config);try{assert.equal(r.document.metadata.createdAt,round===1?'2022-01-01T00:00:00.123Z':undefined);assert.equal(r.document.metadata.modifiedAt,round===1?'2023-01-01T00:00:00.000Z':undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}}
 }
});

test('date removal combines with string removal, text replacement and insertion',async()=>{
 const f=metadataFixture();f.add('core',date('created','2020-01-01T00:00:00Z')+'<d:subject>S</d:subject><d:description>D</d:description>');const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,createdAt:undefined,subject:undefined,description:'Changed',modifiedAt:'2024-01-01T00:00:00.000Z'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const r=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(r.document.metadata.createdAt,undefined);assert.equal(r.document.metadata.subject,undefined);assert.equal(r.document.metadata.description,'Changed');assert.equal(r.document.metadata.modifiedAt,'2024-01-01T00:00:00.000Z');}finally{r.dispose();}}finally{s.dispose();}
});

test('old profile, unrepresentable dates and compatibility markup cannot authorize removal',async()=>{
 const f=metadataFixture();f.add('core',date('created','2020-01-01T00:00:00Z'));const old=await importPptx(encodeMetadata(f),{projection:'content',metadataSaveProfile:'metadata-remove-v13'});try{assert.equal(old.report.coreMetadataSources.find(x=>x.field==='createdAt').removable,false);old.document.setMetadata({...old.document.metadata,createdAt:undefined});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
 const bad=metadataFixture();bad.add('core',date('created','2020-01-01T00:00:00.0001Z'));const s=await importPptx(encodeMetadata(bad),config);try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='createdAt').status,'unmapped');}finally{s.dispose();}
 const m=metadataFixture(),part=m.add('core',date('created','2020-01-01T00:00:00Z'));m.parts[part]=m.core(date('created','2020-01-01T00:00:00Z'),`xmlns:mc="${MC}" xmlns:x="urn:future" mc:Ignorable="x"`);await assert.rejects(importPptx(encodeMetadata(m),config),{code:'import-invalid-package'});
});

test('date removal rejects cancellation, stale plans and exhausted budgets',async()=>{
 const f=metadataFixture();f.add('core',date('created','2020-01-01T00:00:00Z'));const source=encodeMetadata(f),s=await importPptx(source,config);try{s.document.setMetadata({...s.document.metadata,createdAt:undefined});const c=new AbortController();c.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:c.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,createdAt:'2022-01-01T00:00:00.000Z'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}const limited=await importPptx(source,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,createdAt:undefined});await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});
