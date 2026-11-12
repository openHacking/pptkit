import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata} from './metadata-fixture.mjs';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'metadata-revision-remove-v16'};

test('explicit removal distinguishes revision one from an absent source across three reopens',async()=>{
 for(const strict of [false,true])for(const initial of ['1','7']){const f=metadataFixture(strict);f.add('core',`<d:subject>S</d:subject><!--sentinel--><c:revision>${initial}</c:revision>`);let bytes=encodeMetadata(f);
  for(let round=0;round<3;round++){const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{if(round!==1)assert.equal(s.report.coreMetadataSources.find(x=>x.field==='revision').removable,true);s.document.setMetadata({...s.document.metadata,revision:round===1?3:undefined});const p=await s.prepareSave(round===1?{mode:'edit'}:{mode:'edit',coreMetadataRemovals:['revision']});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/core.data').toString(),/<!--sentinel-->/);}finally{s.dispose();}const r=await importPptx(bytes,config);try{assert.equal(r.document.metadata.revision,round===1?3:undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}}
 }
});

test('revision removal combines with other metadata operations and rejects contradictory edits',async()=>{
 const f=metadataFixture();f.add('core','<c:revision>7</c:revision><d:subject>S</d:subject><c:keywords>one</c:keywords>');const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,revision:2,subject:undefined,keywords:[]});let p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['revision']});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);s.document.setMetadata({...s.document.metadata,revision:undefined,createdAt:'2020-01-01T00:00:00.000Z'});p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['revision']});assert.equal(p.status,'ready',JSON.stringify(p.report));const r=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(r.document.metadata.revision,undefined);assert.equal(r.document.metadata.subject,undefined);assert.deepEqual(r.document.metadata.keywords??[],[]);assert.equal(r.document.metadata.createdAt,'2020-01-01T00:00:00.000Z');}finally{r.dispose();}}finally{s.dispose();}
});

test('removal rejects missing, ambiguous and noncanonical sources and older profiles',async()=>{
 for(const body of ['', '<c:revision>01</c:revision>', '<c:revision>1<!--unknown--></c:revision>']){const f=metadataFixture();f.add('core',body);const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,revision:undefined});const p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['revision']});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}}
 const f=metadataFixture();f.add('core','<c:revision>1</c:revision>');const old=await importPptx(encodeMetadata(f),{projection:'content',metadataSaveProfile:'metadata-keywords-remove-v15'});try{await assert.rejects(old.prepareSave({mode:'edit',coreMetadataRemovals:['revision']}),TypeError);}finally{old.dispose();}
});

test('removal validates explicit selections, cancellation, stale plans and limits',async()=>{
 const f=metadataFixture();f.add('core','<c:revision>1</c:revision>');const bytes=encodeMetadata(f),s=await importPptx(bytes,config);try{await assert.rejects(s.prepareSave({mode:'no-op',coreMetadataRemovals:['revision']}),TypeError);await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['revision','revision']}),TypeError);await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['revision'],coreMetadataInsertions:['revision']}),TypeError);const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['revision'],signal:ac.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['revision']});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,revision:2});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}
 const limited=await importPptx(bytes,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{await assert.rejects(limited.prepareSave({mode:'edit',coreMetadataRemovals:['revision']}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});
