import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture,encodeMetadata,MC } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'extended-company-insert-v11'};
function fixture(strict=false,body='<Manager>Sentinel</Manager>'){const f=metadataFixture(strict);f.add('extended',body);return encodeMetadata(f);}

test('missing company inserts and remains editable across three rounds',async()=>{
 for(const strict of [false,true]){let bytes=fixture(strict);for(const [round,value] of ['', '中文 & text','final'].entries()){
  const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{if(round===0)assert.equal(s.report.coreMetadataSources.find(x=>x.field==='company').status,'insertable');s.document.setMetadata({...s.document.metadata,company:value});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/extended.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/extended.data').toString(),/<Manager>Sentinel<\/Manager>/);assert.match(after.get('inventory/extended.data').toString(),/<ep:Company xmlns:ep=/);}finally{s.dispose();}
  const reopened=await importPptx(bytes,config);try{assert.equal(reopened.document.metadata.company,value);assert.equal(reopened.report.coreMetadataSources.find(x=>x.field==='company').status,'bound');}finally{reopened.dispose();}
 }}
});

test('company insertion combines with new core and custom parts',async()=>{
 const s=await importPptx(fixture(),{...config,customPropertySaveProfile:'create-and-append-v5'});try{s.document.setMetadata({...s.document.metadata,company:'Company',title:'Title',keywords:['keyword'],customProperties:[{name:'A',type:'integer',value:2}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(reopened.document.metadata.company,'Company');assert.equal(reopened.document.metadata.title,'Title');assert.deepEqual(reopened.document.metadata.keywords,['keyword']);assert.equal(reopened.document.metadata.customProperties[0].value,2);}finally{reopened.dispose();}}finally{s.dispose();}
});

test('company insertion rejects missing part, ambiguous sources and older profiles',async()=>{
 const absent=await importPptx(encodeMetadata(metadataFixture()),config);try{assert.equal(absent.report.coreMetadataSources.find(x=>x.field==='company').status,'unmapped');absent.document.setMetadata({...absent.document.metadata,company:'New'});assert.equal((await absent.prepareSave({mode:'edit'})).status,'blocked');}finally{absent.dispose();}
 for(const body of ['<Company>old<!--opaque-->value</Company>']){const s=await importPptx(fixture(false,body),config);try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='company').status,'unmapped');s.document.setMetadata({...s.document.metadata,company:'New'});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}}
 const f=metadataFixture(),part=f.add('extended','<Manager>Sentinel</Manager>');f.parts[part]=f.properties('extended','<Manager>Sentinel</Manager>',`xmlns:mc="${MC}"`);const guard=await importPptx(encodeMetadata(f),config);try{assert.equal(guard.report.coreMetadataSources.find(x=>x.field==='company').status,'unmapped');guard.document.setMetadata({...guard.document.metadata,company:'New'});assert.equal((await guard.prepareSave({mode:'edit'})).status,'blocked');}finally{guard.dispose();}
 const old=await importPptx(fixture(),{projection:'content',metadataSaveProfile:'core-keywords-insert-v10'});try{old.document.setMetadata({...old.document.metadata,company:'New'});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
});


test('company insertion respects source guards, cancellation and plan freshness',async()=>{
 const source=fixture(),s=await importPptx(source,config);try{const no=await s.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await s.save(no.plan)).bytes),Buffer.from(source));s.document.setMetadata({...s.document.metadata,company:'New'});const ac=new AbortController();ac.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:ac.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,company:'Other'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}
 const f=metadataFixture();f.add('extended','');f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');const guard=await importPptx(encodeMetadata(f),config);try{guard.document.setMetadata({...guard.document.metadata,company:'New'});const p=await guard.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));}finally{guard.dispose();}
});

test('company insertion rejects invalid XML text and enforces analysis limits',async()=>{
 const source=fixture(),invalid=await importPptx(source,config);try{invalid.document.setMetadata({...invalid.document.metadata,company:'bad\u0000text'});const p=await invalid.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{invalid.dispose();}
 const limited=await importPptx(source,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,company:'Company'});await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});

test('company insertion preserves UTF-16 and handles a self-closing root',async()=>{
 for(const endian of ['le','be']){const f=metadataFixture(),part=f.add('extended',''),text=f.parts[part].replace('></Properties>','/>');let bytes=Buffer.from(text,'utf16le');if(endian==='be')for(let i=0;i<bytes.length;i+=2)[bytes[i],bytes[i+1]]=[bytes[i+1],bytes[i]];f.parts[part]=Buffer.concat([Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]),bytes]);const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,company:'中文 & Co'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const saved=(await s.save(p.plan)).bytes;assert.deepEqual(readGeneratedZip(saved).get(part).subarray(0,2),Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]));const reopened=await importPptx(saved,config);try{assert.equal(reopened.document.metadata.company,'中文 & Co');}finally{reopened.dispose();}}finally{s.dispose();}}
});
