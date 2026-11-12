import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',metadataSaveProfile:'core-date-values-v6'};
const date=(name,value)=>`<t:${name} xsi:type="t:W3CDTF">${value}</t:${name}>`;
function fixture(strict=false,created='2020-02-29T10:00:00.1200+08:00',modified='2021-01-02T03:04:05Z'){const f=metadataFixture(strict);f.add('core',date('created',created)+date('modified',modified)+'<c:revision>7</c:revision>');return encodeMetadata(f);}

test('existing dates project to UTC and save with source type and neighboring bytes preserved over three rounds',async()=>{
 for(const strict of [false,true]) {
  let bytes=fixture(strict);
  for(let round=0;round<3;round++) {
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,options);
   try {if(round===0){assert.equal(s.document.metadata.createdAt,'2020-02-29T02:00:00.120Z');assert.equal(s.report.coreMetadataSources.find(x=>x.field==='createdAt').status,'bound');}const createdAt=`2022-01-0${round+1}T00:00:00.000Z`;s.document.setMetadata({...s.document.metadata,createdAt});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);const xml=after.get('inventory/core.data').toString();assert.match(xml,/<t:created xsi:type="t:W3CDTF">/);assert.match(xml,/<t:modified xsi:type="t:W3CDTF">2021-01-02T03:04:05Z<\/t:modified>/);}finally{s.dispose();}
   const reopened=await importPptx(bytes,options);try{assert.equal(reopened.document.metadata.createdAt,`2022-01-0${round+1}T00:00:00.000Z`);assert.equal(reopened.document.metadata.modifiedAt,'2021-01-02T03:04:05.000Z');}finally{reopened.dispose();}
  }
 }
});

test('unsupported date lexical forms remain source-only and missing date insertion is not admitted',async()=>{
 for(const value of ['2020-01-01T00:00:00','2020-01-01T00:00:00.0001Z','2020-01-01T00:00:00-00:00','2021-02-29T00:00:00Z']) {const s=await importPptx(fixture(false,value),options);try{assert.equal(s.document.metadata.createdAt,undefined,value);assert.equal(s.report.coreMetadataSources.find(x=>x.field==='createdAt').status,'unmapped');s.document.setMetadata({...s.document.metadata,createdAt:'2023-01-01T00:00:00.000Z'});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}}
 const f=metadataFixture();f.add('core','<c:revision>1</c:revision>');const s=await importPptx(encodeMetadata(f),options);try{s.document.setMetadata({...s.document.metadata,createdAt:'2023-01-01T00:00:00.000Z'});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}
});


test('date edits enforce profile, mixed edits, signature guards and snapshot freshness',async()=>{
 const source=fixture(), old=await importPptx(source,{projection:'content',metadataSaveProfile:'core-insert-v5'});
 try {old.document.setMetadata({...old.document.metadata,createdAt:'2023-01-01T00:00:00.000Z'});assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');}finally{old.dispose();}
 const s=await importPptx(source,options);try{
  const noop=await s.prepareSave({mode:'no-op'});assert.equal(noop.status,'ready');assert.deepEqual(Buffer.from((await s.save(noop.plan)).bytes),Buffer.from(source));
  s.document.setMetadata({...s.document.metadata,createdAt:'2023-01-01T00:00:00.000Z',revision:8,title:'New'});
  const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,options);try{assert.equal(reopened.document.metadata.createdAt,'2023-01-01T00:00:00.000Z');assert.equal(reopened.document.metadata.revision,8);assert.equal(reopened.document.metadata.title,'New');}finally{reopened.dispose();}
  s.document.setMetadata({...s.document.metadata,createdAt:'2024-01-01T00:00:00.000Z'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});
 }finally{s.dispose();}
 const signed=metadataFixture();signed.add('core',date('created','2020-01-01T00:00:00Z'));signed.parts['deck/_rels/z.xml.rels']=signed.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
 const guard=await importPptx(encodeMetadata(signed),options);try{guard.document.setMetadata({...guard.document.metadata,createdAt:'2023-01-01T00:00:00.000Z'});const p=await guard.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));}finally{guard.dispose();}
});


test('v7 inserts missing dates with explicit W3CDTF type and keeps them editable over three rounds',async()=>{
 for(const strict of [false,true]){
  const f=metadataFixture(strict);f.add('core','<c:revision>7</c:revision>');let bytes=encodeMetadata(f),config={projection:'content',metadataSaveProfile:'core-date-insert-v7'};
  for(let round=0;round<3;round++){
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);
   try {if(round===0)for(const field of ['createdAt','modifiedAt'])assert.equal(s.report.coreMetadataSources.find(x=>x.field===field).status,'insertable');
    s.document.setMetadata({...s.document.metadata,createdAt:`2022-01-0${round+1}T00:00:00.000Z`,modifiedAt:'2021-01-02T03:04:05.123Z',title:`title ${round}`});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);
    for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);
    const xml=after.get('inventory/core.data').toString();assert.match(xml,/<dct:created xmlns:dct="http:\/\/purl.org\/dc\/terms\/" xmlns:xsi="http:\/\/www.w3.org\/2001\/XMLSchema-instance" xsi:type="dct:W3CDTF">/);assert.match(xml,/<dct:modified /);
   }finally{s.dispose();}
   const reopened=await importPptx(bytes,config);try{assert.equal(reopened.document.metadata.createdAt,`2022-01-0${round+1}T00:00:00.000Z`);assert.equal(reopened.document.metadata.modifiedAt,'2021-01-02T03:04:05.123Z');assert.equal(reopened.report.coreMetadataSources.find(x=>x.field==='createdAt').status,'bound');}finally{reopened.dispose();}
  }
 }
});

test('v7 rejects absent core part, unrepresentable existing dates and protected source changes',async()=>{
 const config={projection:'content',metadataSaveProfile:'core-date-insert-v7'}, missing=await importPptx(encodeMetadata(metadataFixture()),config);
 try{missing.document.setMetadata({...missing.document.metadata,createdAt:'2020-01-01T00:00:00.000Z'});assert.equal((await missing.prepareSave({mode:'edit'})).status,'blocked');}finally{missing.dispose();}
 const existing=await importPptx(fixture(false,'2020-01-01T00:00:00.0001Z'),config);try{assert.equal(existing.report.coreMetadataSources.find(x=>x.field==='createdAt').status,'unmapped');existing.document.setMetadata({...existing.document.metadata,createdAt:'2020-01-01T00:00:00.000Z'});assert.equal((await existing.prepareSave({mode:'edit'})).status,'blocked');}finally{existing.dispose();}
 const f=metadataFixture();f.add('core','');f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
 const guard=await importPptx(encodeMetadata(f),config);try{guard.document.setMetadata({...guard.document.metadata,createdAt:'2020-01-01T00:00:00.000Z'});const p=await guard.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));}finally{guard.dispose();}
});


test('date insertion handles self-closing root and UTF-16 source encoding',async()=>{
 for(const endian of ['le','be']){
  const f=metadataFixture(),part=f.add('core',''),xml=f.parts[part].replace('></c:coreProperties>','/>');let raw=Buffer.from(xml,'utf16le');if(endian==='be'){for(let i=0;i<raw.length;i+=2)[raw[i],raw[i+1]]=[raw[i+1],raw[i]];}f.parts[part]=Buffer.concat([Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]),raw]);
  const s=await importPptx(encodeMetadata(f),{projection:'content',metadataSaveProfile:'core-date-insert-v7'});try{s.document.setMetadata({...s.document.metadata,createdAt:'2020-01-01T00:00:00.000Z'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const bytes=(await s.save(p.plan)).bytes,partBytes=readGeneratedZip(bytes).get(part);assert.deepEqual(partBytes.subarray(0,2),Buffer.from(endian==='le'?[0xff,0xfe]:[0xfe,0xff]));const reopened=await importPptx(bytes,{projection:'content'});try{assert.equal(reopened.document.metadata.createdAt,'2020-01-01T00:00:00.000Z');}finally{reopened.dispose();}}finally{s.dispose();}
 }
});
