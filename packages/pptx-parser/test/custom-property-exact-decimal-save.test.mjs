import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom, MC } from './metadata-fixture.mjs';
import { shape, body } from './content-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { importPptxFile } from '../dist/node.js';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
const options={projection:'content',customDecimalProjection:'exact-save-v2',customPropertySaveProfile:'scalar-v1'};
function fixture(strict=false){const f=metadataFixture(strict);f.add('custom',custom(2,'decimal',' +0001.2500 ','name="Decimal"')+custom(3,'vector','<v:lpwstr>x</v:lpwstr>','name="Opaque"')+custom(4,'lpwstr','before','name="Control"'));return encodeMetadata(f);}
for(const strict of [false,true])test(`exact decimal source edit preserves variant and sibling parts across three rounds (strict=${strict})`,async()=>{
 let bytes=fixture(strict);for(const value of ['0.1','9007199254740993.0001','-0.0000000000000000001']){
  const before=readGeneratedZip(bytes),session=await importPptx(bytes,options);try{assert.equal(session.report.customPropertySources[0].status,'bound');assert.equal(session.report.customPropertySources[0].projectedIndex,0);const no=await session.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes),Buffer.from(bytes));
  session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Decimal'?{...p,value}:p.name==='Control'?{...p,value:'after'}:p)});const ready=await session.prepareSave({mode:'edit'});assert.equal(ready.status,'ready',JSON.stringify(ready.report));bytes=(await session.save(ready.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload]of before)if(name!=='inventory/custom.data')assert.deepEqual(after.get(name),payload,name);const xml=after.get('inventory/custom.data').toString();assert.ok(xml.includes(`<v:decimal>${value}</v:decimal>`));assert.match(xml,/<v:vector><v:lpwstr>x<\/v:lpwstr><\/v:vector>/);
  }finally{session.dispose();}const reopened=await importPptx(bytes,options);try{assert.equal(reopened.document.metadata.customProperties[0].value,value);assert.equal(reopened.report.inspection.metadata.custom.properties[0].valueType,'decimal');assert.equal(reopened.report.inspection.metadata.custom.properties[0].id,2);}finally{reopened.dispose();}
 }
});
test('exact decimal edits reject invalid values and stale plans',async()=>{
 const session=await importPptx(fixture(),options);try{const original=session.document.metadata;for(const value of ['01','1.0','-0','1e2','0'.repeat(10001)])assert.throws(()=>session.document.setMetadata({...original,customProperties:original.customProperties.map(p=>p.name==='Decimal'?{...p,value}:p)}),TypeError);
 session.document.setMetadata({...original,customProperties:original.customProperties.map(p=>p.name==='Decimal'?{...p,value:'0.1'}:p)});const ready=await session.prepareSave({mode:'edit'});assert.equal(ready.status,'ready');session.document.setMetadata({...original,customProperties:original.customProperties.map(p=>p.name==='Decimal'?{...p,value:'0.2'}:p)});await assert.rejects(session.save(ready.plan),{code:'save-snapshot-stale'});
 }finally{session.dispose();}
});
test('exact decimal source save refuses ambiguous raw structures and incompatible profiles',async()=>{
 const f=metadataFixture(),part=f.add('custom',custom(2,'decimal','1.25','name="Value"'));f.parts[part]=f.properties('custom',custom(2,'decimal','1.25','name="Value"'),`xmlns:mc="${MC}"`);const bytes=encodeMetadata(f),session=await importPptx(bytes,options);try{assert.equal(session.report.customPropertySources[0].status,'unmapped');session.document.setMetadata({...session.document.metadata,customProperties:[{name:'Value',type:'decimal',value:'0.1'}]});assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');}finally{session.dispose();}
 for(const bad of [{projection:'content',customDecimalProjection:'exact-save-v2'},{...options,customPropertySaveProfile:'scalar-and-name-v2'},{...options,customCurrencyProjection:'exact-save-v2'},{...options,diagramPreview:'cache'}])await assert.rejects(importPptx(bytes,bad),TypeError);
});

test('unmodified decimal keeps its source spelling when a sibling changes',async()=>{
 const bytes=fixture(),session=await importPptx(bytes,options);try{session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Control'?{...p,value:'changed'}:p)});const ready=await session.prepareSave({mode:'edit'});assert.equal(ready.status,'ready',JSON.stringify(ready.report));const output=(await session.save(ready.plan)).bytes;assert.ok(readGeneratedZip(output).get('inventory/custom.data').toString().includes('<v:decimal> +0001.2500 </v:decimal>'));}finally{session.dispose();}
});
test('file import captures exact decimal save options before asynchronous IO',async()=>{
 const dir=await mkdtemp(path.join(tmpdir(),'pptkit-decimal-save-'));try{const file=path.join(dir,'source.pptx');await writeFile(file,fixture());const opts={...options},pending=importPptxFile(file,opts);opts.customDecimalProjection=undefined;opts.customPropertySaveProfile=undefined;const {session}=await pending;try{session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Decimal'?{...p,value:'0.1'}:p)});assert.equal((await session.prepareSave({mode:'edit'})).status,'ready');}finally{session.dispose();}}finally{await rm(dir,{recursive:true,force:true});}
});

for(const strict of [false,true])test(`exact decimal edits combine atomically with text, metadata and visibility (strict=${strict})`,async()=>{
 const f=metadataFixture(strict);f.add('core','<d:title>Original</d:title>');f.add('custom',custom(2,'decimal','0.1','name="Decimal"'));f.parts['deck/z.xml']=f.p('sld',`<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({id:'2',txBox:true,text:body()})}</p:spTree></p:cSld>`,`xmlns:a="${f.A}" show="0"`);let bytes=encodeMetadata(f);const mixed={...options,metadataSaveProfile:'core-strings-v1',slideVisibilitySaveProfile:'attribute-v1',editSaveProfile:'text-runs-v3'};
 for(const [round,value] of ['1.0001','9007199254740993.0001','-0.0001'].entries()){const session=await importPptx(bytes,mixed);try{const slide=session.document.slides[0],original=slide.elements[0],changed=structuredClone(original);changed.content[0].runs[0].text+='!';slide.replaceElement(original.id,changed);slide.setHidden(!slide.hidden);session.document.setMetadata({...session.document.metadata,title:`Title ${round}`,customProperties:[{name:'Decimal',type:'decimal',value}]});const invalid=structuredClone(changed);invalid.box.x++;slide.replaceElement(original.id,invalid);const blocked=await session.prepareSave({mode:'edit'});assert.equal(blocked.status,'blocked');assert.equal(blocked.plan,undefined);slide.replaceElement(original.id,changed);const ready=await session.prepareSave({mode:'edit'});assert.equal(ready.status,'ready',JSON.stringify(ready.report));bytes=(await session.save(ready.plan)).bytes;}finally{session.dispose();}const reopened=await importPptx(bytes,mixed);try{assert.equal(reopened.document.metadata.title,`Title ${round}`);assert.equal(reopened.document.metadata.customProperties[0].value,value);assert.equal(reopened.document.slides[0].hidden,round%2===1);assert.ok(reopened.document.slides[0].elements[0].content[0].runs[0].text.endsWith('!'.repeat(round+1)));}finally{reopened.dispose();}}
});

test('oversized source-only decimals remain untouched during other property edits',async()=>{
 const value='1'.repeat(10001),f=metadataFixture();f.add('custom',custom(2,'decimal',value,'name="Large"')+custom(3,'lpwstr','old','name="Control"'));const bytes=encodeMetadata(f),session=await importPptx(bytes,{...options,xmlLimits:{maxPendingCharacters:20000}});try{assert.deepEqual(session.document.metadata.customProperties,[{name:'Control',type:'string',value:'old'}]);session.document.setMetadata({...session.document.metadata,customProperties:[{name:'Control',type:'string',value:'new'}]});const ready=await session.prepareSave({mode:'edit'});assert.equal(ready.status,'ready',JSON.stringify(ready.report));const output=(await session.save(ready.plan)).bytes;assert.ok(readGeneratedZip(output).get('inventory/custom.data').toString().includes(`<v:decimal>${value}</v:decimal>`));}finally{session.dispose();}
});
