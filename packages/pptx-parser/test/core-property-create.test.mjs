import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { unzipSync, zipSync } from 'fflate';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'core-create-v8'};

test('create missing core part, preserve unrelated parts, and edit across three reopens',async()=>{
 for(const strict of [false,true]){
  let bytes=encodeMetadata(metadataFixture(strict));
  for(let round=0;round<3;round++){
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);
   try {if(round===0){for(const field of ['title','createdAt'])assert.equal(s.report.coreMetadataSources.find(x=>x.field===field).status,'insertable');}
    s.document.setMetadata({...s.document.metadata,title:`Title ${round}`,createdAt:`2020-01-0${round+1}T00:00:00.000Z`});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);
    for(const [name,payload] of before)if(!['_rels/.rels','[Content_Types].xml','pptkit-core-properties-1.xml'].includes(name))assert.deepEqual(after.get(name),payload,name);
   }finally{s.dispose();}
   const reopened=await importPptx(bytes,config);try{assert.equal(reopened.document.metadata.title,`Title ${round}`);assert.equal(reopened.document.metadata.createdAt,`2020-01-0${round+1}T00:00:00.000Z`);assert.equal(reopened.report.inspection.metadata.core.part,'/pptkit-core-properties-1.xml');}finally{reopened.dispose();}
  }
 }
});

test('simultaneous core and custom creation share one declaration transaction',async()=>{
 for(const strict of [false,true]){
  const source=encodeMetadata(metadataFixture(strict)),s=await importPptx(source,{...config,customPropertySaveProfile:'create-and-append-v5'});
  try{s.document.setMetadata({...s.document.metadata,title:'New',createdAt:'2020-01-01T00:00:00.000Z',customProperties:[{name:'A',type:'integer',value:3}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});try{assert.equal(reopened.document.metadata.title,'New');assert.equal(reopened.document.metadata.createdAt,'2020-01-01T00:00:00.000Z');assert.equal(reopened.document.metadata.customProperties[0].value,3);}finally{reopened.dispose();}}finally{s.dispose();}
 }
});

test('core creation supports explicit defaults and rejects stale plans',async()=>{
 const s=await importPptx(encodeMetadata(metadataFixture()),config);try{const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['title']});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(reopened.document.metadata.title,'PPTKit Presentation');}finally{reopened.dispose();}s.document.setMetadata({...s.document.metadata,title:'changed'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}
});


test('core creation handles no-op, cancellation, name collisions and source-policy guards',async()=>{
 const f=metadataFixture();f.parts['PPTKIT-CORE-PROPERTIES-1.XML']='<extra/>';f.metadataTypes['PPTKIT-CORE-PROPERTIES-1.XML']='application/xml';f.rootRelations.push({Id:'pptkitCoreProperties2',Type:'urn:test',Target:'deck/main.data'});const source=encodeMetadata(f),s=await importPptx(source,config);
 try{assert.equal(s.report.coreMetadataSources.find(x=>x.field==='title').source.part,'/pptkit-core-properties-3.xml');const noop=await s.prepareSave({mode:'no-op'});assert.equal(noop.status,'ready');assert.deepEqual(Buffer.from((await s.save(noop.plan)).bytes),Buffer.from(source));s.document.setMetadata({...s.document.metadata,title:'New'});const abort=new AbortController();abort.abort();await assert.rejects(s.prepareSave({mode:'edit',signal:abort.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await s.save(p.plan)).bytes,config);try{assert.equal(reopened.report.inspection.metadata.core.part,'/pptkit-core-properties-3.xml');}finally{reopened.dispose();}}finally{s.dispose();}
 const {sourceAlternates}=await import('./closure-fixture.mjs');const blocked=metadataFixture();blocked.parts['deck/_rels/z.xml.rels']=sourceAlternates('<Relationship Id="hidden" Type="urn:asset" Target="../pptkit-core-properties-1.xml"/>',`<Relationship Id="layout" Type="${blocked.R}/slideLayout" Target="layout.xml"/>`);const b=await importPptx(encodeMetadata(blocked),config);try{assert.equal(b.report.coreMetadataSources.find(x=>x.field==='title').status,'unmapped');b.document.setMetadata({...b.document.metadata,title:'New'});assert.equal((await b.prepareSave({mode:'edit'})).status,'blocked');}finally{b.dispose();}
});

test('invalid combined core creation and custom insertion produce no partial plan',async()=>{
 const s=await importPptx(encodeMetadata(metadataFixture()),{...config,customPropertySaveProfile:'create-and-append-v5'});try{s.document.setMetadata({...s.document.metadata,title:'New',customProperties:[{name:'A',type:'string',value:'_x005F_x0041_'}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}
});


test('core and custom creation preserve pre-existing OPC namespace prefixes',async()=>{
 const rel='http://schemas.openxmlformats.org/package/2006/relationships',ct='http://schemas.openxmlformats.org/package/2006/content-types';
 const entries=unzipSync(encodeMetadata(metadataFixture()));
 function prefix(name,namespace,p,names){let text=Buffer.from(entries[name]).toString();assert.equal(text.split(`xmlns="${namespace}"`).length,2);text=text.replace(`xmlns="${namespace}"`,`xmlns:${p}="${namespace}"`);for(const n of names)text=text.replaceAll(`<${n}`,`<${p}:${n}`).replaceAll(`</${n}>`,`</${p}:${n}>`);entries[name]=Buffer.from(text);return text;}
 const originalRel=prefix('_rels/.rels',rel,'r',['Relationships','Relationship']),originalTypes=prefix('[Content_Types].xml',ct,'t',['Types','Default','Override']);
 const s=await importPptx(zipSync(entries),{...config,customPropertySaveProfile:'create-and-append-v5'});try{
  s.document.setMetadata({...s.document.metadata,title:'New',customProperties:[{name:'A',type:'integer',value:3}]});const p=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(p.status,'ready',JSON.stringify(p.report));const bytes=(await s.save(p.plan)).bytes,after=readGeneratedZip(bytes),newRel=after.get('_rels/.rels').toString(),newTypes=after.get('[Content_Types].xml').toString();
  assert.ok(newRel.startsWith(originalRel.slice(0,originalRel.indexOf('</r:Relationships>'))));assert.ok(newTypes.startsWith(originalTypes.slice(0,originalTypes.indexOf('</t:Types>'))));assert.match(newRel,/<r:Relationship Id="pptkitCoreProperties1"/);assert.match(newRel,/<r:Relationship Id="pptkitCustomProperties1"/);assert.match(newTypes,/<t:Override PartName="\/pptkit-core-properties-1.xml"/);assert.match(newTypes,/<t:Override PartName="\/pptkit-custom-properties-1.xml"/);
  const reopened=await importPptx(bytes,{projection:'content'});try{assert.equal(reopened.document.metadata.title,'New');assert.equal(reopened.document.metadata.customProperties[0].value,3);}finally{reopened.dispose();}
 }finally{s.dispose();}
});
