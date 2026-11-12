import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata} from './metadata-fixture.mjs';
const profiles=['core-strings-v1','core-values-v2','core-keywords-v3','extended-company-v4','core-insert-v5','core-date-values-v6','core-date-insert-v7','core-create-v8','core-revision-insert-v9','core-keywords-insert-v10','extended-company-insert-v11','extended-create-v12','metadata-remove-v13','metadata-date-remove-v14','metadata-keywords-remove-v15','metadata-revision-remove-v16','metadata-default-remove-v17'];

test('all metadata profiles preserve their exact binding, insertion and removal boundaries',async()=>{
 for(const strict of [false,true]){
  const present=metadataFixture(strict);present.add('core','<d:title>T</d:title><d:creator>A</d:creator><d:subject>S</d:subject><d:description>D</d:description><d:language>en-US</d:language><c:revision>7</c:revision><c:keywords>K</c:keywords><t:created xsi:type="t:W3CDTF">2020-01-01T00:00:00Z</t:created><t:modified xsi:type="t:W3CDTF">2021-01-01T00:00:00Z</t:modified>');present.add('extended','<Company>C</Company>');const empty=metadataFixture(strict);empty.add('core','');empty.add('extended','');const missing=metadataFixture(strict);
  for(const [index,profile] of profiles.entries()){const level=index+1;for(const [kind,bytes] of [['present',encodeMetadata(present)],['empty',encodeMetadata(empty)],['missing',encodeMetadata(missing)]]){
   const s=await importPptx(bytes,{projection:'content',metadataSaveProfile:profile});try{const by=new Map(s.report.coreMetadataSources.map(x=>[x.field,x]));
    for(const [field,min] of [['title',1],['author',1],['subject',1],['description',1],['language',1],['revision',2],['keywords',3],['company',4],['createdAt',6],['modifiedAt',6]]){
     if(level<min){assert.equal(by.has(field),false,`${profile} ${kind} ${field}`);continue;}
     const item=by.get(field);assert.ok(item,`${profile} ${kind} ${field}`);
     const insertAt=field==='revision'?9:field==='keywords'?10:field==='company'?11:field==='createdAt'||field==='modifiedAt'?7:5;
     const expected=kind==='present'?'bound':kind==='empty'?(level>=insertAt?'insertable':'unmapped'):(field==='company'?level>=12:level>=Math.max(8,insertAt))?'insertable':'unmapped';assert.equal(item.status,expected,`${profile} ${kind} ${field}`);
     if(kind==='present'&&level>=13){const removeAt=field==='title'||field==='author'||field==='language'?17:field==='revision'?16:field==='keywords'?15:field==='createdAt'||field==='modifiedAt'?14:13;assert.equal(item.removable,level>=removeAt,`${profile} ${field} removal`);}else assert.equal(item.removable,undefined,`${profile} ${kind} ${field} report shape`);
    }
   }finally{s.dispose();}
  }}
 }
});

test('unknown metadata profile cannot gain authority from a prototype key or malformed value',async()=>{
 const bytes=encodeMetadata(metadataFixture()),defaults=await importPptx(bytes,{projection:'content'});try{assert.equal(defaults.report.coreMetadataSources,undefined);defaults.document.setMetadata({...defaults.document.metadata,company:'New'});assert.equal((await defaults.prepareSave({mode:'edit'})).status,'blocked');}finally{defaults.dispose();}for(const profile of ['__proto__','constructor','metadata-default-remove-v18','',17,null])await assert.rejects(importPptx(bytes,{projection:'content',metadataSaveProfile:profile}),TypeError);
});
