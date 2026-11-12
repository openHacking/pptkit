import test from 'node:test';
import assert from 'node:assert/strict';
import { planPropertyCreation } from '../dist/preservation/plan-property-creation.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
const plan=(bytes,dialect='transitional')=>planPropertyCreation(bytes,'core',dialect,{},1024*1024,128);

test('core creation planning avoids occupied part names and relationship IDs in both dialects',async()=>{
 for(const strict of [false,true]){
  const f=metadataFixture(strict);f.parts['pptkit-core-properties-1.xml']='<unused/>';f.rootRelations.push({Id:'pptkitCoreProperties2',Type:'urn:example',Target:'deck/main.data'});const p=await plan(encodeMetadata(f),strict?'strict':'transitional');
  assert.equal(p.part,'/pptkit-core-properties-3.xml');assert.equal(p.relationshipId,'pptkitCoreProperties3');assert.match(Buffer.from(p.relationships).toString(),/metadata\/core-properties/);assert.match(Buffer.from(p.contentTypes).toString(),/application\/vnd.openxmlformats-package.core-properties\+xml/);
 }
});

test('core creation planning rejects existing declarations and dangling source relationships',async()=>{
 const existing=metadataFixture();existing.add('core','');await assert.rejects(plan(encodeMetadata(existing)),{code:'xml-edit-conflict'});
 const dangling=metadataFixture();dangling.rootRelations.push({Id:'rMissing',Type:'urn:example',Target:'missing.xml'});await assert.rejects(plan(encodeMetadata(dangling)),{code:'package-invalid'});
});
