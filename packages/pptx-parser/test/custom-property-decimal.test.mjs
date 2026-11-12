import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',customPropertySaveProfile:'scalar-v1'};
function fixture(strict=false,value='1.25') {const f=metadataFixture(strict);f.add('custom',custom(2,'decimal',value,'name="Decimal"')+custom(3,'r8','0.1','name="Control"'));return encodeMetadata(f);}

test('exact decimal values retain original type, identity and neighboring parts across three rounds',async()=>{
 for(const strict of [false,true]) {
  let bytes=fixture(strict);
  for(const value of [0.5,-123.75,9007199254740992]) {
   const before=readGeneratedZip(bytes),s=await importPptx(bytes,options);
   try {s.document.setMetadata({...s.document.metadata,customProperties:s.document.metadata.customProperties.map(p=>p.name==='Decimal'?{...p,value}:p)});const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await s.save(plan.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,data] of before) if(name!=='inventory/custom.data')assert.deepEqual(after.get(name),data,name);assert.match(after.get('inventory/custom.data').toString(),/<v:decimal>/);} finally {s.dispose();}
   const reopened=await importPptx(bytes,options);try {assert.equal(reopened.document.metadata.customProperties[0].value,value);assert.equal(reopened.report.inspection.metadata.custom.properties[0].valueType,'decimal');assert.equal(reopened.report.inspection.metadata.custom.properties[0].id,2);}finally{reopened.dispose();}
  }
 }
});

test('inexact and noncanonical decimals stay source-only and unsafe edits are rejected',async()=>{
 for(const value of ['0.1','0.2','1.0','+1','01','-0','0.000001','1.0000000000000001','9007199254740993']) {
  const s=await importPptx(fixture(false,value),options);try{assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name),['Control'],value);assert.equal(s.report.inspection.metadata.custom.properties[0].value.value,value);}finally{s.dispose();}
 }
 const s=await importPptx(fixture(),options);try{
  for(const value of [0.1,0.2,-0,0.000001,1e21]) {s.document.setMetadata({...s.document.metadata,customProperties:s.document.metadata.customProperties.map(p=>p.name==='Decimal'?{...p,value}:p)});const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'blocked',String(value));}
 }finally{s.dispose();}
});


test('binary64 exactness admits fractional powers of two and exact large integers',async()=>{
 for(const [lexical,expected] of [['0',0],['0.125',0.125],['-0.0009765625',-0.0009765625],['100000000000000000000',100000000000000000000]]) {
  const s=await importPptx(fixture(false,lexical),options);try{assert.equal(s.document.metadata.customProperties[0].value,expected,lexical);}finally{s.dispose();}
 }
});
