import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx, assessPptxCurrencyNumber } from '../dist/index.js';
import { formatExactCurrency } from '../dist/metadata/currency-text.js';
import { importPptxFile } from '../dist/node.js';
import { fileURLToPath } from 'node:url';
import { metadataFixture, encodeMetadata, custom, MC } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = {projection:'content',customCurrencyProjection:'exact-save-v2',customPropertySaveProfile:'scalar-v1'};
function fixture(strict=false, value='1.2500') { const f=metadataFixture(strict);f.add('custom',custom(2,'cy',value,'name="Currency"')+custom(3,'cy','0.1000','name="Opaque"')+custom(4,'lpwstr','control','name="Control"'));return encodeMetadata(f); }
for(const strict of [false,true]) test(`currency source saves retain type and unrelated parts across three rounds (strict=${strict})`,async()=>{
 let bytes=fixture(strict);
 for(const [value,text] of [[.5,'0.5000'],[.0625,'0.0625'],[9007199254740992,'9007199254740992.0000']]) {
  const before=readGeneratedZip(bytes), session=await importPptx(bytes,options);
  try {
   assert.deepEqual(session.document.metadata.customProperties.map(p=>p.name),['Currency','Control']);
   const binding=session.report.customPropertySources[0];assert.equal(binding.status,'bound');assert.equal(binding.projectedIndex,0);
   const noop=await session.prepareSave({mode:'no-op'});assert.equal(noop.status,'ready');assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes),Buffer.from(bytes));
   session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Currency'?{...p,value}:p.name==='Control'?{...p,value:'changed'}:p)});
   const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await session.save(plan.plan)).bytes;
   const after=readGeneratedZip(bytes);for(const [name,data] of before)if(name!=='inventory/custom.data')assert.deepEqual(after.get(name),data,name);
   assert.ok(after.get('inventory/custom.data').toString().includes(`<v:cy>${text}</v:cy>`));assert.ok(after.get('inventory/custom.data').toString().includes('<v:cy>0.1000</v:cy>'));
  } finally {session.dispose();}
  const next=await importPptx(bytes,options);try{assert.equal(next.document.metadata.customProperties[0].value,value);assert.equal(next.report.inspection.metadata.custom.properties[0].valueType,'cy');}finally{next.dispose();}
 }
});
test('currency saves reject invalid precision and keep the source unchanged',async()=>{
 const bytes=fixture(),session=await importPptx(bytes,options);
 try {for(const value of [.1,.0001,-1,-0,Number.MIN_VALUE]){session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Currency'?{...p,value}:p)});const result=await session.prepareSave({mode:'edit'});assert.equal(result.status,'blocked',String(value));assert.equal(result.plan,undefined);}}
 finally{session.dispose();}
});
test('currency edit source rejects MCE and mixed or attributed values',async()=>{
 const f=metadataFixture(); f.add('custom',custom(2,'cy','1.25<!--x-->00','name="Mixed"')+custom(3,'cy','1.2500','name="Direct"')+custom(4,'cy','1.2500','name="Attributed"').replace('<v:cy>','<v:cy extra="x">'));
 const session=await importPptx(encodeMetadata(f),options);try{assert.deepEqual(session.document.metadata.customProperties.map(p=>p.name),['Direct']);}finally{session.dispose();}
 const g=metadataFixture();const part=g.add('custom',custom(2,'cy','1.2500','name="Currency"'));g.parts[part]=g.properties('custom',custom(2,'cy','1.2500','name="Currency"'),`xmlns:mc="${MC}"`);
 const s=await importPptx(encodeMetadata(g),options);try{assert.equal(s.report.customPropertySources[0].status,'unmapped');s.document.setMetadata({...s.document.metadata,customProperties:[{name:'Currency',type:'real',value:.5}]});assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');}finally{s.dispose();}
});
test('currency edit option requires its explicit scalar profile',async()=>{
 const bytes=fixture();for(const opts of [{projection:'content',customCurrencyProjection:'exact-save-v2'}, {...options, customPropertySaveProfile:'scalar-and-name-v2'}, {...options, metadataSaveProfile:'core-strings-v1'}]) await assert.rejects(importPptx(bytes,opts),TypeError);
});

test('file import snapshots currency save options before filesystem IO',async()=>{
 const opts={...options}, pending=importPptxFile(fileURLToPath(new URL('../../../tests/fixtures/custom-currency-source.pptx',import.meta.url)),opts);opts.customCurrencyProjection=undefined;opts.customPropertySaveProfile=undefined;
 const {session}=await pending;try{session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Exact'?{...p,value:.5}:p)});assert.equal((await session.prepareSave({mode:'edit'})).status,'ready');}finally{session.dispose();}
});

test('largest finite binary64 writes full decimal digits without exponent or rounding', async()=>{
 const session=await importPptx(fixture(),options);
 try{session.document.setMetadata({...session.document.metadata,customProperties:session.document.metadata.customProperties.map(p=>p.name==='Currency'?{...p,value:Number.MAX_VALUE}:p)});const prepared=await session.prepareSave({mode:'edit'});assert.equal(prepared.status,'ready',JSON.stringify(prepared.report));const output=(await session.save(prepared.plan)).bytes;const expected=((2n**53n-1n)<<971n).toString()+'.0000';assert.ok(readGeneratedZip(output).get('inventory/custom.data').toString().includes(`<v:cy>${expected}</v:cy>`));const next=await importPptx(output,options);try{assert.equal(next.document.metadata.customProperties[0].value,Number.MAX_VALUE);}finally{next.dispose();}}finally{session.dispose();}
});

test('currency formatter matches a binary denominator criterion over varied binary64 values',()=>{
 const view=new DataView(new ArrayBuffer(8));let seed=1234567;
 for(let i=0;i<1000;i++){seed=(Math.imul(seed,1664525)+1013904223)>>>0;view.setUint32(0,seed & 0x7fffffff);seed=(Math.imul(seed,1664525)+1013904223)>>>0;view.setUint32(4,seed);const value=view.getFloat64(0);if(!Number.isFinite(value))continue;const text=formatExactCurrency(value),representable=value>=2**49||Number.isInteger(value*16);assert.equal(text!==undefined,representable,String(value));if(text!==undefined)assert.deepEqual(assessPptxCurrencyNumber(text),{status:'exact',value});}
});
