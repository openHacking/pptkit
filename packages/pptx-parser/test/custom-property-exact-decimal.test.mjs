import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../../core/dist/index.js';
import { generatePptx } from '../../pptx-exporter/dist/index.js';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',customDecimalProjection:'exact-v1'};
for(const strict of [false,true]) test(`exact decimal import retains numeric precision and source bytes (strict=${strict})`,async()=>{
 const f=metadataFixture(strict);f.add('custom',custom(2,'decimal',' +0001.2500 ','name="Exact"')+custom(3,'decimal','.1000','name="Tenth"')+custom(4,'decimal','-000.000','name="Zero"')+custom(5,'decimal','9007199254740993.0001','name="Wide"')+custom(6,'decimal','1.0','name="Linked" linkTarget="elsewhere"'));
 const bytes=encodeMetadata(f),session=await importPptx(bytes,options),ordinary=await importPptx(bytes,{projection:'content'});
 try{assert.deepEqual(session.document.metadata.customProperties,[{name:'Exact',type:'decimal',value:'1.25'},{name:'Tenth',type:'decimal',value:'0.1'},{name:'Zero',type:'decimal',value:'0'},{name:'Wide',type:'decimal',value:'9007199254740993.0001'}]);assert.deepEqual(ordinary.document.metadata.customProperties,[]);
 assert.equal(session.report.inspection.metadata.custom.properties[0].value.value,' +0001.2500 ');const no=await session.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes),Buffer.from(bytes));assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');
 const generated=await generatePptx(session.document);const xml=readGeneratedZip(generated.bytes).get('docProps/custom.xml').toString();for(const text of ['1.25','0.1','0','9007199254740993.0001'])assert.ok(xml.includes(`<vt:decimal>${text}</vt:decimal>`));const again=await importPptx(generated.bytes,options);try{assert.deepEqual(again.document.metadata.customProperties,session.document.metadata.customProperties);}finally{again.dispose();}
 }finally{session.dispose();ordinary.dispose();}
});
test('fresh decimal authoring retains exact values through exporter and parser',async()=>{
 const properties=[{name:'Small',type:'decimal',value:'0.1'},{name:'Wide',type:'decimal',value:'9007199254740993.0001'},{name:'Negative',type:'decimal',value:'-0.0001'}];
 let doc=createPresentation({metadata:{customProperties:properties}});for(let round=0;round<3;round++){const out=await generatePptx(doc),session=await importPptx(out.bytes,options);try{assert.deepEqual(session.document.metadata.customProperties,properties);doc=createPresentation({metadata:{customProperties:session.document.metadata.customProperties}});}finally{session.dispose();}}
});
test('exact decimal projection excludes save profiles and oversized sources',async()=>{
 const f=metadataFixture();f.add('custom',custom(2,'decimal','1'.repeat(10001),'name="Large"'));const bytes=encodeMetadata(f);
 const session=await importPptx(bytes,{...options,xmlLimits:{maxPendingCharacters:20000}});try{assert.deepEqual(session.document.metadata.customProperties,[]);}finally{session.dispose();}
 for(const bad of [{projection:'structure',customDecimalProjection:'exact-v1'},{...options,customPropertySaveProfile:'scalar-v1'},{...options,customCurrencyProjection:'exact-v1'}])await assert.rejects(importPptx(bytes,bad),TypeError);
});

test('file imports snapshot the exact-decimal option before asynchronous IO',async()=>{
 const f=metadataFixture();f.add('custom',custom(2,'decimal','0.1','name="Value"'));const dir=await mkdtemp(path.join(tmpdir(),'pptkit-decimal-'));try{const file=path.join(dir,'input.pptx');await writeFile(file,encodeMetadata(f));const opts={...options},pending=importPptxFile(file,opts);opts.customDecimalProjection=undefined;const {session}=await pending;try{assert.deepEqual(session.document.metadata.customProperties,[{name:'Value',type:'decimal',value:'0.1'}]);}finally{session.dispose();}}finally{await rm(dir,{recursive:true,force:true});}
});

test('legacy preservation insertion blocks the new decimal type before producing a plan',async()=>{
 const f=metadataFixture();f.add('custom','');const session=await importPptx(encodeMetadata(f),{projection:'content',customPropertySaveProfile:'append-and-remove-v4'});
 try{session.document.setMetadata({...session.document.metadata,customProperties:[{name:'Decimal',type:'decimal',value:'0.1'}]});const plan=await session.prepareSave({mode:'edit',customPropertyAdditions:[0]});assert.equal(plan.status,'blocked');assert.equal(plan.plan,undefined);}finally{session.dispose();}
});

test('exact decimals do not project when their decoded names collide with opaque properties',async()=>{
 const f=metadataFixture();f.add('custom',custom(2,'decimal','0.1','name="A"')+custom(3,'vector','<v:lpstr>x</v:lpstr>','name="_x0041_"'));const session=await importPptx(encodeMetadata(f),options);try{assert.deepEqual(session.document.metadata.customProperties,[]);}finally{session.dispose();}
});
