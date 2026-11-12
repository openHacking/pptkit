import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',metadataSaveProfile:'core-insert-v5'};
function fixture(strict=false,body='<c:revision>7</c:revision><d:creator>Existing</d:creator>') {const f=metadataFixture(strict); f.add('core',body); return encodeMetadata(f);}

test('missing direct core strings insert and remain editable across three rounds',async()=>{
  for(const strict of [false,true]) {
    let source=fixture(strict);
    for(let round=0;round<3;round++) {
      const before=readGeneratedZip(source),s=await importPptx(source,options);
      try {
        if(round===0) for(const field of ['title','subject','description','language']) assert.equal(s.report.coreMetadataSources.find(item=>item.field===field).status,'insertable');
        s.document.setMetadata({...s.document.metadata,title:`Title ${round} & 中文`,subject:round===0?'':`Subject ${round}`,description:`Desc ${round}\r\n`,language:'fr-FR',author:`Author ${round}`});
        const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); source=(await s.save(ready.plan)).bytes;
        const after=readGeneratedZip(source); for(const [name,payload] of before) if(name!=='inventory/core.data') assert.deepEqual(after.get(name),payload,name); assert.match(after.get('inventory/core.data').toString(),/<c:revision>7<\/c:revision>/);
      } finally {s.dispose();}
      const reopened=await importPptx(source,options); try {assert.equal(reopened.document.metadata.title,`Title ${round} & 中文`); assert.equal(reopened.document.metadata.subject,round===0?'':`Subject ${round}`); assert.equal(reopened.document.metadata.description,`Desc ${round}\r\n`); assert.equal(reopened.document.metadata.author,`Author ${round}`);} finally {reopened.dispose();}
    }
  }
});

test('insertion requires existing core part, supported profile and valid XML characters',async()=>{
  const absent=await importPptx(encodeMetadata(metadataFixture()),options);
  try {assert.equal(absent.report.coreMetadataSources.find(item=>item.field==='title').status,'unmapped'); absent.document.setMetadata({...absent.document.metadata,title:'Changed'}); assert.equal((await absent.prepareSave({mode:'edit'})).status,'blocked');} finally {absent.dispose();}
  const old=await importPptx(fixture(),{projection:'content',metadataSaveProfile:'extended-company-v4'}); try {old.document.setMetadata({...old.document.metadata,title:'Changed'}); assert.equal((await old.prepareSave({mode:'edit'})).status,'blocked');} finally {old.dispose();}
  const s=await importPptx(fixture(),options); try {s.document.setMetadata({...s.document.metadata,title:'bad\u0000text'}); const p=await s.prepareSave({mode:'edit'}); assert.equal(p.status,'blocked'); assert.equal(p.plan,undefined);} finally {s.dispose();}
});

test('core insertion and custom creation share a single candidate',async()=>{
  const s=await importPptx(fixture(),{...options,customPropertySaveProfile:'create-and-append-v5'});
  try {s.document.setMetadata({...s.document.metadata,title:'New title',subject:'New subject',customProperties:[{name:'A',type:'integer',value:3}]}); const ready=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); const reopened=await importPptx((await s.save(ready.plan)).bytes,options); try {assert.equal(reopened.document.metadata.title,'New title'); assert.equal(reopened.document.metadata.subject,'New subject'); assert.equal(reopened.document.metadata.customProperties[0].value,3);} finally {reopened.dispose();}} finally {s.dispose();}
});

test('self-closing root, signed package, cancellation and stale plans are handled safely',async()=>{
  const f=metadataFixture(); const part=f.add('core',''); f.parts[part]=f.parts[part].replace('></c:coreProperties>','/>'); const source=encodeMetadata(f), s=await importPptx(source,options);
  try {
    const noOp=await s.prepareSave({mode:'no-op'}); assert.equal(noOp.status,'ready'); assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),Buffer.from(source));
    s.document.setMetadata({...s.document.metadata,subject:'New'}); const abort=new AbortController(); abort.abort(); await assert.rejects(s.prepareSave({mode:'edit',signal:abort.signal}),{code:'operation-aborted'});
    const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); s.document.setMetadata({...s.document.metadata,subject:'Changed'}); await assert.rejects(s.save(ready.plan),{code:'save-snapshot-stale'});
  } finally {s.dispose();}
  const signed=metadataFixture(); signed.add('core',''); signed.parts['deck/_rels/z.xml.rels']=signed.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const guard=await importPptx(encodeMetadata(signed),options); try {const explicit=await guard.prepareSave({mode:'edit',coreMetadataInsertions:['title']}); assert.equal(explicit.status,'blocked'); assert.ok(explicit.report.diagnostics.some(d=>d.code==='save-text-package-policy')); guard.document.setMetadata({...guard.document.metadata,subject:'New'}); const p=await guard.prepareSave({mode:'edit'}); assert.equal(p.status,'blocked'); assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));} finally {guard.dispose();}
  const limited=await importPptx(source,{...options,sessionLimits:{maxSaveValidationWork:100}}); try {limited.document.setMetadata({...limited.document.metadata,subject:'New'}); await assert.rejects(limited.prepareSave({mode:'edit'}),{code:'save-limit-exceeded'});} finally {limited.dispose();}
});

test('core insertion preserves UTF-8 BOM and UTF-16 endian encoding',async()=>{
  const encodeBE=text=>{const data=Buffer.from(text,'utf16le'); for(let i=0;i<data.length;i+=2) [data[i],data[i+1]]=[data[i+1],data[i]]; return Buffer.concat([Buffer.from([0xfe,0xff]),data]);};
  const encoders=[text=>Buffer.concat([Buffer.from([0xef,0xbb,0xbf]),Buffer.from(text)]),text=>Buffer.concat([Buffer.from([0xff,0xfe]),Buffer.from(text,'utf16le')]),encodeBE];
  for(const encode of encoders) {
    const f=metadataFixture(), part=f.add('core','<c:revision>7</c:revision>'); const original=f.parts[part]; f.parts[part]=encode(original); const source=encodeMetadata(f), s=await importPptx(source,options);
    try {s.document.setMetadata({...s.document.metadata,subject:'New & 中文'}); const p=await s.prepareSave({mode:'edit'}); assert.equal(p.status,'ready',JSON.stringify(p.report)); const xml=readGeneratedZip((await s.save(p.plan)).bytes).get(part); const expected=original.replace('</c:coreProperties>','<dc:subject xmlns:dc="http://purl.org/dc/elements/1.1/">New &amp; 中文</dc:subject></c:coreProperties>'); assert.deepEqual(xml,encode(expected));} finally {s.dispose();}
  }
});


test('explicit insertion writes normalized defaults and reopens without changing other parts',async()=>{
  const source=fixture(false,'<c:revision>7</c:revision>'), before=readGeneratedZip(source), s=await importPptx(source,options);
  try {
    const defaults={title:'PPTKit Presentation',author:'PPTKit',language:'en-US'};
    const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['title','author','language']});
    assert.equal(p.status,'ready',JSON.stringify(p.report));
    assert.deepEqual(p.plan.report.changes.map(c=>c.documentPath),['/metadata/author','/metadata/language','/metadata/title']);
    const bytes=(await s.save(p.plan)).bytes, after=readGeneratedZip(bytes);
    for(const [name,payload] of before) if(name!=='inventory/core.data') assert.deepEqual(after.get(name),payload,name);
    const reopened=await importPptx(bytes,options); try {for(const field of ['title','author','language']) {assert.equal(reopened.document.metadata[field],defaults[field]); assert.equal(reopened.report.coreMetadataSources.find(item=>item.field===field).status,'bound');}} finally {reopened.dispose();}
  } finally {s.dispose();}
});

test('explicit insertion validates source, value, profile, duplicate selection, and mixed edits',async()=>{
  const s=await importPptx(fixture(),options);
  try {
    await assert.rejects(s.prepareSave({mode:'edit',coreMetadataInsertions:['title','title']}),TypeError);
    await assert.rejects(s.prepareSave({mode:'no-op',coreMetadataInsertions:['title']}),TypeError);
    assert.equal((await s.prepareSave({mode:'edit',coreMetadataInsertions:['author']})).status,'blocked');
    assert.equal((await s.prepareSave({mode:'edit',coreMetadataInsertions:['subject']})).status,'blocked');
    s.document.setMetadata({...s.document.metadata,subject:'',description:'New'});
    const p=await s.prepareSave({mode:'edit',coreMetadataInsertions:['title','subject','description']});
    assert.equal(p.status,'ready',JSON.stringify(p.report));
    assert.equal(p.plan.report.changes.filter(c=>c.documentPath==='/metadata/description').length,1);
    const reopened=await importPptx((await s.save(p.plan)).bytes,options); try {assert.equal(reopened.document.metadata.subject,'');assert.equal(reopened.document.metadata.description,'New');} finally {reopened.dispose();}
  } finally {s.dispose();}
  const old=await importPptx(fixture(),{projection:'content',metadataSaveProfile:'extended-company-v4'}); try {await assert.rejects(old.prepareSave({mode:'edit',coreMetadataInsertions:['title']}),TypeError);} finally {old.dispose();}
});
