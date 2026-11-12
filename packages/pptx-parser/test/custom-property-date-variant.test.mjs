import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options={projection:'content',customPropertySaveProfile:'scalar-v1'};
function fixture(strict=false,value='2020-02-29T12:34:56Z') { const f=metadataFixture(strict); f.add('custom',custom(2,'date',value,'name="Date"')+custom(3,'filetime','2021-01-02T03:04:05Z','name="Other"')); return encodeMetadata(f); }

test('UTC date variant retains source tag and identity across three rounds',async()=>{
  for(const strict of [false,true]) {
    let bytes=fixture(strict);
    for(const value of ['2000-02-29T00:00:00.000Z','2024-12-31T23:59:59.123Z','1970-01-01T00:00:00.000Z']) {
      const before=readGeneratedZip(bytes), s=await importPptx(bytes,options);
      try {
        assert.equal(s.document.metadata.customProperties[0].type,'date'); const props=s.document.metadata.customProperties.map(p=>p.name==='Date'?{...p,value}:p); s.document.setMetadata({...s.document.metadata,customProperties:props});
        const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); bytes=(await s.save(ready.plan)).bytes;
        const after=readGeneratedZip(bytes); for(const [name,payload] of before) if(name!=='inventory/custom.data') assert.deepEqual(after.get(name),payload,name);
        assert.match(after.get('inventory/custom.data').toString(),/<v:date>/); assert.match(after.get('inventory/custom.data').toString(),/<v:filetime>2021-01-02T03:04:05Z<\/v:filetime>/);
      } finally {s.dispose();}
      const reopened=await importPptx(bytes,options); try { assert.equal(reopened.document.metadata.customProperties[0].value,value); assert.equal(reopened.report.inspection.metadata.custom.properties[0].valueType,'date'); } finally {reopened.dispose();}
    }
  }
});

test('unknown offsets, excess precision, invalid dates and local times remain source-only',async()=>{
  for(const value of ['2020-01-02T03:04:05-00:00','2020-01-02T03:04:05.1234Z','2021-02-29T03:04:05Z','2020-01-02T03:04:05','2020-01-02','not-a-date']) {
    const s=await importPptx(fixture(false,value),options); try { assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name),['Other'],value); } finally {s.dispose();}
  }
});

test('date variants support mixed values but reject invalid Core dates atomically',async()=>{
  const s=await importPptx(fixture(),options);
  try {
    const bad=s.document.metadata.customProperties.map(p=>p.name==='Date'?{...p,value:'2024-02-30T00:00:00.000Z'}:p); assert.throws(()=>s.document.setMetadata({...s.document.metadata,customProperties:bad}),TypeError);
    const props=s.document.metadata.customProperties.map(p=>({...p,value:p.name==='Date'?'2022-01-01T00:00:00.000Z':'2023-01-01T00:00:00.000Z'})); s.document.setMetadata({...s.document.metadata,customProperties:props});
    const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); const xml=readGeneratedZip((await s.save(ready.plan)).bytes).get('inventory/custom.data').toString(); assert.match(xml,/<v:date>2022-01-01T00:00:00.000Z<\/v:date>/); assert.match(xml,/<v:filetime>2023-01-01T00:00:00.000Z<\/v:filetime>/);
  } finally {s.dispose();}
});

test('explicit offsets project exact UTC instants and retain source lexical bytes until edited',async()=>{
  const cases=[['2020-01-01T00:30:00+08:00','2019-12-31T16:30:00.000Z'],['2020-02-29T23:30:00.123-05:30','2020-03-01T05:00:00.123Z'],['2024-01-01T00:00:00+14:00','2023-12-31T10:00:00.000Z'],['2024-01-01T00:00:00-14:00','2024-01-01T14:00:00.000Z'],['2024-01-01T00:00:00+00:00','2024-01-01T00:00:00.000Z']];
  for(const strict of [false,true]) for(const kind of ['date','filetime']) for(const [lexical,utc] of cases) {
    const f=metadataFixture(strict); f.add('custom',custom(2,kind,lexical,'name="Offset"')+custom(3,'lpwstr','old','name="Other"')); const source=encodeMetadata(f), s=await importPptx(source,options);
    try {
      assert.equal(s.document.metadata.customProperties[0].value,utc); const properties=s.document.metadata.customProperties.map(p=>p.name==='Other'?{...p,value:'changed'}:p); s.document.setMetadata({...s.document.metadata,customProperties:properties});
      const plan=await s.prepareSave({mode:'edit'}); assert.equal(plan.status,'ready',JSON.stringify(plan.report)); const output=(await s.save(plan.plan)).bytes, xml=readGeneratedZip(output).get('inventory/custom.data').toString(); assert.ok(xml.includes(`<v:${kind}>${lexical}</v:${kind}>`));
      const reopened=await importPptx(output,options); try {assert.equal(reopened.document.metadata.customProperties[0].value,utc);} finally {reopened.dispose();}
    } finally {s.dispose();}
  }
});

test('invalid offset values and UTC rollover outside four-digit years stay source-only',async()=>{
  const values=['2020-02-30T12:00:00+01:00','2020-01-01T24:00:00+01:00','2020-01-01T00:00:00+14:01','2020-01-01T00:00:00+15:00','2020-01-01T00:00:00+01:60','2020-01-01T00:00:00-00:00','0000-01-01T00:00:00+14:00','9999-12-31T23:59:59-14:00'];
  for(const value of values) {const s=await importPptx(fixture(false,value),options); try {assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name),['Other'],value);} finally {s.dispose();}}
});

test('offset source edits bind semantic baseline and write canonical UTC over three rounds',async()=>{
  for(const strict of [false,true]) for(const kind of ['date','filetime']) {
    const f=metadataFixture(strict); f.add('custom',custom(2,kind,'2020-01-01T00:30:00+08:00','name="Offset"')); let bytes=encodeMetadata(f);
    for(const value of ['2021-02-03T04:05:06.123Z','2022-03-04T05:06:07.000Z','2023-04-05T06:07:08.999Z']) {
      const s=await importPptx(bytes,options);
      try {s.document.setMetadata({...s.document.metadata,customProperties:[{...s.document.metadata.customProperties[0],value}]}); const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); bytes=(await s.save(ready.plan)).bytes; const xml=readGeneratedZip(bytes).get('inventory/custom.data').toString(); assert.ok(xml.includes(`<v:${kind}>${value}</v:${kind}>`));} finally {s.dispose();}
      const reopened=await importPptx(bytes,options); try {assert.equal(reopened.document.metadata.customProperties[0].value,value);} finally {reopened.dispose();}
    }
  }
});

test('fractional forms map only when exactly representable in milliseconds',async()=>{
  const valid=[['2020-01-01T00:00:00.1Z','2020-01-01T00:00:00.100Z'],['2020-01-01T00:00:00.12+08:00','2019-12-31T16:00:00.120Z'],['2020-01-01T00:00:00.123000000-05:30','2020-01-01T05:30:00.123Z'],['2020-01-01T00:00:00.0000Z','2020-01-01T00:00:00.000Z'],['2020-01-01T00:00:00.1230000000Z','2020-01-01T00:00:00.123Z']];
  for(const strict of [false,true]) for(const kind of ['date','filetime']) for(const [lexical,utc] of valid) {
    const f=metadataFixture(strict); f.add('custom',custom(2,kind,lexical,'name="Fraction"')+custom(3,'lpwstr','old','name="Other"')); const source=encodeMetadata(f), s=await importPptx(source,options);
    try {
      assert.equal(s.document.metadata.customProperties[0].value,utc); s.document.setMetadata({...s.document.metadata,customProperties:s.document.metadata.customProperties.map(p=>p.name==='Other'?{...p,value:'new'}:p)});
      const plan=await s.prepareSave({mode:'edit'}); assert.equal(plan.status,'ready',JSON.stringify(plan.report)); const xml=readGeneratedZip((await s.save(plan.plan)).bytes).get('inventory/custom.data').toString(); assert.ok(xml.includes(`<v:${kind}>${lexical}</v:${kind}>`));
    } finally {s.dispose();}
  }
  for(const value of ['2020-01-01T00:00:00.0001Z','2020-01-01T00:00:00.123000001Z','2020-01-01T00:00:00.1230000001Z']) {
    const s=await importPptx(fixture(false,value),options); try {assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name),['Other'],value);} finally {s.dispose();}
  }
});
