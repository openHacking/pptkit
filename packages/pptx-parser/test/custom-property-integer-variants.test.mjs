import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', customPropertySaveProfile: 'scalar-v1' };
const variants = [['i1',-128,127],['i2',-32768,32767],['i4',-2147483648,2147483647],['int',-2147483648,2147483647],['i8',-2147483648,2147483647],['ui1',0,255],['ui2',0,65535],['ui4',0,2147483647],['uint',0,2147483647],['ui8',0,2147483647]];
function fixture(strict) { const f=metadataFixture(strict); f.add('custom', variants.map(([type,min],index)=>custom(index+2,type,String(min),`name="${type}"`)).join('') + custom(20,'vector','<v:lpwstr>opaque</v:lpwstr>','name="Opaque"')); return encodeMetadata(f); }

test('integer variants retain source tags and identities through three rounds', async () => {
  for (const strict of [false,true]) {
    let bytes=fixture(strict);
    for (const round of [0,1,2]) {
      const before=readGeneratedZip(bytes), s=await importPptx(bytes,options);
      try {
        assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name), variants.map(v=>v[0]));
        const values=s.document.metadata.customProperties.map((property,index)=>({...property,value:round===0?variants[index][2]:round===1?variants[index][1]:variants[index][1]===0?1:0}));
        s.document.setMetadata({...s.document.metadata,customProperties:values}); const p=await s.prepareSave({mode:'edit'}); assert.equal(p.status,'ready',JSON.stringify(p.report));
        bytes=(await s.save(p.plan)).bytes; const after=readGeneratedZip(bytes);
        for (const [name,payload] of before) if (name!=='inventory/custom.data') assert.deepEqual(after.get(name),payload,name);
        for (const [type] of variants) assert.match(after.get('inventory/custom.data').toString(),new RegExp(`<v:${type}>`));
        assert.match(after.get('inventory/custom.data').toString(),/<v:vector><v:lpwstr>opaque<\/v:lpwstr><\/v:vector>/);
      } finally { s.dispose(); }
      const reopened=await importPptx(bytes,options); try { assert.deepEqual(reopened.report.inspection.metadata.custom.properties.slice(0,10).map(p=>p.valueType),variants.map(v=>v[0])); } finally { reopened.dispose(); }
    }
  }
});

test('out-of-source-range writes are blocked without partial publication', async () => {
  const s=await importPptx(fixture(false),options);
  try {
    for (const [type,min,max] of variants.filter(([type])=>['i1','i2','ui1','ui2','ui4','uint','ui8'].includes(type))) {
      const index=s.document.metadata.customProperties.findIndex(p=>p.name===type);
      for (const value of [min-1,max+1]) {
        if (value < -2147483648 || value > 2147483647) continue;
        const entries=s.document.metadata.customProperties.map((p,i)=>i===index?{...p,value}:p); s.document.setMetadata({...s.document.metadata,customProperties:entries});
        const prepared=await s.prepareSave({mode:'edit'}); assert.equal(prepared.status,'blocked',`${type} ${value}`); assert.equal(prepared.plan,undefined);
        s.document.setMetadata({...s.document.metadata,customProperties:s.document.metadata.customProperties.map((p,i)=>i===index?{...p,value:min}:p)});
      }
    }
  } finally { s.dispose(); }
});

test('wide and noncanonical integer source lexemes remain unprojected', async () => {
  const f=metadataFixture(); f.add('custom',custom(2,'i8','2147483648','name="Wide"')+custom(3,'ui8','18446744073709551615','name="Huge"')+custom(4,'i2','+01','name="Lexical"'));
  const s=await importPptx(encodeMetadata(f),options); try { assert.deepEqual(s.document.metadata.customProperties,[]); assert.equal(s.report.customPropertySources.length,3); } finally { s.dispose(); }
});
