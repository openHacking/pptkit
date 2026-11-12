import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', customPropertySaveProfile: 'scalar-v1' };
function fixture(strict = false, value = '中文 &amp; _x005F_x0041_') { const f = metadataFixture(strict); f.add('custom', custom(2,'lpstr',value,'name="Basic"')+custom(3,'lpwstr','sentinel','name="Other"')); return encodeMetadata(f); }

test('LPSTR Unicode strings retain type and surrounding payloads through three rounds', async () => {
  for (const strict of [false,true]) {
    let bytes=fixture(strict);
    for (const value of ['emoji 😀 & _x0042_', '', 'final\r\ntext']) {
      const before=readGeneratedZip(bytes), s=await importPptx(bytes,options);
      try {
        assert.equal(s.report.customPropertySources.find(item=>item.projectedIndex===0).status,'bound');
        const properties=s.document.metadata.customProperties.map(p=>p.name==='Basic'?{...p,value}:p); s.document.setMetadata({...s.document.metadata,customProperties:properties});
        const ready=await s.prepareSave({mode:'edit'}); assert.equal(ready.status,'ready',JSON.stringify(ready.report)); bytes=(await s.save(ready.plan)).bytes;
        const after=readGeneratedZip(bytes); for (const [name,payload] of before) if (name!=='inventory/custom.data') assert.deepEqual(after.get(name),payload,name);
        assert.match(after.get('inventory/custom.data').toString(),/<v:lpstr>/); assert.match(after.get('inventory/custom.data').toString(),/<v:lpwstr>sentinel<\/v:lpwstr>/);
      } finally { s.dispose(); }
      const reopened=await importPptx(bytes,options); try { assert.equal(reopened.document.metadata.customProperties[0].value,value); assert.equal(reopened.report.inspection.metadata.custom.properties[0].valueType,'lpstr'); } finally { reopened.dispose(); }
    }
  }
});

test('opaque escape sequences and linked values remain unprojected and unsupported edits block', async () => {
  for (const value of ['_x0008_','_x005F_x0041_x0042_']) {
    const s=await importPptx(fixture(false,value),options); try { assert.deepEqual(s.document.metadata.customProperties.map(p=>p.name),['Other']); } finally { s.dispose(); }
  }
  const f=metadataFixture(); f.add('custom',custom(2,'lpstr','value','name="Linked" linkTarget="elsewhere"')); const linked=await importPptx(encodeMetadata(f),options);
  try { assert.deepEqual(linked.document.metadata.customProperties,[]); } finally { linked.dispose(); }
  const s=await importPptx(fixture(),options); try { const props=s.document.metadata.customProperties.map(p=>p.name==='Basic'?{...p,value:'_x005F_x0041_'}:p); s.document.setMetadata({...s.document.metadata,customProperties:props}); const result=await s.prepareSave({mode:'edit'}); assert.equal(result.status,'blocked'); assert.equal(result.plan,undefined); } finally { s.dispose(); }
});
