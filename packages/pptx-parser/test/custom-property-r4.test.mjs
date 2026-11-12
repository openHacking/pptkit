import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', customPropertySaveProfile: 'scalar-v1' };
function fixture(strict = false, value = '1.25') { const f = metadataFixture(strict); f.add('custom', custom(2, 'r4', value, 'name="Single"') + custom(3, 'r8', '0.1', 'name="Double"')); return encodeMetadata(f); }

test('exact binary32 values preserve source r4 and identity across three rounds', async () => {
  for (const strict of [false,true]) {
    let bytes = fixture(strict);
    for (const value of [Math.fround(3.4028234663852886e38), Math.fround(2 ** -149), -0]) {
      const before = readGeneratedZip(bytes), s = await importPptx(bytes, options);
      try {
        const props = s.document.metadata.customProperties.map(p => p.name === 'Single' ? { ...p, value } : p); s.document.setMetadata({ ...s.document.metadata, customProperties: props });
        const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        bytes = (await s.save(ready.plan)).bytes; const after = readGeneratedZip(bytes);
        for (const [name, data] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), data, name);
        assert.match(after.get('inventory/custom.data').toString(), /<v:r4>/); assert.match(after.get('inventory/custom.data').toString(), /<v:r8>0.1<\/v:r8>/);
      } finally { s.dispose(); }
      const reopened = await importPptx(bytes, options); try { assert.ok(Object.is(reopened.document.metadata.customProperties[0].value, value)); assert.equal(reopened.report.inspection.metadata.custom.properties[0].valueType, 'r4'); } finally { reopened.dispose(); }
    }
  }
});

test('inexact or noncanonical binary32 inputs stay source-only and edits fail atomically', async () => {
  for (const value of ['0.1','1.0','1e-50','3.5e+38','NaN','INF']) {
    const s = await importPptx(fixture(false, value), options); try { assert.deepEqual(s.document.metadata.customProperties.map(p => p.name), ['Double'], value); } finally { s.dispose(); }
  }
  const s = await importPptx(fixture(), options);
  try {
    for (const value of [0.1, 1e-50, 3.5e38]) {
      s.document.setMetadata({ ...s.document.metadata, customProperties: s.document.metadata.customProperties.map(p => p.name === 'Single' ? { ...p, value } : p) });
      const result = await s.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked', String(value)); assert.equal(result.plan, undefined);
    }
  } finally { s.dispose(); }
});
