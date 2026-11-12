import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const options = { projection: 'content', customPropertySaveProfile: 'create-and-append-v5' };
test('create missing custom property part and append across three reopens', async () => {
  for (const strict of [false, true]) {
    let bytes = encodeMetadata(metadataFixture(strict));
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, options);
      try {
        assert.equal(session.report.customPropertyInsertion.status, 'available');
        const old = session.document.metadata.customProperties ?? [];
        session.document.setMetadata({ ...session.document.metadata, customProperties: [...old, { name: `Name ${round}`, type: 'integer', value: round }] });
        const prepared = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [round] });
        assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        bytes = (await session.save(prepared.plan)).bytes;
        const after = readGeneratedZip(bytes);
        for (const [name, payload] of before) if (!(round === 0 ? ['_rels/.rels', '[Content_Types].xml'] : ['pptkit-custom-properties-1.xml']).includes(name)) assert.deepEqual(after.get(name), payload, name);
      } finally { session.dispose(); }
      const reopened = await importPptx(bytes, options);
      try { assert.deepEqual(reopened.document.metadata.customProperties, Array.from({ length: round + 1 }, (_, i) => ({ name: `Name ${i}`, type: 'integer', value: i }))); assert.deepEqual(reopened.report.inspection.metadata.custom.properties.map(p => p.id), Array.from({ length: round + 1 }, (_, i) => i + 2)); }
      finally { reopened.dispose(); }
    }
  }
});

test('creation combines with core metadata, visibility and text without modifying other entries', async () => {
  const { shape, body } = await import('./content-fixture.mjs');
  const f = metadataFixture(); f.add('core', '<d:title>Original</d:title>');
  f.parts['deck/z.xml'] = f.p('sld', `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({ id: '2', txBox: true, text: body() })}</p:spTree></p:cSld>`, `xmlns:a="${f.A}" show="0"`);
  const source = encodeMetadata(f), config = { ...options, metadataSaveProfile: 'core-strings-v1', slideVisibilitySaveProfile: 'attribute-v1', editSaveProfile: 'text-runs-v3' }, session = await importPptx(source, config);
  const values = [{ name: 'Text', type: 'string', value: '中文 & _x0041_' }, { name: 'Flag', type: 'boolean', value: true }, { name: 'Real', type: 'real', value: -0 }, { name: 'Date', type: 'date', value: '2020-01-02T03:04:05.000Z' }];
  try {
    const slide = session.document.slides[0], original = slide.elements[0], changed = structuredClone(original); changed.content[0].runs[0].text += '!'; slide.replaceElement(original.id, changed); slide.setHidden(false);
    session.document.setMetadata({ ...session.document.metadata, title: 'Changed', customProperties: values });
    const ready = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0, 1, 2, 3] }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    const output = (await session.save(ready.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const [name, bytes] of before) if (!['_rels/.rels', '[Content_Types].xml', 'inventory/core.data', 'deck/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
    const reopened = await importPptx(output, config);
    try { assert.deepEqual(reopened.document.metadata.customProperties, values); assert.equal(reopened.document.metadata.title, 'Changed'); assert.equal(reopened.document.slides[0].hidden, false); }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('creation blocks invalid properties and indexes and rejects stale or cancelled plans', async () => {
  const source = encodeMetadata(metadataFixture()), session = await importPptx(source, options);
  try {
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: 'New', type: 'string', value: '_x005F_x0041_' }] });
    assert.equal((await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] })).status, 'blocked');
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'safe' }] });
    assert.equal((await session.prepareSave({ mode: 'edit', customPropertyAdditions: [1] })).status, 'blocked');
    const abort = new AbortController(); abort.abort(); await assert.rejects(session.prepareSave({ mode: 'edit', customPropertyAdditions: [0], signal: abort.signal }), { code: 'operation-aborted' });
    const ready = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'changed' }] });
    await assert.rejects(session.save(ready.plan), { code: 'save-snapshot-stale' });
  } finally { session.dispose(); }
});

test('creation uses a fresh path and relationship when package names collide', async () => {
  const f = metadataFixture(); f.parts['PPTKIT-CUSTOM-PROPERTIES-1.XML'] = '<extra/>'; f.metadataTypes['PPTKIT-CUSTOM-PROPERTIES-1.XML'] = 'application/xml';
  f.rootRelations.push({ Id: 'pptkitCustomProperties2', Type: 'urn:test', Target: 'deck/main.data' });
  const session = await importPptx(encodeMetadata(f), options);
  try {
    assert.equal(session.report.customPropertyInsertion.part, '/pptkit-custom-properties-3.xml');
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: 'A', type: 'string', value: 'v' }] });
    const ready = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    const reopened = await importPptx((await session.save(ready.plan)).bytes, options);
    try { assert.equal(reopened.report.inspection.metadata.custom.part, '/pptkit-custom-properties-3.xml'); }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('creation is unavailable when an inactive source relationship points to the proposed missing part', async () => {
  const { sourceAlternates } = await import('./closure-fixture.mjs');
  const f = metadataFixture();
  f.parts['deck/_rels/z.xml.rels'] = sourceAlternates('<Relationship Id="hidden" Type="urn:asset" Target="../pptkit-custom-properties-1.xml"/>', `<Relationship Id="layout" Type="${f.R}/slideLayout" Target="layout.xml"/>`);
  const session = await importPptx(encodeMetadata(f), options);
  try {
    assert.equal(session.report.customPropertyInsertion.status, 'unavailable');
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: 'A', type: 'string', value: 'v' }] });
    const result = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined);
  } finally { session.dispose(); }
});
