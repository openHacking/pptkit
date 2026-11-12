import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom, GUID, MC } from './metadata-fixture.mjs';

const options = { projection: 'content', customPropertySourceProfile: 'raw-v1' };
test('custom property source inventory binds raw scalar ancestry and does not grant save authority', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add('custom', custom(2, 'lpwstr', 'one', 'name="Simple"') + custom(3, 'lpwstr', 'a<!--marker-->b', 'name="Comment"') + custom(4, 'vector', '<v:lpwstr>x</v:lpwstr>', 'name="Opaque"'));
    const source = encodeMetadata(f), session = await importPptx(source, options);
    try {
      const entries = session.report.customPropertySources;
      assert.deepEqual(entries.map(e => e.status), ['bound', 'unmapped', 'unmapped']);
      assert.equal(entries[0].source.sourceObjectId, `${GUID}:2`);
      assert.equal(typeof entries[0].source.elementIndex, 'number');
      assert.equal(typeof entries[0].valueElementIndex, 'number');
      assert.equal(entries[0].projectedIndex, undefined);
      assert.equal(entries[1].valueElementIndex, undefined);
      assert.ok(Object.isFrozen(entries));
      const prepared = await session.prepareSave({ mode: 'no-op' });
      assert.equal(prepared.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), Buffer.from(source));
      const values = session.document.metadata.customProperties.map(p => p.name === 'Simple' ? { ...p, value: 'changed' } : p);
      session.document.setMetadata({ ...session.document.metadata, customProperties: values });
      const edit = await session.prepareSave({ mode: 'edit' });
      assert.equal(edit.status, 'blocked');
      assert.equal(edit.plan, undefined);
    } finally { session.dispose(); }
  }
});

test('inventory is opt-in, handles absent parts and requires content projection', async () => {
  const bytes = encodeMetadata(metadataFixture());
  const ordinary = await importPptx(bytes, { projection: 'content' });
  try { assert.equal(ordinary.report.customPropertySources, undefined); } finally { ordinary.dispose(); }
  const session = await importPptx(bytes, options);
  try { assert.deepEqual(session.report.customPropertySources, []); } finally { session.dispose(); }
  await assert.rejects(importPptx(bytes, { projection: 'structure', customPropertySourceProfile: 'raw-v1' }), TypeError);
  await assert.rejects(importPptx(bytes, { projection: 'content', customPropertySourceProfile: 'unknown' }), TypeError);
});


test('compatibility-processed custom parts do not receive raw bindings', async () => {
  const f = metadataFixture(), body = custom(2, 'lpwstr', 'value', 'name="A"');
  const part = f.add('custom', body);
  f.parts[part] = f.properties('custom', body, `xmlns:mc="${MC}"`);
  const session = await importPptx(encodeMetadata(f), options);
  try {
    assert.equal(session.report.customPropertySources[0].status, 'unmapped');
    assert.equal(session.report.customPropertySources[0].valueElementIndex, undefined);
  } finally { session.dispose(); }
});


test('save profile reports source-to-Core index only for editable bound properties', async () => {
  const f = metadataFixture(); f.add('custom', custom(2, 'vector', '<v:lpwstr>x</v:lpwstr>', 'name="Opaque"') + custom(3, 'lpwstr', 'a<!--comment-->b', 'name="Unbound"') + custom(4, 'lpwstr', 'value', 'name="_x005F_x0041_"'));
  const session = await importPptx(encodeMetadata(f), { projection: 'content', customPropertySaveProfile: 'scalar-v1' });
  try {
    const entries = session.report.customPropertySources;
    assert.deepEqual(entries.map(e => e.projectedIndex), [undefined, undefined, 1]);
    assert.equal(session.document.metadata.customProperties[1].name, '_x0041_');
    assert.ok(Object.isFrozen(entries[2]));
  } finally { session.dispose(); }
});
