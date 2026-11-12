import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom, GUID } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', customPropertySaveProfile: 'remove-v3' };
const id = n => `${GUID}:${n}`;
function fixture(strict) {
  const f = metadataFixture(strict);
  f.add('custom', custom(2, 'lpwstr', 'one', 'name="A"') + custom(3, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"') + custom(4, 'bool', 'false', 'name="B"') + custom(5, 'i4', '7', 'name="C"'));
  return encodeMetadata(f);
}
test('explicit custom source deletion preserves unknown neighbors and supports simultaneous scalar edits', async () => {
  for (const strict of [false, true]) {
    const source = fixture(strict), before = readGeneratedZip(source), session = await importPptx(source, options);
    try {
      const entries = session.document.metadata.customProperties.filter(p => p.name !== 'A' && p.name !== 'C').map(p => ({ ...p, name: 'Renamed', value: true }));
      session.document.setMetadata({ ...session.document.metadata, customProperties: entries });
      const prepared = await session.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2), id(5)] });
      assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
      const result = (await session.save(prepared.plan)).bytes, after = readGeneratedZip(result);
      for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
      const xml = after.get('inventory/custom.data').toString();
      assert.match(xml, /name="Opaque"/); assert.match(xml, /<v:vector><v:lpwstr>opaque<\/v:lpwstr><\/v:vector>/);
      assert.doesNotMatch(xml, /name="A"|name="C"/); assert.match(xml, /name="Renamed"/);
      const reopened = await importPptx(result, options);
      try { assert.deepEqual(reopened.document.metadata.customProperties, entries); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test('removal requires explicit bound identity and matching document state', async () => {
  const bytes = fixture(false);
  for (const scenario of ['missing', 'unknown', 'opaque', 'duplicate', 'wrong-profile']) {
    const session = await importPptx(bytes, scenario === 'wrong-profile' ? { projection: 'content', customPropertySaveProfile: 'scalar-v1' } : options);
    try {
      session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.filter(p => p.name !== 'A') });
      const ids = scenario === 'missing' ? [] : scenario === 'unknown' ? ['unknown'] : scenario === 'opaque' ? [id(3)] : scenario === 'duplicate' ? [id(2), id(2)] : [id(2)];
      if (scenario === 'duplicate' || scenario === 'wrong-profile') await assert.rejects(session.prepareSave({ mode: 'edit', customPropertyRemovals: ids }), TypeError);
      else { const result = await session.prepareSave({ mode: 'edit', customPropertyRemovals: ids }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined); }
    } finally { session.dispose(); }
  }
});


test('three sequential removals retain source IDs and leave the opaque property and empty projection', async () => {
  for (const strict of [false, true]) {
    let bytes = fixture(strict);
    for (const [propertyId, name] of [[2, 'A'], [4, 'B'], [5, 'C']]) {
      const session = await importPptx(bytes, options);
      try {
        session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.filter(p => p.name !== name) });
        const prepared = await session.prepareSave({ mode: 'edit', customPropertyRemovals: [id(propertyId)] });
        assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        assert.equal(prepared.plan.report.changes[0].source.sourceObjectId, id(propertyId));
        bytes = (await session.save(prepared.plan)).bytes;
      } finally { session.dispose(); }
    }
    const session = await importPptx(bytes, options);
    try { assert.deepEqual(session.document.metadata.customProperties, []); assert.deepEqual(session.report.inspection.metadata.custom.properties.map(p => p.name), ['Opaque']); } finally { session.dispose(); }
  }
});

test('custom removals obey cancellation, stale plans and shared change budget', async () => {
  const bytes = fixture(false), session = await importPptx(bytes, options);
  try {
    session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.filter(p => p.name !== 'A') });
    const abort = new AbortController(); abort.abort();
    await assert.rejects(session.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2)], signal: abort.signal }), { code: 'operation-aborted' });
    const plan = await session.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2)] }); assert.equal(plan.status, 'ready');
    session.document.setMetadata({ ...session.document.metadata, title: 'changed after plan' });
    await assert.rejects(session.save(plan.plan), { code: 'save-snapshot-stale' });
    await assert.rejects(session.prepareSave({ mode: 'edit', customPropertyRemovals: new Array(1) }), TypeError);
  } finally { session.dispose(); }
  const limited = await importPptx(bytes, { ...options, sessionLimits: { maxSaveChanges: 1 } });
  try {
    limited.document.setMetadata({ ...limited.document.metadata, customProperties: limited.document.metadata.customProperties.filter(p => p.name !== 'A' && p.name !== 'C') });
    await assert.rejects(limited.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2), id(5)] }), { code: 'save-limit-exceeded' });
  } finally { limited.dispose(); }
});


test('explicit deletion identity accounting respects the shared analysis budget', async () => {
  const session = await importPptx(fixture(false), { ...options, sessionLimits: { maxSaveAnalysisWork: 10 } });
  try {
    session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.filter(p => p.name !== 'A') });
    await assert.rejects(session.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2)] }), { code: 'save-limit-exceeded' });
  } finally { session.dispose(); }
});


test('removal identifiers are captured before asynchronous preparation', async () => {
  const session = await importPptx(fixture(false), options);
  try {
    session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.filter(p => p.name !== 'A') });
    const ids = [id(2)], pending = session.prepareSave({ mode: 'edit', customPropertyRemovals: ids });
    ids[0] = id(5);
    const plan = await pending; assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
    assert.equal(plan.plan.report.changes[0].source.sourceObjectId, id(2));
    const output = (await session.save(plan.plan)).bytes, reopened = await importPptx(output, options);
    try { assert.deepEqual(reopened.document.metadata.customProperties.map(p => p.name), ['B', 'C']); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});
