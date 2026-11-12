import test from 'node:test';
import assert from 'node:assert/strict';
import { planCustomPropertyCreation as plan } from '../dist/preservation/plan-custom-property-creation.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { alternateParts, closureBytes } from './closure-fixture.mjs';

const run = (bytes, dialect = 'transitional', options = {}) => plan(bytes, dialect, options, 1_000_000, 1000);

test('creation plans collision-free identities for both OOXML dialects', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.parts['PPTKIT-CUSTOM-PROPERTIES-1.XML'] = '<extra/>';
    f.metadataTypes['PPTKIT-CUSTOM-PROPERTIES-1.XML'] = 'application/xml';
    f.rootRelations.push({ Id: 'pptkitCustomProperties2', Type: 'urn:test', Target: 'deck/main.data' });
    const result = await run(encodeMetadata(f), strict ? 'strict' : 'transitional');
    assert.equal(result.part, '/pptkit-custom-properties-3.xml');
    assert.equal(result.relationshipId, 'pptkitCustomProperties3');
    assert.match(Buffer.from(result.relationships).toString(), /Id="pptkitCustomProperties3"/);
    assert.match(Buffer.from(result.contentTypes).toString(), /PartName="\/pptkit-custom-properties-3.xml"/);
  }
});

test('existing custom properties and cancellation cannot produce a creation plan', async () => {
  const f = metadataFixture(); f.add('custom', custom(2, 'lpwstr', 'x', 'name="Existing"'));
  await assert.rejects(run(encodeMetadata(f)), { code: 'xml-edit-conflict' });
  const abort = new AbortController(); abort.abort();
  await assert.rejects(run(encodeMetadata(metadataFixture()), 'transitional', { signal: abort.signal }), { code: 'operation-aborted' });
});

test('planning captures input and options before yielding', async () => {
  const bytes = Buffer.from(encodeMetadata(metadataFixture())), options = {}, pending = run(bytes, 'transitional', options);
  bytes.fill(0); options.signal = AbortSignal.abort();
  assert.equal((await pending).part, '/pptkit-custom-properties-1.xml');
});

test('an inactive raw relationship to a missing part blocks creation', async () => {
  const parts = alternateParts('<Relationship Id="hidden" Type="urn:asset" Target="pptkit-custom-properties-1.xml"/>');
  await assert.rejects(run(closureBytes(parts)), { code: 'xml-edit-unsupported' });
});
