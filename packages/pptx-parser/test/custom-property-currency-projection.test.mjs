import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { importPptxFile } from '../dist/node.js';
import { fileURLToPath } from 'node:url';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';

const options = { projection: 'content', customCurrencyProjection: 'exact-v1' };
for (const strict of [false, true]) test(`currency projection is exact and read-only (strict=${strict})`, async () => {
  const f = metadataFixture(strict);
  f.add('custom', custom(2,'cy','1.2500','name="Exact"') + custom(3,'cy','0.1000','name="Inexact"') + custom(4,'cy','-12.5000','name="Invalid"') + custom(5,'cy','9007199254740992.0000','name="Wide"') + custom(6,'cy','9007199254740993.0000','name="Rounded"') + custom(7,'cy','1.25<!--comment-->00','name="Comment"') + custom(8,'cy','1.2500','name="Linked" linkTarget="elsewhere"') + custom(9,'lpwstr','before','name="Control"'));
  const bytes = encodeMetadata(f);
  const ordinary = await importPptx(bytes, { projection: 'content' });
  const session = await importPptx(bytes, options);
  try {
    assert.deepEqual(ordinary.document.metadata.customProperties.map(p => p.name), ['Control']);
    assert.deepEqual(session.document.metadata.customProperties, [{name:'Exact',type:'real',value:1.25},{name:'Wide',type:'real',value:9007199254740992},{name:'Control',type:'string',value:'before'}]);
    assert.ok(session.report.inspection.metadata.custom.properties.filter(p => p.valueType === 'cy').every(p => p.value.status === 'uninterpreted'));
    assert.equal(session.report.findings.filter(f => f.feature === 'metadata.custom-property-currency' && f.tier === 'fallback').length, 2);
    const no = await session.prepareSave({ mode: 'no-op' });
    assert.equal(no.status, 'ready');
    assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes), Buffer.from(bytes));
    assert.equal((await session.prepareSave({mode:'edit'})).status, 'blocked');
  } finally { ordinary.dispose(); session.dispose(); }
});

test('currency opt-in rejects save profiles and non-content projection', async () => {
  const bytes = encodeMetadata(metadataFixture());
  for (const opts of [{projection:'structure',customCurrencyProjection:'exact-v1'}, {...options, customPropertySaveProfile:'scalar-v1'}, {...options, customCurrencyProjection:'other'}]) await assert.rejects(importPptx(bytes, opts), TypeError);
});

test('file adapter captures currency opt-in before asynchronous IO', async () => {
  const opts = { ...options };
  const pending = importPptxFile(fileURLToPath(new URL('../../../tests/fixtures/custom-currency-source.pptx', import.meta.url)), opts);
  opts.customCurrencyProjection = undefined;
  const { session } = await pending;
  try { assert.deepEqual(session.document.metadata.customProperties.map(p => p.name), ['Exact', 'Control']); }
  finally { session.dispose(); }
});

test('oversized direct text cannot be projected from a valid prefix', async () => {
  const f = metadataFixture();
  f.add('custom', custom(2, 'cy', '1.2500' + ' '.repeat(1000000), 'name="Large"'));
  const session = await importPptx(encodeMetadata(f), { ...options, xmlLimits: { maxPendingCharacters: 2000000 } });
  try { assert.deepEqual(session.document.metadata.customProperties, []); } finally { session.dispose(); }
});
