import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspectPptx, inspectPptxOpaqueCustomProperties } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom, MC } from './metadata-fixture.mjs';

test('opaque source inventory exposes direct text without changing standard metadata inspection', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/custom-currency-source.pptx', import.meta.url));
  const original = await inspectPptx(bytes), items = await inspectPptxOpaqueCustomProperties(bytes);
  assert.deepEqual(items.map(item => [item.name, item.valueType, item.content]), [
    ['Exact', 'cy', { status: 'text', value: '1.2500' }],
    ['Inexact', 'cy', { status: 'text', value: '0.1000' }],
    ['Negative', 'cy', { status: 'text', value: '-12.5000' }],
  ]);
  assert.ok(items.every(item => item.part === '/docProps/custom.xml' && Object.isFrozen(item.content)));
  assert.ok(Object.isFrozen(items));
  assert.ok(original.metadata.custom.properties.filter(item => item.valueType === 'cy').every(item => item.value.status === 'uninterpreted'));
});

test('both dialects preserve text, nested content and effective MCE branches separately', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add('custom', '');
    const choice = custom(2, 'cy', 'wrong', 'name="Choice"');
    const fallback = custom(2, 'cy', '<![CDATA[1.]]>2500', 'name="Fallback" linkTarget="https://example.invalid/value"');
    f.parts['inventory/custom.data'] = f.properties('custom', `<mc:AlternateContent><mc:Choice Requires="u">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>${custom(3, 'vector', '<v:lpstr>nested</v:lpstr>', 'name="Vector"')}${custom(4, 'cy', '1<!--unknown-->2', 'name="Mixed"')}${custom(5, 'lpwstr', 'ordinary', 'name="Mapped"')}`, `xmlns:mc="${MC}" xmlns:u="urn:unknown"`);
    const items = await inspectPptxOpaqueCustomProperties(encodeMetadata(f));
    assert.deepEqual(items.map(item => [item.id, item.name, item.valueType, item.content]), [[2, 'Fallback', 'cy', { status: 'text', value: '1.2500' }], [3, 'Vector', 'vector', { status: 'non-text' }], [4, 'Mixed', 'cy', { status: 'non-text' }]]);
    assert.ok(items.every(item => item.valueNamespace === f.VT));
    assert.equal(items[0].linkTarget, 'https://example.invalid/value');
  }
});

test('opaque inventory respects aggregate text limits, cancellation and input isolation', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/custom-currency-source.pptx', import.meta.url));
  const mutable = new Uint8Array(source), pending = inspectPptxOpaqueCustomProperties(mutable); mutable.fill(0);
  assert.equal((await pending)[0].content.value, '1.2500');
  const f = metadataFixture(); f.add('custom', custom(2, 'cy', '1'.repeat(1000001), 'name="Large"'));
  const bytes = encodeMetadata(f);
  await assert.rejects(inspectPptxOpaqueCustomProperties(bytes, { xmlLimits: { maxPendingCharacters: 2000000 } }), { code: 'package-limit-exceeded' });
  const controller = new AbortController(); controller.abort();
  await assert.rejects(inspectPptxOpaqueCustomProperties(bytes, { signal: controller.signal }), { code: 'operation-aborted' });
  await assert.rejects(inspectPptxOpaqueCustomProperties(new Uint8Array(new SharedArrayBuffer(4))), TypeError);
});
