import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { assessPptxCurrencyText, assessPptxCurrencyNumber, inspectPptx, inspectPptxOpaqueCustomProperties } from '../dist/index.js';
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
  assert.deepEqual(items.map(item => assessPptxCurrencyText(item.content.value).status), ['valid', 'valid', 'invalid']);
});

test('currency text assessment preserves exact decimal digits and rejects non-schema forms', () => {
  for (const [text, canonical, units] of [
    ['1.2500', '1.2500', '12500'], [' .1000\t', '0.1000', '1000'], ['0000.0000', '0.0000', '0'],
    ['000123.0040', '123.0040', '1230040'], ['9007199254740993.9999', '9007199254740993.9999', '90071992547409939999'],
  ]) {
    const result = assessPptxCurrencyText(text);
    assert.deepEqual(result, { status: 'valid', canonical, tenThousandths: units }); assert.ok(Object.isFrozen(result));
  }
  for (const text of ['-12.5000', '+1.0000', '1', '1.', '1.000', '1.00000', '1e0', '１.0000', '1.0000\u00a0', '1.0000\u2028', '1.0000x', '']) assert.deepEqual(assessPptxCurrencyText(text), { status: 'invalid' }, text);
  assert.throws(() => assessPptxCurrencyText(1), TypeError);
  assert.throws(() => assessPptxCurrencyText('0'.repeat(1000001)), RangeError);
});

test('currency number assessment refuses rounding and overflow', () => {
  for (const [text, expected] of [
    ['0.0000', { status: 'exact', value: 0 }], ['1.2500', { status: 'exact', value: 1.25 }],
    ['.0625', { status: 'exact', value: .0625 }], ['0.1000', { status: 'inexact' }],
    ['0.0001', { status: 'inexact' }], ['9007199254740992.0000', { status: 'exact', value: 9007199254740992 }],
    ['9007199254740993.0000', { status: 'inexact' }], ['-1.0000', { status: 'invalid' }],
    [`${'9'.repeat(310)}.0000`, { status: 'out-of-range' }],
  ]) { const result = assessPptxCurrencyNumber(text); assert.deepEqual(result, expected, text); assert.ok(Object.isFrozen(result)); }
  const maxFiniteInteger = ((2n ** 53n - 1n) << 971n).toString();
  assert.deepEqual(assessPptxCurrencyNumber(`${maxFiniteInteger}.0000`), { status: 'exact', value: Number.MAX_VALUE });
  assert.deepEqual(assessPptxCurrencyNumber(`${'9'.repeat(309)}.0000`), { status: 'out-of-range' });
  // For these bounded integers, units/10000 has an exact binary64 form iff
  // the factor 5^4 in 10000 divides units; no binary64 significand can overflow.
  let state = 123456789;
  for (let i = 0; i < 500; i++) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    const raw = state % 1000000000, units = i % 2 === 0 ? raw - raw % 625 : raw, text = `${Math.floor(units / 10000)}.${String(units % 10000).padStart(4, '0')}`;
    assert.equal(assessPptxCurrencyNumber(text).status, units % 625 === 0 ? 'exact' : 'inexact', text);
  }
  assert.throws(() => assessPptxCurrencyNumber('0'.repeat(1000001)), RangeError);
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
