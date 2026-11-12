import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg, SvgRenderLimitError } from '../dist/index.js';

const bytes = await readFile(new URL('./fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url));
const font = { family: 'Sample "Face"', weight: 400, italic: false, format: 'woff2', base64: bytes.toString('base64') };
function layout() {
  const d = createPresentation(); d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [{ text: 'Test', style: { fontFamily: font.family } }] }] }] });
  return resolveNormalizedLayout(normalizePresentation(d));
}

test('explicit font embedding preserves bytes, style and caller input', async () => {
  const d = layout(), before = structuredClone(d), plain = await renderLayoutToSvg(d);
  const result = await renderLayoutToSvg(d, { embeddedFonts: [font], portabilityDiagnostics: 'warn' });
  assert.deepEqual(d, before); assert.ok(!plain.slides[0].svg.includes('@font-face'));
  assert.match(result.slides[0].svg, /@font-face\{font-family:&quot;Sample \\&quot;Face\\&quot;&quot;;/);
  assert.ok(result.slides[0].svg.includes(`data:font/woff2;base64,${font.base64}`));
  assert.ok(result.warnings.some(w => w.code === 'svg-embedded-font-unverified'));
  assert.deepEqual(result.warnings.find(w => w.code === 'svg-system-font-dependency').fontDependencies, [{ family: font.family, bold: false, italic: false }]);
  await assert.rejects(renderLayoutToSvg(d, { embeddedFonts: [font], maxSvgBytes: 20 }), SvgRenderLimitError);
  assert.deepEqual((await renderLayoutToSvg(d, { embeddedFonts: [font], slideIndices: [] })).slides, []);
});

test('malformed, ambiguous, oversized and mismatched font inputs reject atomically', async () => {
  const d = layout();
  for (const faces of [null, {}, Array(33).fill(font), [font, font], [font, { ...font, family: font.family.toUpperCase() }], [{ ...font, family: '' }], [{ ...font, family: 'bad\nname' }], [{ ...font, weight: 500 }], [{ ...font, italic: 'false' }], [{ ...font, format: 'woff' }], [{ ...font, base64: 'AAAA' }], [{ ...font, base64: 'wOF2===' }], [{ ...font, base64: 'A'.repeat(11200000) }]]) {
    await assert.rejects(renderLayoutToSvg(d, { embeddedFonts: faces }), error => error instanceof TypeError || error instanceof RangeError);
  }
  assert.deepEqual(await renderLayoutToSvg(d, { embeddedFonts: [] }), await renderLayoutToSvg(d));
});

test('all declared formats require their own header and reject noncanonical padding', async () => {
  const d = layout();
  for (const [format, signature] of [['woff2', 'wOF2'], ['woff', 'wOFF'], ['truetype', '\x00\x01\x00\x00'], ['opentype', 'OTTO']]) {
    const face = { ...font, format, base64: Buffer.from(signature, 'latin1').toString('base64') };
    const result = await renderLayoutToSvg(d, { embeddedFonts: [face] });
    assert.ok(result.slides[0].svg.includes(`format('${format}')`));
    assert.ok(result.warnings.some(w => w.code === 'svg-embedded-font-unverified'));
  }
  await assert.rejects(renderLayoutToSvg(d, { embeddedFonts: [{ ...font, base64: 'd09GMh==' }] }), TypeError);
});

test('incremental rendering invalidates for font changes and snapshots mutable input', async () => {
  const d = layout(), face = { ...font }, options = { environmentRevision: 'v', embeddedFonts: [face] };
  const first = await renderIncrementalSvg(d, options);
  assert.deepEqual((await renderIncrementalSvg(d, options, first.cache)).renderedSlideIds, []);
  face.family = 'Renamed Face';
  const second = await renderIncrementalSvg(d, options, first.cache);
  assert.deepEqual(second.renderedSlideIds, ['s']);
  assert.deepEqual(second.result, await renderLayoutToSvg(d, options));
  assert.ok(first.result.slides[0].svg.includes('Sample'));
});

test('font inputs are captured before waiting for asset resolution', async () => {
  const d = createPresentation();
  d.registerAsset({ id: 'asset', kind: 'image', source: { type: 'path', value: 'virtual' } });
  d.addSlide({ id: 's', elements: [
    { type: 'image', id: 'i', assetId: 'asset', box: { x: 0, y: 0, width: 10, height: 10 } },
    { type: 'text', id: 't', box: { x: 0, y: 10, width: 100, height: 30 }, content: 'Text' },
  ] });
  const normalized = resolveNormalizedLayout(normalizePresentation(d));
  for (const incremental of [false, true]) {
    const face = { ...font }, faces = [face]; let entered, release;
    const called = new Promise(resolve => { entered = resolve; });
    const options = { embeddedFonts: faces, environmentRevision: 'v', resolveAsset: () => { entered(); return new Promise(resolve => { release = resolve; }); } };
    const pending = incremental ? renderIncrementalSvg(normalized, options) : renderLayoutToSvg(normalized, options);
    await called; face.family = 'Changed'; faces.length = 0; release(undefined);
    const output = await pending, result = incremental ? output.result : output;
    assert.ok(result.slides[0].svg.includes('Sample'));
    assert.ok(!result.slides[0].svg.includes('Changed'));
  }
});
