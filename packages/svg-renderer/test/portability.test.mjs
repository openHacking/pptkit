import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { renderIncrementalSvg, renderLayoutToSvg } from '../dist/index.js';

test('portability warnings follow emitted content and resolver results without changing SVG', async () => {
  const d = createPresentation();
  d.registerAsset({ id: 'asset', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a.png' } });
  d.addSlide({ id: 'one', elements: [
    { type: 'text', id: 'simple', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'font-family: literal' },
    { type: 'text', id: 'rich', box: { x: 0, y: 30, width: 100, height: 30 }, content: [{ runs: [{ text: 'A' }, { text: 'B', style: { bold: true } }] }] },
    { type: 'image', id: 'remote', assetId: 'asset', box: { x: 0, y: 60, width: 20, height: 20 } },
    { type: 'image', id: 'hidden', hidden: true, assetId: 'asset', box: { x: 0, y: 80, width: 20, height: 20 } },
  ] });
  d.addSlide({ id: 'empty' });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  const plain = await renderLayoutToSvg(layout);
  const checked = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn' });
  assert.deepEqual(checked.slides, plain.slides);
  assert.deepEqual(checked.warnings.filter(w => w.code.startsWith('svg-')).map(w => [w.code, w.slideId, w.elementId, w.assetId]), [
    ['svg-system-font-dependency', 'one', 'simple', undefined],
    ['svg-foreign-object-dependency', 'one', 'rich', undefined],
    ['svg-system-font-dependency', 'one', 'rich', undefined],
    ['svg-external-image-dependency', 'one', 'remote', 'asset'],
  ]);
  const embedded = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', resolveAsset: () => 'data:image/png;base64,AA==' });
  assert.ok(!embedded.warnings.some(w => w.code === 'svg-external-image-dependency'));
  const blob = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', resolveAsset: () => 'blob:https://example.invalid/123' });
  assert.ok(blob.warnings.some(w => w.code === 'svg-external-image-dependency'));
  const relative = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', resolveAsset: () => 'images/a.png' });
  assert.ok(relative.warnings.some(w => w.code === 'svg-external-image-dependency'));
  const rejected = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', resolveAsset: () => 'javascript:alert(1)' });
  assert.ok(rejected.warnings.some(w => w.code === 'unsafe-asset-url'));
  assert.ok(!rejected.warnings.some(w => w.code === 'svg-external-image-dependency'));
  assert.deepEqual((await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', slideIndices: [1] })).warnings, []);
  await assert.rejects(renderLayoutToSvg(layout, { portabilityDiagnostics: null }), TypeError);
});

test('table and group descendants report their own dependencies once', async () => {
  const d = createPresentation(); d.addSlide({ id: 's', elements: [
    { type: 'group', id: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [
      { type: 'text', id: 'child', box: { x: 0, y: 0, width: 50, height: 20 }, content: 'child' },
    ] },
    { type: 'table', id: 'table', box: { x: 0, y: 100, width: 100, height: 30 }, columns: [100], rows: [{ height: 30, cells: [{ content: 'cell' }] }] },
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  const warnings = (await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn' })).warnings;
  assert.deepEqual(warnings.filter(w => w.code.startsWith('svg-')).map(w => [w.code, w.elementId]), [
    ['svg-system-font-dependency', 'child'], ['svg-foreign-object-dependency', 'table'], ['svg-system-font-dependency', 'table'],
  ]);
});

test('incremental rendering invalidates cached diagnostics when the mode changes', async () => {
  const d = createPresentation(); d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'Hello' }] });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  const first = await renderIncrementalSvg(layout, { environmentRevision: '1' });
  const next = await renderIncrementalSvg(layout, { environmentRevision: '1', portabilityDiagnostics: 'warn' }, first.cache);
  assert.deepEqual(next.renderedSlideIds, ['s']);
  assert.deepEqual(next.result, await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn' }));
  assert.deepEqual((await renderIncrementalSvg(layout, { environmentRevision: '1', portabilityDiagnostics: 'warn' }, next.cache)).renderedSlideIds, []);
  await assert.rejects(renderIncrementalSvg(layout, { environmentRevision: '1', portabilityDiagnostics: null }), TypeError);
});

test('font dependencies identify emitted families and styles for text, tables and charts', async () => {
  const d = createPresentation({ theme: { fonts: { body: 'Body Face', heading: 'Heading Face' } } });
  d.addSlide({ id: 's', elements: [
    { type: 'text', id: 'text', box: { x: 0, y: 0, width: 200, height: 40 }, content: [{ runs: [{ text: 'A中', style: { fontFamily: 'Latin Face', eastAsianFontFamily: 'CJK Face', bold: true } }] }] },
    { type: 'text', id: 'html', box: { x: 0, y: 50, width: 200, height: 40 }, content: [{ runs: [{ text: 'A', style: { fontFamily: 'Family "Quoted"' } }, { text: 'B', style: { fontFamily: 'Family "Quoted"', italic: true } }] }] },
    { type: 'table', id: 'table', box: { x: 0, y: 100, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: [{ runs: [{ text: 'X', style: { fontFamily: 'Cell Face' } }] }] }] }] },
    { type: 'chart', id: 'chart', chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }], title: 'Chart', box: { x: 220, y: 0, width: 200, height: 150 } },
  ] });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  for (const tableMode of ['browser', 'native']) {
    const result = await renderLayoutToSvg(layout, { portabilityDiagnostics: 'warn', tableMode });
    const fonts = Object.fromEntries(result.warnings.filter(w => w.code === 'svg-system-font-dependency').map(w => [w.elementId, w.fontDependencies]));
    assert.deepEqual(fonts.text, [{ family: 'Latin Face', bold: true, italic: false }, { family: 'CJK Face', bold: true, italic: false }]);
    assert.deepEqual(fonts.html, [{ family: 'Family "Quoted"', bold: false, italic: false }, { family: 'Family "Quoted"', bold: false, italic: true }]);
    assert.deepEqual(fonts.table, [{ family: 'Cell Face', bold: false, italic: false }]);
    assert.ok(fonts.chart.length > 0 && fonts.chart.every(font => typeof font.family === 'string' && typeof font.bold === 'boolean' && typeof font.italic === 'boolean'));
  }
});

test('incremental cache isolates nested font dependency values', async () => {
  const d = createPresentation(); d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'Hello' }] });
  const layout = resolveNormalizedLayout(normalizePresentation(d)), options = { environmentRevision: 'v', portabilityDiagnostics: 'warn' };
  const first = await renderIncrementalSvg(layout, options);
  const font = first.result.warnings.find(w => w.code === 'svg-system-font-dependency').fontDependencies[0];
  const expected = font.family;
  font.family = 'changed';
  const again = await renderIncrementalSvg(layout, options, first.cache);
  assert.equal(again.result.warnings.find(w => w.code === 'svg-system-font-dependency').fontDependencies[0].family, expected);
  const cached = first.cache.entries[0].warnings.find(w => w.code === 'svg-system-font-dependency').fontDependencies;
  assert.ok(Object.isFrozen(cached) && Object.isFrozen(cached[0]));
});
