import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { BrowserRasterError, renderIncrementalSvg, renderLayoutToBrowserRaster, renderLayoutToSvg } from '../dist/index.js';

const deck = createPresentation(); deck.addSlide({ id: 's', background: { type: 'solid', color: 'FF0000' } });
const layout = resolveNormalizedLayout(normalizePresentation(deck));

test('transparent SVG background is explicit and invalidates the incremental cache', async () => {
  const normal = await renderLayoutToSvg(layout), transparent = await renderLayoutToSvg(layout, { backgroundMode: 'transparent' });
  assert.match(normal.slides[0].svg, /fill="#FF0000"/);
  assert.doesNotMatch(transparent.slides[0].svg, /fill="#FF0000"/);
  const first = await renderIncrementalSvg(layout, { environmentRevision: 'v' });
  const next = await renderIncrementalSvg(layout, { environmentRevision: 'v', backgroundMode: 'transparent' }, first.cache);
  assert.deepEqual(next.renderedSlideIds, ['s']); assert.deepEqual(next.result, transparent);
  await assert.rejects(renderLayoutToSvg(layout, { backgroundMode: null }), TypeError);
  await assert.rejects(renderIncrementalSvg(layout, { environmentRevision: 'v', backgroundMode: null }), TypeError);
});

test('browser raster entry rejects unsupported environments and invalid options', async () => {
  await assert.rejects(renderLayoutToBrowserRaster(layout, { format: 'png' }), error => error instanceof BrowserRasterError && error.code === 'raster-environment-unavailable');
  for (const options of [{ format: 'bad' }, { format: 'png', scale: 0 }, { format: 'png', jpegBackground: '#FFFFFF' }, { format: 'jpeg', jpegQuality: 2 }, { format: 'png', maxTotalPixels: -1 }, { format: 'png', fontPolicy: 'unknown' }, { format: 'png', fontPolicy: 'require-covered' }, { format: 'png', fontCoverage: () => true }, { format: 'png', imageAnimationPolicy: 'bad' }, { format: 'png', imageAnimationPolicy: 'require-static' }, { format: 'png', isStaticImage: () => true }, { format: 'png', jpegDecodePolicy: 'bad' }, { format: 'png', maxJpegInspectionPixels: 1 }, { format: 'png', jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: -1 }, { format: 'png', crop: { x: -1, y: 0, width: 10, height: 10 } }]) {
    await assert.rejects(renderLayoutToBrowserRaster(layout, options), error => error instanceof TypeError || error instanceof RangeError);
  }
});
