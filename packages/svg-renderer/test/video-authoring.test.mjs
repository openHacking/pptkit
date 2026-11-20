import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { renderPresentationToSvg } from '../dist/index.js';

test('video SVG is a static poster or placeholder and never resolves video', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'video', kind: 'video', source: { type: 'url', value: 'https://example.invalid/video.mp4' } });
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [
    { type: 'video', id: 'with-poster', box: { x: 0, y: 0, width: 160, height: 90 }, transform: { rotation: 90 }, videoAssetId: 'video', posterAssetId: 'poster' },
    { type: 'video', id: 'without-poster', box: { x: 160, y: 0, width: 90, height: 160 }, videoAssetId: 'video' },
  ] });
  const calls = [], result = await renderPresentationToSvg(p, { resolveAsset: async asset => { calls.push(asset.id); return 'https://example.invalid/poster.png'; } });
  assert.deepEqual(calls, ['poster']); assert.equal(result.warnings.filter(w => w.code === 'video-static-preview').length, 2);
  assert.match(result.slides[0].svg, /poster\.png/); assert.match(result.slides[0].svg, /rotate\(90/); assert.doesNotMatch(result.slides[0].svg, /video\.mp4|<video/);
});
