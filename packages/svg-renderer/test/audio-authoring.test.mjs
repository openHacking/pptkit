import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { renderPresentationToSvg } from '../dist/index.js';

test('audio SVG is a static poster or placeholder and never resolves audio', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'url', value: 'https://example.invalid/audio.wav' } });
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [
    { type: 'audio', id: 'with-poster', box: { x: 0, y: 0, width: 100, height: 50 }, audioAssetId: 'audio', posterAssetId: 'poster' },
    { type: 'audio', id: 'without-poster', box: { x: 100, y: 0, width: 100, height: 50 }, audioAssetId: 'audio' },
  ] });
  const calls = [], result = await renderPresentationToSvg(p, { resolveAsset: async asset => { calls.push(asset.id); return 'https://example.invalid/poster.png'; } });
  assert.deepEqual(calls, ['poster']); assert.equal(result.warnings.filter(w => w.code === 'audio-static-preview').length, 2);
  assert.match(result.slides[0].svg, /poster\.png/); assert.doesNotMatch(result.slides[0].svg, /audio\.wav|<audio/);
});
