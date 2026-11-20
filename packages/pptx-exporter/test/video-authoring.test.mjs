import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';

test('PPTX rejects video before resolving any assets, including hidden nested video', async () => {
  for (const nested of [false, true]) {
    const p = createPresentation();
    p.registerAsset({ id: 'video', kind: 'video', source: { type: 'url', value: 'https://example.invalid/video.mp4' } });
    p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
    const box = { x: 0, y: 0, width: 120, height: 60 }, video = { type: 'video', hidden: nested, videoAssetId: 'video', posterAssetId: 'poster', box };
    p.addSlide({ elements: nested ? [{ type: 'group', box, coordinateSize: { width: 120, height: 60 }, children: [video] }] : [video] });
    let calls = 0;
    await assert.rejects(generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } }), /video export is not supported/);
    assert.equal(calls, 0);
  }
});

test('unreferenced video assets are not passed to the image loader', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'unused', kind: 'video', source: { type: 'url', value: 'https://example.invalid/video.mp4' } });
  p.addSlide({ elements: [] }); let calls = 0;
  await generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } });
  assert.equal(calls, 0);
});
