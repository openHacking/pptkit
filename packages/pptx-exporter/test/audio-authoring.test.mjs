import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';

test('PPTX rejects audio before resolving any assets instead of silently dropping it', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'url', value: 'https://example.invalid/audio.wav' } });
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'audio', posterAssetId: 'poster', box: { x: 0, y: 0, width: 100, height: 50 } }] });
  let calls = 0;
  await assert.rejects(generatePptx(p, { resolveAsset: async () => { calls++; throw Error('unexpected'); } }), /audio export is not implemented/);
  assert.equal(calls, 0);
});

test('hidden nested audio is rejected even though ordinary hidden objects are omitted', async () => {
  const p = createPresentation();
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'bytes', value: Uint8Array.of(1) } });
  p.addSlide({ hidden: true, elements: [{ type: 'group', box: { x: 0, y: 0, width: 100, height: 50 }, coordinateSize: { width: 100, height: 50 }, children: [{ type: 'audio', hidden: true, audioAssetId: 'audio', box: { x: 0, y: 0, width: 20, height: 20 } }] }] });
  await assert.rejects(generatePptx(p), /audio export is not implemented/);
});
