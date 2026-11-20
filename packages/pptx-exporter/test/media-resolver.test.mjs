import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

async function samples() {
  const parts = readGeneratedZip(await readFile(new URL('../../../tests/fixtures/pptkit-mixed-av.pptx', import.meta.url)));
  return { movie: parts.get('ppt/media/movie.mp4'), sound: parts.get('ppt/media/sound.wav') };
}
function document() {
  const p = createPresentation();
  p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'url', value: 'https://example.invalid/movie.mp4' } });
  p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'path', value: '/not-loaded/sound.wav' } });
  p.addSlide({ elements: ['video', 'audio', 'video', 'audio'].map((type, i) => ({ type, id: `m${i}`, ...(type === 'video' ? { videoAssetId: 'movie' } : { audioAssetId: 'sound' }), box: { x: i * 50, y: 0, width: 40, height: 40 } })) });
  return p;
}
const options = { audioMode: 'wave-pcm-v1', videoMode: 'iso-bmff-v1' };

test('explicit resolver embeds external media once per asset without implicit fetching', async () => {
  const data = await samples(), calls = [], p = document();
  await assert.rejects(generatePptx(p, options), /paths and links are not loaded/);
  const result = await generatePptx(p, { ...options, maxAssetBytes: 100000, resolveMedia(asset, context) {
    assert.ok(Object.isFrozen(asset) && Object.isFrozen(asset.source) && Object.isFrozen(asset.accessibility) && Object.isFrozen(context));
    assert.equal(context.maxBytes, 100000); calls.push(asset.id); return { bytes: data[asset.id], mimeType: asset.mimeType };
  } });
  assert.deepEqual(calls, ['movie', 'sound']);
  const zip = readGeneratedZip(result.bytes);
  assert.deepEqual(zip.get('ppt/media/movie.mp4'), data.movie); assert.deepEqual(zip.get('ppt/media/sound.wav'), data.sound);
  assert.equal([...zip.keys()].filter(name => name.endsWith('.mp4')).length, 1);
  assert.equal([...zip.keys()].filter(name => name.endsWith('.wav')).length, 1);
  assert.equal(result.warnings.filter(x => x.code === 'video-playback-unverified').length, 2);
  assert.equal(result.warnings.filter(x => x.code === 'audio-playback-unverified').length, 2);
});

test('resolution rejects missing, oversized, mismatched and damaged data and allows retry', async () => {
  const data = await samples(), p = document();
  for (const mode of ['missing', 'mime', 'oversize', 'damaged', 'shared']) {
    await assert.rejects(generatePptx(p, { ...options, maxAssetBytes: 100000, resolveMedia(asset) {
      if (asset.id !== 'movie') return { bytes: data.sound };
      if (mode === 'missing') return undefined;
      if (mode === 'mime') return { bytes: data.movie, mimeType: 'audio/wav' };
      if (mode === 'oversize') return { bytes: new Uint8Array(100001) };
      if (mode === 'shared') return { bytes: new Uint8Array(new SharedArrayBuffer(16)) };
      return { bytes: Uint8Array.of(1, 2, 3) };
    } }), undefined, mode);
  }
  const retry = await generatePptx(p, { ...options, resolveMedia: asset => ({ bytes: data[asset.id] }) });
  assert.deepEqual(readGeneratedZip(retry.bytes).get('ppt/media/movie.mp4'), data.movie);
});

test('invalid resolved audio prevents output before image resolver is invoked', async () => {
  const data = await samples(), p = document(); let images = 0;
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'url', value: 'https://example.invalid/poster.png' } });
  p.addSlide({ elements: [{ type: 'image', assetId: 'poster', box: { x: 0, y: 0, width: 20, height: 20 } }] });
  for (const mode of ['mime', 'damaged']) {
    await assert.rejects(generatePptx(p, { ...options, resolveAsset() { images++; }, resolveMedia(asset) {
      return asset.id === 'movie' ? { bytes: data.movie } : mode === 'mime' ? { bytes: data.sound, mimeType: 'video/mp4' } : { bytes: Uint8Array.of(1, 2, 3) };
    } }), undefined, mode);
    assert.equal(images, 0);
  }
});

test('invalid profile is rejected before external resolution and cancellation interrupts a pending resolver', async () => {
  const p = document(); let calls = 0;
  await assert.rejects(generatePptx(p, { audioMode: 'wave-pcm-v1', resolveMedia() { calls++; } }), /requires videoMode/); assert.equal(calls, 0);
  const controller = new AbortController(), reason = new Error('stopped');
  const pending = generatePptx(p, { ...options, signal: controller.signal, resolveMedia(_asset, context) { calls++; assert.equal(context.signal, controller.signal); return new Promise(() => {}); } });
  await new Promise(resolve => setTimeout(resolve, 0)); controller.abort(reason);
  await assert.rejects(pending, error => error === reason); assert.equal(calls, 1);
});
