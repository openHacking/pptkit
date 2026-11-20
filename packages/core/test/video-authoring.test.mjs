import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation } from '../dist/index.js';

const box = { x: 10, y: 20, width: 120, height: 60 };
function fixture() {
  const p = createPresentation(), bytes = Uint8Array.of(1, 2, 3);
  p.registerAsset({ id: 'video', kind: 'video', source: { type: 'bytes', value: bytes }, mimeType: 'video/mp4' }); bytes.fill(0);
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(4, 5) } });
  p.addSlide({ id: 'source', elements: [{ type: 'video', id: 'clip', box, transform: { rotation: 90 }, videoAssetId: 'video', posterAssetId: 'poster' }] });
  return p;
}

test('video bytes and poster are independent immutable assets protected by references', () => {
  const p = fixture(); assert.equal(p.getAsset('video').source.value, 'AQID'); assert.deepEqual(validatePresentation(p), []);
  const video = normalizePresentation(p).slides[0].elements[0];
  assert.equal(video.type, 'video'); assert.equal(video.videoAssetId, 'video'); assert.equal(video.posterAssetId, 'poster'); assert.equal(video.transform.rotation, 90);
  assert.equal(p.getAssetReferences('video')[0].path, 'slides.0.elements.0.videoAssetId');
  assert.equal(p.getAssetReferences('poster')[0].path, 'slides.0.elements.0.posterAssetId');
  assert.throws(() => p.removeAsset('video'), /referenced/); assert.throws(() => p.removeAsset('poster'), /referenced/);
  assert.deepEqual(p.removeUnusedAssets(), []); p.removeSlide('source'); assert.deepEqual(p.removeUnusedAssets().map(x => x.id), ['video', 'poster']);
});

test('video and poster references require matching asset kinds', () => {
  const p = fixture(), slide = p.slides[0];
  slide.replaceElement('clip', { type: 'video', id: 'clip', box, videoAssetId: 'poster', posterAssetId: 'video' });
  assert.deepEqual(validatePresentation(p).map(x => x.code).sort(), ['missing-video-asset', 'missing-video-poster']);
  slide.replaceElement('clip', { type: 'image', id: 'clip', box, assetId: 'video' });
  assert.ok(validatePresentation(p).some(x => x.code === 'missing-image-asset'));
});

test('slide copy remaps video and poster and protects hidden nested uses', () => {
  const source = fixture(); source.addSlide({ id: 'hidden', hidden: true, elements: [{ type: 'group', id: 'g', box, coordinateSize: { width: 120, height: 60 }, children: [{ type: 'video', id: 'nested', box, videoAssetId: 'video' }] }] });
  const destination = createPresentation();
  const result = destination.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  const ids = new Map(result.assets.map(x => [x.sourceId, x.destinationId])), video = result.slide.elements[0];
  assert.equal(video.videoAssetId, ids.get('video')); assert.equal(video.posterAssetId, ids.get('poster'));
  assert.equal(destination.getAsset(video.videoAssetId).kind, 'video'); assert.deepEqual(validatePresentation(destination), []);
  source.removeSlide('source'); assert.throws(() => source.removeAsset('video'), /referenced/); assert.deepEqual(source.removeUnusedAssets().map(x => x.id), ['poster']);
});
