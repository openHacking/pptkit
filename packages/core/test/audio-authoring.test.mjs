import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation } from '../dist/index.js';

const box = { x: 10, y: 20, width: 80, height: 50 };
function fixture() {
  const p = createPresentation(), bytes = Uint8Array.of(1, 2, 3);
  p.registerAsset({ id: 'audio', kind: 'audio', source: { type: 'bytes', value: bytes }, mimeType: 'audio/wav' }); bytes.fill(0);
  p.registerAsset({ id: 'poster', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(4, 5) } });
  p.addSlide({ id: 'source', elements: [{ type: 'audio', id: 'speaker', box, audioAssetId: 'audio', posterAssetId: 'poster' }] });
  return p;
}

test('audio owns immutable bytes and separate poster references through normalization and cleanup', () => {
  const p = fixture();
  assert.equal(p.getAsset('audio').source.value, 'AQID');
  assert.deepEqual(validatePresentation(p), []);
  const audio = normalizePresentation(p).slides[0].elements[0];
  assert.equal(audio.type, 'audio'); assert.equal(audio.audioAssetId, 'audio'); assert.equal(audio.posterAssetId, 'poster');
  assert.equal(p.getAssetReferences('audio')[0].path, 'slides.0.elements.0.audioAssetId');
  assert.equal(p.getAssetReferences('poster')[0].path, 'slides.0.elements.0.posterAssetId');
  assert.throws(() => p.removeAsset('audio'), /referenced/); assert.throws(() => p.removeAsset('poster'), /referenced/);
  assert.deepEqual(p.removeUnusedAssets(), []);
  p.removeSlide('source'); assert.deepEqual(p.removeUnusedAssets().map(x => x.id), ['audio', 'poster']);
});

test('audio and poster must refer to the correct asset kinds', () => {
  const p = fixture(), slide = p.slides[0];
  slide.replaceElement('speaker', { type: 'audio', id: 'speaker', box, audioAssetId: 'poster', posterAssetId: 'audio' });
  assert.deepEqual(validatePresentation(p).map(x => x.code).sort(), ['missing-audio-asset', 'missing-audio-poster']);
  slide.replaceElement('speaker', { type: 'image', id: 'speaker', box, assetId: 'audio' });
  assert.ok(validatePresentation(p).some(x => x.code === 'missing-image-asset'));
});

test('cross-document copying remaps both audio and poster atomically', () => {
  const source = fixture(), destination = createPresentation();
  const result = destination.copySlideFrom(source, 'source', { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
  const audio = result.slide.elements[0], ids = new Map(result.assets.map(x => [x.sourceId, x.destinationId]));
  assert.equal(audio.audioAssetId, ids.get('audio')); assert.equal(audio.posterAssetId, ids.get('poster'));
  assert.equal(destination.getAsset(audio.audioAssetId).kind, 'audio'); assert.equal(destination.getAsset(audio.posterAssetId).kind, 'image');
  assert.deepEqual(validatePresentation(destination), []);
  assert.equal(source.getAsset('audio').source.value, 'AQID');
});

test('nested and hidden audio references remain protected', () => {
  const p = fixture(), slide = p.addSlide({ id: 'hidden', hidden: true });
  slide.addElement({ type: 'group', id: 'group', box, coordinateSize: { width: 80, height: 50 }, children: [{ type: 'audio', id: 'nested', box, audioAssetId: 'audio' }] });
  assert.deepEqual(p.getAssetReferences('audio').map(x => x.elementId), ['speaker', 'nested']);
  p.removeSlide('source'); assert.throws(() => p.removeAsset('audio'), /referenced/);
  assert.deepEqual(p.removeUnusedAssets().map(x => x.id), ['poster']);
});
