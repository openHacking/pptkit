import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../dist/index.js';

const bytes = (...value) => ({ kind: 'image', source: { type: 'bytes', value: Uint8Array.from(value) }, mimeType: 'image/png', width: 8, height: 4 });
const options = { deduplicate: 'content', external: 'reject' };

test('cross-document asset copy maps IDs, deduplicates matching bytes and detaches source state', () => {
  const source = createPresentation({ id: 'source' }), target = createPresentation({ id: 'target' });
  source.registerAssets([{ asset: { ...bytes(1, 2), id: 'alpha' } }, { asset: { ...bytes(1, 2), id: 'beta' } }, { asset: { ...bytes(1, 2), id: 'other', width: 9 } }]);
  target.registerAsset({ ...bytes(9), id: 'asset-1' }); const previous = target.assets;
  const map = target.copyAssetsFrom(source, ['alpha', 'beta', 'other'], options);
  assert.deepEqual(map, [{ sourceId: 'alpha', destinationId: 'asset-2', reused: false }, { sourceId: 'beta', destinationId: 'asset-2', reused: true }, { sourceId: 'other', destinationId: 'asset-3', reused: false }]);
  assert.ok(Object.isFrozen(map)); assert.ok(map.every(Object.isFrozen)); assert.equal(previous.length, 1);
  const second = target.copyAssetsFrom(source, ['alpha', 'other'], options); assert.deepEqual(second.map(x => x.reused), [true, true]);
  assert.equal(target.assets.length, 3); const stored = target.getAsset('asset-2');
  source.replaceAsset('alpha', bytes(7, 8)); source.removeAsset('beta'); assert.equal(stored.source.value, 'AQI='); assert.equal(target.getAsset('asset-2'), stored);
});

test('external source policy is explicit and failures never commit earlier items', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAssets([{ asset: { ...bytes(1), id: 'embedded' } }, { asset: { kind: 'image', id: 'file', source: { type: 'path', value: './logo.png' } } }, { asset: { kind: 'image', id: 'remote', source: { type: 'url', value: 'https://example.invalid/a' } } }]);
  assert.throws(() => target.copyAssetsFrom(source, ['embedded', 'file'], options), /explicit preserve/); assert.equal(target.assets.length, 0);
  const map = target.copyAssetsFrom(source, ['embedded', 'file', 'remote'], { deduplicate: 'registry', external: 'preserve' });
  assert.deepEqual(map.map(m => m.destinationId), ['asset-1', 'asset-2', 'asset-3']);
  assert.deepEqual(target.getAsset('asset-2').source, { type: 'path', value: './logo.png' }); assert.equal(target.getAsset('asset-3').source.value, 'https://example.invalid/a');
  const repeated = target.copyAssetsFrom(source, ['embedded'], { deduplicate: 'registry', external: 'reject' }); assert.equal(repeated[0].reused, false); assert.equal(repeated[0].destinationId, 'asset-4');
});

test('cross-document resource conflicts, invalid identity and reentrancy abort before commit', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAssets([{ asset: { ...bytes(1), id: 'a' } }, { asset: { ...bytes(2), id: 'b', dedupeKey: 'key', width: 9 } }]);
  target.registerAsset({ ...bytes(2), id: 'existing', dedupeKey: 'key', width: 8 });
  assert.throws(() => target.copyAssetsFrom(source, ['a', 'b'], options), /different metadata/); assert.deepEqual(target.assets.map(a => a.id), ['existing']);
  for (const ids of [['missing'], ['a', 'a'], [''], [1], new Array(1)]) assert.throws(() => target.copyAssetsFrom(source, ids, options));
  assert.throws(() => target.copyAssetsFrom(target, [], options), TypeError);
  for (const value of [undefined, {}, { deduplicate: 'content', external: 'ignore' }, { ...options, extra: true }]) assert.throws(() => target.copyAssetsFrom(source, ['a'], value), TypeError);
  assert.deepEqual(target.assets.map(a => a.id), ['existing']);
  const hostile = { get assets() { target.registerAsset({ ...bytes(7), id: 'reentrant' }); return source.assets; } };
  assert.throws(() => target.copyAssetsFrom(hostile, ['a'], options), /changed while capturing/);
  assert.deepEqual(target.assets.map(a => a.id), ['existing', 'reentrant']);
  const retry = target.copyAssetsFrom(source, ['a'], options); assert.equal(retry[0].destinationId, 'asset-1');
});

test('cross-document asset copy rejects unknown source fields and duplicate source identities', () => {
  const target = createPresentation(), valid = { id: 'a', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'AQ==' } };
  for (const source of [{ assets: [{ ...valid, unknown: 1 }] }, { assets: [valid, valid] }, { assets: [{ ...valid, source: { ...valid.source, extra: true } }] }, { assets: [{ ...valid, source: { type: 'path', value: 7 } }] }, { assets: [{ ...valid, accessibility: { extra: true } }] }]) {
    assert.throws(() => target.copyAssetsFrom(source, ['a'], options), TypeError); assert.equal(target.assets.length, 0);
  }
  let reads = 0; const changing = { ...valid, get id() { return ++reads === 1 ? 'a' : 'b'; } };
  assert.throws(() => target.copyAssetsFrom({ assets: [changing] }, ['a'], options), /identity changed/); assert.equal(target.assets.length, 0);
});

test('cross-document copy rejects hidden or symbol metadata before cloning and preserves atomic retry', () => {
  const target = createPresentation(), valid = { id: 'a', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'AQ==' } };
  const withHidden = value => Object.defineProperty(value, 'unmodeled', { value: 'retain-me', enumerable: false });
  const cases = [
    withHidden({ ...valid }), { ...valid, [Symbol('opaque')]: 'retain-me' },
    { ...valid, source: withHidden({ ...valid.source }) }, { ...valid, source: { ...valid.source, [Symbol('opaque')]: 'retain-me' } },
    { ...valid, accessibility: withHidden({ description: 'visible' }) }, { ...valid, accessibility: { description: 'visible', [Symbol('opaque')]: 'retain-me' } },
    Object.defineProperty({ ...valid }, 'mimeType', { value: 'image/png', enumerable: false }),
  ];
  for (const asset of cases) {
    assert.throws(() => target.copyAssetsFrom({ assets: [asset] }, ['a'], options), /unsupported fields/);
    assert.equal(target.assets.length, 0);
  }
  assert.equal(target.copyAssetsFrom({ assets: [valid] }, ['a'], options)[0].destinationId, 'asset-1');
});

test('metadata introduced by a getter during asset capture is not silently dropped', () => {
  const target = createPresentation(), src = { type: 'bytes', encoding: 'base64' };
  Object.defineProperty(src, 'value', { enumerable: true, get() { Object.defineProperty(src, 'opaque', { value: 1, enumerable: false }); return 'AQ=='; } });
  assert.throws(() => target.copyAssetsFrom({ assets: [{ id: 'a', kind: 'image', source: src }] }, ['a'], options), /unsupported fields/);
  assert.equal(target.assets.length, 0);
});

test('slide copy rejects an opaque source asset without committing its page or resources', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAsset({ ...bytes(1), id: 'a' });
  source.addSlide({ id: 'slide', elements: [{ id: 'image', type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 10, height: 10 } }] });
  const original = source.assets[0], opaque = { ...original, source: Object.defineProperty({ ...original.source }, 'unknown', { value: true, enumerable: false }) };
  const custom = Object.create(source, { assets: { value: [opaque] } });
  const policy = { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: options };
  assert.throws(() => target.copySlideFrom(custom, 'slide', policy), /unsupported fields/);
  assert.equal(target.assets.length, 0); assert.equal(target.slides.length, 0);
  const result = target.copySlideFrom(source, 'slide', policy);
  assert.equal(result.slide.id, 'slide-1'); assert.equal(result.assets[0].destinationId, 'asset-1');
});

test('source ID list is captured once before policy access can mutate it', () => {
  const source = createPresentation(), target = createPresentation(); source.registerAssets([{ asset: { ...bytes(1), id: 'a' } }, { asset: { ...bytes(2), id: 'b' } }]);
  const ids = ['a']; let reads = 0;
  Object.defineProperty(ids, '0', { configurable: true, get() { reads++; return reads === 1 ? 'a' : 'b'; } });
  const map = target.copyAssetsFrom(source, ids, { get deduplicate() { ids.length = 0; return 'content'; }, external: 'reject' });
  assert.equal(reads, 1); assert.equal(map[0].sourceId, 'a'); assert.equal(target.getAsset(map[0].destinationId).source.value, 'AQ==');
});
