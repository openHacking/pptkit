import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 100, height: 40 };
const options = { id: 'copied', masterId: null, theme: 'destination', assets: { deduplicate: 'content', external: 'reject' } };

test('equivalent layout reuse remaps owned connector identities and preserves existing resources', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.defineSlideLayout({ id: 'old', name: 'Shared', elements: [
    { id: 'picture', type: 'image', assetId: 'image', box },
    { id: 'edge', type: 'connector', start: { elementId: 'picture' }, end: { x: 120, y: 10 } },
  ], placeholders: [{ key: 'body', kind: 'body', box }] });
  const first = dest.copySlideLayoutFrom(source, 'old', options);
  const repeated = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'unused', reuseEquivalent: true });
  assert.equal(repeated.layout, first.layout);
  assert.deepEqual(repeated.elementIds, first.elementIds);
  assert.equal(repeated.assets[0].destinationId, first.assets[0].destinationId);
  assert.equal(repeated.assets[0].reused, true);
  assert.equal(dest.layouts.length, 1);
  assert.equal(dest.assets.length, 1);
  const newLayout = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'unused' });
  assert.equal(newLayout.layout.id, 'unused');
  assert.equal(newLayout.elementIds.picture, 'element-3');
  const third = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'third', reuseEquivalent: true });
  assert.equal(third.layout, first.layout);
});

test('layout reuse distinguishes definitions, target masters and resource metadata', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'old', name: 'Shared', placeholders: [{ key: 'body', kind: 'body', box }] });
  dest.defineSlideLayout({ id: 'different', name: 'Shared', placeholders: [{ key: 'body', kind: 'body', box: { ...box, width: 99 } }] });
  dest.defineSlideMaster({ id: 'master', name: 'Master' });
  const result = dest.copySlideLayoutFrom(source, 'old', { ...options, masterId: 'master', reuseEquivalent: true });
  assert.equal(result.layout.id, 'copied');
  assert.equal(result.layout.masterId, 'master');
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'other', reuseEquivalent: 'yes' }), /boolean/);
});

test('layout reuse checks nested reference topology and never commits staged resources on reuse', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(4) }, width: 8 });
  source.defineSlideLayout({ id: 'old', name: 'Nested', elements: [{ id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 40 }, children: [
    { id: 'image-shape', type: 'image', assetId: 'image', box },
    { id: 'edge', type: 'connector', start: { elementId: 'image-shape' }, end: { x: 20, y: 20 } },
  ] }] });
  const first = dest.copySlideLayoutFrom(source, 'old', options);
  const next = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'second', reuseEquivalent: true });
  assert.equal(next.layout, first.layout);
  assert.deepEqual(next.elementIds, first.elementIds);
  const original = source.layouts.find(layout => layout.id === 'old');
  source.replaceSlideLayout('old', { ...original, elements: [{ ...original.elements[0], children: [original.elements[0].children[0], { ...original.elements[0].children[1], start: { x: 0, y: 0 } }] }] });
  const changedTopology = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'topology', reuseEquivalent: true });
  assert.equal(changedTopology.layout.id, 'topology');
  source.replaceSlideLayout('old', original);
  source.replaceAsset('image', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(4) }, width: 9 });
  const changed = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'changed', reuseEquivalent: true });
  assert.equal(changed.layout.id, 'changed');
  assert.equal(dest.assets.length, 2);
  assert.notEqual(changed.assets[0].destinationId, first.assets[0].destinationId);
});

test('layout reuse distinguishes remapped page actions and external master targets', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'source-page' });
  source.defineSlideMaster({ id: 'source-master', name: 'Master', elements: [{ id: 'source-anchor', type: 'shape', shape: 'rect', box }] });
  source.defineSlideLayout({ id: 'old', name: 'Actions', masterId: 'source-master', elements: [
    { id: 'shape', type: 'shape', shape: 'rect', box, action: { type: 'slide', slideId: 'source-page' } },
    { id: 'edge', type: 'connector', start: { elementId: 'source-anchor' }, end: { x: 120, y: 10 } },
  ] });
  dest.addSlide({ id: 'page-a' }); dest.addSlide({ id: 'page-b' });
  dest.defineSlideMaster({ id: 'target-master', name: 'Master', elements: [
    { id: 'anchor-a', type: 'shape', shape: 'rect', box }, { id: 'anchor-b', type: 'shape', shape: 'rect', box },
  ] });
  const copy = (id, page, anchor) => dest.copySlideLayoutFrom(source, 'old', {
    ...options, id, masterId: 'target-master', reuseEquivalent: true,
    slideLinks: { 'source-page': page }, connectorTargets: { 'source-anchor': anchor },
  });
  const first = copy('first', 'page-a', 'anchor-a');
  assert.equal(copy('different-page', 'page-b', 'anchor-a').layout.id, 'different-page');
  assert.equal(copy('different-anchor', 'page-a', 'anchor-b').layout.id, 'different-anchor');
  assert.equal(copy('reuse', 'page-a', 'anchor-a').layout, first.layout);
  assert.equal(dest.layouts.length, 3);
});

test('layout reuse preserves a large forward-reference graph and detects a late difference', () => {
  const source = createPresentation(), dest = createPresentation();
  const elements = [];
  for (let i = 0; i < 150; i++) {
    elements.push({ id: `edge-${i}`, type: 'connector', start: { elementId: `shape-${i}` }, end: { x: 200, y: i } });
  }
  for (let i = 0; i < 150; i++) elements.push({ id: `shape-${i}`, type: 'shape', shape: 'rect', box: { ...box, y: i } });
  source.defineSlideLayout({ id: 'old', name: 'Graph', elements });
  const first = dest.copySlideLayoutFrom(source, 'old', options);
  const reused = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'next', reuseEquivalent: true });
  assert.equal(reused.layout, first.layout);
  for (let i = 0; i < 150; i++) {
    assert.equal(reused.elementIds[`shape-${i}`], first.elementIds[`shape-${i}`]);
    assert.equal(reused.layout.elements[i].start.elementId, reused.elementIds[`shape-${i}`]);
  }
  const original = source.layouts.find(layout => layout.id === 'old');
  source.replaceSlideLayout('old', { ...original, elements: original.elements.map((element, i) => i === 299 ? { ...element, box: { ...element.box, width: 101 } } : element) });
  const changed = dest.copySlideLayoutFrom(source, 'old', { ...options, id: 'next', reuseEquivalent: true });
  assert.equal(changed.layout.id, 'next');
  assert.equal(dest.layouts.length, 2);
});

test('layout copy reads the source layout collection once before selecting a definition', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'old', name: 'Original', placeholders: [{ key: 'body', kind: 'body', box }] });
  let reads = 0;
  const changing = Object.create(source, { layouts: { get() { reads++; return reads === 1 ? source.layouts : []; } } });
  const result = dest.copySlideLayoutFrom(changing, 'old', options);
  assert.equal(reads, 1);
  assert.equal(result.layout.name, 'Original');
  assert.equal(result.layout.placeholders[0].key, 'body');
});

test('layout copy captures mapping values before reentrant source access', () => {
  const source = createPresentation(), dest = createPresentation({ textStylePresets: { first: {}, second: {} } });
  source.defineSlideLayout({ id: 'old', name: 'Old', placeholders: [{ key: 'body', kind: 'body', box, textStylePreset: 'old' }] });
  let calls = 0;
  const getter = Object.defineProperty({}, 'old', { enumerable: true, get() { calls++; return 'first'; } });
  for (const field of ['slideLinks', 'connectorTargets', 'textPresets']) {
    assert.throws(() => dest.copySlideLayoutFrom(source, 'old', { ...options, [field]: getter }), TypeError);
    assert.equal(dest.layouts.length, 0); assert.equal(dest.assets.length, 0);
  }
  assert.equal(calls, 0);
  const mapping = { old: 'first' };
  const changing = Object.create(source, { assets: { get() { mapping.old = 'second'; return source.assets; } } });
  const result = dest.copySlideLayoutFrom(changing, 'old', { ...options, textPresets: mapping });
  assert.equal(mapping.old, 'second'); assert.equal(result.layout.placeholders[0].textStylePreset, 'first');
});

test('copies a detached layout with remapped artwork, assets and explicit destination master', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.defineSlideLayout({ id: 'old', name: 'Source layout', placeholders: [{ key: 'body', kind: 'body', box }], elements: [
    { id: 'edge', type: 'connector', start: { elementId: 'image' }, end: { x: 90, y: 20 } },
    { id: 'image', type: 'image', assetId: 'picture', box },
  ] });
  dest.defineSlideMaster({ id: 'master', name: 'Destination master' });
  const result = dest.copySlideLayoutFrom(source, 'old', { ...options, masterId: 'master' });
  assert.equal(result.layout.id, 'copied'); assert.equal(result.layout.masterId, 'master');
  assert.equal(result.layout.elements[0].start.elementId, result.elementIds.image);
  assert.equal(result.layout.elements[1].assetId, result.assets[0].destinationId);
  assert.equal(result.layout.placeholders[0].key, 'body');
  source.replaceAsset('picture', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } });
  assert.equal(dest.assets[0].source.value, 'AQI=');
  assert.equal(dest.addSlide({ layoutId: 'copied', elements: [{ id: 'body', type: 'text', placeholderKey: 'body', content: 'Body' }] }).layoutId, 'copied');
  assert.ok(Object.isFrozen(result.layout));
});

test('layout copy rejects missing links and late validation without installing assets or IDs', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.defineSlideLayout({ id: 'old', name: 'Old', elements: [{ id: 'image', type: 'image', box, assetId: 'picture', action: { type: 'slide', slideId: 'other' } }] });
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', options), /mapping/);
  assert.equal(dest.layouts.length, 0); assert.equal(dest.assets.length, 0);
  dest.addSlide({ id: 'target' });
  const result = dest.copySlideLayoutFrom(source, 'old', { ...options, slideLinks: { other: 'target' } });
  assert.equal(result.layout.elements[0].action.slideId, 'target'); assert.equal(result.elementIds.image, 'element-1');
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', { ...options, slideLinks: { other: 'target' } }), /Duplicate/);
  assert.equal(dest.layouts.length, 1); assert.equal(dest.assets.length, 1);
});

test('layout copy maps placeholder presets and rejects missing destination styles', () => {
  const source = createPresentation({ textStylePresets: { old: { run: { fontSize: 20 } } } });
  const dest = createPresentation({ textStylePresets: { next: { run: { fontSize: 22 } } } });
  source.defineSlideLayout({ id: 'old', name: 'Old', placeholders: [{ key: 'body', kind: 'body', box, textStylePreset: 'old' }] });
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', options), /preset/);
  assert.equal(dest.layouts.length, 0);
  const result = dest.copySlideLayoutFrom(source, 'old', { ...options, textPresets: { old: 'next' } });
  assert.equal(result.layout.placeholders[0].textStylePreset, 'next');
});

test('layout validation and reentrant resource edits do not partially commit the copied definition', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.defineSlideLayout({ id: 'bad', name: 'Bad', elements: [{ id: 'bad-image', type: 'image', assetId: 'image', box: { ...box, width: -1 } }] });
  assert.throws(() => dest.copySlideLayoutFrom(source, 'bad', options));
  assert.equal(dest.assets.length, 0); assert.equal(dest.layouts.length, 0);
  source.defineSlideLayout({ id: 'good', name: 'Good', elements: [{ id: 'good-image', type: 'image', assetId: 'image', box }] });
  const hostile = Object.create(source, { assets: { get() { dest.registerAsset({ id: 'independent', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } }); return source.assets; } } });
  assert.throws(() => dest.copySlideLayoutFrom(hostile, 'good', options), /changed/);
  assert.deepEqual(dest.assets.map(item => item.id), ['independent']); assert.equal(dest.layouts.length, 0);
  const result = dest.copySlideLayoutFrom(source, 'good', options);
  assert.equal(result.elementIds['good-image'], 'element-1');
});

test('legacy source layout with empty name cannot install an invalid shared definition', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'invalid-name', name: '' });
  assert.throws(() => dest.copySlideLayoutFrom(source, 'invalid-name', options), /name/);
  assert.equal(dest.layouts.length, 0);
  source.defineSlideLayout({ id: 'valid-name', name: 'Valid' });
  assert.equal(dest.copySlideLayoutFrom(source, 'valid-name', options).layout.id, 'copied');
});

test('external layout connector targets map into the selected destination master only', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideMaster({ id: 'master', name: 'Master', elements: [{ id: 'source-shared', type: 'shape', shape: 'rect', box }] });
  source.defineSlideLayout({ id: 'old', name: 'Old', masterId: 'master', elements: [{ id: 'edge', type: 'connector', start: { elementId: 'source-shared' }, end: { x: 150, y: 40 } }] });
  dest.defineSlideMaster({ id: 'target-master', name: 'Target', elements: [{ id: 'target-shared', type: 'shape', shape: 'rect', box }] });
  dest.addSlide({ id: 'other', elements: [{ id: 'outside', type: 'shape', shape: 'rect', box }] });
  const opts = { ...options, masterId: 'target-master', connectorTargets: { 'source-shared': 'outside' } };
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', opts), /destination/);
  assert.equal(dest.layouts.length, 0);
  const result = dest.copySlideLayoutFrom(source, 'old', { ...opts, connectorTargets: { 'source-shared': 'target-shared' } });
  assert.equal(result.layout.elements[0].start.elementId, 'target-shared');
});

test('layout external mapping cannot guess an identity being allocated in that layout', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'old', name: 'Old', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box }, { id: 'edge', type: 'connector', start: { elementId: 'unknown' }, end: { x: 10, y: 10 } }] });
  assert.throws(() => dest.copySlideLayoutFrom(source, 'old', { ...options, connectorTargets: { unknown: 'element-1' } }), /top-level destination/);
  assert.equal(dest.layouts.length, 0); assert.equal(dest.assets.length, 0);
});

test('layout copy detects unrelated reentrant definition changes even without artwork', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source' });
  const hostile = Object.create(source, { assets: { get() { dest.defineSlideLayout({ id: 'independent', name: 'Independent' }); return source.assets; } } });
  assert.throws(() => dest.copySlideLayoutFrom(hostile, 'source', options), /changed/);
  assert.deepEqual(dest.layouts.map(layout => layout.id), ['independent']);
  assert.equal(dest.copySlideLayoutFrom(source, 'source', options).layout.id, 'copied');
});
