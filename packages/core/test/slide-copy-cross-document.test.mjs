import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 100, height: 40 };
const policies = { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } };

test('cross-document slide copy remaps page, elements and images in one detached transaction', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2, 3) } });
  source.addSlide({ id: 'source', notes: [{ runs: [{ text: 'note', action: { type: 'slide', slideId: 'source' } }] }], elements: [
    { id: 'image', type: 'image', box, assetId: 'picture' },
    { id: 'shape', type: 'shape', shape: 'rect', box, action: { type: 'slide', slideId: 'source' } },
    { id: 'connector', type: 'connector', start: { elementId: 'shape' }, end: { elementId: 'image' } },
  ] });
  const before = JSON.stringify(source.slides[0]);
  const result = dest.copySlideFrom(source, 'source', policies);
  assert.equal(result.slide.id, 'slide-1');
  assert.equal(result.slide.elements[0].assetId, result.assets[0].destinationId);
  assert.equal(result.slide.elements[1].action.slideId, result.slide.id);
  assert.equal(result.slide.notes[0].runs[0].action.slideId, result.slide.id);
  assert.equal(result.slide.elements[2].start.elementId, result.elementIds.shape);
  assert.equal(result.slide.elements[2].end.elementId, result.elementIds.image);
  assert.equal(dest.assets.length, 1); assert.equal(JSON.stringify(source.slides[0]), before);
  source.removeSlide('source'); assert.equal(result.slide.elements.length, 3);
  assert.ok(Object.isFrozen(result)); assert.ok(Object.isFrozen(result.elementIds));
});

test('unmapped external slide links fail before committing assets or generated identities', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.addSlide({ id: 'source', elements: [{ id: 'image', type: 'image', box, assetId: 'picture', action: { type: 'slide', slideId: 'other' } }] });
  assert.throws(() => dest.copySlideFrom(source, 'source', policies), /mapping/);
  assert.equal(dest.assets.length, 0); assert.equal(dest.slides.length, 0);
  dest.addSlide({ id: 'target' });
  const result = dest.copySlideFrom(source, 'source', { ...policies, slideLinks: { other: 'target' } });
  assert.equal(result.slide.id, 'slide-1'); assert.equal(result.slide.elements[0].id, 'element-1');
  assert.equal(result.slide.elements[0].action.slideId, 'target');
});

test('late validation and source reentrancy cannot commit a partial copy', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.addSlide({ id: 'source', elements: [{ id: 'image', type: 'image', box: { ...box, width: -1 }, assetId: 'picture' }] });
  assert.throws(() => dest.copySlideFrom(source, 'source', policies));
  assert.equal(dest.assets.length, 0); assert.equal(dest.slides.length, 0);
  source.slides[0].replaceElement('image', { id: 'image', type: 'image', box, assetId: 'picture' });
  const hostile = Object.create(source, { assets: { get() { dest.registerAsset({ id: 'outside', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(8) } }); return source.assets; } } });
  assert.throws(() => dest.copySlideFrom(hostile, 'source', policies), /changed/);
  assert.deepEqual(dest.assets.map(asset => asset.id), ['outside']); assert.equal(dest.slides.length, 0);
  const result = dest.copySlideFrom(source, 'source', policies);
  assert.equal(result.slide.id, 'slide-1'); assert.equal(result.slide.elements[0].id, 'element-1');
});

test('copy reconciles placeholders, presets and timeline targets against the destination', () => {
  const source = createPresentation({ textStylePresets: { old: { run: { fontSize: 18 } } } });
  const dest = createPresentation({ textStylePresets: { next: { run: { fontSize: 22 } } } });
  dest.defineSlideLayout({ id: 'layout', name: 'Layout', placeholders: [{ key: 'new', role: 'body', box }] });
  const slide = source.addSlide({ id: 'source', elements: [
    { id: 'text', type: 'text', box, placeholderKey: 'old', textStylePreset: 'old', content: 'Hello' },
    { id: 'button', type: 'shape', shape: 'rect', box },
  ] });
  slide.setEventTimeline({ stages: [{ id: 'stage', target: { pageId: 'source', elementId: 'text' }, trigger: { type: 'object-click', elementId: 'button' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] });
  const options = { ...policies, layoutId: 'layout', placeholders: { old: 'new' }, textPresets: { old: 'next' } };
  const result = dest.copySlideFrom(source, 'source', options);
  assert.equal(result.slide.elements[0].placeholderKey, 'new'); assert.equal(result.slide.elements[0].textStylePreset, 'next');
  assert.equal(result.slide.eventTimeline.stages[0].target.pageId, result.slide.id);
  assert.equal(result.slide.eventTimeline.stages[0].target.elementId, result.elementIds.text);
  assert.equal(result.slide.eventTimeline.stages[0].trigger.elementId, result.elementIds.button);
  assert.throws(() => dest.copySlideFrom(source, 'source', { ...options, placeholders: {} }), /placeholder/);
  assert.throws(() => dest.copySlideFrom(source, 'source', { ...options, textPresets: {} }), /preset/);
  assert.equal(dest.slides.length, 1);
});

test('nested copied images share deduplicated destination bytes and local connector identities', () => {
  const source = createPresentation(), dest = createPresentation();
  for (const id of ['a', 'b']) source.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.addSlide({ id: 's', elements: [{ id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 40 }, children: [
    { id: 'edge', type: 'connector', start: { elementId: 'left' }, end: { elementId: 'right' } },
    { id: 'left', type: 'image', box, assetId: 'a' }, { id: 'right', type: 'image', box, assetId: 'b' },
  ] }] });
  const result = dest.copySlideFrom(source, 's', policies), [edge, left, right] = result.slide.elements[0].children;
  assert.equal(dest.assets.length, 1); assert.equal(left.assetId, right.assetId);
  assert.deepEqual(result.assets.map(item => item.reused), [false, true]);
  assert.equal(edge.start.elementId, left.id); assert.equal(edge.end.elementId, right.id);
  source.replaceAsset('a', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } });
  assert.equal(dest.assets[0].source.value, 'AQI=');
});

test('null layout resolves an explicitly defined default layout', () => {
  const source = createPresentation(), dest = createPresentation();
  dest.defineSlideLayout({ id: 'layout-blank', name: 'Default', placeholders: [{ key: 'body', role: 'body', box }] });
  source.addSlide({ id: 's', elements: [{ id: 'text', type: 'text', placeholderKey: 'body', content: 'hello' }] });
  const result = dest.copySlideFrom(source, 's', policies);
  assert.equal(result.slide.elements[0].placeholderKey, 'body');
});

test('occupied destination IDs are avoided and external connector targets fail closed', () => {
  const source = createPresentation(), dest = createPresentation();
  dest.addSlide({ id: 'slide-1', elements: [{ id: 'element-1', type: 'shape', shape: 'rect', box }] });
  source.addSlide({ id: 's', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box }, { id: 'edge', type: 'connector', start: { elementId: 'shape' }, end: { elementId: 'outside' } }] });
  assert.throws(() => dest.copySlideFrom(source, 's', policies), /external connector/);
  assert.equal(dest.slides.length, 1); assert.equal(dest.assets.length, 0);
  source.slides[0].replaceElement('edge', { id: 'edge', type: 'connector', start: { elementId: 'shape' }, end: { x: 100, y: 40 } });
  const copy = dest.copySlideFrom(source, 's', policies);
  assert.equal(copy.slide.id, 'slide-2'); assert.equal(copy.elementIds.shape, 'element-2');
  assert.equal(copy.slide.elements[1].start.elementId, copy.elementIds.shape);
});

test('external slide connector targets map into the selected layout and reject out-of-scope IDs', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source-layout', name: 'Source', elements: [{ id: 'source-shared', type: 'shape', shape: 'rect', box }] });
  source.addSlide({ id: 'source', layoutId: 'source-layout', elements: [{ id: 'edge', type: 'connector', start: { elementId: 'source-shared' }, end: { x: 150, y: 40 } }] });
  dest.defineSlideLayout({ id: 'target-layout', name: 'Target', elements: [{ id: 'target-shared', type: 'shape', shape: 'rect', box }] });
  dest.addSlide({ id: 'other', elements: [{ id: 'outside', type: 'shape', shape: 'rect', box }] });
  const opts = { ...policies, layoutId: 'target-layout', connectorTargets: { 'source-shared': 'outside' } };
  assert.throws(() => dest.copySlideFrom(source, 'source', opts), /destination/);
  assert.equal(dest.slides.length, 1);
  const result = dest.copySlideFrom(source, 'source', { ...opts, connectorTargets: { 'source-shared': 'target-shared' } });
  assert.equal(result.slide.elements[0].start.elementId, 'target-shared');
});

test('connector mapping rejects descendants in a different group coordinate scope', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'source', elements: [{ id: 'edge', type: 'connector', start: { elementId: 'shared' }, end: { x: 150, y: 40 } }] });
  dest.defineSlideLayout({ id: 'layout', name: 'Layout', elements: [{ id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 40 }, children: [{ id: 'nested', type: 'shape', shape: 'rect', box }] }] });
  const opts = { ...policies, layoutId: 'layout', connectorTargets: { shared: 'nested' } };
  assert.throws(() => dest.copySlideFrom(source, 'source', opts), /destination/);
  assert.equal(dest.slides.length, 0);
  const result = dest.copySlideFrom(source, 'source', { ...opts, connectorTargets: { shared: 'group' } });
  assert.equal(result.slide.elements[0].start.elementId, 'group');
});

test('external connector mapping cannot bind to an identity allocated by the same copy', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'source', elements: [{ id: 'shape', type: 'shape', shape: 'rect', box }, { id: 'edge', type: 'connector', start: { elementId: 'unrelated' }, end: { x: 10, y: 10 } }] });
  assert.throws(() => dest.copySlideFrom(source, 'source', { ...policies, connectorTargets: { unrelated: 'element-1' } }), /top-level destination/);
  assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
});
