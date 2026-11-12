import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 100, height: 40 };
const options = { layoutId: null, theme: 'destination', geometry: 'preserve', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } };

test('explicit custom-data policy rejects opaque metadata atomically or preserves it without rewriting', () => {
  const source = createPresentation(), target = createPresentation();
  source.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.addSlide({ id: 'a', elements: [{ id: 'picture', type: 'image', assetId: 'image', box }] });
  source.addSlide({ id: 'b', customData: { reference: 'a', nested: { target: 'picture', values: [null, false, 0] } } });
  assert.throws(() => target.copySlidesFrom(source, ['a', 'b'], { ...options, customData: 'reject' }), /opaque custom data.*b/);
  assert.equal(target.slides.length, 0); assert.equal(target.assets.length, 0);
  assert.throws(() => target.copySlideFrom(source, 'b', { ...options, customData: 'reject' }), /opaque custom data/);
  const copied = target.copySlidesFrom(source, ['a', 'b'], options);
  assert.deepEqual(copied.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
  assert.equal(copied.assets[0].destinationId, 'asset-1');
  assert.deepEqual(copied.slides[1].customData, { reference: 'a', nested: { target: 'picture', values: [null, false, 0] } });
  assert.notEqual(copied.slides[1].customData, source.slides[1].customData);
  const safe = createPresentation();
  assert.equal(safe.copySlideFrom(source, 'a', { ...options, customData: 'reject' }).slide.id, 'slide-1');
  const special = createPresentation();
  special.addSlide({ id: 'special', customData: Object.fromEntries([['__proto__', null]]) });
  assert.throws(() => safe.copySlideFrom(special, 'special', { ...options, customData: 'reject' }), /opaque custom data/);
  assert.throws(() => safe.copySlideFrom(special, 'special', { ...options, customData: 'discard' }), /policies/);
});

test('custom-data rejection sees hidden and symbol fields before cloning custom source slides', () => {
  const source = createPresentation(), target = createPresentation(); source.addSlide({ id: 'a' });
  const hidden = Object.defineProperty({}, 'hidden', { value: 'a', enumerable: false });
  const symbol = { [Symbol('opaque')]: 'a' };
  for (const customData of [hidden, symbol]) {
    const slide = Object.create(source.slides[0], { customData: { value: customData } });
    const view = Object.create(source, { slides: { value: [slide] } });
    assert.throws(() => target.copySlideFrom(view, 'a', { ...options, customData: 'reject' }), /opaque custom data/);
    assert.equal(target.slides.length, 0); assert.equal(target.assets.length, 0);
  }
  assert.equal(target.copySlideFrom(source, 'a', { ...options, customData: 'reject' }).slide.id, 'slide-1');
});

test('fit explicit geometry centers differing page sizes without changing source or group-local coordinates', () => {
  const source = createPresentation({ size: { width: 400, height: 200 } });
  const dest = createPresentation({ size: { width: 200, height: 200 } });
  const point = { x: 40, y: 20 };
  const slide = source.addSlide({ id: 'a', elements: [
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 20, y: 30, width: 40, height: 20 } },
    { id: 'edge', type: 'connector', start: point, end: { elementId: 'shape', anchor: 'right' }, route: [point] },
    { id: 'group', type: 'group', box: { x: 100, y: 40, width: 80, height: 40 }, coordinateSize: { width: 80, height: 40 }, children: [{ id: 'child', type: 'shape', shape: 'rect', box: { x: 4, y: 6, width: 20, height: 10 } }] },
    { id: 'table', type: 'table', box: { x: 200, y: 40, width: 100, height: 40 }, columns: [40, 60], rows: [{ height: 40, cells: [{ content: 'A' }, { content: 'B' }] }] },
    { id: 'text', type: 'text', box: { x: 20, y: 120, width: 100 }, content: 'Unscaled text' },
  ] });
  const original = structuredClone(slide.elements);
  const { slide: copy } = dest.copySlideFrom(source, 'a', { ...options, geometry: 'fit-explicit-geometry' });
  assert.deepEqual(copy.elements[0].box, { x: 10, y: 65, width: 20, height: 10 });
  assert.deepEqual(copy.elements[1].start, { x: 20, y: 60 });
  assert.deepEqual(copy.elements[1].route, [{ x: 20, y: 60 }]);
  assert.equal(copy.elements[1].end.elementId, copy.elements[0].id);
  assert.deepEqual(copy.elements[2].box, { x: 50, y: 70, width: 40, height: 20 });
  assert.deepEqual(copy.elements[2].children[0].box, { x: 4, y: 6, width: 20, height: 10 });
  assert.deepEqual(copy.elements[3].columns, [20, 30]); assert.equal(copy.elements[3].rows[0].height, 20);
  assert.deepEqual(copy.elements[4].box, { x: 10, y: 110, width: 50 });
  assert.deepEqual(slide.elements, original);
});

test('fit geometry rejects invalid source dimensions atomically and supports batch retry', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' }); source.addSlide({ id: 'b' });
  const hostile = Object.create(source, { size: { value: { width: 0, height: 100, unit: 'pt' } } });
  assert.throws(() => dest.copySlidesFrom(hostile, ['a', 'b'], { ...options, geometry: 'fit-explicit-geometry' }), /dimensions/);
  assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
  assert.deepEqual(dest.copySlidesFrom(source, ['a', 'b'], { ...options, geometry: 'fit-explicit-geometry' }).slides.map(slide => slide.id), ['slide-1', 'slide-2']);
});

test('fit geometry rejects a non-finite scale after resource staging without consuming IDs', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2, 3) } });
  source.addSlide({ id: 'a', elements: [{ id: 'image', type: 'image', assetId: 'picture', box: { x: 0, y: 0, width: 10, height: 10 } }] });
  const original = structuredClone(source.slides[0].elements);
  const tiny = Object.create(source, { size: { value: { width: 1e-308, height: 1e-308, unit: 'pt' } } });
  assert.throws(() => dest.copySlideFrom(tiny, 'a', { ...options, geometry: 'fit-explicit-geometry' }), /scale/);
  assert.equal(dest.slides.length, 0);
  assert.equal(dest.assets.length, 0);
  assert.deepEqual(source.slides[0].elements, original);
  const retry = dest.copySlideFrom(source, 'a', { ...options, geometry: 'fit-explicit-geometry' });
  assert.equal(retry.slide.id, 'slide-1');
  assert.equal(retry.elementIds.image, 'element-1');
  assert.equal(retry.assets[0].destinationId, 'asset-1');
});

test('fit geometry centers horizontal spare space and leaves target-inherited placeholder geometry alone', () => {
  const source = createPresentation({ size: { width: 100, height: 200 } });
  const dest = createPresentation({ size: { width: 400, height: 200 } });
  source.addSlide({ id: 'a', elements: [
    { id: 'shape', type: 'shape', shape: 'rect', box: { x: 5, y: 10, width: 20, height: 30 } },
    { id: 'text', type: 'text', placeholderKey: 'body', content: 'Inherited' },
  ] });
  dest.defineSlideLayout({ id: 'l', name: 'L', placeholders: [{ key: 'body', kind: 'body', box: { x: 10, y: 20, width: 80, height: 30 } }] });
  const { slide } = dest.copySlideFrom(source, 'a', { ...options, layoutId: 'l', geometry: 'fit-explicit-geometry' });
  assert.deepEqual(slide.elements[0].box, { x: 155, y: 10, width: 20, height: 30 });
  assert.equal(slide.elements[1].box, undefined);
  assert.equal(slide.elements[1].placeholderKey, 'body');
});

test('fit geometry detects reentrant target edits while reading source size', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' });
  const hostile = Object.create(source, { size: { get() { dest.addSlide({ id: 'independent' }); return source.size; } } });
  assert.throws(() => dest.copySlideFrom(hostile, 'a', { ...options, geometry: 'fit-explicit-geometry' }), /changed/);
  assert.deepEqual(dest.slides.map(slide => slide.id), ['independent']);
  assert.equal(dest.copySlideFrom(source, 'a', { ...options, geometry: 'fit-explicit-geometry' }).slide.id, 'slide-1');
});

test('batch copy remaps forward and backward links, notes, resources and selected order', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'shared', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1, 2) } });
  source.addSlide({ id: 'a', notes: [{ runs: [{ text: 'to b', action: { type: 'slide', slideId: 'b' } }] }], elements: [{ id: 'a-image', type: 'image', assetId: 'shared', box, action: { type: 'slide', slideId: 'b' } }] });
  source.addSlide({ id: 'b', elements: [{ id: 'b-image', type: 'image', assetId: 'shared', box, action: { type: 'slide', slideId: 'a' } }] });
  dest.addSlide({ id: 'existing' });
  const result = dest.copySlidesFrom(source, ['b', 'a'], { ...options, index: 0 });
  assert.deepEqual(dest.slides.map(slide => slide.id), ['slide-1', 'slide-2', 'existing']);
  assert.equal(result.slideIds.b, 'slide-1'); assert.equal(result.slideIds.a, 'slide-2');
  assert.equal(result.slides[0].elements[0].action.slideId, result.slideIds.a);
  assert.equal(result.slides[1].elements[0].action.slideId, result.slideIds.b);
  assert.equal(result.slides[1].notes[0].runs[0].action.slideId, result.slideIds.b);
  assert.equal(result.slides[0].elements[0].assetId, result.slides[1].elements[0].assetId);
  assert.equal(dest.assets.length, 1); assert.ok(Object.isFrozen(result.slides)); assert.ok(Object.isFrozen(result.elementIds.a));
  source.removeSlide('a'); source.removeSlide('b'); assert.equal(result.slides.length, 2);
});

test('a late invalid slide rolls back every staged page, resource and allocator', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'asset', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.addSlide({ id: 'a', elements: [{ id: 'image', type: 'image', assetId: 'asset', box }] });
  source.addSlide({ id: 'b', elements: [{ id: 'bad', type: 'shape', shape: 'rect', box: { ...box, width: -1 } }] });
  assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], options), /validation/);
  assert.equal(dest.assets.length, 0); assert.equal(dest.slides.length, 0);
  source.slides[1].replaceElement('bad', { id: 'bad', type: 'shape', shape: 'rect', box });
  const result = dest.copySlidesFrom(source, ['a', 'b'], options);
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
  assert.equal(result.elementIds.a.image, 'element-1'); assert.equal(result.elementIds.b.bad, 'element-2');
});

test('batch copy selects layouts per page and maps links outside the batch', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'a', elements: [{ id: 'a-text', type: 'text', box, content: 'A', placeholderKey: 'a' }] });
  source.addSlide({ id: 'b', elements: [{ id: 'b-text', type: 'text', box, content: 'B', placeholderKey: 'b', action: { type: 'slide', slideId: 'external' } }] });
  dest.defineSlideLayout({ id: 'layout-a', name: 'A', placeholders: [{ key: 'a', kind: 'body', box }] });
  dest.defineSlideLayout({ id: 'layout-b', name: 'B', placeholders: [{ key: 'b', kind: 'body', box }] });
  dest.addSlide({ id: 'target' });
  const result = dest.copySlidesFrom(source, ['a', 'b'], { ...options, layouts: { a: 'layout-a', b: 'layout-b' }, slideLinks: { external: 'target' } });
  assert.deepEqual(result.slides.map(slide => slide.layoutId), ['layout-a', 'layout-b']);
  assert.equal(result.slides[1].elements[0].action.slideId, 'target');
});

test('optional strict placeholder role matching rejects mismatches atomically, including explicit mappings', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'heading', kind: 'title', index: 5, box }] });
  for (const id of ['a', 'b']) source.addSlide({ id, layoutId: 'source', elements: [
    { id: `text-${id}`, type: 'text', placeholderKey: 'heading', content: id },
    { id: `image-${id}`, type: 'image', assetId: 'picture', box },
  ] });
  dest.defineSlideLayout({ id: 'good', name: 'Good', placeholders: [{ key: 'heading', kind: 'title', index: 8, box }] });
  dest.defineSlideLayout({ id: 'bad', name: 'Bad', placeholders: [{ key: 'heading', kind: 'body', index: 5, box }] });
  const strict = { ...options, layoutId: 'good', layouts: { b: 'bad' }, placeholderRoles: 'require-match' };
  assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], strict), /role mismatch/);
  assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
  assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], { ...strict, perSlide: { b: { placeholders: { heading: 'heading' } } } }), /role mismatch/);
  assert.throws(() => dest.copySlideFrom(source, 'a', { ...options, layoutId: 'good', placeholderRoles: 'other' }), /role policy/);
  const result = dest.copySlidesFrom(source, ['a', 'b'], { ...strict, layouts: { b: 'good' } });
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
  assert.equal(result.assets[0].destinationId, 'asset-1');
  assert.deepEqual(result.slides.map(slide => slide.elements[0].placeholderKey), ['heading', 'heading']);
  const allowChange = createPresentation();
  allowChange.defineSlideLayout({ id: 'bad', name: 'Bad', placeholders: [{ key: 'heading', kind: 'body', box }] });
  assert.equal(allowChange.copySlideFrom(source, 'a', { ...options, layoutId: 'bad' }).slide.elements[0].placeholderKey, 'heading');
});

test('strict placeholder role matching requires a valid source binding and accepts explicit same-role renaming', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'body', box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'new', kind: 'body', box }] });
  const policy = { ...options, layoutId: 'target', placeholderRoles: 'require-match', placeholders: { old: 'new' } };
  assert.equal(dest.copySlideFrom(source, 'a', policy).slide.elements[0].placeholderKey, 'new');
  const invalidSource = createPresentation();
  invalidSource.addSlide({ id: 'a', elements: [{ id: 'text', type: 'text', box, placeholderKey: 'old', content: 'A' }] });
  assert.throws(() => dest.copySlideFrom(invalidSource, 'a', policy), /source layout/);
  assert.equal(dest.slides.length, 1);
  invalidSource.defineSlideLayout({ id: 'layout-blank', name: 'Blank', placeholders: [{ key: 'different', kind: 'body', box }] });
  assert.throws(() => dest.copySlideFrom(invalidSource, 'a', policy), /source placeholder/);
  assert.equal(dest.slides.length, 1);
  const ambiguous = createPresentation();
  ambiguous.defineSlideLayout({ id: 'target', name: 'Ambiguous', placeholders: [{ key: 'new', kind: 'body', box }, { key: 'new', kind: 'title', box }] });
  assert.throws(() => ambiguous.copySlideFrom(source, 'a', policy), /one destination placeholder/);
  assert.equal(ambiguous.slides.length, 0);
});

test('strict role matching rejects duplicate source layout identities and reentrant target changes', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'body', kind: 'body', box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'body', content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'body', kind: 'body', box }] });
  const policy = { ...options, layoutId: 'target', placeholderRoles: 'require-match' };
  const duplicate = Object.create(source, { layouts: { value: [source.layouts[0], source.layouts[0]] } });
  assert.throws(() => dest.copySlideFrom(duplicate, 'a', policy), /one source layout/);
  const hostile = Object.create(source, { layouts: { get() { dest.addSlide({ id: 'independent' }); return source.layouts; } } });
  assert.throws(() => dest.copySlideFrom(hostile, 'a', policy), /changed/);
  assert.deepEqual(dest.slides.map(slide => slide.id), ['independent']); assert.equal(dest.assets.length, 0);
  assert.equal(dest.copySlideFrom(source, 'a', policy).slide.id, 'slide-1');
});

test('strict role and index matching accepts explicit boundary indexes and rejects missing or conflicting identities', () => {
  for (const index of [0, 4294967295]) {
    const source = createPresentation(), dest = createPresentation();
    source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'body', index, box }] });
    source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
    dest.defineSlideLayout({ id: 'good', name: 'Good', placeholders: [{ key: 'new', kind: 'body', index, box }] });
    dest.defineSlideLayout({ id: 'wrong', name: 'Wrong', placeholders: [{ key: 'new', kind: 'body', index: index === 0 ? 1 : 0, box }] });
    dest.defineSlideLayout({ id: 'missing', name: 'Missing', placeholders: [{ key: 'new', kind: 'body', box }] });
    dest.defineSlideLayout({ id: 'duplicate', name: 'Duplicate', placeholders: [{ key: 'new', kind: 'body', index, box }, { key: 'other', kind: 'body', index, box }] });
    const policy = { ...options, layoutId: 'good', placeholderRoles: 'require-match-and-index', placeholders: { old: 'new' } };
    for (const [layoutId, error] of [['wrong', /index mismatch/], ['missing', /explicit valid/], ['duplicate', /unique destination/]]) {
      assert.throws(() => dest.copySlideFrom(source, 'a', { ...policy, layoutId }), error);
      assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
    }
    assert.equal(dest.copySlideFrom(source, 'a', policy).slide.id, 'slide-1');
  }
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'body', box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'old', kind: 'body', index: 0, box }] });
  assert.throws(() => dest.copySlideFrom(source, 'a', { ...options, layoutId: 'target', placeholderRoles: 'require-match-and-index' }), /explicit valid/);
  assert.equal(dest.slides.length, 0);
});

test('strict index matching rejects invalid or duplicate indexes in the source layout', () => {
  for (const index of [-1, 1.5, 4294967296]) {
    const source = createPresentation(), dest = createPresentation();
    source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'body', kind: 'body', index, box }] });
    source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'body', content: 'A' }] });
    dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'body', kind: 'body', index: 0, box }] });
    assert.throws(() => dest.copySlideFrom(source, 'a', { ...options, layoutId: 'target', placeholderRoles: 'require-match-and-index' }), /invalid or duplicate placeholder indexes/);
    assert.equal(dest.slides.length, 0);
  }
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'body', kind: 'body', index: 0, box }, { key: 'other', kind: 'body', index: 0, box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'body', content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'body', kind: 'body', index: 0, box }] });
  assert.throws(() => dest.copySlideFrom(source, 'a', { ...options, layoutId: 'target', placeholderRoles: 'require-match-and-index' }), /invalid or duplicate placeholder indexes/);
});

test('explicit-index matching resolves renamed placeholders per destination layout without guessing', () => {
  const source = createPresentation(), dest = createPresentation();
  source.registerAsset({ id: 'picture', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'heading', kind: 'title', index: 0, box }, { key: 'body', kind: 'body', index: 4294967295, box }] });
  for (const id of ['a', 'b']) source.addSlide({ id, layoutId: 'source', elements: [
    { id: `h-${id}`, type: 'text', placeholderKey: 'heading', content: `Heading ${id}` },
    { id: `b-${id}`, type: 'text', placeholderKey: 'body', content: `Body ${id}` },
    { id: `i-${id}`, type: 'image', assetId: 'picture', box },
  ] });
  for (const id of ['a', 'b']) dest.defineSlideLayout({ id: `layout-${id}`, name: id, placeholders: [{ key: `h-${id}`, kind: 'title', index: 0, box }, { key: `b-${id}`, kind: 'body', index: 4294967295, box }] });
  const policy = { ...options, layoutId: null, layouts: { a: 'layout-a', b: 'layout-b' }, placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index' };
  assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], { ...policy, perSlide: { b: { placeholders: { body: 'h-b' } } } }), /role mismatch/);
  assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
  const original = JSON.stringify(source.slides.map(slide => slide.elements));
  const result = dest.copySlidesFrom(source, ['a', 'b'], policy);
  assert.deepEqual(result.slides.map(slide => slide.elements.slice(0, 2).map(element => element.placeholderKey)), [['h-a', 'b-a'], ['h-b', 'b-b']]);
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']); assert.equal(result.assets[0].destinationId, 'asset-1');
  assert.equal(JSON.stringify(source.slides.map(slide => slide.elements)), original);
});

test('explicit-index matching rejects missing, duplicated or role-conflicting candidates and incompatible policies', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'title', index: 0, box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
  dest.defineSlideLayout({ id: 'missing', name: 'Missing', placeholders: [{ key: 'new', kind: 'title', box }] });
  dest.defineSlideLayout({ id: 'duplicate', name: 'Duplicate', placeholders: [{ key: 'a', kind: 'title', index: 0, box }, { key: 'b', kind: 'title', index: 0, box }] });
  dest.defineSlideLayout({ id: 'role', name: 'Role', placeholders: [{ key: 'new', kind: 'body', index: 0, box }] });
  dest.defineSlideLayout({ id: 'good', name: 'Good', placeholders: [{ key: 'new', kind: 'title', index: 0, box }] });
  const policy = { ...options, layoutId: 'good', placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index' };
  for (const [layoutId, error] of [['missing', /one destination placeholder with index/], ['duplicate', /one destination placeholder with index/], ['role', /role mismatch/]]) {
    assert.throws(() => dest.copySlideFrom(source, 'a', { ...policy, layoutId }), error);
    assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
  }
  for (const role of [undefined, 'require-match']) assert.throws(() => dest.copySlideFrom(source, 'a', { ...policy, placeholderRoles: role }), /requires strict/);
  assert.throws(() => dest.copySlideFrom(source, 'a', { ...policy, placeholderMatching: 'guess' }), /requires strict/);
  assert.equal(dest.copySlideFrom(source, 'a', policy).slide.elements[0].placeholderKey, 'new');
  const missingSource = createPresentation();
  missingSource.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'title', box }] });
  missingSource.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
  assert.throws(() => dest.copySlideFrom(missingSource, 'a', policy), /explicit valid source placeholder index/);
  assert.equal(dest.slides.length, 1);
});

test('explicit-index matching honors per-slide then shared mappings before automatic selection', () => {
  const source = createPresentation(), dest = createPresentation();
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: [{ key: 'old', kind: 'body', index: 0, box }] });
  source.addSlide({ id: 'a', layoutId: 'source', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'matched', kind: 'body', index: 0, box }, { key: 'wrong', kind: 'body', index: 1, box }] });
  const policy = { ...options, layoutId: 'target', placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index' };
  assert.throws(() => dest.copySlidesFrom(source, ['a'], { ...policy, placeholders: { old: 'wrong' } }), /index mismatch/);
  assert.equal(dest.slides.length, 0);
  const copied = dest.copySlidesFrom(source, ['a'], { ...policy, placeholders: { old: 'wrong' }, perSlide: { a: { placeholders: { old: 'matched' } } } });
  assert.equal(copied.slides[0].elements[0].placeholderKey, 'matched');
});

test('large shared-layout batches map every renamed placeholder by its explicit index', () => {
  const source = createPresentation(), dest = createPresentation(), count = 900;
  const indexes = Array.from({ length: count }, (_, i) => i);
  source.defineSlideLayout({ id: 'source', name: 'Source', placeholders: indexes.map(index => ({ key: `source-${index}`, kind: 'body', index, box })) });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [...indexes].reverse().map(index => ({ key: `target-${index}`, kind: 'body', index, box })) });
  for (const id of ['a', 'b']) source.addSlide({ id, layoutId: 'source', elements: indexes.map(index => ({ id: `${id}-${index}`, type: 'text', placeholderKey: `source-${index}`, content: `${id}:${index}` })) });
  const result = dest.copySlidesFrom(source, ['b', 'a'], { ...options, layoutId: 'target', placeholderRoles: 'require-match-and-index', placeholderMatching: 'explicit-index' });
  assert.equal(result.slides.length, 2);
  for (const [page, id] of ['b', 'a'].entries()) {
    for (const index of indexes) {
      const element = result.slides[page].elements[index];
      assert.equal(element.placeholderKey, `target-${index}`); assert.equal(element.content, `${id}:${index}`);
    }
  }
});

test('batch copy rejects duplicate IDs and detects reentrant source capture without partial commit', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' }); source.addSlide({ id: 'b' });
  assert.throws(() => dest.copySlidesFrom(source, ['a', 'a'], options), /distinct/);
  const hostile = Object.create(source, { assets: { get() { dest.addSlide({ id: 'independent' }); return source.assets; } } });
  assert.throws(() => dest.copySlidesFrom(hostile, ['a', 'b'], options), /changed/);
  assert.deepEqual(dest.slides.map(slide => slide.id), ['independent']);
  const result = dest.copySlidesFrom(source, ['a', 'b'], options);
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
});

test('batch copy retains per-page timeline target and trigger ownership', () => {
  const source = createPresentation(), dest = createPresentation();
  for (const id of ['a', 'b']) {
    const slide = source.addSlide({ id, elements: [{ id: `${id}-target`, type: 'shape', shape: 'rect', box }, { id: `${id}-button`, type: 'shape', shape: 'rect', box }] });
    slide.setEventTimeline({ stages: [{ id: 'stage', target: { pageId: id, elementId: `${id}-target` }, trigger: { type: 'object-click', elementId: `${id}-button` }, timing: { kind: 'span', id: 'time', durationMs: 10 } }] });
  }
  const result = dest.copySlidesFrom(source, ['a', 'b'], options);
  for (let index = 0; index < 2; index++) {
    const sourceId = ['a', 'b'][index], slide = result.slides[index], stage = slide.eventTimeline.stages[0];
    assert.equal(stage.target.pageId, slide.id); assert.equal(stage.target.elementId, result.elementIds[sourceId][`${sourceId}-target`]);
    assert.equal(stage.trigger.elementId, result.elementIds[sourceId][`${sourceId}-button`]);
  }
});

test('batch copy detects a reentrantly added empty layout before commit', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'a' }); source.addSlide({ id: 'b' });
  const hostile = Object.create(source, { assets: { get() { dest.defineSlideLayout({ id: 'late', name: 'Late' }); return source.assets; } } });
  assert.throws(() => dest.copySlidesFrom(hostile, ['a', 'b'], { ...options, layoutId: 'late' }), /changed/);
  assert.equal(dest.slides.length, 0); assert.deepEqual(dest.layouts.map(layout => layout.id), ['late']);
  const retry = dest.copySlidesFrom(source, ['a', 'b'], { ...options, layoutId: 'late' });
  assert.deepEqual(retry.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
});

test('per-page overrides reconcile the same source references into different target layouts', () => {
  const source = createPresentation({ textStylePresets: { old: { run: { fontSize: 18 } } } });
  const dest = createPresentation({ textStylePresets: { first: { run: { fontSize: 20 } }, second: { run: { fontSize: 24 } } } });
  source.defineSlideLayout({ id: 'source-layout', name: 'Source', elements: [{ id: 'shared', type: 'shape', shape: 'rect', box }] });
  for (const [id, other] of [['a', 'b'], ['b', 'a']]) source.addSlide({ id, layoutId: 'source-layout', elements: [
    { id: `${id}-text`, type: 'text', box, content: id, placeholderKey: 'title', textStylePreset: 'old', action: { type: 'slide', slideId: other } },
    { id: `${id}-edge`, type: 'connector', start: { elementId: 'shared' }, end: { x: 100, y: 50 } },
  ] });
  for (const id of ['a', 'b']) dest.defineSlideLayout({ id: `layout-${id}`, name: id, placeholders: [{ key: `title-${id}`, kind: 'title', box }], elements: [{ id: `shared-${id}`, type: 'shape', shape: 'rect', box }] });
  const result = dest.copySlidesFrom(source, ['a', 'b'], { ...options, layouts: { a: 'layout-a', b: 'layout-b' }, textPresets: { old: 'first' }, perSlide: {
    a: { placeholders: { title: 'title-a' }, connectorTargets: { shared: 'shared-a' } },
    b: { placeholders: { title: 'title-b' }, textPresets: { old: 'second' }, connectorTargets: { shared: 'shared-b' } },
  } });
  assert.deepEqual(result.slides.map(slide => slide.elements[0].placeholderKey), ['title-a', 'title-b']);
  assert.deepEqual(result.slides.map(slide => slide.elements[0].textStylePreset), ['first', 'second']);
  assert.deepEqual(result.slides.map(slide => slide.elements[1].start.elementId), ['shared-a', 'shared-b']);
  assert.equal(result.slides[0].elements[0].action.slideId, result.slideIds.b);
  assert.equal(result.slides[1].elements[0].action.slideId, result.slideIds.a);
});

test('invalid per-page mappings reject the entire batch and unknown page keys', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' }); source.addSlide({ id: 'b' });
  for (const perSlide of [{ missing: {} }, { a: { unknown: {} } }, { a: { placeholders: { key: '' } } }, { a: null }]) {
    assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], { ...options, perSlide }), TypeError);
    assert.equal(dest.slides.length, 0);
  }
  assert.deepEqual(dest.copySlidesFrom(source, ['a', 'b'], options).slides.map(slide => slide.id), ['slide-1', 'slide-2']);
});

test('per-page overrides reject accessors, hidden fields and symbols before invoking them', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' });
  let calls = 0;
  const getter = Object.defineProperty({}, 'a', { enumerable: true, get() { calls++; return {}; } });
  const hidden = Object.defineProperty({}, 'missing', { value: {}, enumerable: false });
  const symbol = { a: {}, [Symbol('hidden')]: {} };
  const nested = { a: Object.defineProperty({}, 'placeholders', { enumerable: true, get() { calls++; return {}; } }) };
  for (const perSlide of [getter, hidden, symbol, nested]) {
    assert.throws(() => dest.copySlidesFrom(source, ['a'], { ...options, perSlide }), TypeError);
    assert.equal(dest.slides.length, 0);
  }
  assert.equal(calls, 0);
});

test('layout overrides reject unknown pages and non-data properties without partial copy', () => {
  const source = createPresentation(), dest = createPresentation(); source.addSlide({ id: 'a' }); source.addSlide({ id: 'b' });
  let calls = 0;
  const getter = Object.defineProperty({}, 'a', { enumerable: true, get() { calls++; return null; } });
  const hidden = Object.defineProperty({}, 'a', { value: null, enumerable: false });
  const symbol = { a: null, [Symbol('hidden')]: null };
  for (const layouts of [{ missing: null }, getter, hidden, symbol, { a: '' }, { b: undefined }]) {
    assert.throws(() => dest.copySlidesFrom(source, ['a', 'b'], { ...options, layouts }), TypeError);
    assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
  }
  assert.equal(calls, 0);
  const result = dest.copySlidesFrom(source, ['a', 'b'], { ...options, layouts: { a: null } });
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
});

test('shared mappings reject non-data properties and remain stable during source capture', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: 'a', elements: [{ id: 'text', type: 'text', placeholderKey: 'old', box, content: 'A' }] });
  dest.defineSlideLayout({ id: 'target', name: 'Target', placeholders: [{ key: 'first', kind: 'body', box }, { key: 'second', kind: 'body', box }] });
  let calls = 0;
  const getter = Object.defineProperty({}, 'old', { enumerable: true, get() { calls++; return 'first'; } });
  const hidden = Object.defineProperty({}, 'hidden', { value: 'first', enumerable: false });
  const symbol = { old: 'first', [Symbol('extra')]: 'first' };
  for (const field of ['slideLinks', 'connectorTargets', 'placeholders', 'textPresets']) {
    for (const mapping of [getter, hidden, symbol, { old: '' }]) {
      assert.throws(() => dest.copySlideFrom(source, 'a', { ...options, layoutId: 'target', [field]: mapping }), TypeError);
      assert.equal(dest.slides.length, 0); assert.equal(dest.assets.length, 0);
    }
  }
  assert.equal(calls, 0);
  const mapping = { old: 'first' };
  const changing = Object.create(source, { assets: { get() { mapping.old = 'second'; return source.assets; } } });
  const result = dest.copySlideFrom(changing, 'a', { ...options, layoutId: 'target', placeholders: mapping });
  assert.equal(mapping.old, 'second'); assert.equal(result.slide.elements[0].placeholderKey, 'first');
});

test('special source IDs are represented as own batch mapping keys', () => {
  const source = createPresentation(), dest = createPresentation();
  source.addSlide({ id: '__proto__', elements: [{ id: 'item', type: 'text', box, content: 'special' }] });
  const perSlide = Object.fromEntries([['__proto__', {}]]);
  const layouts = Object.fromEntries([['__proto__', null]]);
  const result = dest.copySlidesFrom(source, ['__proto__'], { ...options, layouts, perSlide });
  assert.equal(Object.hasOwn(result.slideIds, '__proto__'), true);
  assert.equal(result.slideIds.__proto__, result.slides[0].id);
  assert.equal(result.elementIds.__proto__.item, result.slides[0].elements[0].id);
  assert.equal(Object.getPrototypeOf(result.slideIds), Object.prototype);
});

test('large reversed batches preserve order, unique identities and a cyclic link graph', () => {
  const source = createPresentation(), dest = createPresentation(), count = 1200, ids = Array.from({ length: count }, (_, index) => `source-${index}`);
  for (const [index, id] of ids.entries()) source.addSlide({ id, elements: [{ id: `text-${index}`, type: 'text', box, content: `${index}`, action: { type: 'slide', slideId: ids[(index + 1) % count] } }] });
  dest.addSlide({ id: 'before' }); dest.addSlide({ id: 'after' });
  const selection = [...ids].reverse(), result = dest.copySlidesFrom(source, selection, { ...options, index: 1 });
  const all = dest.slides;
  assert.equal(all.length, count + 2); assert.equal(all[0].id, 'before'); assert.equal(all[count + 1].id, 'after');
  assert.equal(new Set(result.slides.map(slide => slide.id)).size, count);
  assert.equal(new Set(result.slides.flatMap(slide => slide.elements.map(element => element.id))).size, count);
  for (const [position, id] of selection.entries()) {
    const n = Number(id.slice('source-'.length)), slide = result.slides[position];
    assert.equal(all[position + 1], slide);
    assert.equal(slide.elements[0].action.slideId, result.slideIds[ids[(n + 1) % count]]);
  }
});

test('batch source identities are indexed once, then selected snapshots are captured once', () => {
  const source = createPresentation(), dest = createPresentation();
  for (const id of ['a', 'b', 'unused']) source.addSlide({ id });
  const reads = new Map(), wrappers = source.slides.map(slide => Object.create(slide, { id: { enumerable: true, get() { reads.set(slide.id, (reads.get(slide.id) ?? 0) + 1); return slide.id; } } }));
  const view = Object.create(source, { slides: { get() { return wrappers; } } });
  const result = dest.copySlidesFrom(view, ['b', 'a'], options);
  assert.deepEqual(result.slides.map(slide => slide.id), ['slide-1', 'slide-2']);
  assert.deepEqual(Object.fromEntries(reads), { a: 2, b: 2, unused: 1 });
  const duplicate = Object.create(source, { slides: { get() { return [wrappers[0], wrappers[0]]; } } });
  assert.throws(() => dest.copySlidesFrom(duplicate, ['a'], options), /Expected one source slide/);
  assert.equal(dest.slides.length, 2);
});
