import test from 'node:test';
import assert from 'node:assert/strict';
import { createSvgViewportMapping, createSvgAffineMapping } from '../dist/index.js';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveNormalizedLayout } from '@pptkit/layout';
import { getSvgSelectionFrames, hitTestSvgSelectionFrames } from '../dist/index.js';

test('viewport mapping centers the slide with letterboxing and round trips outside points', () => {
  const map = createSvgViewportMapping({ width: 400, height: 200 }, { x: 10, y: 20, width: 200, height: 200 });
  assert.equal(map.scale, .5);
  assert.deepEqual(map.contentBox, { x: 10, y: 70, width: 200, height: 100 });
  for (const point of [{ x: 0, y: 0 }, { x: 400, y: 200 }, { x: -20, y: 230 }]) assert.deepEqual(map.viewportToSlide(map.slideToViewport(point)), point);
  assert.deepEqual(map.viewportToSlide({ x: 110, y: 20 }), { x: 200, y: -100 });
  assert.ok(Object.isFrozen(map) && Object.isFrozen(map.contentBox) && Object.isFrozen(map.slideToViewport({ x: 0, y: 0 })));
});

test('viewport mapping centers horizontally and rejects singular or overflowing inputs', () => {
  const map = createSvgViewportMapping({ width: 100, height: 200 }, { x: -20, y: 10, width: 400, height: 200 });
  assert.deepEqual(map.contentBox, { x: 130, y: 10, width: 100, height: 200 });
  for (const value of [0, -1, Infinity, NaN]) assert.throws(() => createSvgViewportMapping({ width: value, height: 1 }, { x: 0, y: 0, width: 100, height: 100 }), TypeError);
  assert.throws(() => createSvgViewportMapping({ width: 1, height: 1 }, { x: Number.MAX_VALUE, y: 0, width: Number.MAX_VALUE, height: 1 }), RangeError);
  assert.throws(() => map.slideToViewport({ x: Infinity, y: 0 }), TypeError);
  const tiny = createSvgViewportMapping({ width: 1e308, height: 1e308 }, { x: 0, y: 0, width: 1, height: 1 });
  assert.throws(() => tiny.viewportToSlide({ x: 2, y: 0 }), RangeError);
});

test('viewport points compose with nested frame hit tests and mappings are detached', () => {
  const doc = createPresentation();
  doc.addSlide({ elements: [{ type: 'group', id: 'group', box: { x: 100, y: 100, width: 200, height: 100 }, coordinateSize: { width: 100, height: 100 }, transform: { rotation: 90 }, children: [
    { type: 'shape', shape: 'rect', id: 'target', box: { x: 10, y: 10, width: 20, height: 20 } },
  ] }] });
  const layout = resolveNormalizedLayout(normalizePresentation(doc)), frames = getSvgSelectionFrames(layout.slides[0]);
  const viewport = { x: 40, y: 50, width: 320, height: 300 }, size = { ...layout.size };
  const map = createSvgViewportMapping(size, viewport), corners = frames[0].corners;
  const center = { x: (corners[0].x + corners[2].x) / 2, y: (corners[0].y + corners[2].y) / 2 };
  const client = map.slideToViewport(center);
  assert.equal(hitTestSvgSelectionFrames(frames, map.viewportToSlide(client)).elementId, 'target');
  viewport.x = 900; size.width = 1;
  assert.deepEqual(map.slideToViewport(center), client);
  assert.throws(() => createSvgViewportMapping(layout.size, { x: NaN, y: 0, width: 1, height: 1 }), TypeError);
  assert.throws(() => createSvgViewportMapping(layout.size, { x: 0, y: 0, width: 0, height: 1 }), TypeError);
});

test('affine mappings handle rotation, skew, reflection and nonuniform scale', () => {
  const matrix = { a: 0, b: 2, c: -3, d: 1, e: 10, f: 20 }, map = createSvgAffineMapping(matrix);
  assert.deepEqual(map.slideToViewport({ x: 4, y: 5 }), { x: -5, y: 33 });
  const p = map.viewportToSlide({ x: -5, y: 33 });
  assert.ok(Math.abs(p.x - 4) < 1e-12 && Math.abs(p.y - 5) < 1e-12);
  matrix.e = 900; assert.equal(map.matrix.e, 10);
  assert.ok(Object.isFrozen(map) && Object.isFrozen(map.matrix) && Object.isFrozen(p));
  const reflected = createSvgAffineMapping({ a: -2, b: 0, c: .5, d: 3, e: -9, f: 1 });
  const original = { x: -20, y: 30 }, roundTrip = reflected.viewportToSlide(reflected.slideToViewport(original));
  assert.ok(Math.abs(roundTrip.x - original.x) < 1e-12 && Math.abs(roundTrip.y - original.y) < 1e-12);
});

test('affine mappings reject invalid, singular, 3D and unrepresentable transforms', () => {
  const identity = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  assert.throws(() => createSvgAffineMapping({ ...identity, is2D: false }), TypeError);
  assert.throws(() => createSvgAffineMapping({ ...identity, e: NaN }), TypeError);
  assert.throws(() => createSvgAffineMapping({ ...identity, a: 0, d: 0 }), RangeError);
  assert.throws(() => createSvgAffineMapping({ ...identity, a: 1, b: 2, c: 2, d: 4 }), RangeError);
  assert.throws(() => createSvgAffineMapping({ ...identity, a: Number.MIN_VALUE, d: Number.MIN_VALUE }), RangeError);
  const huge = createSvgAffineMapping({ ...identity, a: 1e308, d: 1e308 });
  assert.throws(() => huge.slideToViewport({ x: 2, y: 0 }), RangeError);
  assert.throws(() => huge.viewportToSlide({ x: Infinity, y: 0 }), TypeError);
});
