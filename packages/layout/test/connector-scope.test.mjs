import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, PresentationValidationError } from '@pptkit/core';
import { resolveLayout, resolveNormalizedLayout } from '../dist/index.js';

test('both authored and direct normalized paths reject connector references outside their coordinate scope', () => {
  const doc = createPresentation(), box = { x: 10, y: 20, width: 100, height: 100 };
  const slide = doc.addSlide({ elements: [
    { id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 100 }, children: [{ id: 'nested', type: 'shape', shape: 'rect', box: { x: 5, y: 10, width: 20, height: 20 } }] },
    { id: 'edge', type: 'connector', start: { elementId: 'group', anchor: 'right' }, end: { x: 200, y: 30 } },
  ] });
  assert.deepEqual(resolveLayout(doc).slides[0].elements[1].start, { x: 110, y: 70 });
  const normalized = normalizePresentation(doc);
  normalized.slides[0].elements[1].start = { elementId: 'nested', anchor: 'right' };
  assert.throws(() => resolveNormalizedLayout(normalized), /outside the current coordinate scope/);
  slide.replaceElement('edge', { id: 'edge', type: 'connector', start: { elementId: 'nested' }, end: { x: 200, y: 30 } });
  assert.throws(() => resolveLayout(doc), PresentationValidationError);
});

test('missing direct normalized references fail rather than becoming an origin point', () => {
  const doc = createPresentation(), box = { x: 0, y: 0, width: 20, height: 20 };
  doc.addSlide({ elements: [{ id: 'shape', type: 'shape', shape: 'rect', box }, { id: 'edge', type: 'connector', start: { elementId: 'shape' }, end: { x: 10, y: 10 } }] });
  const normalized = normalizePresentation(doc);
  normalized.slides[0].elements[1].start = { elementId: 'deleted' };
  assert.throws(() => resolveNormalizedLayout(normalized), /Connector target "deleted"/);
});
