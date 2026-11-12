import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '../dist/index.js';

test('managed slides change visibility explicitly, reject invalid values and preserve copies', () => {
  const deck = createPresentation(), first = deck.addSlide({ id: 'first' }), second = deck.addSlide({ id: 'second', hidden: true });
  assert.deepEqual(normalizePresentation(deck).slides.map(slide => slide.hidden), [false, true]);
  assert.ok(Object.keys(first).includes("hidden"));
  first.setHidden(true); second.setHidden(false); assert.equal({ ...first }.hidden, true); assert.deepEqual(normalizePresentation(deck).slides.map(slide => slide.hidden), [true, false]);
  for (const bad of [undefined, null, 0, 1, 'true', {}]) { assert.throws(() => first.setHidden(bad), TypeError); assert.equal(first.hidden, true); }
  const copy = deck.duplicateSlide('first'); assert.equal(copy.hidden, true); first.setHidden(false); assert.equal(copy.hidden, true);
});
