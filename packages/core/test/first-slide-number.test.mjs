import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation, PresentationValidationError } from '../dist/index.js';

test('authored first slide number is optional, validated, replaceable and detached', () => {
  const deck = createPresentation(); deck.addSlide({ id: 'a' });
  assert.equal(deck.firstSlideNumber, undefined); assert.equal('firstSlideNumber' in normalizePresentation(deck), false);
  for (const value of [0, 1, 9999, 10000, Number.MAX_SAFE_INTEGER]) {
    deck.setFirstSlideNumber(value); const normalized = normalizePresentation(deck);
    assert.equal(normalized.firstSlideNumber, value); normalized.firstSlideNumber = 4; assert.equal(deck.firstSlideNumber, value);
  }
  for (const value of [-1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, null, '1']) {
    assert.throws(() => deck.setFirstSlideNumber(value), RangeError); assert.equal(deck.firstSlideNumber, Number.MAX_SAFE_INTEGER);
  }
  deck.setFirstSlideNumber(); assert.equal('firstSlideNumber' in normalizePresentation(deck), false);
  const invalid = { ...normalizePresentation(deck), textStylePresets: {}, firstSlideNumber: -1 };
  assert.ok(validatePresentation(invalid).some(item => item.code === 'invalid-first-slide-number' && item.path === 'firstSlideNumber'));
  assert.throws(() => normalizePresentation(invalid), PresentationValidationError);
});
