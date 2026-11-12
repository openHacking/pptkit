import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, fieldContextForSlide, evaluatePresentationField, PresentationFieldError } from '../dist/index.js';

test('document slide context snapshots hidden policy, copies and actual order', () => {
  const deck = createPresentation(); deck.addSlide({ id: 'a' }); deck.addSlide({ id: 'hidden', hidden: true }); deck.addSlide({ id: 'b' });
  const copy = deck.duplicateSlide('a');
  const before = fieldContextForSlide(deck, { slideId: 'b', firstPageNumber: 4, includeHidden: true });
  const visible = fieldContextForSlide(deck, { slideId: 'b', firstPageNumber: 4, includeHidden: false });
  assert.deepEqual(before.pageOrder, deck.slides.map(slide => slide.id)); assert.ok(before.pageOrder.includes(copy.id));
  assert.equal(evaluatePresentationField({ id: 'n', kind: 'slide-number' }, before).value, '7');
  assert.equal(evaluatePresentationField({ id: 'n', kind: 'slide-number' }, visible).value, '6');
  deck.moveSlide('b', 0);
  assert.equal(evaluatePresentationField({ id: 'n', kind: 'slide-number' }, before).value, '7');
  assert.equal(evaluatePresentationField({ id: 'n', kind: 'slide-number' }, fieldContextForSlide(deck, { slideId: 'b', firstPageNumber: 4, includeHidden: true })).value, '4');
  assert.ok(Object.isFrozen(before)); assert.ok(Object.isFrozen(before.pageOrder));
});

test('explicit custom order is checked against real slides and hidden exclusions', () => {
  const deck = createPresentation(); deck.addSlide({ id: 'a' }); deck.addSlide({ id: 'b' }); deck.addSlide({ id: 'hidden', hidden: true });
  const options = { slideId: 'a', firstPageNumber: 1, includeHidden: false, order: ['hidden', 'b', 'a'] };
  assert.deepEqual(fieldContextForSlide(deck, options).pageOrder, ['b', 'a']);
  const fails = input => assert.throws(() => fieldContextForSlide(deck, input), e => e instanceof PresentationFieldError && e.code === 'invalid-field-context');
  for (const input of [{ ...options, order: ['a', 'a'] }, { ...options, order: ['a', 'missing'] }, { ...options, slideId: 'hidden' }, { ...options, includeHidden: undefined }, { ...options, firstPageNumber: Number.MAX_SAFE_INTEGER }]) fails(input);
  let invoked = false; const unsafe = { ...options }; Object.defineProperty(unsafe, 'order', { enumerable: true, get() { invoked = true; return ['a']; } });
  fails(unsafe); assert.equal(invoked, false);
});
