import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, fieldContextForOccurrences, evaluatePresentationField, PresentationFieldError } from '../dist/index.js';

test('custom shows can number repeated visits to the same slide independently', () => {
  const deck = createPresentation(); deck.addSlide({ id: 'a' }); deck.addSlide({ id: 'b' }); deck.addSlide({ id: 'h', hidden: true });
  const occurrences = [{ id: 'a-first', slideId: 'a' }, { id: 'hidden', slideId: 'h' }, { id: 'b', slideId: 'b' }, { id: 'a-again', slideId: 'a' }];
  const options = { occurrences, occurrenceId: 'a-again', firstPageNumber: 10, includeHidden: false };
  const result = fieldContextForOccurrences(deck, options);
  assert.deepEqual(result.pageOrder, ['a-first', 'b', 'a-again']);
  assert.equal(evaluatePresentationField({ id: 'page', kind: 'slide-number' }, result).value, '12');
  assert.equal(evaluatePresentationField({ id: 'page', kind: 'slide-number' }, fieldContextForOccurrences(deck, { ...options, includeHidden: true })).value, '13');
  occurrences[0].id = 'changed'; assert.deepEqual(result.pageOrder, ['a-first', 'b', 'a-again']);
});

test('occurrence identities, slide links and excluded targets are checked atomically', () => {
  const deck = createPresentation(); deck.addSlide({ id: 'a' }); deck.addSlide({ id: 'h', hidden: true });
  const base = { occurrences: [{ id: 'first', slideId: 'a' }], occurrenceId: 'first', firstPageNumber: 1, includeHidden: true };
  for (const input of [{ ...base, occurrences: [{ id: 'first', slideId: 'missing' }] }, { ...base, occurrences: [...base.occurrences, ...base.occurrences] }, { ...base, occurrences: [{ id: 'first', slideId: 'a', extra: 1 }] }, { ...base, occurrences: [{ id: 'first', slideId: 'h' }], includeHidden: false }, { ...base, occurrenceId: 'absent' }]) assert.throws(() => fieldContextForOccurrences(deck, input), e => e instanceof PresentationFieldError && e.code === 'invalid-field-context');
});
