import test from 'node:test';
import assert from 'node:assert/strict';
import { evaluatePresentationField as evaluate, PresentationFieldError } from '../dist/index.js';

const field = (kind, extra = {}) => ({ id: 'field-1', kind, ...extra });
const error = (fn, code) => assert.throws(fn, e => e instanceof PresentationFieldError && e.code === code);

test('slide numbers follow the explicit final occurrence order and starting number', () => {
  const context = { pageOrder: ['original', 'hidden', 'copy', 'continuation'], pageId: 'copy', firstPageNumber: 7 };
  const result = evaluate(field('slide-number'), context);
  assert.deepEqual(result, { id: 'field-1', kind: 'slide-number', value: '9' }); assert.ok(Object.isFrozen(result));
  assert.equal(evaluate(field('slide-number'), { ...context, pageOrder: ['copy', 'original'], pageId: 'original' }).value, '8');
  assert.equal(evaluate(field('slide-number'), { ...context, pageOrder: ['continuation'], pageId: 'continuation', firstPageNumber: 0 }).value, '0');
  assert.deepEqual(context.pageOrder, ['original', 'hidden', 'copy', 'continuation']);
});

test('slide numbering rejects missing, ambiguous, absent and overflowing context', () => {
  error(() => evaluate(field('slide-number')), 'missing-field-context');
  for (const context of [
    { pageOrder: ['a', 'a'], pageId: 'a', firstPageNumber: 1 },
    { pageOrder: ['a'], pageId: 'b', firstPageNumber: 1 },
    { pageOrder: ['a'], pageId: 'a', firstPageNumber: -1 },
    { pageOrder: ['a', 'b'], pageId: 'b', firstPageNumber: Number.MAX_SAFE_INTEGER },
    { pageOrder: [''], pageId: '', firstPageNumber: 1 },
  ]) error(() => evaluate(field('slide-number'), context), 'invalid-field-context');
});

test('static text preserves authored text without consulting date or locale', () => {
  for (const kind of ['static-date', 'footer']) assert.deepEqual(evaluate(field(kind, { text: '  文本\n' })), { id: 'field-1', kind, value: '  文本\n' });
  error(() => evaluate(field('footer', { text: 1 })), 'invalid-field');
  error(() => evaluate(field('unknown')), 'invalid-field');
  error(() => evaluate({ id: '', kind: 'footer', text: 'x' }), 'invalid-field');
});

test('dynamic date uses the supplied instant, locale and time zone without a clock', () => {
  const date = field('dynamic-date', { dateStyle: 'short' }), instantMs = Date.UTC(2025, 0, 1, 0, 30), context = { instantMs, locale: 'en-US', timeZone: 'UTC' };
  assert.equal(evaluate(date, context).value, '1/1/25');
  assert.equal(evaluate(date, { ...context, timeZone: 'America/Los_Angeles' }).value, '12/31/24');
  assert.deepEqual(evaluate(date, context), evaluate(date, context));
  error(() => evaluate(date, { locale: 'en-US', timeZone: 'UTC' }), 'missing-field-context');
  for (const bad of [{ ...context, instantMs: NaN }, { ...context, instantMs: 1.1 }, { ...context, locale: 'invalid_locale!' }, { ...context, timeZone: 'Invalid/Zone' }]) error(() => evaluate(date, bad), 'invalid-field-context');
  error(() => evaluate(field('dynamic-date', { dateStyle: 'other' }), context), 'invalid-field');
});

test('date evaluation does not consult the current clock or default locale and time zone', () => {
  const original = Date.now; Date.now = () => { throw new Error('unexpected clock access'); };
  try {
    assert.equal(evaluate(field('dynamic-date', { dateStyle: 'short' }), { instantMs: 0, locale: 'en-US', timeZone: 'UTC' }).value, '1/1/70');
    error(() => evaluate(field('dynamic-date', { dateStyle: 'short' }), { instantMs: 0, locale: 'en-US' }), 'missing-field-context');
    error(() => evaluate(field('dynamic-date', { dateStyle: 'short' }), { instantMs: 0, timeZone: 'UTC' }), 'missing-field-context');
  } finally { Date.now = original; }
});

test('date-time cache explicitly supports time-only and date plus time', () => {
  const instantMs = Date.UTC(2025, 0, 1, 12, 34, 56), context = { instantMs, locale: 'en-US', timeZone: 'UTC' };
  assert.equal(evaluate(field('dynamic-date', { timeStyle: 'medium', updateFormat: 'time-12-seconds' }), context).value, '12:34:56 PM');
  const combined = evaluate(field('dynamic-date', { dateStyle: 'short', timeStyle: 'short', updateFormat: 'month-day-year-time-12' }), context).value;
  assert.match(combined, /1\/1\/25/); assert.match(combined, /12:34/);
  for (const invalid of [{}, { dateStyle: 'short', timeStyle: 'bad' }, { timeStyle: 'bad' }, { timeStyle: 'short', updateFormat: 'missing' }]) error(() => evaluate(field('dynamic-date', invalid), context), 'invalid-field');
});
