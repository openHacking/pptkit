import test from 'node:test';
import assert from 'node:assert/strict';
import { materializeFieldText, PresentationFieldError } from '../dist/index.js';

test('field templates materialize detached rich text with field locations', () => {
  const template = [{ style: { align: 'center' }, runs: [{ text: 'Page ' }, { field: { id: 'number', kind: 'slide-number' }, style: { bold: true } }, { text: ' / ' }, { field: { id: 'footer', kind: 'footer', text: 'Final' }, action: { type: 'url', url: 'https://example.com' } }] }];
  const result = materializeFieldText(template, { pageOrder: ['a', 'b'], pageId: 'b', firstPageNumber: 3 });
  assert.deepEqual(result.content[0].runs.map(run => run.text), ['Page ', '4', ' / ', 'Final']);
  assert.deepEqual(result.fields.map(f => [f.id, f.value, f.paragraphIndex, f.runIndex]), [['number', '4', 0, 1], ['footer', 'Final', 0, 3]]);
  assert.equal(result.content[0].runs[1].style.bold, true); assert.equal(result.content[0].runs[3].action.url, 'https://example.com');
  template[0].runs[1].style.bold = false; assert.equal(result.content[0].runs[1].style.bold, true);
  assert.ok(Object.isFrozen(result.fields)); assert.ok(result.fields.every(Object.isFrozen));
});

test('bad or missing field context fails the whole operation without input mutation', () => {
  const template = [{ runs: [{ text: 'before' }, { field: { id: 'page', kind: 'slide-number' } }] }], original = structuredClone(template);
  assert.throws(() => materializeFieldText(template), e => e instanceof PresentationFieldError && e.code === 'missing-field-context');
  assert.deepEqual(template, original);
  for (const input of [[{ runs: [{ text: 'x', field: { id: 'x', kind: 'footer', text: 'y' } }] }], [{ runs: [{}] }], [{ runs: [{ text: 'x', extra: true }] }]]) assert.throws(() => materializeFieldText(input), e => e instanceof PresentationFieldError && e.code === 'invalid-field');
  let invoked = false; const run = { text: 'x' }; Object.defineProperty(run, 'style', { enumerable: true, get() { invoked = true; return {}; } });
  assert.throws(() => materializeFieldText([{ runs: [run] }])); assert.equal(invoked, false);
});

test('repeated page fields retain independent identities and all are validated', () => {
  const page = id => ({ field: { id, kind: 'slide-number' } }), context = { pageOrder: ['a', 'b'], pageId: 'b', firstPageNumber: 1 };
  const result = materializeFieldText([{ runs: [page('one'), page('two')] }], context);
  assert.deepEqual(result.fields.map(item => [item.id, item.value]), [['one', '2'], ['two', '2']]);
  assert.throws(() => materializeFieldText([{ runs: [page('valid'), page('')] }], context), e => e instanceof PresentationFieldError && e.code === 'invalid-field');
});

test('native output is explicit, retains slide-number identity and rejects other kinds', () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', context = { pageOrder: ['a'], pageId: 'a', firstPageNumber: 3 };
  const result = materializeFieldText([{ runs: [{ field: { id, kind: 'slide-number' }, output: 'native' }, { field: { id: 'plain', kind: 'slide-number' } }] }], context);
  assert.deepEqual(result.content[0].runs[0], { text: '3', field: { id, kind: 'slide-number' } }); assert.deepEqual(result.content[0].runs[1], { text: '3' });
  for (const run of [{ field: { id: 'not-uuid', kind: 'slide-number' }, output: 'native' }, { field: { id, kind: 'footer', text: 'x' }, output: 'native' }, { text: 'x', output: 'native' }, { field: { id, kind: 'slide-number' }, output: 'other' }]) assert.throws(() => materializeFieldText([{ runs: [run] }], context), e => e instanceof PresentationFieldError && e.code === 'invalid-field');
});

test('native dynamic dates require an explicit update format separate from the cache locale', () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', context = { instantMs: 0, locale: 'en-US', timeZone: 'UTC' };
  const date = { id, kind: 'dynamic-date', dateStyle: 'short', updateFormat: 'day-month-year' };
  const result = materializeFieldText([{ runs: [{ field: date, output: 'native' }] }], context);
  assert.deepEqual(result.content[0].runs[0], { text: '1/1/70', field: { id, kind: 'dynamic-date', updateFormat: 'day-month-year' } });
  assert.throws(() => materializeFieldText([{ runs: [{ field: { id, kind: 'dynamic-date', dateStyle: 'short' }, output: 'native' }] }], context), e => e instanceof PresentationFieldError && e.code === 'invalid-field');
});

test('time-only cache and native update type remain independently authored', () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', field = { id, kind: 'dynamic-date', timeStyle: 'medium', updateFormat: 'time-24-seconds' };
  const result = materializeFieldText([{ runs: [{ field, output: 'native' }] }], { instantMs: Date.UTC(2025, 0, 1, 12, 34, 56), locale: 'en-US', timeZone: 'UTC' });
  assert.equal(result.content[0].runs[0].text, '12:34:56 PM'); assert.equal(result.content[0].runs[0].field.updateFormat, 'time-24-seconds');
});
