import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, validatePresentation } from '../dist/index.js';

const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a';
const run = () => ({ text: '2', field: { kind: 'slide-number', id }, style: { bold: true } });
test('native page-number field and cache survive owned state, duplication and normalization', () => {
  const input = run(), deck = createPresentation(), slide = deck.addSlide({ id: 's', notes: [{ runs: [input] }], elements: [{ id: 't', type: 'text', content: [{ runs: [input] }], box: { x: 1, y: 2, width: 100, height: 30 } }] });
  input.field.id = 'changed'; const copied = deck.duplicateSlide(slide.id);
  assert.equal(slide.notes[0].runs[0].field.id, id); assert.equal(copied.notes[0].runs[0].field.id, id);
  const normalized = normalizePresentation(deck);
  for (const s of normalized.slides) { assert.deepEqual(s.notes[0].runs[0].field, { kind: 'slide-number', id }); assert.deepEqual(s.elements[0].content[0].runs[0].field, { kind: 'slide-number', id }); assert.equal(s.elements[0].plainText, '2'); }
  assert.deepEqual(validatePresentation(deck).filter(d => d.code === 'invalid-run-field'), []);
});

test('native field kind, UUID and extra properties fail normal validation', () => {
  for (const field of [{ id: 'bad', kind: 'slide-number' }, { id, kind: 'dynamic-date' }, { id, kind: 'slide-number', extra: 1 }, null]) {
    const deck = createPresentation(); deck.addSlide({ notes: [{ runs: [{ text: 'x', field }] }] });
    assert.ok(validatePresentation(deck).some(d => d.code === 'invalid-run-field'));
  }
});

test('native date update intent stays distinct from its supplied display cache', () => {
  const deck = createPresentation(); deck.addSlide({ notes: [{ runs: [{ text: '2025-01-01', field: { id, kind: 'dynamic-date', updateFormat: 'month-day-year' } }] }] });
  assert.deepEqual(normalizePresentation(deck).slides[0].notes[0].runs[0].field, { id, kind: 'dynamic-date', updateFormat: 'month-day-year' });
  const bad = createPresentation(); bad.addSlide({ notes: [{ runs: [{ text: 'x', field: { id, kind: 'dynamic-date', updateFormat: 'unsupported' } }] }] });
  assert.ok(validatePresentation(bad).some(d => d.code === 'invalid-run-field'));
});
