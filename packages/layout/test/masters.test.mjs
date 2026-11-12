import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveLayout, resolveNormalizedLayout } from '../dist/index.js';
import { multiMasterDocument } from '../../../tests/fixtures/multi-master.mjs';

test('intrinsic and table measurements use the effective owner theme for slides and reusable layout objects', () => {
  const document = multiMasterDocument(createPresentation), families = [];
  document.defineSlideLayout({ id: 'table-layout', name: 'Table', masterId: 'blue', elements: [{ id: 'layout-table', type: 'table', box: { x: 0, y: 0, width: 200, height: 20 }, columns: [200], rows: [{ cells: [{ content: [{ runs: [{ text: 'Table', style: { fontFamily: { theme: 'body' } } }] }] }] }] }] });
  const layout = resolveLayout(document, { intrinsicSizing: document.slides.map(s => ({ elementId: s.elements[0].id, height: { type: 'content' } })), tableSizing: [{ elementId: 'layout-table', autoRows: [{ row: 0 }] }], textMeasurement: { measure(request) { families.push(request.paragraphs[0].runs[0].style.fontFamily); return { bounds: { x: 0, y: 0, width: 50, height: 20 } }; } } });
  assert.deepEqual(families, ['Times New Roman', 'Georgia', 'Times New Roman', 'Georgia', 'Arial', 'Times New Roman', 'Georgia', 'Times New Roman', 'Georgia', 'Arial']);
  assert.equal(layout.layouts.find(l => l.id === 'layout-0').masterId, 'red');
  assert.equal(layout.layouts.find(l => l.id === 'layout-1').theme.colors.accent1, '0044CC');
  layout.masters[0].theme.colors.accent1 = '999999'; assert.equal(document.masters[0].theme.colors.accent1, 'CC2200');
});

test('incomplete transported master/theme domains reject instead of borrowing the global theme', () => {
  for (const change of [n => { delete n.layouts.find(l => l.masterId).theme; }, n => { n.masters = []; }]) {
    const n = normalizePresentation(multiMasterDocument(createPresentation)); change(n);
    assert.throws(() => resolveNormalizedLayout(n), TypeError);
  }
});
