import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation, PresentationFieldError } from '../../core/dist/index.js';
import { resolveLayout, resolveNormalizedLayout } from '../dist/index.js';
import { captureSizingFields, assertSizingFieldsStable } from '../dist/fields/sizing-consistency.js';
import { refreshLayoutFields } from '../dist/fields/refresh.js';

const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a';
const number = () => ({ text: 'stale', field: { id, kind: 'slide-number' } });
const date = () => ({ text: 'old-date', field: { id, kind: 'dynamic-date', updateFormat: 'day-month-year' } });
const refresh = { firstPageNumber: 5, includeHidden: false, instantMs: Date.UTC(2025, 0, 1, 0, 30), locale: 'en-US', timeZone: 'America/Los_Angeles', dateStyle: 'short' };

test('field refresh traverses a wide group without an argument-count failure', () => {
  const field = { type: 'text', hidden: false, content: [{ runs: [number()] }], plainText: 'stale' };
  const children = Array(150000).fill({ type: 'shape', hidden: false });
  children[0] = field;
  const input = { layouts: [], masters: [], slides: [{ hidden: false, elements: [{ type: 'group', hidden: false, children }], notes: [] }] };
  refreshLayoutFields(input, { firstPageNumber: 17, includeHidden: true });
  assert.equal(field.plainText, '17');
  assert.equal(field.content[0].runs[0].text, '17');
});

test('field refresh uses final pagination order, preserves source and refreshes repeated headers', () => {
  const deck = createPresentation({ size: { width: 300, height: 200 } }), slide = deck.addSlide({ id: 'source', notes: [{ runs: [number(), date()] }] });
  slide.addElement({ type: 'text', id: 'text', box: { x: 10, y: 5, width: 100, height: 20 }, content: [{ runs: [{ text: 'Page ' }, number(), date()] }] });
  slide.addElement({ type: 'table', id: 'table', box: { x: 20, y: 30, width: 100, height: 90 }, columns: [100], rows: [{ height: 10, cells: [{ content: [{ runs: [number()] }] }] }, ...Array.from({ length: 4 }, (_, i) => ({ height: 20, cells: [{ content: `row ${i}` }] }))] });
  deck.addSlide({ id: 'hidden', hidden: true, elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [number()] }] }] });
  deck.addSlide({ id: 'after', elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [number()] }] }] });
  const before = normalizePresentation(deck), layout = resolveLayout(deck, { fieldRefresh: refresh, tablePagination: [{ elementId: 'table', area: { x: 20, y: 30, width: 100, height: 55 }, headerRows: 1 }] });
  assert.equal(layout.slides.length, 4);
  const first = layout.slides[0], next = layout.slides[1], hidden = layout.slides[2], after = layout.slides[3];
  assert.equal(first.elements.find(e => e.type === 'text').plainText, 'Page 512/31/24');
  assert.deepEqual(first.notes[0].runs.map(run => run.text), ['5', '12/31/24']);
  assert.equal(first.elements.find(e => e.type === 'table').rows[0].cells[0].content[0].runs[0].text, '5');
  assert.equal(next.elements.find(e => e.type === 'table').rows[0].cells[0].content[0].runs[0].text, '6');
  assert.equal(hidden.elements[0].plainText, 'stale'); assert.equal(after.elements[0].plainText, '7');
  assert.deepEqual(normalizePresentation(deck), before);
  const counted = resolveLayout(deck, { fieldRefresh: { ...refresh, includeHidden: true }, tablePagination: [{ elementId: 'table', area: { x: 20, y: 30, width: 100, height: 55 }, headerRows: 1 }] });
  assert.equal(counted.slides[2].elements[0].plainText, '7'); assert.equal(counted.slides[3].elements[0].plainText, '8');
});

test('missing date context, shared fields, overflow and sizing conflicts fail explicitly', () => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [date()] }] }] });
  assert.throws(() => resolveLayout(deck, { fieldRefresh: { firstPageNumber: 1, includeHidden: true } }), e => e instanceof PresentationFieldError && e.code === 'missing-field-context');
  assert.throws(() => resolveLayout(deck, { fieldRefresh: { ...refresh, dateStyle: 'invalid' } }), e => e instanceof PresentationFieldError && e.code === 'invalid-field-context');
  assert.doesNotThrow(() => resolveLayout(deck, { fieldRefresh: refresh, intrinsicSizing: [], tablePagination: [{ elementId: 'missing', area: { x: 0, y: 0, width: 100, height: 100 } }] }));
  const two = createPresentation(); two.addSlide(); two.addSlide();
  assert.throws(() => resolveLayout(two, { fieldRefresh: { firstPageNumber: Number.MAX_SAFE_INTEGER, includeHidden: true } }), e => e instanceof PresentationFieldError && e.code === 'invalid-field-context');
  const shared = createPresentation(); shared.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [number()] }] }] }); shared.addSlide();
  assert.doesNotThrow(() => resolveLayout(shared, { fieldRefresh: refresh }));
  shared.defineSlideLayout({ id: 'with-master', name: 'With master', masterId: 'm' }); shared.addSlide({ layoutId: 'with-master' });
  assert.throws(() => resolveLayout(shared, { fieldRefresh: refresh }), /shared master or layout artwork/);
});

test('pagination accepts sizing when measured field text remains stable and rejects changed continuation or later page values', () => {
  const makeDeck = (header, later = false) => {
    const deck = createPresentation({ size: { width: 300, height: 200 } });
    deck.addSlide({ id: 'first', elements: [{ type: 'table', id: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [
      { height: 10, cells: [{ content: [{ runs: [header()] }] }] },
      ...Array.from({ length: 3 }, (_, i) => ({ height: 20, cells: [{ content: `row ${i}` }] })),
    ] }] });
    deck.addSlide({ id: 'later', elements: [{ id: 'later-text', type: 'text', box: { x: 0, y: 0, width: 100, height: 20 }, content: [{ runs: [number()] }] }] });
    const sizing = [{ elementId: 'table', autoRows: [{ row: 0 }] }];
    const options = { fieldRefresh: refresh, tableSizing: sizing, tablePagination: [{ elementId: 'table', area: { x: 0, y: 0, width: 100, height: 45 }, headerRows: 1 }], textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 10 } }) }, ...(later ? { intrinsicSizing: [{ elementId: 'later-text', width: { type: 'content' } }] } : {}) };
    return { deck, options };
  };
  const stable = makeDeck(date), snapshot = normalizePresentation(stable.deck), result = resolveLayout(stable.deck, stable.options);
  assert.ok(result.slides.length > 2);
  for (const slide of result.slides.slice(0, -1)) assert.equal(slide.elements.find(e => e.type === 'table').rows[0].cells[0].content[0].runs[0].text, '12/31/24');
  assert.equal(result.slides.at(-1).elements[0].plainText, String(5 + result.slides.length - 1));
  assert.deepEqual(normalizePresentation(stable.deck), snapshot);
  const numberedHeader = makeDeck(number), headerSnapshot = normalizePresentation(numberedHeader.deck);
  assert.throws(() => resolveLayout(numberedHeader.deck, numberedHeader.options), /Field text in sizing target "table" changed after pagination/);
  assert.deepEqual(normalizePresentation(numberedHeader.deck), headerSnapshot);
  const later = makeDeck(date, true);
  assert.throws(() => resolveLayout(later.deck, later.options), /Field text in sizing target "later-text" changed after pagination/);
  const continuation = makeDeck(date);
  continuation.deck.defineSlideLayout({ id: 'future', name: 'Future', elements: [{ id: 'future-date', type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [date()] }] }] });
  continuation.options.intrinsicSizing = [{ elementId: 'future-date', width: { type: 'content' } }];
  continuation.options.tablePagination[0].continuation = { layoutId: 'future' };
  assert.throws(() => resolveLayout(continuation.deck, continuation.options), /Field text in sizing target "future-date" changed after pagination/);
});

test('pagination tracks merged source rows and repeated field IDs by location', () => {
  const fixture = trailing => {
    const deck = createPresentation();
    deck.addSlide({ elements: [{ id: 'merged', type: 'table', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [100], rows: [
      { height: 10, cells: [{ content: 'header' }] },
      { height: 10, cells: [{ rowSpan: 2, content: [{ runs: [date(), number()] }] }] },
      { height: 10, cells: [] },
      { height: 10, cells: [{ content: [{ runs: [date(), trailing()] }] }] },
    ] }] });
    return deck;
  };
  const options = { fieldRefresh: refresh, tableSizing: [{ elementId: 'merged', autoRows: [{ row: 0 }] }], tablePagination: [{ elementId: 'merged', area: { x: 0, y: 0, width: 100, height: 35 }, headerRows: 1 }], textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 5 } }) } };
  const changed = fixture(number), before = normalizePresentation(changed);
  assert.throws(() => resolveLayout(changed, options), /Field text in sizing target "merged" changed after pagination/);
  assert.deepEqual(normalizePresentation(changed), before);
  const stable = fixture(date), result = resolveLayout(stable, options);
  assert.equal(result.slides.length, 2);
  assert.deepEqual(result.pagination.fragments.map(fragment => fragment.sourceRowIndices), [[0, 1, 2], [0, 3]]);
  assert.equal(result.slides[0].elements[0].rows[1].cells[0].rowSpan, 2);
  assert.equal(result.slides[1].elements[0].rows[1].cells[0].content[0].runs[1].text, '12/31/24');
});

test('sizing guard rejects missing fields or missing fragments, including repeated identifiers', () => {
  const deck = createPresentation();
  deck.addSlide({ elements: [{ id: 'text', type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [number(), number()] }] }] });
  const options = { intrinsicSizing: [{ elementId: 'text', width: { type: 'content' } }] };
  const normalized = normalizePresentation(deck), snapshot = captureSizingFields(normalized, options), layout = resolveNormalizedLayout(normalized);
  assert.doesNotThrow(() => assertSizingFieldsStable(layout, snapshot));
  const lost = structuredClone(layout); delete lost.slides[0].elements[0].content[0].runs[1].field;
  assert.throws(() => assertSizingFieldsStable(lost, snapshot), /changed after pagination/);
  const missing = structuredClone(layout); missing.slides[0].elements = [];
  assert.throws(() => assertSizingFieldsStable(missing, snapshot), /changed after pagination/);
  const tableDeck = createPresentation();
  tableDeck.addSlide({ elements: [{ id: 'table', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ height: 10, cells: [{ content: [{ runs: [date()] }] }] }, ...Array.from({ length: 2 }, () => ({ height: 20, cells: [{ content: 'row' }] }))] }] });
  const input = normalizePresentation(tableDeck), captured = captureSizingFields(input, { tableSizing: [{ elementId: 'table', autoRows: [{ row: 0 }] }] });
  const pages = resolveNormalizedLayout(input, { tablePagination: [{ elementId: 'table', area: { x: 0, y: 0, width: 100, height: 35 }, headerRows: 1 }] });
  assert.equal(pages.slides.length, 2); assert.doesNotThrow(() => assertSizingFieldsStable(pages, captured));
  delete pages.slides[1].elements[0].rows[0].cells[0].content[0].runs[0].field;
  assert.throws(() => assertSizingFieldsStable(pages, captured), /changed after pagination/);
});

test('sizing without pagination measures refreshed fields, including shared and hidden slide policy', () => {
  const deck = createPresentation();
  deck.defineSlideLayout({ id: 'l', name: 'Layout', elements: [{ id: 'shared-date', type: 'text', box: { x: 0, y: 0, width: 10, height: 20 }, frame: { margin: 0 }, content: [{ runs: [date()] }] }] });
  deck.addSlide({ id: 'one', layoutId: 'l', elements: [
    { id: 'number', type: 'text', box: { x: 0, y: 0, width: 10, height: 20 }, frame: { margin: 0 }, content: [{ runs: [number()] }] },
    { id: 'table', type: 'table', box: { x: 0, y: 30, width: 100, height: 10 }, columns: [100], rows: [{ height: 10, cells: [{ content: [{ runs: [date()] }] }] }] },
  ] });
  deck.addSlide({ id: 'hidden', hidden: true, elements: [{ id: 'hidden-number', type: 'text', box: { x: 0, y: 0, width: 10, height: 20 }, frame: { margin: 0 }, content: [{ runs: [number()] }] }] });
  const before = normalizePresentation(deck), seen = [];
  const result = resolveLayout(deck, {
    fieldRefresh: refresh, tablePagination: [],
    intrinsicSizing: ['number', 'shared-date', 'hidden-number'].map(elementId => ({ elementId, width: { type: 'content' }, includeHidden: true })),
    tableSizing: [{ elementId: 'table', autoRows: [{ row: 0 }] }],
    textMeasurement: { measure(request) { const text = request.paragraphs.flatMap(p => p.runs.map(r => r.text)).join(''); seen.push(text); return { bounds: { x: 0, y: 0, width: text.length * 2, height: text.length * 3 } }; } },
  });
  assert.ok(seen.includes('5')); assert.ok(seen.includes('12/31/24')); assert.ok(seen.includes('stale'));
  assert.equal(result.slides[0].elements.find(e => e.id === 'number').box.width, 2);
  assert.equal(result.layouts.find(l => l.id === 'l').elements[0].box.width, 16);
  assert.equal(result.slides[1].elements[0].box.width, 10);
  assert.ok(result.slides[0].elements.find(e => e.id === 'table').rows[0].height >= 24);
  assert.deepEqual(normalizePresentation(deck), before);
  const normalized = normalizePresentation(deck), snapshot = structuredClone(normalized);
  let calls = 0;
  assert.throws(() => resolveNormalizedLayout(normalized, { fieldRefresh: { firstPageNumber: 1, includeHidden: true }, intrinsicSizing: [{ elementId: 'number', width: { type: 'content' } }], textMeasurement: { measure() { calls++; return { bounds: { x: 0, y: 0, width: 1, height: 1 } }; } } }), e => e instanceof PresentationFieldError && e.code === 'missing-field-context');
  assert.equal(calls, 0); assert.deepEqual(normalized, snapshot);
});

test('shared date caches refresh only for visible effective owners; shared page numbers still reject', () => {
  const deck = createPresentation();
  deck.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ id: 'master-date', type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [date()] }] }] });
  deck.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [{ id: 'layout-date', type: 'text', box: { x: 0, y: 40, width: 100, height: 30 }, content: [{ runs: [date()] }] }] });
  deck.addSlide({ id: 'one', layoutId: 'l' }); deck.addSlide({ id: 'two', layoutId: 'l' });
  const before = normalizePresentation(deck), layout = resolveLayout(deck, { fieldRefresh: refresh });
  assert.equal(layout.masters.find(master => master.id === 'm').elements[0].plainText, '12/31/24');
  assert.equal(layout.layouts.find(item => item.id === 'l').elements[0].plainText, '12/31/24'); assert.deepEqual(normalizePresentation(deck), before);
  assert.throws(() => resolveLayout(deck, { fieldRefresh: { firstPageNumber: 1, includeHidden: true } }), e => e instanceof PresentationFieldError && e.code === 'missing-field-context');
  const hidden = createPresentation(); hidden.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [number()] }] }] });
  hidden.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', showMasterShapes: false }); hidden.addSlide({ layoutId: 'l' });
  assert.doesNotThrow(() => resolveLayout(hidden, { fieldRefresh: refresh }));
  const suppressed = createPresentation(); suppressed.defineSlideLayout({ id: 'l', name: 'Layout', elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [number()] }] }] }); suppressed.addSlide({ layoutId: 'l', showInheritedShapes: false });
  assert.doesNotThrow(() => resolveLayout(suppressed, { fieldRefresh: refresh }));
  const hiddenArt = createPresentation(); hiddenArt.defineSlideLayout({ id: 'l', name: 'Layout', elements: [{ type: 'group', hidden: true, box: { x: 0, y: 0, width: 100, height: 30 }, coordinateSize: { width: 100, height: 30 }, children: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [number()] }] }] }] }); hiddenArt.addSlide({ layoutId: 'l' });
  assert.doesNotThrow(() => resolveLayout(hiddenArt, { fieldRefresh: refresh }));
  const mixed = createPresentation();
  mixed.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [date()] }] }] });
  mixed.defineSlideLayout({ id: 'off', name: 'Off', masterId: 'm', showMasterShapes: false });
  mixed.defineSlideLayout({ id: 'on', name: 'On', masterId: 'm' });
  mixed.addSlide({ layoutId: 'off' }); mixed.addSlide({ layoutId: 'on', hidden: true });
  assert.equal(resolveLayout(mixed, { fieldRefresh: refresh }).masters.find(master => master.id === 'm').elements[0].plainText, 'old-date');
  assert.equal(resolveLayout(mixed, { fieldRefresh: { ...refresh, includeHidden: true } }).masters.find(master => master.id === 'm').elements[0].plainText, '12/31/24');
});

test('nested group and shape fields refresh while the default remains the authored cache', () => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'group', box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 30 }, text: { content: [{ runs: [{ text: 'N=' }, number()] }] } }] }] });
  const original = resolveLayout(deck), updated = resolveLayout(deck, { fieldRefresh: { firstPageNumber: 9, includeHidden: true } });
  assert.equal(original.slides[0].elements[0].children[0].text.plainText, 'N=stale');
  assert.equal(updated.slides[0].elements[0].children[0].text.plainText, 'N=9');
});
