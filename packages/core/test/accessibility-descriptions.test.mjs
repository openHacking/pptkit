import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, inspectMissingDescriptions, inspectEffectiveMissingDescriptions, EffectiveDescriptionLimitError, normalizePresentation, PresentationValidationError, validatePresentation } from '../dist/index.js';

const box = { x: 0, y: 0, width: 100, height: 50 };
const image = (id, extra = {}) => ({ id, type: 'image', assetId: 'asset', box, ...extra });
const group = (id, children, extra = {}) => ({ id, type: 'group', box, coordinateSize: { width: 100, height: 50 }, children, ...extra });
const table = (id, extra = {}) => ({ id, type: 'table', box, columns: [100], rows: [{ cells: [{ content: 'A' }] }], ...extra });
const chart = (id, extra = {}) => ({ id, type: 'chart', box, chartType: 'bar', categories: ['A'], series: [{ name: 'Series', values: [1] }], ...extra });

test('description inspection reports direct and nested objects by owner without conflating names or asset metadata', () => {
  const deck = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [image('master-image')] }] });
  deck.registerAsset({ id: 'asset', kind: 'image', source: { type: 'path', value: 'image.png' }, accessibility: { description: 'Asset description' } });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [table('layout-table')] });
  deck.addSlide({ id: 'slide', layoutId: 'layout', elements: [
    image('named', { name: 'File name' }),
    group('outer', [chart('chart'), group('inner', [table('nested')])]),
    image('described', { accessibility: { description: 'A useful description' } }),
    image('blank', { accessibility: { description: '  ' } }),
    image('decorative', { accessibility: { decorative: true } }),
    image('hidden', { hidden: true }),
    group('decorative-group', [chart('suppressed-chart')], { accessibility: { decorative: true } }),
    group('hidden-group', [table('suppressed-table')], { hidden: true }),
  ] });
  deck.addSlide({ id: 'hidden-slide', hidden: true, elements: [image('suppressed-slide')] });
  const before = JSON.stringify(normalizePresentation(deck)), result = inspectMissingDescriptions(deck);
  assert.equal(result.scope, 'authored-definitions');
  assert.deepEqual(result.findings.map(({ elementId, path, slideId, layoutId, masterId }) => Object.fromEntries(Object.entries({ elementId, path, slideId, layoutId, masterId }).filter(([, value]) => value !== undefined))), [
    { elementId: 'named', path: 'slides.0.elements.0.accessibility.description', slideId: 'slide' },
    { elementId: 'chart', path: 'slides.0.elements.1.children.0.accessibility.description', slideId: 'slide' },
    { elementId: 'nested', path: 'slides.0.elements.1.children.1.children.0.accessibility.description', slideId: 'slide' },
    { elementId: 'blank', path: 'slides.0.elements.3.accessibility.description', slideId: 'slide' },
    { elementId: 'layout-table', path: 'layouts.0.elements.0.accessibility.description', layoutId: 'layout' },
    { elementId: 'master-image', path: 'masters.0.elements.0.accessibility.description', masterId: 'master' },
  ]);
  assert.equal(JSON.stringify(normalizePresentation(deck)), before);
  assert.ok(Object.isFrozen(result) && Object.isFrozen(result.findings) && Object.isFrozen(result.findings[0]));
  assert.throws(() => result.findings.push({}), TypeError);
});

test('description inspection rejects invalid documents instead of returning incomplete results', () => {
  const deck = createPresentation();
  deck.addSlide({ elements: [image('missing')] });
  assert.throws(() => inspectMissingDescriptions(deck), error => error instanceof PresentationValidationError && error.diagnostics.some(item => item.code === 'missing-image-asset'));
});

test('inspection follows copied identities and removes findings after deletion', () => {
  const deck = createPresentation();
  const original = deck.addSlide({ id: 'original', elements: [group('group', [table('table')])] });
  const copy = deck.duplicateSlide(original.id);
  const originalFinding = inspectMissingDescriptions(deck).findings.find(item => item.slideId === original.id);
  const copyFinding = inspectMissingDescriptions(deck).findings.find(item => item.slideId === copy.id);
  assert.equal(originalFinding.elementId, 'table');
  assert.notEqual(copyFinding.elementId, originalFinding.elementId);
  deck.removeSlide(original.id);
  assert.deepEqual(inspectMissingDescriptions(deck).findings.map(item => [item.slideId, item.elementId]), [[copy.id, copyFinding.elementId]]);
  copy.removeElement(copy.elements[0].id);
  assert.deepEqual(inspectMissingDescriptions(deck).findings, []);
});

test('effective description findings follow shared owners, visibility and live slide identity', () => {
  const deck = createPresentation({ masters: [{ id: 'master', name: 'Master', elements: [group('mg', [table('mt')]), table('hidden-master', { hidden: true })] }] });
  deck.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master', elements: [chart('lc'), group('hidden-group', [table('suppressed')], { accessibility: { decorative: true } })] });
  deck.defineSlideLayout({ id: 'no-master', name: 'No master', masterId: 'master', showMasterShapes: false, elements: [table('lt')] });
  deck.defineSlideLayout({ id: 'unused', name: 'Unused', elements: [table('unused-table')] });
  const slide = deck.addSlide({ id: 'one', layoutId: 'layout', elements: [table('st')] });
  deck.addSlide({ id: 'two', layoutId: 'layout', showInheritedShapes: false, elements: [chart('sc')] });
  deck.addSlide({ id: 'three', layoutId: 'no-master', showInheritedShapes: true });
  deck.addSlide({ id: 'hidden', layoutId: 'layout', hidden: true, elements: [table('hidden-slide')] });
  deck.addSlide({ id: 'four', layoutId: 'layout' });
  const before = JSON.stringify(normalizePresentation(deck));
  const report = inspectEffectiveMissingDescriptions(deck);
  assert.equal(report.scope, 'modeled-slide-occurrences');
  assert.deepEqual(report.findings.map(f => [f.slideId, f.source.kind, f.source.id, f.elementId, f.path]), [
    ['one', 'master', 'master', 'mt', 'masters.0.elements.0.children.0.accessibility.description'],
    ['one', 'layout', 'layout', 'lc', 'layouts.0.elements.0.accessibility.description'],
    ['one', 'slide', 'one', 'st', 'slides.0.elements.0.accessibility.description'],
    ['two', 'slide', 'two', 'sc', 'slides.1.elements.0.accessibility.description'],
    ['three', 'layout', 'no-master', 'lt', 'layouts.1.elements.0.accessibility.description'],
    ['four', 'master', 'master', 'mt', 'masters.0.elements.0.children.0.accessibility.description'],
    ['four', 'layout', 'layout', 'lc', 'layouts.0.elements.0.accessibility.description'],
  ]);
  assert.ok(Object.isFrozen(report) && Object.isFrozen(report.findings) && Object.isFrozen(report.findings[0]) && Object.isFrozen(report.findings[0].source));
  assert.equal(JSON.stringify(normalizePresentation(deck)), before);
  assert.equal(inspectEffectiveMissingDescriptions(deck, { maxFindings: 7 }).findings.length, 7);
  assert.throws(() => inspectEffectiveMissingDescriptions(deck, { maxFindings: 6 }), error => error instanceof EffectiveDescriptionLimitError && error.maxFindings === 6);
  for (const maxFindings of [-1, 0.1, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) assert.throws(() => inspectEffectiveMissingDescriptions(deck, { maxFindings }), TypeError);
  const copied = deck.duplicateSlide(slide.id);
  const copy = inspectEffectiveMissingDescriptions(deck).findings.filter(f => f.slideId === copied.id);
  assert.deepEqual(copy.map(f => f.source.kind), ['master', 'layout', 'slide']);
  assert.notEqual(copy[2].elementId, 'st');
  deck.removeSlide('one');
  assert.ok(inspectEffectiveMissingDescriptions(deck).findings.every(f => f.slideId !== 'one'));
  copied.setShowInheritedShapes(false);
  assert.deepEqual(inspectEffectiveMissingDescriptions(deck).findings.filter(f => f.slideId === copied.id).map(f => f.source.kind), ['slide']);
  const invalid = createPresentation(); invalid.addSlide({ elements: [image('missing')] });
  assert.throws(() => inspectEffectiveMissingDescriptions(invalid), PresentationValidationError);
});

test('decorative actions produce located warnings for objects and nested text runs', () => {
  const deck = createPresentation();
  const action = { type: 'url', url: 'https://example.com' };
  deck.addSlide({ id: 'page', elements: [
    group('g', [
      { id: 'child', type: 'shape', shape: 'rect', box, action },
      { id: 'text', type: 'text', box, content: [{ runs: [{ text: 'Action', action }] }] },
      table('table', { rows: [{ cells: [{ content: [{ runs: [{ text: 'Cell', action }] }] }] }] }),
    ], { accessibility: { decorative: true }, action }),
    { id: 'ordinary', type: 'shape', shape: 'rect', box, action },
  ] });
  const warnings = validatePresentation(deck).filter(item => item.code === 'decorative-action');
  assert.deepEqual(warnings.map(item => [item.severity, item.elementId, item.path]), [
    ['warning', 'g', 'slides.0.elements.0.action'],
    ['warning', 'child', 'slides.0.elements.0.children.0.action'],
    ['warning', 'text', 'slides.0.elements.0.children.1.content.0.runs.0.action'],
    ['warning', 'table', 'slides.0.elements.0.children.2.rows.0.cells.0.content.0.runs.0.action'],
  ]);
  assert.doesNotThrow(() => normalizePresentation(deck));
});
