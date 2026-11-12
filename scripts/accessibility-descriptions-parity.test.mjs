import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import * as core from '../packages/core/dist/index.js';

function makeDeck(api) {
  const deck = api.createPresentation();
  const slide = deck.addSlide({ id: 'slide', elements: [
    { id: 'table', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] },
    { id: 'chart', type: 'chart', chartType: 'bar', box: { x: 100, y: 0, width: 100, height: 50 }, categories: ['A'], series: [{ name: 'S', values: [1] }], accessibility: { description: 'Sales' } },
  ] });
  slide.reorderElements(['chart', 'table']);
  const grouped = deck.addSlide({ id: 'grouped-slide', elements: [{ id: 'group', type: 'group', box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [{ id: 'table-a', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] }, { id: 'table-b', type: 'table', box: { x: 100, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'B' }] }] }] }] });
  grouped.reorderGroupElements('group', ['table-b', 'table-a']);
  deck.addSlide({ id: 'text-slide', elements: [{ id: 'text', type: 'text', box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [{ text: 'Contrast', style: { color: '777777', fontSize: 18 } }] }] }] });
  deck.addSlide({ id: 'alpha-slide', elements: [{ id: 'alpha-text', type: 'text', box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [{ text: 'Alpha', style: { color: '000000', colorOpacity: 0.534681, fontSize: 12 } }] }] }] });
  deck.addSlide({ id: 'cover-slide', background: { type: 'solid', color: 'FFFFFF' }, elements: [{ id: 'cover', type: 'shape', shape: 'rect', box: { x: 0, y: 0, ...deck.size }, style: { fill: { type: 'solid', color: '002244' }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [{ text: 'White', style: { color: 'FFFFFF', fontSize: 12 } }] }] } }] });
  deck.addSlide({ id: 'translucent-cover-slide', background: { type: 'solid', color: 'FFFFFF' }, elements: [{ id: 'translucent-cover', type: 'shape', shape: 'rect', box: { x: 0, y: 0, ...deck.size }, style: { fill: { type: 'solid', color: '000000', opacity: 0.5446693 }, stroke: { paint: { type: 'none' }, width: 0 } }, text: { content: [{ runs: [{ text: 'Black', style: { color: '000000', fontSize: 12 } }] }] } }] });
  deck.defineSlideLayout({ id: 'art-layout', name: 'Art', elements: [{ id: 'art', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 } }, { id: 'inherited-table', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] }] });
  deck.addSlide({ id: 'suppressed-slide', layoutId: 'art-layout', showInheritedShapes: false, elements: [{ id: 'suppressed-text', type: 'text', box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [{ text: 'Black', style: { color: '000000', fontSize: 12 } }] }] }] });
  deck.addSlide({ id: 'inherited-slide', layoutId: 'art-layout' });
  return deck;
}

test('description inspection agrees across Node and browser public exports', async t => {
  const deck = makeDeck(core), expected = [core.inspectMissingDescriptions(deck), core.inspectSlideTextContrast(deck), core.inspectEffectiveMissingDescriptions(deck)];
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(); const requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(new URL('../packages/core/dist/browser/global.js', import.meta.url), 'utf8') });
  const bundle = await readFile(new URL('../packages/core/dist/browser/index.js', import.meta.url), 'utf8');
  const result = await page.evaluate(async ({ source, factory }) => {
    const esm = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`);
    const create = (0, eval)(`(${factory})`);
    return [esm, globalThis.PPTKit].map(api => { const deck = create(api); return [api.inspectMissingDescriptions(deck), api.inspectSlideTextContrast(deck), api.inspectEffectiveMissingDescriptions(deck)]; });
  }, { source: bundle, factory: makeDeck.toString() });
  for (const actual of result) assert.deepEqual(actual, expected);
  assert.equal(expected[1].findings.find(finding => finding.slideId === 'text-slide').reason, 'output-size-dependent');
  assert.equal(expected[1].findings.find(finding => finding.slideId === 'alpha-slide').reason, 'output-opacity-dependent');
  assert.equal(expected[1].findings.find(finding => finding.slideId === 'suppressed-slide').status, 'meets-threshold');
  assert.ok(expected[1].findings.find(finding => finding.slideId === 'cover-slide').ratio > 15);
  assert.equal(expected[1].findings.find(finding => finding.slideId === 'translucent-cover-slide').reason, 'output-opacity-dependent');
  assert.deepEqual(expected[2].findings.filter(finding => finding.elementId === 'inherited-table').map(finding => finding.slideId), ['inherited-slide']);
  assert.deepEqual(expected[2].findings.filter(finding => finding.slideId === 'grouped-slide').map(finding => finding.elementId), ['table-b', 'table-a']);
  assert.deepEqual(requests, []);
});
