import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, evaluatePresentationField, fieldContextForOccurrences, fieldContextForSlide, materializeFieldText } from '../packages/core/dist/index.js';

test('field evaluation and failures agree in Node, browser ESM and global', async t => {
  const cases = [
    [{ id: 'page', kind: 'slide-number' }, { pageOrder: ['a', 'b'], pageId: 'b', firstPageNumber: 2 }],
    [{ id: 'date', kind: 'dynamic-date', dateStyle: 'short' }, { instantMs: 0, locale: 'en-US', timeZone: 'UTC' }],
    [{ id: 'time', kind: 'dynamic-date', timeStyle: 'medium', updateFormat: 'time-24-seconds' }, { instantMs: 0, locale: 'en-US', timeZone: 'UTC' }],
    [{ id: 'footer', kind: 'footer', text: 'text' }, {}],
    [{ id: 'bad', kind: 'slide-number' }, {}],
  ];
  const run = (api, cases) => { const doc = api.createPresentation(); doc.addSlide({ id: 'a' }); doc.addSlide({ id: 'hidden', hidden: true }); doc.addSlide({ id: 'b' }); return { values: cases.map(([field, context]) => { try { return api.evaluatePresentationField(field, context); } catch (error) { return { error: error.code }; } }), snapshot: api.materializeFieldText([{ runs: [{ text: 'Page ' }, { field: { id: 'page', kind: 'slide-number' }, style: { bold: true } }, { field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' }, output: 'native' }] }], { pageOrder: ['a', 'b'], pageId: 'b', firstPageNumber: 2 }), context: api.fieldContextForSlide(doc, { slideId: 'b', firstPageNumber: 1, includeHidden: false }), occurrences: api.fieldContextForOccurrences(doc, { occurrences: [{ id: 'a1', slideId: 'a' }, { id: 'b', slideId: 'b' }, { id: 'a2', slideId: 'a' }], occurrenceId: 'a2', firstPageNumber: 1, includeHidden: true }) }; };
  const expected = run({ createPresentation, evaluatePresentationField, fieldContextForOccurrences, fieldContextForSlide, materializeFieldText }, cases), esm = await readFile('packages/core/dist/browser/index.js'), global = await readFile('packages/core/dist/browser/global.js', 'utf8');
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: global });
  const results = await page.evaluate(async ({ url, cases, source }) => { const module = await import(url), run = (0, eval)(`(${source})`); return [run(module, cases), run(PPTKit, cases)]; }, { url: `data:text/javascript;base64,${esm.toString('base64')}`, cases, source: run.toString() });
  for (const result of results) assert.deepEqual(result, expected); assert.deepEqual(requests, []);
});
