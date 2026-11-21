import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { inspectPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('atomic section structure edits agree across Node and browser output', async t => {
  async function exercise(api) {
    const doc = api.createPresentation({ id: 'deck' });
    for (const id of ['a', 'b', 'c']) doc.addSlide({ id });
    doc.setSections([{ id: 'one', name: 'same', slideIds: ['a', 'b'] }, { id: 'two', name: 'same', slideIds: ['c'] }]);
    doc.setSectionStructure(['c', 'a', 'b'], [{ id: 'two', name: 'same', slideIds: ['c', 'a'] }, { id: 'one', name: 'same', slideIds: ['b'] }]);
    let rejected = false;
    try { doc.setSectionStructure(['a', 'a', 'b'], []); } catch { rejected = true; }
    return { order: doc.slides.map(s => s.id), sections: doc.sections, rejected, bytes: [...(await api.generatePptx(doc)).bytes] };
  }
  const expected = await exercise({ createPresentation, generatePptx });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const pkg of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${pkg}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(`(${exercise.toString()})(PPTKit)`);
  assert.deepEqual(actual, expected); assert.equal(actual.rejected, true);
});

test('serialized edit batch and inverse agree across Node, browser and PPTX output', async t => {
  const operation = { schemaVersion: 'pptkit-edit-v1', documentId: 'deck', changes: [
    { type: 'slide-hidden', slideId: 'a', expected: false, value: true },
    { type: 'first-slide-number', expected: null, value: 17 },
    { type: 'slide-order', expected: ['a', 'b'], value: ['b', 'a'] },
    { type: 'text-run', slideId: 'a', elementId: 'text-a', paragraphIndex: 0, runIndex: 0, expected: 'A', value: 'Edited A' },
    { type: 'text-run', slideId: 'a', elementId: 'shape-a', paragraphIndex: 0, runIndex: 0, expected: 'Shape', value: 'Edited shape' },
    { type: 'table-cell-text-run', slideId: 'b', elementId: 'table', rowIndex: 0, cellIndex: 0, paragraphIndex: 0, runIndex: 0, expected: 'Cell', value: 'Edited cell' },
    { type: 'text-string', slideId: 'b', elementId: 'plain', expected: 'Plain', value: 'Edited plain' },
    { type: 'table-cell-text-string', slideId: 'b', elementId: 'table', rowIndex: 0, cellIndex: 1, expected: 'Short', value: 'Edited short' },
    { type: 'element-box', slideId: 'a', elementId: 'group', expected: { x: 100, y: 10, width: 100, height: 100 }, value: { x: 130, y: 10, width: 100, height: 100 } },
    { type: 'element-hidden', slideId: 'b', elementId: 'plain', expected: null, value: true },
    { type: 'element-opacity', slideId: 'a', elementId: 'group', expected: null, value: 0.5 },
    { type: 'element-name', slideId: 'a', elementId: 'shape-a', expected: null, value: 'Edited & named' },
    { type: 'element-transform', slideId: 'a', elementId: 'shape-a', expected: null, value: { rotation: 30, flipH: true, flipV: false } },
    { type: 'element-accessibility', slideId: 'a', elementId: 'shape-a', expected: null, value: { description: 'Decorative detail', decorative: true } },
    { type: 'element-action', slideId: 'a', elementId: 'shape-a', expected: null, value: { type: 'slide', slideId: 'b', tooltip: 'Go to B' } },
    { type: 'text-run-action', slideId: 'a', elementId: 'text-a', paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'url', url: 'https://example.org/path', tooltip: 'Example' } },
    { type: 'table-cell-text-run-action', slideId: 'b', elementId: 'table', rowIndex: 0, cellIndex: 0, paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'slide', slideId: 'a', tooltip: 'Go to A' } },
    { type: 'notes-run', slideId: 'a', paragraphIndex: 0, runIndex: 0, expected: 'Notes', value: 'Edited notes' },
    { type: 'notes-run-action', slideId: 'a', paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'slide', slideId: 'b', tooltip: 'Notes link' } },
    { type: 'chart-point', slideId: 'b', elementId: 'chart', seriesIndex: 0, pointIndex: 0, coordinate: 'values', expected: 1, value: '-0' },
    { type: 'chart-category', slideId: 'b', elementId: 'chart', categoryIndex: 0, expected: 'A', value: '二 & two' },
    { type: 'chart-series-name', slideId: 'b', elementId: 'chart', seriesIndex: 0, expected: 'S', value: 'Revenue' },
  ] };
  const text = value => ({ id: `text-${value.toLowerCase()}`, type: 'text', box: { x: 10, y: 10, width: 80, height: 30 }, content: [{ runs: [{ text: value }] }] });
  const group = () => ({ id: 'group', type: 'group', box: { x: 100, y: 10, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [{ id: 'shape-a', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 90, height: 60 }, text: { content: [{ runs: [{ text: 'Shape' }] }] } }] });
  const table = () => ({ id: 'table', type: 'table', box: { x: 100, y: 100, width: 100, height: 40 }, columns: [50, 50], rows: [{ cells: [{ content: [{ runs: [{ text: 'Cell' }] }] }, { content: 'Short' }] }] });
  const plain = () => ({ id: 'plain', type: 'text', box: { x: 10, y: 50, width: 80, height: 30 }, content: 'Plain' });
  const chart = () => ({ id: 'chart', type: 'chart', chartType: 'bar', box: { x: 250, y: 100, width: 150, height: 100 }, categories: ['A'], series: [{ name: 'S', values: [1] }] });
  const doc = createPresentation({ id: 'deck' }); doc.addSlide({ id: 'a', notes: [{ runs: [{ text: 'Notes' }] }], elements: [text('A'), group()] }); doc.addSlide({ id: 'b', elements: [text('B'), table(), plain(), chart()] });
  doc.slides[1].setReadingOrder([{ elementIds: ['plain', 'chart', 'table', 'text-b'] }]);
  const inverse = doc.applyEditOperation(operation).inverse;
  const chartSvg = (await renderPresentationToSvg(doc)).slides[0].svg;
  assert.match(chartSvg, /二 &amp; two/);
  const bytes = (await generatePptx(doc)).bytes, inspected = await inspectPptx(bytes);
  assert.equal(inspected.firstSlideNumber, 17);
  const editedParts = readGeneratedZip(bytes);
  assert.match(editedParts.get('ppt/slides/slide1.xml').toString(), /<a:t>B<\/a:t>/);
  assert.match(editedParts.get('ppt/slides/slide1.xml').toString(), /<a:t>Edited cell<\/a:t>/);
  assert.doesNotMatch(editedParts.get('ppt/slides/slide1.xml').toString(), /<a:t>Edited plain<\/a:t>/);
  assert.match(editedParts.get('ppt/slides/slide1.xml').toString(), /<a:t>Edited short<\/a:t>/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /\bshow="0"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:t>Edited A<\/a:t>/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:t>Edited shape<\/a:t>/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /name="Edited &amp; named"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /descr="Decorative detail"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<adec:decorative[^>]+val="1"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /rot="1800000" flipH="1"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:hlinkClick[^>]*tooltip="Go to B"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:hlinkClick[^>]*tooltip="Example"/);
  assert.match(editedParts.get('ppt/slides/slide1.xml').toString(), /<a:hlinkClick[^>]*tooltip="Go to A"/);
  assert.match(editedParts.get('ppt/slides/_rels/slide2.xml.rels').toString(), /Target="https:\/\/example\.org\/path" TargetMode="External"/);
  assert.match(editedParts.get('ppt/slides/_rels/slide2.xml.rels').toString(), /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide" Target="slide1\.xml"/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:off x="1651000" y="127000"\/>/);
  assert.match(editedParts.get('ppt/slides/slide2.xml').toString(), /<a:alpha val="50000"\/>/);
  assert.match(editedParts.get('ppt/notesSlides/notesSlide2.xml').toString(), /<a:t>Edited notes<\/a:t>/);
  assert.match(editedParts.get('ppt/notesSlides/notesSlide2.xml').toString(), /tooltip="Notes link"/);
  assert.match(editedParts.get('ppt/charts/chart1.xml').toString(), /<c:v>-0<\/c:v>/);
  assert.match(editedParts.get('ppt/charts/chart1.xml').toString(), /二 &amp; two/);
  assert.match(editedParts.get('ppt/charts/chart1.xml').toString(), /Revenue/);
  const chartBook = readGeneratedZip(editedParts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'));
  assert.match(chartBook.get('xl/worksheets/sheet1.xml').toString(), /二 &amp; two/);
  assert.match(chartBook.get('xl/worksheets/sheet1.xml').toString(), /Revenue/);
  doc.applyEditOperation(inverse);
  assert.deepEqual(doc.slides[1].readingOrder, [{ elementIds: ['plain', 'chart', 'table', 'text-b'] }]);
  const reverted = readGeneratedZip((await generatePptx(doc)).bytes);
  assert.doesNotMatch(reverted.get('ppt/slides/slide1.xml').toString(), /\bshow="0"/);
  assert.match(reverted.get('ppt/slides/slide1.xml').toString(), /<a:t>A<\/a:t>/);
  assert.match(reverted.get('ppt/slides/slide1.xml').toString(), /<a:t>Shape<\/a:t>/);
  assert.doesNotMatch(reverted.get('ppt/slides/slide1.xml').toString(), /name="Edited &amp; named"/);
  assert.doesNotMatch(reverted.get('ppt/slides/slide1.xml').toString(), /<adec:decorative[^>]+val="1"/);
  assert.doesNotMatch(reverted.get('ppt/slides/slide1.xml').toString(), /rot="1800000"/);
  assert.match(reverted.get('ppt/slides/slide1.xml').toString(), /<a:off x="1270000" y="127000"\/>/);
  assert.doesNotMatch(reverted.get('ppt/presentation.xml').toString(), /\bfirstSlideNum=/);
  assert.match(reverted.get('ppt/slides/slide2.xml').toString(), /<a:t>Plain<\/a:t>/);
  assert.match(reverted.get('ppt/notesSlides/notesSlide1.xml').toString(), /<a:t>Notes<\/a:t>/);
  assert.doesNotMatch(reverted.get('ppt/notesSlides/notesSlide1.xml').toString(), /tooltip="Notes link"/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const pkg of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${pkg}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async operation => {
    const text = value => ({ id: `text-${value.toLowerCase()}`, type: 'text', box: { x: 10, y: 10, width: 80, height: 30 }, content: [{ runs: [{ text: value }] }] });
    const group = () => ({ id: 'group', type: 'group', box: { x: 100, y: 10, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [{ id: 'shape-a', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 90, height: 60 }, text: { content: [{ runs: [{ text: 'Shape' }] }] } }] });
    const table = () => ({ id: 'table', type: 'table', box: { x: 100, y: 100, width: 100, height: 40 }, columns: [50, 50], rows: [{ cells: [{ content: [{ runs: [{ text: 'Cell' }] }] }, { content: 'Short' }] }] });
    const plain = () => ({ id: 'plain', type: 'text', box: { x: 10, y: 50, width: 80, height: 30 }, content: 'Plain' });
    const chart = () => ({ id: 'chart', type: 'chart', chartType: 'bar', box: { x: 250, y: 100, width: 150, height: 100 }, categories: ['A'], series: [{ name: 'S', values: [1] }] });
    const doc = PPTKit.createPresentation({ id: 'deck' }); doc.addSlide({ id: 'a', notes: [{ runs: [{ text: 'Notes' }] }], elements: [text('A'), group()] }); doc.addSlide({ id: 'b', elements: [text('B'), table(), plain(), chart()] });
    doc.slides[1].setReadingOrder([{ elementIds: ['plain', 'chart', 'table', 'text-b'] }]);
    const inverse = doc.applyEditOperation(operation).inverse, bytes = (await PPTKit.generatePptx(doc)).bytes;
    doc.applyEditOperation(inverse);
    return { inverse, bytes: Array.from(bytes), visible: doc.slides.map(slide => slide.hidden), number: doc.firstSlideNumber ?? null, readingOrder: doc.slides[1].readingOrder };
  }, operation);
  assert.deepEqual(actual.inverse, inverse); assert.deepEqual(new Uint8Array(actual.bytes), bytes);
  assert.deepEqual(actual.visible, [false, false]); assert.equal(actual.number, null);
  assert.deepEqual(actual.readingOrder, [{ elementIds: ['plain', 'chart', 'table', 'text-b'] }]);
});
