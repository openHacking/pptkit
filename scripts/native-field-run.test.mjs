import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, materializeFieldText, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx, inspectPptxNotes } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('authored slide-number fields retain identity and caller cache in slide and notes XML', async () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', content = materializeFieldText([{ runs: [{ text: 'Page ' }, { field: { id, kind: 'slide-number' }, output: 'native', style: { bold: true } }] }], { pageOrder: ['a'], pageId: 'a', firstPageNumber: 42 }).content;
  const shapeContent = structuredClone(content), tableContent = structuredClone(content); shapeContent[0].runs[1].field.id = '11111111-2222-3333-4444-555555555555'; tableContent[0].runs[1].field.id = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
  const deck = createPresentation(); deck.addSlide({ id: 'a', notes: content, elements: [{ type: 'text', id: 't', box: { x: 10, y: 10, width: 200, height: 40 }, content }, { type: 'shape', id: 'shape', shape: 'rect', box: { x: 10, y: 60, width: 200, height: 40 }, text: { content: shapeContent } }, { type: 'table', id: 'table', box: { x: 10, y: 110, width: 200, height: 40 }, columns: [200], rows: [{ height: 40, cells: [{ content: tableContent }] }] }] }); deck.duplicateSlide('a');
  const generated = await generatePptx(deck), parts = readGeneratedZip(generated.bytes);
  for (const file of ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/notesSlides/notesSlide1.xml', 'ppt/notesSlides/notesSlide2.xml']) {
    const xml = parts.get(file).toString(); assert.match(xml, /<a:fld id="\{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A\}" type="slidenum">/); assert.match(xml, /<a:t>42<\/a:t><\/a:fld>/);
    assert.equal((xml.match(/<a:fld\b/g) ?? []).length, file.includes('notesSlides') ? 1 : 3);
  }
  const notes = await inspectPptxNotes(generated.bytes);
  for (const page of notes.pages) { const part = notes.parts.find(item => item.part === page.notesPart), field = part.fields.find(item => item.type === 'slidenum'); assert.equal(field.id, `{${id.toUpperCase()}}`); assert.equal(field.cachedTexts[0].value, '42'); }
  const svg = await renderPresentationToSvg(deck); assert.equal(svg.slides.length, 2); for (const slide of svg.slides) assert.match(slide.svg, />42</);
  let bytes = generated.bytes;
  for (let round = 0; round < 3; round++) { const session = await importPptx(bytes, { projection: 'structure' }); try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); const next = (await session.save(plan.plan)).bytes; assert.deepEqual(next, bytes); bytes = next; } finally { session.dispose(); } }
  bytes = generated.bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content', fieldTextProfile: 'slide-number-v1' });
    try {
      for (const slide of normalizePresentation(session.document).slides) {
        const text = slide.elements.find(element => element.type === 'text'), shape = slide.elements.find(element => element.type === 'shape'), table = slide.elements.find(element => element.type === 'table');
        for (const runs of [text.content[0].runs, shape.text.content[0].runs, table.rows[0].cells[0].content[0].runs, slide.notes[0].runs]) { assert.equal(runs[1].text, '42'); assert.equal(runs[1].field.kind, 'slide-number'); }
      }
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test('native field output matches the browser global without network access', async t => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', deck = createPresentation();
  deck.addSlide({ id: 'a', elements: [{ type: 'text', id: 't', box: { x: 10, y: 10, width: 200, height: 40 }, content: [{ runs: [{ text: '42', field: { id, kind: 'slide-number' } }] }] }] });
  const expected = (await generatePptx(deck)).bytes, browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const result = await page.evaluate(async id => { const deck = PPTKit.createPresentation(); deck.addSlide({ id: 'a', elements: [{ type: 'text', id: 't', box: { x: 10, y: 10, width: 200, height: 40 }, content: [{ runs: [{ text: '42', field: { id, kind: 'slide-number' } }] }] }] }); return Array.from((await PPTKit.generatePptx(deck)).bytes); }, id);
  assert.deepEqual(new Uint8Array(result), expected); assert.deepEqual(requests, []);
});

test('native page-number fields resolve the cached script font without splitting field identity', async () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: [{ runs: [{ text: '١', field: { id, kind: 'slide-number' }, style: { scriptFontFamilies: { Arab: 'FieldArabic' } } }] }] }] });
  const result = await generatePptx(deck), xml = readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((xml.match(/<a:fld\b/g) ?? []).length, 1); assert.match(xml, /<a:latin typeface="FieldArabic"\/>/); assert.ok(result.warnings.some(warning => warning.code === 'script-font-map-materialized'));
});

test('adjacent text saves across owners without writing field caches or IDs', async () => {
  const id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', content = [{ runs: [{ text: 'before' }, { text: '42', field: { id, kind: 'slide-number' } }] }], deck = createPresentation();
  deck.addSlide({ id: 'a', notes: content, elements: [{ type: 'text', id: 't', box: { x: 10, y: 10, width: 200, height: 40 }, content }, { type: 'shape', id: 's', shape: 'rect', box: { x: 10, y: 60, width: 200, height: 40 }, text: { content } }, { type: 'table', id: 'table', box: { x: 10, y: 110, width: 200, height: 40 }, columns: [200], rows: [{ height: 40, cells: [{ content }] }] }] });
  const source = (await generatePptx(deck)).bytes, options = { projection: 'content', fieldTextProfile: 'slide-number-v1', editSaveProfile: 'text-runs-v3' }, session = await importPptx(source, options);
  try {
    const slide = session.document.slides[0], notes = structuredClone(slide.notes); notes[0].runs[0].text = 'edited'; slide.setNotes(notes);
    for (const element of [...slide.elements]) {
      const next = structuredClone(element), paragraphs = next.type === 'text' ? next.content : next.type === 'shape' ? next.text.content : next.rows[0].cells[0].content;
      paragraphs[0].runs[0].text = 'edited'; slide.replaceElement(element.id, next);
    }
    const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
    const result = (await session.save(prepared.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(result);
    for (const [name, payload] of before) {
      if (['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'].includes(name)) {
        const a = payload.toString(), b = after.get(name).toString();
        assert.equal(b, a.replaceAll('<a:t>before</a:t>', '<a:t>edited</a:t>'));
        assert.deepEqual([...b.matchAll(/<a:fld\b[\s\S]*?<\/a:fld>/g)].map(match => match[0]), [...a.matchAll(/<a:fld\b[\s\S]*?<\/a:fld>/g)].map(match => match[0]));
      } else assert.deepEqual(after.get(name), payload);
    }
  } finally { session.dispose(); }
});

test('native date update formats and caller caches serialize independently', async t => {
  const formats = [['application-default', 'datetime'], ['month-day-year', 'datetime1'], ['weekday-month-day-year', 'datetime2'], ['day-month-year', 'datetime3'], ['month-day-year-long', 'datetime4'], ['automatic-month-day-year', 'datetimeFigureOut'], ['day-short-month-year', 'datetime5'], ['month-year', 'datetime6'], ['short-month-year', 'datetime7'], ['month-day-year-time-12', 'datetime8'], ['month-day-year-time-12-seconds', 'datetime9'], ['time-24', 'datetime10'], ['time-24-seconds', 'datetime11'], ['time-12', 'datetime12'], ['time-12-seconds', 'datetime13']];
  const content = [{ runs: formats.map(([updateFormat], i) => ({ text: `cache-${i}`, field: { id: `38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4${i.toString(16)}`, kind: 'dynamic-date', updateFormat } })) }], deck = createPresentation();
  deck.addSlide({ id: 'date', notes: content, elements: [{ type: 'text', box: { x: 10, y: 10, width: 400, height: 50 }, content }] });
  const bytes = (await generatePptx(deck)).bytes, parts = readGeneratedZip(bytes);
  for (const name of ['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml']) {
    const xml = parts.get(name).toString();
    const actual = [...xml.matchAll(/<a:fld\b[^>]*type="([^"]+)"[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t><\/a:fld>/g)].map(match => [match[1], match[2]]);
    assert.deepEqual(actual, formats.map(([, type], i) => [type, `cache-${i}`]));
  }
  const inventory = await inspectPptxNotes(bytes), page = inventory.pages[0], notes = inventory.parts.find(item => item.part === page.notesPart);
  assert.deepEqual(notes.fields.map(field => field.type), formats.map(([, type]) => type));
  assert.deepEqual(notes.fields.map(field => field.cachedTexts[0].value), formats.map((_, i) => `cache-${i}`));
  let preserved = bytes;
  for (let round = 0; round < 3; round++) { const session = await importPptx(preserved, { projection: 'structure' }); try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); const next = (await session.save(plan.plan)).bytes; assert.deepEqual(next, preserved); preserved = next; } finally { session.dispose(); } }
  let fresh = bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(fresh, { projection: 'content', fieldTextProfile: 'date-time-and-number-v1' });
    try { const slide = normalizePresentation(session.document).slides[0]; for (const runs of [slide.elements[0].content[0].runs, slide.notes[0].runs]) { assert.deepEqual(runs.map(run => run.field?.updateFormat), formats.map(([format]) => format)); assert.deepEqual(runs.map(run => run.text), formats.map((_, i) => `cache-${i}`)); } fresh = (await generatePptx(session.document)).bytes; } finally { session.dispose(); }
  }
  const browser = await chromium.launch(); t.after(() => browser.close()); const browserPage = await browser.newPage(), requests = [];
  await browserPage.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter']) await browserPage.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const browserBytes = await browserPage.evaluate(async content => { const doc = PPTKit.createPresentation(); doc.addSlide({ id: 'date', notes: content, elements: [{ type: 'text', box: { x: 10, y: 10, width: 400, height: 50 }, content }] }); return Array.from((await PPTKit.generatePptx(doc)).bytes); }, content);
  assert.deepEqual(new Uint8Array(browserBytes), bytes); assert.deepEqual(requests, []);
});
