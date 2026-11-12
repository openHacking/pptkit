import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const create = api => {
  const deck = api.createPresentation({ size: { width: 300, height: 200 } }), slide = deck.addSlide({ id: 'source' });
  const field = { text: 'stale', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' } };
  slide.addElement({ type: 'table', id: 'table', box: { x: 20, y: 30, width: 100, height: 90 }, columns: [100], rows: [{ height: 10, cells: [{ content: [{ runs: [field] }] }] }, ...Array.from({ length: 4 }, (_, i) => ({ height: 20, cells: [{ content: `row ${i}` }] }))] });
  deck.addSlide({ id: 'after', elements: [{ type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [field] }] }] }); return deck;
};
const options = { fieldRefresh: { firstPageNumber: 5, includeHidden: true }, tablePagination: [{ elementId: 'table', area: { x: 20, y: 30, width: 100, height: 55 }, headerRows: 1 }] };

test('post-pagination field caches reach SVG, native PPTX and reimport without mutating author state', async t => {
  const deck = create({ createPresentation }), baseline = normalizePresentation(deck), output = await generatePptx(deck, { layout: options }), preview = await renderPresentationToSvg(deck, { layout: options }), parts = readGeneratedZip(output.bytes);
  assert.equal(output.slideCount, 3);
  for (let i = 0; i < 3; i++) { const xml = parts.get(`ppt/slides/slide${i + 1}.xml`).toString(); assert.match(xml, new RegExp(`<a:t>${i + 5}</a:t></a:fld>`)); assert.match(preview.slides[i].svg, new RegExp(`>${i + 5}<`)); }
  assert.deepEqual(normalizePresentation(deck), baseline);
  let source = output.bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(source, { projection: 'content', fieldTextProfile: 'date-time-and-number-v1' });
    try { const doc = normalizePresentation(session.document); for (let i = 0; i < 3; i++) { const element = doc.slides[i].elements.find(e => e.type === 'table' || e.type === 'text'), runs = element.type === 'table' ? element.rows[0].cells[0].content[0].runs : element.content[0].runs; assert.equal(runs[0].text, String(i + 5)); assert.equal(runs[0].field.kind, 'slide-number'); } const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); const saved = (await session.save(plan.plan)).bytes; assert.deepEqual(saved, source); source = saved; } finally { session.dispose(); }
  }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const result = await page.evaluate(async ({ source, options }) => { const create = (0, eval)(`(${source})`), deck = create(PPTKit), generated = await PPTKit.generatePptx(deck, { layout: options }), rendered = await PPTKit.renderPresentationToSvg(deck, { layout: options }); return { bytes: Array.from(generated.bytes), svgs: rendered.slides.map(slide => slide.svg) }; }, { source: create.toString(), options });
  assert.deepEqual(new Uint8Array(result.bytes), output.bytes); assert.deepEqual(result.svgs, preview.slides.map(slide => slide.svg)); assert.deepEqual(requests, []);
});

test('shared date refresh retains master/layout ownership and appears on each inheriting page', async t => {
  const createShared = api => { const deck = api.createPresentation(), date = (id, text) => [{ runs: [{ text }, { text: 'old', field: { id, kind: 'dynamic-date', updateFormat: 'day-month-year' } }] }]; deck.defineSlideMaster({ id: 'm', name: 'Master', elements: [{ id: 'master-date', type: 'text', box: { x: 10, y: 10, width: 250, height: 30 }, content: date('11111111-2222-3333-4444-555555555555', 'M ') }] }); deck.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [{ id: 'layout-date', type: 'text', box: { x: 10, y: 50, width: 250, height: 30 }, content: date('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', 'L ') }] }); deck.addSlide({ id: 'one', layoutId: 'l' }); deck.addSlide({ id: 'two', layoutId: 'l' }); return deck; };
  const deck = createShared({ createPresentation });
  const baseline = normalizePresentation(deck), opts = { fieldRefresh: { firstPageNumber: 1, includeHidden: true, instantMs: 0, locale: 'en-US', timeZone: 'UTC', dateStyle: 'short' } };
  const output = await generatePptx(deck, { layout: opts }), rendered = await renderPresentationToSvg(deck, { layout: opts }), parts = readGeneratedZip(output.bytes);
  for (const slide of rendered.slides) assert.equal((slide.svg.match(/>1\/1\/70</g) ?? []).length, 2);
  const masters = [...parts].filter(([name]) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(name)).map(([, bytes]) => bytes.toString()), layouts = [...parts].filter(([name]) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)).map(([, bytes]) => bytes.toString());
  const master = masters.find(xml => xml.includes('11111111-2222-3333-4444-555555555555')), layout = layouts.find(xml => xml.includes('AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'));
  assert.ok(master); assert.ok(layout); assert.match(master, /<a:t>1\/1\/70<\/a:t><\/a:fld>/); assert.match(layout, /<a:t>1\/1\/70<\/a:t><\/a:fld>/);
  for (const name of ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml']) assert.doesNotMatch(parts.get(name).toString(), /<a:fld\b/);
  assert.deepEqual(normalizePresentation(deck), baseline);
  let source = output.bytes; for (let round = 0; round < 3; round++) { const session = await importPptx(source, { projection: 'structure' }); try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); const saved = (await session.save(plan.plan)).bytes; assert.deepEqual(saved, source); source = saved; } finally { session.dispose(); } }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ source, opts }) => { const create = (0, eval)(`(${source})`), deck = create(PPTKit), generated = await PPTKit.generatePptx(deck, { layout: opts }), rendered = await PPTKit.renderPresentationToSvg(deck, { layout: opts }); return { bytes: Array.from(generated.bytes), svgs: rendered.slides.map(slide => slide.svg) }; }, { source: createShared.toString(), opts });
  assert.deepEqual(new Uint8Array(actual.bytes), output.bytes); assert.deepEqual(actual.svgs, rendered.slides.map(slide => slide.svg)); assert.deepEqual(requests, []);
});

test('refreshed cache is used for intrinsic sizing in Node and the browser', async t => {
  const createSized = api => { const deck = api.createPresentation(); deck.addSlide({ elements: [{ id: 'number', type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: 'stale-cache', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' } }] }] }] }); return deck; };
  const makeOptions = () => ({ fieldRefresh: { firstPageNumber: 123, includeHidden: true }, intrinsicSizing: [{ elementId: 'number', width: { type: 'content' } }], textMeasurement: { measure: request => ({ bounds: { x: 0, y: 0, width: request.paragraphs.flatMap(p => p.runs).reduce((sum, run) => sum + run.text.length * 10, 0), height: 10 } }) } });
  const deck = createSized({ createPresentation }), baseline = normalizePresentation(deck), opts = makeOptions();
  const output = await generatePptx(deck, { layout: opts }), preview = await renderPresentationToSvg(deck, { layout: opts });
  const xml = readGeneratedZip(output.bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:ext cx="381000"/); assert.match(xml, /<a:t>123<\/a:t><\/a:fld>/); assert.match(preview.slides[0].svg, />123</);
  assert.deepEqual(normalizePresentation(deck), baseline);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ source, optionsSource }) => { const create = (0, eval)(`(${source})`), makeOptions = (0, eval)(`(${optionsSource})`), deck = create(PPTKit), opts = makeOptions(); const generated = await PPTKit.generatePptx(deck, { layout: opts }), rendered = await PPTKit.renderPresentationToSvg(deck, { layout: opts }); return { bytes: Array.from(generated.bytes), svgs: rendered.slides.map(slide => slide.svg) }; }, { source: createSized.toString(), optionsSource: makeOptions.toString() });
  assert.deepEqual(new Uint8Array(actual.bytes), output.bytes); assert.deepEqual(actual.svgs, preview.slides.map(slide => slide.svg)); assert.deepEqual(requests, []);
});

test('pagination and table sizing accept stable dates and reject changing measured page fields across runtimes', async t => {
  const createDeck = (api, kind) => { const deck = api.createPresentation(), field = kind === 'date' ? { text: 'old', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'dynamic-date', updateFormat: 'day-month-year' } } : { text: 'old', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' } }; deck.addSlide({ elements: [{ id: 'table', type: 'table', box: { x: 10, y: 10, width: 100, height: 70 }, columns: [100], rows: [{ height: 10, cells: [{ content: [{ runs: [field] }] }] }, ...Array.from({ length: 3 }, (_, i) => ({ height: 20, cells: [{ content: `row ${i}` }] }))] }] }); return deck; };
  const makeOptions = () => ({ fieldRefresh: { firstPageNumber: 1, includeHidden: true, instantMs: 0, locale: 'en-US', timeZone: 'UTC', dateStyle: 'short' }, tableSizing: [{ elementId: 'table', autoRows: [{ row: 0 }] }], tablePagination: [{ elementId: 'table', area: { x: 10, y: 10, width: 100, height: 45 }, headerRows: 1 }], textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 10 } }) } });
  const output = await generatePptx(createDeck({ createPresentation }, 'date'), { layout: makeOptions() }), preview = await renderPresentationToSvg(createDeck({ createPresentation }, 'date'), { layout: makeOptions() });
  assert.ok(output.slideCount > 1); for (const slide of preview.slides) assert.match(slide.svg, />1\/1\/70</);
  await assert.rejects(generatePptx(createDeck({ createPresentation }, 'number'), { layout: makeOptions() }), /changed after pagination/);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const actual = await page.evaluate(async ({ source, optionsSource }) => { const create = (0, eval)(`(${source})`), options = (0, eval)(`(${optionsSource})`); const generated = await PPTKit.generatePptx(create(PPTKit, 'date'), { layout: options() }), rendered = await PPTKit.renderPresentationToSvg(create(PPTKit, 'date'), { layout: options() }); let error = ''; try { await PPTKit.generatePptx(create(PPTKit, 'number'), { layout: options() }); } catch (e) { error = e.message; } return { bytes: Array.from(generated.bytes), svgs: rendered.slides.map(slide => slide.svg), error }; }, { source: createDeck.toString(), optionsSource: makeOptions.toString() });
  assert.deepEqual(new Uint8Array(actual.bytes), output.bytes); assert.deepEqual(actual.svgs, preview.slides.map(slide => slide.svg)); assert.match(actual.error, /changed after pagination/); assert.deepEqual(requests, []);
});

test('copied and reordered slides refresh their own caches without changing field identities or the authoring state', async () => {
  const deck = createPresentation(), id = '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a';
  const field = { text: 'authored', field: { id, kind: 'slide-number' } };
  deck.addSlide({ id: 'original', notes: [{ runs: [field] }], elements: [{ id: 'label', type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [field] }] }] });
  deck.addSlide({ id: 'hidden', hidden: true, elements: [{ id: 'hidden-label', type: 'text', box: { x: 10, y: 10, width: 100, height: 30 }, content: [{ runs: [field] }] }] });
  const copy = deck.duplicateSlide('original'); deck.moveSlide(copy.id, 2);
  const before = normalizePresentation(deck), layout = { fieldRefresh: { firstPageNumber: 8, includeHidden: false } };
  const output = await generatePptx(deck, { layout }), preview = await renderPresentationToSvg(deck, { layout }), parts = readGeneratedZip(output.bytes);
  for (const [index, value] of [[1, '8'], [3, '9']]) {
    const xml = parts.get(`ppt/slides/slide${index}.xml`).toString(), note = parts.get(`ppt/notesSlides/notesSlide${index}.xml`).toString();
    assert.match(xml, new RegExp(`id="\\{${id.toUpperCase()}\\}" type="slidenum"`));
    assert.match(xml, new RegExp(`<a:t>${value}</a:t></a:fld>`)); assert.match(note, new RegExp(`<a:t>${value}</a:t></a:fld>`));
    assert.match(preview.slides[index - 1].svg, new RegExp(`>${value}<`));
  }
  assert.match(parts.get('ppt/slides/slide2.xml').toString(), /<a:t>authored<\/a:t><\/a:fld>/);
  assert.match(preview.slides[1].svg, />authored</);
  assert.deepEqual(normalizePresentation(deck), before);
  const session = await importPptx(output.bytes, { projection: 'content', fieldTextProfile: 'date-time-and-number-v1' });
  try {
    const imported = normalizePresentation(session.document);
    assert.deepEqual([0, 2].map(i => imported.slides[i].elements.find(e => e.type === 'text').content[0].runs[0].field.id.toLowerCase()), [id, id]);
    assert.deepEqual([0, 2].map(i => imported.slides[i].elements.find(e => e.type === 'text').plainText), ['8', '9']);
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual((await session.save(plan.plan)).bytes, output.bytes);
  } finally { session.dispose(); }
});
