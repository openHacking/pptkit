import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

function deckFor(core) {
  const deck = core.createPresentation({ id: 'decorative-example' });
  deck.addSlide({ id: 'page', elements: [{ id: 'shape', name: 'Example', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { decorative: true, description: 'Alt text' } }] });
  return deck;
}

test('decorative export matches browser and survives projection, fresh export and original-byte no-op', async t => {
  const result = await generatePptx(deckFor({ createPresentation }));
  let roundBytes = result.bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(roundBytes, { projection: 'content' });
    try {
      assert.equal(session.document.slides[0].elements[0].accessibility.decorative, true);
      assert.equal(session.document.slides[0].elements[0].accessibility.description, 'Alt text');
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
      assert.deepEqual((await session.save(plan.plan)).bytes, roundBytes);
      roundBytes = (await generatePptx(session.document)).bytes;
      assert.match(readGeneratedZip(roundBytes).get('ppt/slides/slide1.xml').toString(), /<adec:decorative[^>]+val="1"/);
    } finally { session.dispose(); }
  }
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js'));
  const layoutUrl = url(await readFile('packages/layout/dist/browser/index.js'));
  const exporterUrl = url(await readFile('packages/pptx-exporter/dist/browser/index.js'));
  const parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl, '@pptkit/layout': layoutUrl } })}</script>`);
  const browserResult = await page.evaluate(async ({ coreUrl, exporterUrl, parserUrl, factory }) => {
    const core = await import(coreUrl), exporter = await import(exporterUrl), parser = await import(parserUrl), make = (0, eval)(`(${factory})`);
    const bytes = (await exporter.generatePptx(make(core))).bytes, session = await parser.importPptx(bytes, { projection: 'content' });
    try { return { bytes: Array.from(bytes), accessibility: session.document.slides[0].elements[0].accessibility }; }
    finally { session.dispose(); }
  }, { coreUrl, exporterUrl, parserUrl, factory: deckFor.toString() });
  assert.deepEqual(Uint8Array.from(browserResult.bytes), result.bytes);
  assert.deepEqual(browserResult.accessibility, { description: 'Alt text', decorative: true });
  assert.match(readGeneratedZip(Uint8Array.from(browserResult.bytes)).get('ppt/slides/slide1.xml').toString(), /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('fresh projections retain decorative flags across mapped element types and nested groups', async () => {
  const deck = createPresentation(), box = { x: 0, y: 0, width: 100, height: 50 }, accessibility = { decorative: true, description: 'Description' };
  deck.registerAsset({ id: 'asset', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC' }, mimeType: 'image/png', width: 8, height: 4 });
  deck.addSlide({ elements: [
    { id: 'text', type: 'text', box, content: 'Text', accessibility },
    { id: 'image', type: 'image', assetId: 'asset', box, accessibility },
    { id: 'shape', type: 'shape', shape: 'rect', box, accessibility },
    { id: 'connector', type: 'connector', start: { x: 10, y: 10 }, end: { x: 30, y: 30 }, accessibility },
    { id: 'chart', type: 'chart', box, chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }], accessibility },
    { id: 'table', type: 'table', box, columns: [100], rows: [{ cells: [{ content: 'Cell' }] }], accessibility },
    { id: 'group', type: 'group', box, coordinateSize: { width: 100, height: 50 }, accessibility, children: [{ id: 'child', type: 'shape', shape: 'rect', box, accessibility }] },
  ] });
  let bytes = (await generatePptx(deck, { connectorMode: 'native' })).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const elements = session.document.slides[0].elements;
      assert.deepEqual(elements.map(e => e.type), ['text', 'image', 'shape', 'connector', 'chart', 'table', 'group']);
      assert.ok(elements.every(e => e.accessibility.decorative === true && e.accessibility.description === 'Description'));
      assert.equal(elements[6].children[0].accessibility.decorative, true);
      assert.equal(session.report.findings.filter(f => f.feature === 'drawing.decorative').length, 8);
      bytes = (await generatePptx(session.document, { connectorMode: 'native' })).bytes;
    } finally { session.dispose(); }
  }
});

test('decorative text preservation agrees between Node and browser for three rounds', async t => {
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 50 }, content: 'before', accessibility: { decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 1; round <= 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v1' });
      try {
        const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
        if (next.accessibility.decorative !== true) throw new Error('Decorative flag missing');
        next.content[0].runs[0].text = `round-${round}`; slide.replaceElement(original.id, next);
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const result = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(result), expected);
  assert.match(readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString(), /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('existing description preservation matches Node and browser for three rounds', async t => {
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { description: 'Original', decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (const value of ['A & B', '中文😀', 'tab\tend']) {
      const session = await api.importPptx(bytes, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
      try {
        const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
        next.accessibility.description = value; slide.replaceElement(original.id, next);
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected);
  const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /descr="tab&#x9;end"/); assert.match(xml, /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('description attribute insertion and removal match Node and browser', async t => {
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { description: 'Original', decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (const value of [null, '', 'A & 中文😀']) {
      const session = await api.importPptx(bytes, { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' });
      try {
        const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
        if (value === null) delete next.accessibility.description; else next.accessibility.description = value;
        slide.replaceElement(original.id, next);
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected);
  const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /descr="A &amp; 中文😀"/); assert.match(xml, /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('batch description edits match Node and browser for three rounds', async t => {
  const deck = createPresentation(), box = { x: 0, y: 0, width: 100, height: 50 };
  deck.addSlide({ elements: [{ type: 'shape', shape: 'rect', box, accessibility: { description: 'A', decorative: true } }, { type: 'shape', shape: 'rect', box, accessibility: { description: 'B' } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 1; round <= 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', descriptionEditSaveProfile: 'batch-v3' });
      try {
        const slide = session.document.slides[0];
        for (const [index, original] of [...slide.elements].entries()) {
          const next = structuredClone(original); next.accessibility.description = `${round} & ${index} 中文😀`; slide.replaceElement(original.id, next);
        }
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected);
  const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /descr="3 &amp; 0 中文😀"/); assert.match(xml, /descr="3 &amp; 1 中文😀"/); assert.match(xml, /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('mixed text and description preservation matches Node and browser', async t => {
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 150, height: 50 }, content: 'Original', accessibility: { description: 'Original alt', decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 1; round <= 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v4' });
      try {
        const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
        next.content[0].runs[0].text = `Text ${round} 中文`; next.accessibility.description = `Alt & ${round}`; slide.replaceElement(original.id, next);
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected);
  const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /descr="Alt &amp; 3"/); assert.match(xml, />Text 3 中文<\/a:t>/); assert.match(xml, /<adec:decorative[^>]+val="1"/);
  assert.deepEqual(requests, []);
});

test('mixed-v5 chart, text, descriptions and numbering match Node and browser over three rounds', async t => {
  const { chartFixture, chartXml, series, workbook, encodeChart, frame } = await import('../packages/pptx-parser/test/chart-fixture.mjs');
  const { shape, body } = await import('../packages/pptx-parser/test/content-fixture.mjs');
  const initial = encodeChart(chartFixture({ xml: chartXml({ family: 'bar', data: series({ refs: true, values: [0, -2.5] }), external: true }), book: workbook({ values: [0, -2.5] }), content: frame() + shape({ id: '20', txBox: true, text: body() }) }));
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 1; round <= 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2' });
      try {
        session.document.setFirstSlideNumber(round);
        const slide = session.document.slides[0];
        for (const original of [...slide.elements]) {
          const next = structuredClone(original);
          next.name = `Name ${round} ${next.type}`; next.accessibility = { ...next.accessibility, description: `Alt ${round} 中文` };
          if (next.type === 'chart') next.series[0].values[1] = round + 10;
          else next.content[0].runs[0].text = `Text ${round}`;
          slide.replaceElement(original.id, next);
        }
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => {
    const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`);
    return Array.from(await edit(parser, Uint8Array.from(initial)));
  }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected); assert.deepEqual(requests, []);
});

test('drawing name preservation matches Node and browser over three rounds', async t => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'shape', shape: 'rect', name: 'Original', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { description: 'Original alt', decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (const name of ['First & 中文', '', 'Third 😀']) {
      const session = await api.importPptx(bytes, { projection: 'content', drawingNameSaveProfile: 'existing-v1' });
      try { const slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original); next.name = name; slide.replaceElement(original.id, next); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial), browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`;
  const coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => { const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`); return Array.from(await edit(parser, Uint8Array.from(initial))); }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected); assert.deepEqual(requests, []);
  const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString(); assert.match(xml, /name="Third 😀"/); assert.match(xml, /descr="Original alt"/); assert.match(xml, /<adec:decorative[^>]+val="1"/);
});

test('decorative source-value saves match Node and browser over three rounds', async t => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: 'shape', shape: 'rect', name: 'Original', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { description: 'Original alt', decorative: true } }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (const [round, value] of [false, true, false].entries()) {
      const session = await api.importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', decorativeSaveProfile: 'existing-v1' });
      try {
        const doc = session.document, slide = doc.slides[0], original = slide.elements[0], target = { slideId: slide.id, elementId: original.id };
        const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
          { type: 'element-name', ...target, expected: original.name, value: `Name ${round}` },
          { type: 'element-accessibility', ...target, expected: { description: original.accessibility.description, decorative: original.accessibility.decorative }, value: { description: `Alt ${round}`, decorative: value } },
        ] }).inverse;
        const redo = doc.applyEditOperation(undo).inverse; doc.applyEditOperation(redo);
        const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes;
      } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial), browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`, coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => { const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`); return Array.from(await edit(parser, Uint8Array.from(initial))); }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected); assert.deepEqual(requests, []); const xml = readGeneratedZip(expected).get('ppt/slides/slide1.xml').toString(); assert.match(xml, /<adec:decorative[^>]+val="0"/); assert.match(xml, /name="Name 2"/); assert.match(xml, /descr="Alt 2"/);
});

test('slide visibility source saves match Node and browser over three rounds', async t => {
  const deck = createPresentation(); deck.addSlide({ id: 'one', hidden: false, elements: [{ type: 'text', name: 'Target', box: { x: 0, y: 0, width: 100, height: 50 }, content: 'Original', accessibility: { description: 'Alt' } }] }); deck.addSlide({ id: 'two', hidden: true });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 0; round < 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', slideVisibilitySaveProfile: 'attribute-v1' });
      try { const [first, second] = session.document.slides, original = first.elements[0], next = structuredClone(original); first.setHidden(round % 2 === 0); second.setHidden(round % 2 !== 0); next.accessibility.description = `Alt ${round}`; next.content[0].runs[0].text = `Text ${round}`; first.replaceElement(original.id, next); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial), browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`, coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => { const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`); return Array.from(await edit(parser, Uint8Array.from(initial))); }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected); assert.deepEqual(requests, []);
  const reopened = await importPptx(expected, { projection: 'content' }); try { assert.deepEqual(reopened.document.slides.map(slide => slide.hidden), [true, false]); assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, 'Text 2'); } finally { reopened.dispose(); }
});

test('core metadata source saves match Node and browser across three rounds', async t => {
  const deck = createPresentation({ metadata: { title: 'Original', author: 'Original author', revision: 7 } }); deck.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 50 }, content: 'Original text' }] });
  const initial = (await generatePptx(deck)).bytes;
  async function edit(api, initialBytes) {
    let bytes = initialBytes;
    for (let round = 0; round < 3; round++) {
      const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-strings-v1', editSaveProfile: 'text-runs-v3', slideVisibilitySaveProfile: 'attribute-v1' });
      try { const slide = session.document.slides[0], old = slide.elements[0], next = structuredClone(old); session.document.setMetadata({ ...session.document.metadata, title: `Title ${round} & 中文`, author: `Author ${round}` }); slide.setHidden(round % 2 === 0); next.content[0].runs[0].text = `Text ${round}`; slide.replaceElement(old.id, next); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes; } finally { session.dispose(); }
    }
    return bytes;
  }
  const expected = await edit({ importPptx }, initial), browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const url = bytes => `data:text/javascript;base64,${bytes.toString('base64')}`, coreUrl = url(await readFile('packages/core/dist/browser/index.js')), parserUrl = url(await readFile('packages/pptx-parser/dist/browser/index.js'));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: { '@pptkit/core': coreUrl } })}</script>`);
  const actual = await page.evaluate(async ({ parserUrl, initial, operation }) => { const parser = await import(parserUrl), edit = (0, eval)(`(${operation})`); return Array.from(await edit(parser, Uint8Array.from(initial))); }, { parserUrl, initial: Array.from(initial), operation: edit.toString() });
  assert.deepEqual(Uint8Array.from(actual), expected); assert.deepEqual(requests, []);
  const reopened = await importPptx(expected, { projection: 'content' }); try { assert.equal(reopened.document.metadata.title, 'Title 2 & 中文'); assert.equal(reopened.document.metadata.author, 'Author 2'); assert.equal(reopened.document.metadata.revision, 7); assert.equal(reopened.document.slides[0].hidden, true); } finally { reopened.dispose(); }
});
