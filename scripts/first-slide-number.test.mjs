import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx, inspectPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('native first slide number is explicit, bounded, and independent of cached page numbers', async t => {
  const makeDeck = api => { const d = api.createPresentation(); d.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: [{ runs: [{ text: 'cache', field: { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' } }] }] }] }); return d; };
  const deck = makeDeck({ createPresentation });
  const defaultOutput = await generatePptx(deck), defaultXml = readGeneratedZip(defaultOutput.bytes).get('ppt/presentation.xml').toString();
  assert.doesNotMatch(defaultXml, /firstSlideNum=/);
  const defaultInspection = await inspectPptx(defaultOutput.bytes); assert.equal(defaultInspection.firstSlideNumber, 1); assert.equal(defaultInspection.firstSlideNumberDeclared, false);
  assert.deepEqual((await generatePptx(deck, { firstSlideNumber: undefined })).bytes, defaultOutput.bytes);
  const cacheOptions = { fieldRefresh: { firstPageNumber: 42, includeHidden: true } };
  const baselineParts = readGeneratedZip((await generatePptx(deck, { layout: cacheOptions })).bytes);
  const outputs = [];
  for (const value of [0, 1, 17, 9999]) {
    const generated = await generatePptx(deck, { firstSlideNumber: value, layout: { fieldRefresh: { firstPageNumber: 42, includeHidden: true } } });
    const parts = readGeneratedZip(generated.bytes);
    const inspected = await inspectPptx(generated.bytes); assert.equal(inspected.firstSlideNumber, value); assert.equal(inspected.firstSlideNumberDeclared, true);
    assert.match(parts.get('ppt/presentation.xml').toString(), new RegExp(` firstSlideNum="${value}"`));
    assert.match(parts.get('ppt/slides/slide1.xml').toString(), /<a:t>42<\/a:t><\/a:fld>/);
    assert.deepEqual([...parts.keys()], [...baselineParts.keys()]);
    for (const [name, bytes] of parts) if (name !== 'ppt/presentation.xml') assert.deepEqual(bytes, baselineParts.get(name));
    const session = await importPptx(generated.bytes, { projection: 'structure' });
    try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual((await session.save(plan.plan)).bytes, generated.bytes); } finally { session.dispose(); }
    outputs.push(generated.bytes);
  }
  for (const value of [-1, 10000, 0.5, NaN, Infinity, '1', null]) await assert.rejects(generatePptx(deck, { firstSlideNumber: value }), RangeError);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'layout', 'pptx-exporter', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const browserResults = await page.evaluate(async ({ source }) => { const create = (0, eval)(`(${source})`); return await Promise.all([0, 1, 17, 9999].map(async value => { const bytes = (await PPTKit.generatePptx(create(PPTKit), { firstSlideNumber: value, layout: { fieldRefresh: { firstPageNumber: 42, includeHidden: true } } })).bytes; const inspection = await PPTKit.inspectPptx(bytes), session = await PPTKit.importPptx(bytes, { projection: 'structure' }); try { const projected = session.document.firstSlideNumber, fresh = await PPTKit.inspectPptx((await PPTKit.generatePptx(session.document)).bytes); return { bytes: Array.from(bytes), value: inspection.firstSlideNumber, declared: inspection.firstSlideNumberDeclared, projected, fresh: fresh.firstSlideNumber, freshDeclared: fresh.firstSlideNumberDeclared }; } finally { session.dispose(); } })); }, { source: makeDeck.toString() });
  for (let i = 0; i < outputs.length; i++) { assert.deepEqual(new Uint8Array(browserResults[i].bytes), outputs[i]); assert.equal(browserResults[i].value, [0, 1, 17, 9999][i]); assert.equal(browserResults[i].declared, true); assert.equal(browserResults[i].projected, browserResults[i].value); assert.equal(browserResults[i].fresh, browserResults[i].value); assert.equal(browserResults[i].freshDeclared, true); }
  assert.deepEqual(requests, []);
});

test('ordinary text edits retain source firstSlideNum through three preservation saves and fresh export', async () => {
  for (const value of [undefined, 0, 1, 17, 9999]) {
    const deck = createPresentation(), field = { id: '38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a', kind: 'slide-number' };
    deck.addSlide({ elements: [{ id: 'text', type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: [{ runs: [{ text: 'before' }, { text: 'cache', field }] }] }] });
    let source = (await generatePptx(deck, { firstSlideNumber: value })).bytes;
    const presentation = readGeneratedZip(source).get('ppt/presentation.xml');
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(source, { projection: 'content', fieldTextProfile: 'slide-number-v1', editSaveProfile: 'text-runs-v3' });
      try {
        assert.equal(session.document.firstSlideNumber, value); assert.equal(normalizePresentation(session.document).firstSlideNumber, value);
        const slide = session.document.slides[0], element = slide.elements.find(item => item.type === 'text'), next = structuredClone(element);
        next.content[0].runs[0].text = `edited-${round}`; slide.replaceElement(element.id, next);
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const saved = (await session.save(plan.plan)).bytes, parts = readGeneratedZip(saved), inspected = await inspectPptx(saved);
        assert.deepEqual(parts.get('ppt/presentation.xml'), presentation);
        assert.equal(inspected.firstSlideNumber, value ?? 1); assert.equal(inspected.firstSlideNumberDeclared, value !== undefined);
        assert.match(parts.get('ppt/slides/slide1.xml').toString(), /<a:t>cache<\/a:t><\/a:fld>/);
        const fresh = await generatePptx(session.document, { firstSlideNumber: inspected.firstSlideNumberDeclared ? inspected.firstSlideNumber : undefined });
        const freshInspection = await inspectPptx(fresh.bytes); assert.equal(freshInspection.firstSlideNumber, value ?? 1); assert.equal(freshInspection.firstSlideNumberDeclared, value !== undefined);
        const automatic = await inspectPptx((await generatePptx(session.document)).bytes);
        assert.equal(automatic.firstSlideNumber, value ?? 1); assert.equal(automatic.firstSlideNumberDeclared, value !== undefined);
        source = saved;
      } finally { session.dispose(); }
    }
  }
});

test('author numbering remains independent of cache, with an explicit export override', async () => {
  const deck = createPresentation({ firstSlideNumber: 17 }); deck.addSlide();
  assert.equal(deck.firstSlideNumber, 17); assert.equal(normalizePresentation(deck).firstSlideNumber, 17);
  assert.equal((await inspectPptx((await generatePptx(deck)).bytes)).firstSlideNumber, 17);
  assert.equal((await inspectPptx((await generatePptx(deck, { firstSlideNumber: 3 })).bytes)).firstSlideNumber, 3);
  assert.equal(deck.firstSlideNumber, 17);
  deck.setFirstSlideNumber(0); assert.equal((await inspectPptx((await generatePptx(deck)).bytes)).firstSlideNumber, 0);
  assert.throws(() => deck.setFirstSlideNumber(-1), RangeError); assert.equal(deck.firstSlideNumber, 0);
  deck.setFirstSlideNumber(); assert.equal((await inspectPptx((await generatePptx(deck)).bytes)).firstSlideNumberDeclared, false);
  const large = createPresentation({ firstSlideNumber: 10000 }); large.addSlide();
  await assert.rejects(generatePptx(large), /0 to 9999/);
  assert.equal((await inspectPptx((await generatePptx(large, { firstSlideNumber: 1 })).bytes)).firstSlideNumber, 1);
  for (const value of [-1, 0.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, null, '1']) assert.throws(() => createPresentation({ firstSlideNumber: value }), RangeError);
});

test('changing projected start number blocks source-preserving saves until restored', async () => {
  const deck = createPresentation({ firstSlideNumber: 17 }); deck.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'before' }] });
  const source = (await generatePptx(deck)).bytes;
  const session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3' });
  try {
    session.document.setFirstSlideNumber(18);
    assert.equal((await session.prepareSave({ mode: 'no-op' })).status, 'blocked');
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    session.document.setFirstSlideNumber(17);
    const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual((await session.save(plan.plan)).bytes, source);
  } finally { session.dispose(); }
});

test('out-of-range source declarations remain explicit and never silently become a native setting', async () => {
  for (const strict of [false, true]) for (const value of [-1, 10000]) {
    const fixture = structureFixture(strict);
    fixture.parts['deck/main.data'] = fixture.parts['deck/main.data'].replace('<p:presentation ', `<p:presentation firstSlideNum="${value}" `);
    const bytes = encodeStructure(fixture), inspection = await inspectPptx(bytes);
    assert.equal(inspection.firstSlideNumber, value); assert.equal(inspection.firstSlideNumberDeclared, true);
    const session = await importPptx(bytes, { projection: 'structure' });
    try {
      assert.equal(session.document.firstSlideNumber, value < 0 ? undefined : value);
      assert.ok(session.report.findings.some(f => f.feature === 'first-slide-number' && f.operation === 'import' && f.tier === (value < 0 ? 'unsupported' : 'editable')));
      assert.ok(session.report.findings.some(f => f.feature === 'first-slide-number' && f.operation === 'export' && f.tier === 'unsupported'));
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes));
      if (value === 10000) await assert.rejects(generatePptx(session.document), /0 to 9999/);
    } finally { session.dispose(); }
  }
});
