import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { chromium } from '@playwright/test';
import { createPresentation } from '@pptkit/core';
import { resolveLayout } from '@pptkit/layout';
import { PdfRenderError, renderLayoutToPdf } from '../dist/index.js';

function layout() {
  const d = createPresentation({ size: { width: 400, height: 200 } });
  d.addSlide({ id: 'target', hidden: true });
  d.addSlide({ id: 'source', elements: [{ type: 'text', box: { x: 10, y: 10, width: 200, height: 50 }, content: [{ runs: [{ text: 'Navigate', action: { type: 'slide', slideId: 'target' } }] }] }] });
  return resolveLayout(d);
}
function fake(overrides = {}) {
  const state = { closed: 0, contexts: 0, routes: 0 };
  const page = { setContent: async html => { state.html = html; }, evaluate: async expression => expression.includes('querySelectorAll') ? 0 : true, pdf: async () => new TextEncoder().encode('%PDF-test'), ...overrides.page };
  const context = { route: async (_, handler) => { state.routes++; state.handler = handler; }, newPage: async () => page, close: async () => { state.closed++; if (overrides.closeError) throw new Error('close'); } };
  return { state, browser: { newContext: async () => { state.contexts++; return context; } } };
}

test('PDF conversion isolates context, preserves selected order, and returns detached bytes', async () => {
  const source = new TextEncoder().encode('%PDF-test');
  const { browser, state } = fake({ page: { pdf: async () => source } });
  const result = await renderLayoutToPdf(layout(), { browser, slideIndices: [1, 0] });
  assert.deepEqual(result.pages, [{ slideId: 'source', index: 1, hidden: false }, { slideId: 'target', index: 0, hidden: true }]);
  assert.equal(state.closed, 1); assert.equal(state.routes, 1);
  assert.match(state.html, /Content-Security-Policy/); assert.match(state.html, /href="#pptkit-link-slide-0"/);
  assert.equal(result.warnings.at(-1).code, 'pdf-page-size-rounding');
  source[0] = 0; assert.equal(result.bytes[0], 37);
  const aborts = []; await state.handler({ abort: async () => aborts.push(true) }); assert.deepEqual(aborts, [true]);
});

test('PDF failures are atomic and close contexts without masking the original error', async () => {
  for (const [overrides, expected] of [
    [{ page: { pdf: async () => new Uint8Array([1]) } }, 'pdf-print-failed'],
    [{ page: { pdf: async () => { throw new Error('failure'); } } }, 'pdf-print-failed'],
    [{ page: { evaluate: async () => false } }, 'pdf-resource-unavailable'],
    [{ closeError: true }, 'pdf-print-failed'],
    [{ closeError: true, page: { evaluate: async () => false } }, 'pdf-resource-unavailable'],
  ]) {
    const { browser, state } = fake(overrides);
    await assert.rejects(renderLayoutToPdf(layout(), { browser }), error => error instanceof PdfRenderError && error.code === expected);
    assert.equal(state.closed, 1);
  }
  const { browser, state } = fake();
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxOutputBytes: 1 }), error => error.code === 'pdf-limit-exceeded');
  assert.equal(state.closed, 1);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, slideIndices: [] }), error => error.code === 'pdf-empty-selection');
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxOutputBytes: -1 }), RangeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxSvgBytes: null }), RangeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, imageAnimationPolicy: 'invalid' }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, imageAnimationPolicy: 'require-static' }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxStaticInspectionBytes: 1 }), RangeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, imageAnimationPolicy: 'require-structurally-static', isStaticImage: () => true }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, jpegDecodePolicy: 'invalid' }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, jpegPrecisionPolicy: 'invalid' }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, jpegStructurePolicy: 'invalid' }), TypeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxStructureInspectionBytes: 1 }), RangeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, maxJpegInspectionPixels: 1 }), RangeError);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 64_000_001 }), RangeError);
});

test('cancellation during browser work closes the isolated context', async () => {
  let ready; const started = new Promise(resolve => { ready = resolve; });
  const { browser, state } = fake({ page: { setContent: () => { ready(); return new Promise(() => {}); } } });
  const controller = new AbortController(), pending = renderLayoutToPdf(layout(), { browser, signal: controller.signal });
  await started; controller.abort();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(state.closed, 1);
});

test('abort while a context is being created closes it when creation settles', async () => {
  let finish; const gate = new Promise(resolve => { finish = resolve; });
  const { browser, state } = fake();
  const delayed = { newContext: async () => { await gate; return browser.newContext(); } };
  const controller = new AbortController(), pending = renderLayoutToPdf(layout(), { browser: delayed, signal: controller.signal });
  await new Promise(resolve => setImmediate(resolve)); controller.abort(); finish();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(state.closed, 1);
});

test('a browser request is aborted and prevents publication', async () => {
  let handler, closed = false, aborted = false;
  const browser = { newContext: async () => ({
    route: async (_, callback) => { handler = callback; },
    newPage: async () => ({ setContent: async () => handler({ abort: async () => { aborted = true; } }), evaluate: async expression => expression.includes('querySelectorAll') ? 0 : true, pdf: async () => new TextEncoder().encode('%PDF-test') }),
    close: async () => { closed = true; },
  }) };
  await assert.rejects(renderLayoutToPdf(layout(), { browser }), error => error.code === 'pdf-resource-unavailable');
  assert.equal(aborted, true); assert.equal(closed, true);
});

test('static policy accepts selected pages with no image uses without attestation warning', async () => {
  const { browser } = fake({ page: { evaluate: async expression => expression.startsWith('Array.from(') ? [[], []] : expression.includes('querySelectorAll') ? 0 : true } });
  let calls = 0;
  const result = await renderLayoutToPdf(layout(), { browser, imageAnimationPolicy: 'require-static', isStaticImage: () => { calls++; return false; } });
  assert.equal(calls, 0);
  assert.ok(!result.warnings.some(w => w.code === 'pdf-image-static-attested'));
});

test('static policy fails closed for invalid inventory or callback exceptions', async () => {
  const invalid = fake({ page: { evaluate: async expression => expression.startsWith('Array.from(') ? null : true } });
  await assert.rejects(renderLayoutToPdf(layout(), { browser: invalid.browser, imageAnimationPolicy: 'require-static', isStaticImage: () => true }), error => error.code === 'pdf-resource-unavailable');
  assert.equal(invalid.state.closed, 1);
  const throwing = fake({ page: { evaluate: async expression => expression.startsWith('Array.from(') ? [['data:image/png;base64,AAAA'], []] : true } });
  await assert.rejects(renderLayoutToPdf(layout(), { browser: throwing.browser, imageAnimationPolicy: 'require-static', isStaticImage: () => { throw new Error('failure'); } }), error => error.code === 'pdf-image-not-static');
  assert.equal(throwing.state.closed, 1);
});

test('strict font policy rejects missing declarations before opening a browser', async () => {
  const { browser, state } = fake();
  await assert.rejects(renderLayoutToPdf(layout(), { browser, fontPolicy: 'require-embedded' }), error => error.code === 'pdf-font-unavailable');
  assert.equal(state.contexts, 0);
  await assert.rejects(renderLayoutToPdf(layout(), { browser, fontPolicy: 'invalid' }), TypeError);
  const blank = createPresentation(); blank.addSlide();
  assert.equal((await renderLayoutToPdf(resolveLayout(blank), { browser, fontPolicy: 'require-embedded' })).pages.length, 1);
});

test('strict font policy loads supplied faces and rejects invalid font programs', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const family = 'PPTKit "PDF" \\ Test Font', d = createPresentation({ size: { width: 400, height: 200 }, theme: { fonts: { body: family, heading: family } } });
  d.addSlide({ elements: [{ type: 'text', box: { x: 10, y: 10, width: 300, height: 60 }, content: [{ runs: [{ text: 'Hello PDF', style: { fontFamily: family } }] }] }] });
  const base64 = readFileSync(new URL('../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url)).toString('base64');
  const face = { family, weight: 400, italic: false, format: 'woff2', base64 };
  face.unrelated = () => {};
  const result = await renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts: [face] });
  assert.equal(Buffer.from(result.bytes.subarray(0, 5)).toString(), '%PDF-');
  assert.ok(result.warnings.some(w => w.code === 'pdf-font-selection-unverified'));
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts: [{ ...face, weight: 700 }] }), error => error.code === 'pdf-font-unavailable');
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts: [{ ...face, base64: Buffer.from('wOF2invalid-font-data').toString('base64') }] }), error => error.code === 'pdf-font-unavailable');
  assert.equal(browser.contexts().length, 0);
});

test('strict font declarations are captured before awaiting an asset resolver', async () => {
  const family = 'PDF Snapshot Font', d = createPresentation({ theme: { fonts: { body: family, heading: family } } });
  d.registerAsset({ id: 'image', kind: 'image', source: { type: 'url', value: 'data:image/png;base64,AAAA' } });
  d.addSlide({ elements: [
    { type: 'text', box: { x: 0, y: 0, width: 200, height: 40 }, content: [{ runs: [{ text: 'Text', style: { fontFamily: family } }] }] },
    { type: 'image', assetId: 'image', box: { x: 0, y: 50, width: 10, height: 10 } },
  ] });
  const base64 = readFileSync(new URL('../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url)).toString('base64');
  const face = { family, weight: 400, italic: false, format: 'woff2', base64 };
  let start, finish; const started = new Promise(resolve => { start = resolve; });
  const gate = new Promise(resolve => { finish = resolve; }), { browser, state } = fake();
  const pending = renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts: [face], resolveAsset: async () => { start(); await gate; return 'data:image/png;base64,AAAA'; } });
  await started; face.family = 'Mutated'; face.base64 = 'AAAA'; finish();
  assert.equal((await pending).pages.length, 1);
  assert.match(state.html, /PDF Snapshot Font/);
});

test('strict font load failure closes its context', async () => {
  const family = 'PDF Face', d = createPresentation({ theme: { fonts: { body: family, heading: family } } });
  d.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'Test' }] });
  const base64 = readFileSync(new URL('../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url)).toString('base64');
  const { browser, state } = fake({ page: { evaluate: async expression => expression.includes('const faces =') ? false : true } });
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts: [{ family, weight: 400, italic: false, format: 'woff2', base64 }] }), error => error.code === 'pdf-font-unavailable');
  assert.equal(state.closed, 1);
});

test('strict fonts follow selected slides rather than unselected dependencies', async () => {
  const d = createPresentation();
  for (const family of ['Selected Face', 'Unselected Face']) d.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 120, height: 40 }, content: [{ runs: [{ text: 'Text', style: { fontFamily: family } }] }] }] });
  const base64 = readFileSync(new URL('../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url)).toString('base64');
  const { browser, state } = fake(), embeddedFonts = [{ family: 'Selected Face', weight: 400, italic: false, format: 'woff2', base64 }];
  assert.equal((await renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts, slideIndices: [0] })).pages.length, 1);
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, fontPolicy: 'require-embedded', embeddedFonts }), error => error.code === 'pdf-font-unavailable');
  assert.equal(state.contexts, 1);
});

test('cancellation during strict font loading closes context and permits retry', async () => {
  const family = 'PDF Abort Face', d = createPresentation({ theme: { fonts: { body: family, heading: family } } });
  d.addSlide({ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 30 }, content: 'Test' }] });
  const base64 = readFileSync(new URL('../../svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url)).toString('base64');
  const embeddedFonts = [{ family, weight: 400, italic: false, format: 'woff2', base64 }];
  let notify; const started = new Promise(resolve => { notify = resolve; });
  const slow = fake({ page: { evaluate: async expression => { if (expression.includes('const faces =')) { notify(); return new Promise(() => {}); } return true; } } });
  const controller = new AbortController();
  const pending = renderLayoutToPdf(resolveLayout(d), { browser: slow.browser, fontPolicy: 'require-embedded', embeddedFonts, signal: controller.signal });
  await started; controller.abort();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(slow.state.closed, 1);
  const retry = fake();
  assert.equal((await renderLayoutToPdf(resolveLayout(d), { browser: retry.browser, fontPolicy: 'require-embedded', embeddedFonts })).pages.length, 1);
});

test('real Chromium prints pages and rejects corrupt and external images', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const originalContexts = browser.contexts().length;
  const result = await renderLayoutToPdf(layout(), { browser, slideIndices: [1, 0] });
  assert.equal(Buffer.from(result.bytes.subarray(0, 5)).toString(), '%PDF-');
  assert.deepEqual(result.pages.map(page => page.slideId), ['source', 'target']);
  assert.equal(browser.contexts().length, originalContexts);
  const png = readFileSync(new URL('../../../tests/fixtures/transparent-replacement.png', import.meta.url)).toString('base64');
  const picture = createPresentation(); picture.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/png;base64,${png}` } });
  picture.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
  assert.ok((await renderLayoutToPdf(resolveLayout(picture), { browser })).warnings.some(w => w.code === 'pdf-image-frame-unverified'));
  for (const value of ['data:image/png;base64,AAAA', 'https://example.invalid/x.png']) {
    const d = createPresentation(); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value } });
    d.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser }), error => error.code === 'pdf-resource-unavailable');
    assert.equal(browser.contexts().length, originalContexts);
  }
});

test('static image attestation checks every selected use and rejects animated input', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const png = `data:image/png;base64,${readFileSync(new URL('../../../tests/fixtures/transparent-replacement.png', import.meta.url)).toString('base64')}`;
  const apng = `data:image/png;base64,${readFileSync(new URL('../../../tests/fixtures/two-frame-apng.png', import.meta.url)).toString('base64')}`;
  const d = createPresentation();
  d.registerAsset({ id: 'still', kind: 'image', source: { type: 'url', value: png } });
  d.registerAsset({ id: 'animated', kind: 'image', source: { type: 'url', value: apng } });
  const image = assetId => ({ type: 'image', assetId, box: { x: 0, y: 0, width: 20, height: 20 } });
  d.addSlide({ id: 'first', elements: [image('still'), image('still')] });
  d.addSlide({ id: 'animated', elements: [image('animated')] });
  d.addSlide({ id: 'last', elements: [image('still')] });
  const calls = [], isStaticImage = request => { calls.push({ ...request, frozen: Object.isFrozen(request) }); return request.dataUrl === png; };
  const result = await renderLayoutToPdf(resolveLayout(d), { browser, slideIndices: [2, 0], imageAnimationPolicy: 'require-static', isStaticImage });
  assert.deepEqual(calls.map(({ slideId, slideIndex, mimeType, frozen }) => [slideId, slideIndex, mimeType, frozen]), [['last', 2, 'image/png', true], ['first', 0, 'image/png', true], ['first', 0, 'image/png', true]]);
  assert.ok(result.warnings.some(w => w.code === 'pdf-image-static-attested'));
  assert.ok(result.warnings.some(w => w.code === 'pdf-image-frame-unverified'));
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, imageAnimationPolicy: 'require-static', isStaticImage }), error => error.code === 'pdf-image-not-static');
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, slideIndices: [0], imageAnimationPolicy: 'require-static', isStaticImage: () => Promise.reject(new Error('late')) }), TypeError);
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, imageAnimationPolicy: 'allow', isStaticImage }), TypeError);
  const controller = new AbortController();
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, slideIndices: [0], imageAnimationPolicy: 'require-static', signal: controller.signal, isStaticImage: () => { controller.abort(); return true; } }), error => error.code === 'operation-aborted');
  assert.equal(browser.contexts().length, 0);
});

test('automatic static profile accepts named still structures and rejects animation and unknown formats', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const stillPng = readFileSync(new URL('../../../tests/fixtures/transparent-replacement.png', import.meta.url));
  const animatedPng = readFileSync(new URL('../../../tests/fixtures/two-frame-apng.png', import.meta.url));
  const jpeg = readFileSync(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const stillGif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
  const animatedGif = Buffer.from('R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7', 'base64');
  const make = (bytes, mime, count = 1) => {
    const d = createPresentation(); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:${mime};base64,${bytes.toString('base64')}` } });
    for (let i = 0; i < count; i++) d.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    return resolveLayout(d);
  };
  for (const [bytes, mime] of [[stillPng, 'image/png'], [stillGif, 'image/gif'], [jpeg, 'image/jpeg']]) {
    const result = await renderLayoutToPdf(make(bytes, mime, 2), { browser, imageAnimationPolicy: 'require-structurally-static', maxStaticInspectionBytes: bytes.length });
    assert.equal(result.pages.length, 2);
    assert.ok(result.warnings.some(w => w.code === 'pdf-image-static-structure-checked'));
    assert.ok(result.warnings.some(w => w.code === 'pdf-image-frame-unverified'));
    await assert.rejects(renderLayoutToPdf(make(bytes, mime), { browser, imageAnimationPolicy: 'require-structurally-static', maxStaticInspectionBytes: bytes.length - 1 }), error => error.code === 'pdf-limit-exceeded');
  }
  for (const [bytes, mime] of [[animatedPng, 'image/png'], [animatedPng, 'image/jpeg'], [animatedGif, 'image/gif']]) {
    await assert.rejects(renderLayoutToPdf(make(bytes, mime), { browser, imageAnimationPolicy: 'require-structurally-static' }), error => error.code === 'pdf-image-not-static');
  }
  const loopExtension = Buffer.concat([Buffer.from([0x21, 0xff, 11]), Buffer.from('NETSCAPE2.0'), Buffer.from([3, 1, 0, 0, 0])]);
  const loopingStill = Buffer.concat([stillGif.subarray(0, 19), loopExtension, stillGif.subarray(19)]);
  assert.equal((await renderLayoutToPdf(make(loopingStill, 'image/gif'), { browser })).pages.length, 1);
  await assert.rejects(renderLayoutToPdf(make(loopingStill, 'image/gif'), { browser, imageAnimationPolicy: 'require-structurally-static' }), error => error.code === 'pdf-image-static-unverified');
  await assert.rejects(renderLayoutToPdf(make(Buffer.from('RIFFxxxxWEBP'), 'image/webp'), { browser, imageAnimationPolicy: 'require-structurally-static' }), error => error.code === 'pdf-image-static-unverified');
  const page = await browser.newPage();
  const webpUrl = await page.evaluate(() => { const canvas = document.createElement('canvas'); canvas.width = 2; canvas.height = 2; const context = canvas.getContext('2d'); context.fillStyle = 'red'; context.fillRect(0, 0, 2, 2); return canvas.toDataURL('image/webp'); });
  await page.close();
  assert.match(webpUrl, /^data:image\/webp;base64,/);
  const webp = Buffer.from(webpUrl.split(',')[1], 'base64');
  assert.equal((await renderLayoutToPdf(make(webp, 'image/webp'), { browser })).pages.length, 1);
  await assert.rejects(renderLayoutToPdf(make(webp, 'image/webp'), { browser, imageAnimationPolicy: 'require-structurally-static' }), error => error.code === 'pdf-image-static-unverified');
  await assert.rejects(renderLayoutToPdf(make(stillPng.subarray(0, -1), 'image/png'), { browser, imageAnimationPolicy: 'require-structurally-static' }), error => error.code === 'pdf-image-static-unverified');
  const both = await renderLayoutToPdf(make(jpeg, 'image/jpeg'), { browser, imageAnimationPolicy: 'require-structurally-static', jpegStructurePolicy: 'require-valid-8-bit', maxStaticInspectionBytes: jpeg.length, maxStructureInspectionBytes: jpeg.length });
  assert.equal(both.pages.length, 1);
  await assert.rejects(renderLayoutToPdf(make(jpeg, 'image/jpeg'), { browser, imageAnimationPolicy: 'require-structurally-static', jpegStructurePolicy: 'require-valid-8-bit', maxStaticInspectionBytes: jpeg.length - 1, maxStructureInspectionBytes: jpeg.length }), error => error.code === 'pdf-limit-exceeded');
  await assert.rejects(renderLayoutToPdf(make(jpeg, 'image/jpeg'), { browser, imageAnimationPolicy: 'require-structurally-static', jpegStructurePolicy: 'require-valid-8-bit', maxStaticInspectionBytes: jpeg.length, maxStructureInspectionBytes: jpeg.length - 1 }), error => error.code === 'pdf-limit-exceeded');
  const high = readFileSync(new URL('../../../tests/fixtures/jpeg-12bit-sequential.jpg', import.meta.url));
  const staticHigh = await renderLayoutToPdf(make(high, 'image/jpeg'), { browser, imageAnimationPolicy: 'require-structurally-static' });
  assert.ok(staticHigh.warnings.some(w => w.code === 'pdf-image-frame-unverified'));
  await assert.rejects(renderLayoutToPdf(make(high, 'image/jpeg'), { browser, imageAnimationPolicy: 'require-structurally-static', jpegStructurePolicy: 'require-valid-8-bit' }), error => error.code === 'pdf-jpeg-unsupported-precision');
  assert.equal(browser.contexts().length, 0);
});

test('JPEG opacity inspection rejects transparent decodes and counts unique source pixels', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const make = (name, count = 1) => {
    const data = readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url)).toString('base64');
    const d = createPresentation({ size: { width: 6, height: 4 } });
    d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${data}` } });
    for (let i = 0; i < count; i++) d.addSlide({ id: `s${i}`, elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 6, height: 4 }, fit: 'stretch' }] });
    return resolveLayout(d);
  };
  const good = make('exif-orientation6.jpg', 2);
  assert.equal((await renderLayoutToPdf(good, { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 6 })).pages.length, 2);
  let calls = 0;
  const combined = await renderLayoutToPdf(good, { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 6, imageAnimationPolicy: 'require-static', isStaticImage: () => { calls++; return true; } });
  assert.equal(calls, 2); assert.ok(combined.warnings.some(w => w.code === 'pdf-image-static-attested'));
  await assert.rejects(renderLayoutToPdf(good, { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 5 }), error => error.code === 'pdf-limit-exceeded');
  const blank = createPresentation(); blank.addSlide();
  assert.equal((await renderLayoutToPdf(resolveLayout(blank), { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 0 })).pages.length, 1);
  const raw = readFileSync(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url)).toString('base64'), separate = createPresentation();
  for (const [id, mime] of [['a', 'jpeg'], ['b', 'jpg']]) separate.registerAsset({ id, kind: 'image', source: { type: 'url', value: `data:image/${mime};base64,${raw}` } });
  separate.addSlide({ elements: ['a', 'b'].map(assetId => ({ type: 'image', assetId, box: { x: 0, y: 0, width: 6, height: 4 } })) });
  assert.equal((await renderLayoutToPdf(resolveLayout(separate), { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 12 })).pages.length, 1);
  await assert.rejects(renderLayoutToPdf(resolveLayout(separate), { browser, jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 11 }), error => error.code === 'pdf-limit-exceeded');
  for (const name of ['jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg']) {
    const input = make(name);
    assert.equal((await renderLayoutToPdf(input, { browser })).pages.length, 1);
    await assert.rejects(renderLayoutToPdf(input, { browser, jpegDecodePolicy: 'require-opaque' }), error => error.code === 'pdf-resource-unavailable');
  }
  assert.equal(browser.contexts().length, 0);
});

test('JPEG inspection failures and cancellation close the context', async () => {
  const unreadable = fake({ page: { evaluate: async expression => expression.includes('const seen =') ? 'unavailable' : expression.includes('querySelectorAll') ? 0 : true } });
  await assert.rejects(renderLayoutToPdf(layout(), { browser: unreadable.browser, jpegDecodePolicy: 'require-opaque' }), error => error.code === 'pdf-resource-unavailable');
  assert.equal(unreadable.state.closed, 1);
  let notify; const started = new Promise(resolve => { notify = resolve; });
  const slow = fake({ page: { evaluate: async expression => { if (expression.includes('const seen =')) { notify(); return new Promise(() => {}); } return expression.includes('querySelectorAll') ? 0 : true; } } });
  const controller = new AbortController(), pending = renderLayoutToPdf(layout(), { browser: slow.browser, jpegDecodePolicy: 'require-opaque', signal: controller.signal });
  await started; controller.abort();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(slow.state.closed, 1);
});

test('real browser pixel-read failure is rejected and isolated context closes', async t => {
  const chromiumBrowser = await chromium.launch(); t.after(() => chromiumBrowser.close());
  const browser = { newContext: async () => {
    const context = await chromiumBrowser.newContext();
    return {
      route: (...args) => context.route(...args), close: () => context.close(),
      newPage: async () => {
        const page = await context.newPage();
        return {
          setContent: html => page.setContent(html), pdf: options => page.pdf(options),
          evaluate: async expression => {
            if (expression.includes('const seen =')) await page.evaluate(() => { CanvasRenderingContext2D.prototype.getImageData = () => { throw new Error('unreadable'); }; });
            return page.evaluate(expression);
          },
        };
      },
    };
  } };
  const raw = readFileSync(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url)).toString('base64'), d = createPresentation();
  d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${raw}` } });
  d.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
  await assert.rejects(renderLayoutToPdf(resolveLayout(d), { browser, jpegDecodePolicy: 'require-opaque' }), error => error.code === 'pdf-resource-unavailable');
  assert.equal(chromiumBrowser.contexts().length, 0);
  assert.equal((await renderLayoutToPdf(resolveLayout(d), { browser: chromiumBrowser, jpegDecodePolicy: 'require-opaque' })).pages.length, 1);
});

test('JPEG precision policy rejects high precision and unverifiable prefixes independently of MIME', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const data = name => readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url));
  const make = url => {
    const d = createPresentation(); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: url } });
    d.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    return resolveLayout(d);
  };
  const normal = data('exif-orientation6.jpg'), high = data('jpeg-12bit-sequential.jpg');
  const standard = `data:image/jpeg;base64,${normal.toString('base64')}`;
  const blank = createPresentation(); blank.addSlide();
  assert.equal((await renderLayoutToPdf(resolveLayout(blank), { browser, jpegPrecisionPolicy: 'require-8-bit' })).pages.length, 1);
  const png = `data:image/png;base64,${data('transparent-replacement.png').toString('base64')}`;
  assert.equal((await renderLayoutToPdf(make(png), { browser, jpegPrecisionPolicy: 'require-8-bit' })).pages.length, 1);
  assert.equal((await renderLayoutToPdf(make(standard), { browser, jpegPrecisionPolicy: 'require-8-bit', imageAnimationPolicy: 'require-static', isStaticImage: () => true })).pages.length, 1);
  assert.equal((await renderLayoutToPdf(make(standard), { browser, jpegPrecisionPolicy: 'require-8-bit', jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 6 })).pages.length, 1);
  const percentEncoded = `data:image/jpeg,${[...normal].map(byte => `%${byte.toString(16).padStart(2, '0')}`).join('')}`;
  assert.equal((await renderLayoutToPdf(make(percentEncoded), { browser, jpegPrecisionPolicy: 'require-8-bit' })).pages.length, 1);
  for (const mime of ['image/jpeg', 'image/png']) {
    const url = `data:${mime};base64,${high.toString('base64')}`;
    await assert.rejects(renderLayoutToPdf(make(url), { browser, jpegPrecisionPolicy: 'require-8-bit' }), error => error.code === 'pdf-jpeg-unsupported-precision');
  }
  const delayed = Buffer.concat([normal.subarray(0, 2), Buffer.from([0xff, 0xe1, 0xff, 0xff]), Buffer.alloc(65533), normal.subarray(2)]);
  await assert.rejects(renderLayoutToPdf(make(`data:image/jpeg;base64,${delayed.toString('base64')}`), { browser, jpegPrecisionPolicy: 'require-8-bit' }), error => error.code === 'pdf-jpeg-precision-unavailable');
  assert.equal(browser.contexts().length, 0);
});

test('cancellation while collecting JPEG precision inventory closes context', async () => {
  let notify; const started = new Promise(resolve => { notify = resolve; });
  const slow = fake({ page: { evaluate: async expression => { if (expression.startsWith('Array.from(')) { notify(); return new Promise(() => {}); } return true; } } });
  const controller = new AbortController(), pending = renderLayoutToPdf(layout(), { browser: slow.browser, jpegPrecisionPolicy: 'require-8-bit', signal: controller.signal });
  await started; controller.abort();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(slow.state.closed, 1);
});

test('cancellation while collecting static image inventory closes context', async () => {
  let notify; const started = new Promise(resolve => { notify = resolve; });
  const slow = fake({ page: { evaluate: async expression => { if (expression.startsWith('Array.from(')) { notify(); return new Promise(() => {}); } return true; } } });
  const controller = new AbortController(), pending = renderLayoutToPdf(layout(), { browser: slow.browser, imageAnimationPolicy: 'require-structurally-static', signal: controller.signal });
  await started; controller.abort();
  await assert.rejects(pending, error => error.code === 'operation-aborted');
  assert.equal(slow.state.closed, 1);
});

test('JPEG precision mode rejects an image inventory that changes before decode', async () => {
  const browserState = fake({ page: { evaluate: async expression => expression.startsWith('Array.from(') ? [['data:image/png;base64,AAAA'], []] : expression.includes('querySelectorAll') ? 0 : true } });
  await assert.rejects(renderLayoutToPdf(layout(), { browser: browserState.browser, jpegPrecisionPolicy: 'require-8-bit' }), error => error.code === 'pdf-resource-unavailable');
  assert.equal(browserState.state.closed, 1);
});

test('full JPEG structure policy rejects duplicate frames and enforces unique-image byte budgets', async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const normal = readFileSync(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const high = readFileSync(new URL('../../../tests/fixtures/jpeg-12bit-sequential.jpg', import.meta.url));
  const make = (bytes, mime = 'image/jpeg', count = 1) => {
    const d = createPresentation(); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:${mime};base64,${bytes.toString('base64')}` } });
    for (let i = 0; i < count; i++) d.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    return resolveLayout(d);
  };
  assert.equal((await renderLayoutToPdf(make(normal, 'image/jpeg', 2), { browser, jpegStructurePolicy: 'require-valid-8-bit', maxStructureInspectionBytes: normal.length })).pages.length, 2);
  await assert.rejects(renderLayoutToPdf(make(normal), { browser, jpegStructurePolicy: 'require-valid-8-bit', maxStructureInspectionBytes: normal.length - 1 }), error => error.code === 'pdf-limit-exceeded');
  const blank = createPresentation(); blank.addSlide();
  assert.equal((await renderLayoutToPdf(resolveLayout(blank), { browser, jpegStructurePolicy: 'require-valid-8-bit', maxStructureInspectionBytes: 0 })).pages.length, 1);
  const png = readFileSync(new URL('../../../tests/fixtures/transparent-replacement.png', import.meta.url)), mixed = createPresentation();
  for (const [id, mime, bytes] of [['jpeg', 'image/jpeg', normal], ['png', 'image/png', png]]) mixed.registerAsset({ id, kind: 'image', source: { type: 'url', value: `data:${mime};base64,${bytes.toString('base64')}` } });
  mixed.addSlide({ elements: ['jpeg', 'png'].map(assetId => ({ type: 'image', assetId, box: { x: 0, y: 0, width: 20, height: 20 } })) });
  assert.equal((await renderLayoutToPdf(resolveLayout(mixed), { browser, jpegStructurePolicy: 'require-valid-8-bit', maxStructureInspectionBytes: normal.length + png.length })).pages.length, 1);
  await assert.rejects(renderLayoutToPdf(resolveLayout(mixed), { browser, jpegStructurePolicy: 'require-valid-8-bit', maxStructureInspectionBytes: normal.length + png.length - 1 }), error => error.code === 'pdf-limit-exceeded');
  for (const mime of ['image/jpeg', 'image/png']) await assert.rejects(renderLayoutToPdf(make(high, mime), { browser, jpegStructurePolicy: 'require-valid-8-bit' }), error => error.code === 'pdf-jpeg-unsupported-precision');
  let at = 2, malformed;
  while (at < normal.length) {
    assert.equal(normal[at++], 0xff); while (normal[at] === 0xff) at++;
    const marker = normal[at++], length = normal.readUInt16BE(at), start = at - 2, end = at + length;
    if ([0xc0, 0xc1, 0xc2].includes(marker)) { malformed = Buffer.concat([normal.subarray(0, end), normal.subarray(start, end), normal.subarray(end)]); break; }
    at = end;
  }
  assert.ok(malformed);
  await assert.rejects(renderLayoutToPdf(make(malformed), { browser, jpegStructurePolicy: 'require-valid-8-bit' }), error => error.code === 'pdf-jpeg-structure-unavailable');
  const delayed = Buffer.concat([normal.subarray(0, 2), Buffer.from([0xff, 0xe1, 0xff, 0xff]), Buffer.alloc(65533), normal.subarray(2)]);
  assert.equal((await renderLayoutToPdf(make(delayed), { browser, jpegStructurePolicy: 'require-valid-8-bit' })).pages.length, 1);
  assert.equal(browser.contexts().length, 0);
});
