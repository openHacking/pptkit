import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";

import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";
import { createFixture, fixtureCatalog, pixelPng } from "../../../tests/fixtures/presentations.mjs";

const fixtureFont = readFileSync(new URL("./fixtures/roboto-mono-latin-400-normal.woff2", import.meta.url)).toString("base64");

test('document slide links navigate between inline SVG pages', async ({ page }) => {
  const deck = createPresentation({ size: { width: 400, height: 200 } });
  deck.addSlide({ id: 'target' });
  deck.addSlide({ id: 'source', elements: [
    { type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: [{ runs: [{ text: 'Navigate', action: { type: 'slide', slideId: 'target' } }] }] },
    { type: 'shape', shape: 'rect', box: { x: 10, y: 80, width: 100, height: 50 }, style: { fill: { type: 'solid', color: 'FF0000' } }, action: { type: 'slide', slideId: 'target' } },
  ] });
  const result = await renderPresentationToSvg(deck, { slideIndices: [1, 0], slideActionMode: 'document-links', idNamespace: 'navigation' });
  await page.setContent(result.slides.map(slide => slide.svg).join(''));
  const target = await page.locator('svg[data-pptkit-slide-id="target"]').getAttribute('id');
  const anchors = page.locator('a[data-pptkit-slide-target="target"]');
  await expect(anchors).toHaveCount(2);
  for (let index = 0; index < 2; index++) {
    await anchors.nth(index).click();
    await expect.poll(() => page.evaluate(() => location.hash)).toBe(`#${target}`);
    await page.evaluate(() => { location.hash = ''; });
  }
});

test('browser raster adapter enforces dimensions, transparency, JPEG matte and failure budgets', async ({ page }) => {
  const requests = []; await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const output = await page.evaluate(async pixelPng => {
    const d = PPTKit.createPresentation({ size: { width: 40, height: 20 } });
    for (const id of ['first', 'second']) d.addSlide({ id, hidden: id === 'second', elements: [{ type: 'shape', id: `rect-${id}`, shape: 'rect', box: { x: 0, y: 0, width: 10, height: 10 }, style: { fill: { type: 'solid', color: 'FF0000' }, stroke: { paint: { type: 'none' } } } }] });
    const layout = PPTKit.resolveLayout(d);
    const inspect = async image => {
      const url = URL.createObjectURL(new Blob([image.bytes], { type: image.mimeType }));
      try {
        const decoded = new Image(); decoded.src = url; await decoded.decode();
        const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
        const ctx = canvas.getContext('2d'); ctx.drawImage(decoded, 0, 0);
        return { corner: [...ctx.getImageData(image.width - 2, image.height - 2, 1, 1).data], red: [...ctx.getImageData(5, 5, 1, 1).data] };
      } finally { URL.revokeObjectURL(url); }
    };
    const png = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', scale: 2, backgroundMode: 'transparent', slideIndices: [1, 0] });
    const jpeg = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'jpeg', scale: 2, backgroundMode: 'transparent', jpegBackground: '#00FF00', jpegQuality: 1, slideIndices: [0] });
    const crop = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', scale: 2, backgroundMode: 'transparent', crop: { x: 5, y: 0, width: 20, height: 20 }, slideIndices: [0] });
    const error = async options => { try { await PPTKit.renderLayoutToBrowserRaster(layout, options); return 'accepted'; } catch (e) { return e.code ?? e.name; } };
    const controller = new AbortController(); controller.abort();
    const remote = PPTKit.createPresentation(); remote.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a.png' } }); remote.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 10, height: 10 } }] });
    let resource; try { await PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(remote), { format: 'png' }); } catch (e) { resource = e.code; }
    const imageResult = async value => {
      const p = PPTKit.createPresentation({ size: { width: 20, height: 20 } }); p.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value } }); p.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
      try { const result = await PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(p), { format: 'png' }); return { count: result.slides.length, warnings: result.warnings.map(w => w.code) }; } catch (e) { return e.code; }
    };
    return { png: png.slides.map(x => ({ id: x.slideId, hidden: x.hidden, width: x.width, height: x.height, mime: x.mimeType, signature: [...x.bytes.slice(0, 8)] })), pngPixels: await inspect(png.slides[0]), crop: { width: crop.slides[0].width, height: crop.slides[0].height, pixels: await inspect(crop.slides[0]) }, jpeg: { mime: jpeg.slides[0].mimeType, signature: [...jpeg.slides[0].bytes.slice(0, 3)], pixels: await inspect(jpeg.slides[0]) }, errors: [await error({ format: 'png', maxTotalPixels: 1 }), await error({ format: 'png', maxOutputBytes: 1 }), await error({ format: 'png', signal: controller.signal }), await error({ format: 'png', jpegQuality: 0.5 }), await error({ format: 'png', crop: { x: 30, y: 0, width: 20, height: 20 } })], resource, embedded: await imageResult(pixelPng), corrupt: await imageResult('data:image/png;base64,AAAA'), nestedSvg: await imageResult('data:image/svg+xml,%3Csvg%20xmlns=%22http://www.w3.org/2000/svg%22/%3E'), empty: (await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', slideIndices: [], maxTotalPixels: 0 })).slides.length };
  }, pixelPng);
  expect(output.png.map(x => [x.id, x.hidden, x.width, x.height, x.mime])).toEqual([['second', true, 80, 40, 'image/png'], ['first', false, 80, 40, 'image/png']]);
  expect(output.png[0].signature).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect(output.pngPixels.corner).toEqual([0, 0, 0, 0]); expect(output.pngPixels.red).toEqual([255, 0, 0, 255]);
  expect([output.crop.width, output.crop.height]).toEqual([40, 40]); expect(output.crop.pixels.corner).toEqual([0, 0, 0, 0]); expect(output.crop.pixels.red).toEqual([255, 0, 0, 255]);
  expect(output.jpeg.signature).toEqual([255, 216, 255]); expect(output.jpeg.mime).toBe('image/jpeg');
  expect(output.jpeg.pixels.corner[1]).toBeGreaterThan(245); expect(output.jpeg.pixels.corner[0]).toBeLessThan(10);
  expect(output.jpeg.pixels.red[0]).toBeGreaterThan(240); expect(output.jpeg.pixels.red[1]).toBeLessThan(15);
  expect(output.errors).toEqual(['raster-limit-exceeded', 'raster-limit-exceeded', 'operation-aborted', 'TypeError', 'RangeError']);
  expect(output.resource).toBe('raster-resource-unavailable'); expect(output.embedded.count).toBe(1); expect(output.embedded.warnings).toContain('raster-image-frame-unverified');
  expect(output.corrupt).toBe('raster-resource-unavailable'); expect(output.nestedSvg).toBe('raster-resource-unavailable');
  expect(output.empty).toBe(0); expect(requests).toEqual([]);
});

test('browser raster cancellation during encoding releases URLs and permits retry', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async () => {
    const d = PPTKit.createPresentation({ size: { width: 40, height: 20 } }); d.addSlide(); const layout = PPTKit.resolveLayout(d);
    const controller = new AbortController(), original = HTMLCanvasElement.prototype.toBlob;
    const create = URL.createObjectURL, revoke = URL.revokeObjectURL, opened = [], closed = [];
    URL.createObjectURL = blob => { const value = create.call(URL, blob); opened.push(value); return value; };
    URL.revokeObjectURL = value => { closed.push(value); return revoke.call(URL, value); };
    HTMLCanvasElement.prototype.toBlob = function (callback, type, quality) { setTimeout(() => original.call(this, callback, type, quality), 30); controller.abort(); };
    let error;
    try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', signal: controller.signal }); } catch (e) { error = e.code; }
    finally { HTMLCanvasElement.prototype.toBlob = original; URL.createObjectURL = create; URL.revokeObjectURL = revoke; }
    await new Promise(resolve => setTimeout(resolve, 50));
    const retry = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png' });
    return { error, opened, closed, retry: retry.slides.length };
  });
  expect(result.error).toBe('operation-aborted'); expect(result.opened).toHaveLength(1);
  expect(result.closed).toEqual(result.opened); expect(result.retry).toBe(1);
});

test('browser raster rejects absent or wrong-format encodings and permits retry', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async () => {
    const d = PPTKit.createPresentation({ size: { width: 40, height: 20 } }); d.addSlide(); const layout = PPTKit.resolveLayout(d);
    const original = HTMLCanvasElement.prototype.toBlob; const errors = [];
    try {
      for (const blob of [null, new Blob(['not a jpeg'], { type: 'image/png' })]) {
        HTMLCanvasElement.prototype.toBlob = function (callback) { callback(blob); };
        try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'jpeg' }); errors.push('accepted'); } catch (e) { errors.push(e.code); }
      }
    } finally { HTMLCanvasElement.prototype.toBlob = original; }
    return { errors, retry: (await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'jpeg' })).slides.length };
  });
  expect(result.errors).toEqual(['raster-encoding-failed', 'raster-encoding-failed']); expect(result.retry).toBe(1);
});

test('strict raster font policy checks selected face declarations and browser loading', async ({ page }) => {
  const requests = []; await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async base64 => {
    const family = 'Raster Required Face';
    const d = PPTKit.createPresentation({ size: { width: 160, height: 50 } }); d.addSlide({ id: 'blank' });
    d.addSlide({ id: 'text', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 150, height: 40 }, content: [{ runs: [{ text: 'Sample', style: { fontFamily: family } }] }] }] });
    const layout = PPTKit.resolveLayout(d), face = { family, weight: 400, italic: false, format: 'woff2', base64 };
    const attempt = async options => {
      try { const r = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded', ...options }); return { count: r.slides.length, warnings: r.warnings.map(w => w.code) }; }
      catch (e) { return { code: e.code ?? e.name, slideIndex: e.slideIndex }; }
    };
    const blank = await attempt({ slideIndices: [0] });
    const missing = await attempt({ slideIndices: [1] });
    const wrong = await attempt({ slideIndices: [1], embeddedFonts: [{ ...face, weight: 700 }] });
    const invalid = await attempt({ slideIndices: [1], embeddedFonts: [{ ...face, base64: 'd09GMg==' }] });
    const valid = await attempt({ slideIndices: [1], embeddedFonts: [face] });
    const unused = await attempt({ slideIndices: [1], embeddedFonts: [face, { ...face, family: 'Unused Face', base64: 'd09GMg==' }] });
    const fallback = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', slideIndices: [1] });
    return { blank, missing, wrong, invalid, valid, unused, fallback: fallback.slides.length };
  }, fixtureFont);
  expect(result.blank.count).toBe(1);
  for (const item of [result.missing, result.wrong, result.invalid]) expect(item).toEqual({ code: 'raster-font-unavailable', slideIndex: 1 });
  expect(result.valid.count).toBe(1); expect(result.valid.warnings).toContain('svg-embedded-font-unverified');
  expect(result.unused.count).toBe(1);
  expect(result.fallback).toBe(1); expect(requests).toEqual([]);
});

test('strict raster font loading can be cancelled without late rejection or poisoned retry', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async base64 => {
    const family = 'Raster Abort Face', d = PPTKit.createPresentation({ size: { width: 100, height: 40 } });
    d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 90, height: 30 }, content: [{ runs: [{ text: 'Test', style: { fontFamily: family } }] }] }] });
    const layout = PPTKit.resolveLayout(d), face = { family, weight: 400, italic: false, format: 'woff2', base64 };
    const original = globalThis.FontFace, controller = new AbortController(), late = [];
    const listener = event => late.push(String(event.reason)); window.addEventListener('unhandledrejection', listener);
    let entered, rejectLater; const called = new Promise(resolve => { entered = resolve; });
    globalThis.FontFace = class { load() { entered(); return new Promise((_, reject) => { rejectLater = reject; }); } };
    let error;
    try {
      const pending = PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded', embeddedFonts: [face], signal: controller.signal });
      await called; controller.abort();
      try { await pending; } catch (e) { error = e.code; }
      rejectLater(new Error('late font failure'));
    } finally { globalThis.FontFace = original; }
    await new Promise(resolve => setTimeout(resolve, 20)); window.removeEventListener('unhandledrejection', listener);
    const retry = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded', embeddedFonts: [face] });
    return { error, late, retry: retry.slides.length };
  }, fixtureFont);
  expect(result).toEqual({ error: 'operation-aborted', late: [], retry: 1 });
});

test('strict raster captures embedded font inputs before awaiting asset resolution', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async ({ base64, png }) => {
    const family = 'Raster Captured Face', d = PPTKit.createPresentation({ size: { width: 120, height: 60 } });
    d.registerAsset({ id: 'asset', kind: 'image', source: { type: 'path', value: 'virtual' } });
    d.addSlide({ id: 's', elements: [
      { type: 'image', id: 'image', assetId: 'asset', box: { x: 0, y: 0, width: 10, height: 10 } },
      { type: 'text', id: 'text', box: { x: 0, y: 10, width: 100, height: 40 }, content: [{ runs: [{ text: 'Test', style: { fontFamily: family } }] }] },
    ] });
    const face = { family, weight: 400, italic: false, format: 'woff2', base64 }, faces = [face];
    let entered, release; const called = new Promise(resolve => { entered = resolve; });
    const pending = PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png', fontPolicy: 'require-embedded', embeddedFonts: faces, resolveAsset: () => { entered(); return new Promise(resolve => { release = resolve; }); } });
    await called; face.base64 = 'd09GMg=='; faces.length = 0; release(png);
    const output = await pending;
    return { count: output.slides.length, warnings: output.warnings.map(w => w.code) };
  }, { base64: fixtureFont, png: pixelPng });
  expect(result.count).toBe(1); expect(result.warnings).toContain('svg-embedded-font-unverified');
});

test('covered raster policy checks emitted text, bullets, tables and chart labels', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async base64 => {
    const family = 'Raster Coverage Face', face = { family, weight: 400, italic: false, format: 'woff2', base64 };
    const make = element => { const d = PPTKit.createPresentation({ size: { width: 240, height: 160 }, theme: { fonts: { body: family, heading: family } } }); d.addSlide({ id: 's', elements: [element] }); return PPTKit.resolveLayout(d); };
    const text = value => ({ type: 'text', id: 'text', box: { x: 0, y: 0, width: 150, height: 50 }, content: [{ runs: [{ text: value, style: { fontFamily: family, fontSize: 18, language: 'en-US' } }] }] });
    const bullet = { type: 'text', id: 'bullet', box: { x: 0, y: 0, width: 150, height: 50 }, content: [{ style: { bullet: { type: 'bullet' } }, runs: [{ text: 'Text', style: { fontFamily: family } }] }] };
    const table = { type: 'table', id: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ height: 50, cells: [{ content: [{ runs: [{ text: '中', style: { fontFamily: family } }] }] }] }] };
    const chart = { type: 'chart', id: 'chart', chartType: 'bar', categories: ['中'], series: [{ name: 'S', values: [1] }], box: { x: 0, y: 0, width: 220, height: 150 } };
    const bold = { type: 'text', id: 'bold', box: { x: 0, y: 0, width: 150, height: 50 }, content: [{ runs: [{ text: 'Bold', style: { fontFamily: family, bold: true } }] }] };
    const split = { type: 'text', id: 'split', box: { x: 0, y: 0, width: 150, height: 50 }, content: [{ runs: [{ text: 'A中', style: { fontFamily: family, eastAsianFontFamily: 'CJK Face' } }] }] };
    const run = async (element, coverage, extra = {}) => {
      const seen = [];
      try {
        const output = await PPTKit.renderLayoutToBrowserRaster(make(element), { format: 'png', fontPolicy: 'require-covered', embeddedFonts: [face], tableMode: 'native', ...extra, fontCoverage: request => { seen.push({ ...request, frozen: Object.isFrozen(request) }); return coverage(request); } });
        return { count: output.slides.length, seen };
      } catch (e) { return { code: e.code ?? e.name, slideIndex: e.slideIndex, elementId: e.elementId, seen }; }
    };
    const ascii = request => /^[\x00-\x7f]*$/.test(request.text);
    const good = await run(text('Hello'), ascii);
    const missing = await run(text('Hello中'), ascii);
    const bulletMissing = await run(bullet, ascii);
    const tableMissing = await run(table, ascii);
    const chartMissing = await run(chart, ascii);
    const boldMissing = await run(bold, request => !request.bold, { embeddedFonts: [face, { ...face, weight: 700 }] });
    const splitMissing = await run(split, ascii, { embeddedFonts: [face, { ...face, family: 'CJK Face' }] });
    const thrown = await run(text('Hello'), () => { throw new Error('coverage failure'); });
    const asyncAnswer = await run(text('Hello'), () => Promise.reject(new Error('late')));
    return { good, missing, bulletMissing, tableMissing, chartMissing, boldMissing, splitMissing, thrown, asyncAnswer };
  }, fixtureFont);
  expect(result.good.count).toBe(1);
  expect(result.good.seen).toEqual([{ family: 'Raster Coverage Face', bold: false, italic: false, text: 'Hello', fontSize: 18, slideId: 's', elementId: 'text', language: 'en-US', frozen: true }]);
  for (const name of ['missing', 'bulletMissing', 'tableMissing', 'chartMissing', 'boldMissing', 'splitMissing', 'thrown']) expect(result[name].code).toBe('raster-font-unavailable');
  expect(result.bulletMissing.seen.some(request => request.text.includes('•'))).toBe(true);
  expect(result.tableMissing.seen.some(request => request.text.includes('中'))).toBe(true);
  expect(result.chartMissing.seen.some(request => request.text.includes('中'))).toBe(true);
  expect(result.boldMissing.seen.some(request => request.bold)).toBe(true);
  expect(result.splitMissing.seen.some(request => request.family === 'CJK Face' && request.text.includes('中'))).toBe(true);
  expect(result.tableMissing.elementId).toBe('table'); expect(result.chartMissing.elementId).toBe('chart');
  expect(result.asyncAnswer.code).toBe('TypeError');
});

test('coverage callback abort absorbs a late promise rejection and leaves rendering retryable', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async base64 => {
    const family = 'Coverage Abort Face', d = PPTKit.createPresentation({ size: { width: 100, height: 40 } });
    d.addSlide({ id: 's', elements: [{ type: 'text', id: 't', box: { x: 0, y: 0, width: 90, height: 30 }, content: [{ runs: [{ text: 'Test', style: { fontFamily: family } }] }] }] });
    const layout = PPTKit.resolveLayout(d), face = { family, weight: 400, italic: false, format: 'woff2', base64 }, controller = new AbortController(), late = [];
    const listener = event => late.push(String(event.reason)); window.addEventListener('unhandledrejection', listener);
    let code;
    try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-covered', embeddedFonts: [face], signal: controller.signal, fontCoverage: () => { controller.abort(); return Promise.reject(new Error('late')); } }); }
    catch (error) { code = error.code; }
    await new Promise(resolve => setTimeout(resolve, 20)); window.removeEventListener('unhandledrejection', listener);
    const retry = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-covered', embeddedFonts: [face], fontCoverage: () => true });
    return { code, late, retry: retry.slides.length };
  }, fixtureFont);
  expect(result).toEqual({ code: 'operation-aborted', late: [], retry: 1 });
});

test('coverage requests follow inherited, grouped, hidden and selected artwork', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async base64 => {
    const family = 'Coverage Layer Face', run = text => [{ runs: [{ text, style: { fontFamily: family } }] }];
    const d = PPTKit.createPresentation({ size: { width: 200, height: 120 }, masters: [{ id: 'master', name: 'Master', elements: [{ type: 'text', id: 'master-text', box: { x: 0, y: 0, width: 100, height: 20 }, content: run('Master') }] }] });
    d.defineSlideLayout({ id: 'layout', name: 'Layout', masterId: 'master' });
    d.addSlide({ id: 'selected', layoutId: 'layout', elements: [
      { type: 'group', id: 'group', box: { x: 0, y: 30, width: 100, height: 40 }, coordinateSize: { width: 100, height: 40 }, children: [{ type: 'text', id: 'child', box: { x: 0, y: 0, width: 100, height: 20 }, content: run('Child') }] },
      { type: 'text', id: 'hidden', hidden: true, box: { x: 0, y: 90, width: 100, height: 20 }, content: run('Hidden') },
    ] });
    d.addSlide({ id: 'other', layoutId: 'layout', elements: [{ type: 'text', id: 'other-text', box: { x: 0, y: 0, width: 100, height: 20 }, content: run('Other') }] });
    const requests = [], output = await PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png', slideIndices: [0], fontPolicy: 'require-covered', embeddedFonts: [{ family, weight: 400, italic: false, format: 'woff2', base64 }], fontCoverage: request => { requests.push([request.slideId, request.elementId, request.text]); return true; } });
    return { slides: output.slides.map(slide => slide.slideId), requests };
  }, fixtureFont);
  expect(result.slides).toEqual(['selected']);
  expect(result.requests).toEqual([['selected', 'master-text', 'Master'], ['selected', 'child', 'Child']]);
});

test('browser raster preserves native table, chart and semi-transparent overlay pixels', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async () => {
    const d = PPTKit.createPresentation({ size: { width: 240, height: 120 } });
    d.addSlide({ elements: [
      { type: 'table', id: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [50, 50], rows: [{ height: 50, cells: [
        { content: 'A', style: { fill: { type: 'solid', color: 'FFFF00' } } }, { content: 'B', style: { fill: { type: 'solid', color: '00FFFF' } } },
      ] }] },
      { type: 'chart', id: 'chart', chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [3], color: '0000FF' }], legend: { visible: false }, box: { x: 120, y: 0, width: 110, height: 100 } },
      { type: 'shape', id: 'red', shape: 'rect', box: { x: 0, y: 70, width: 40, height: 40 }, style: { fill: { type: 'solid', color: 'FF0000' }, stroke: { paint: { type: 'none' } } } },
      { type: 'shape', id: 'blue', shape: 'rect', box: { x: 20, y: 70, width: 40, height: 40 }, opacity: 0.5, style: { fill: { type: 'solid', color: '0000FF' }, stroke: { paint: { type: 'none' } } } },
    ] });
    const raster = await PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png', tableMode: 'native' });
    const item = raster.slides[0], url = URL.createObjectURL(new Blob([item.bytes], { type: item.mimeType }));
    try {
      const img = new Image(); img.src = url; await img.decode();
      const canvas = document.createElement('canvas'); canvas.width = item.width; canvas.height = item.height;
      const c = canvas.getContext('2d'); c.drawImage(img, 0, 0);
      const sample = (x, y) => [...c.getImageData(x, y, 1, 1).data];
      const chart = c.getImageData(120, 0, 110, 100).data; let blue = 0;
      for (let i = 0; i < chart.length; i += 4) if (chart[i] < 100 && chart[i + 1] < 100 && chart[i + 2] > 180) blue++;
      return { yellow: sample(5, 45), cyan: sample(95, 45), overlay: sample(30, 90), blue, warnings: raster.warnings.map(w => w.code) };
    } finally { URL.revokeObjectURL(url); }
  });
  expect(result.yellow).toEqual([255, 255, 0, 255]); expect(result.cyan).toEqual([0, 255, 255, 255]);
  expect(Math.abs(result.overlay[0] - 127)).toBeLessThanOrEqual(1); expect(result.overlay.slice(1)).toEqual([0, 128, 255]); expect(result.blue).toBeGreaterThan(50);
  expect(result.warnings).toContain('svg-native-table-approximation');
});

test('a later undecodable image rejects the entire multi-slide raster operation', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async () => {
    const d = PPTKit.createPresentation({ size: { width: 20, height: 20 } }); d.registerAsset({ id: 'bad', kind: 'image', source: { type: 'url', value: 'data:image/png;base64,AAAA' } });
    d.addSlide({ id: 'first' }); d.addSlide({ id: 'second', elements: [{ type: 'image', assetId: 'bad', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    let encoded = 0; const original = HTMLCanvasElement.prototype.toBlob;
    HTMLCanvasElement.prototype.toBlob = function (...args) { encoded++; return original.apply(this, args); };
    try { await PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), { format: 'png' }); return { error: 'accepted', encoded }; }
    catch (e) { return { error: e.code, slideIndex: e.slideIndex, encoded }; }
    finally { HTMLCanvasElement.prototype.toBlob = original; }
  });
  expect(result).toEqual({ error: 'raster-resource-unavailable', slideIndex: 1, encoded: 1 });
});

test('static image attestation rejects an APNG and checks every selected use', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const animated = `data:image/png;base64,${readFileSync(new URL('../../../tests/fixtures/two-frame-apng.png', import.meta.url)).toString('base64')}`;
  const result = await page.evaluate(async ({ png, animated }) => {
    const d = PPTKit.createPresentation({ size: { width: 20, height: 20 } });
    for (const [id, value] of [['static', png], ['animated', animated]]) d.registerAsset({ id, kind: 'image', source: { type: 'url', value } });
    for (const [id, assetId] of [['one', 'static'], ['two', 'animated'], ['three', 'static']]) {
      const elements = [{ type: 'image', assetId, box: { x: 0, y: 0, width: 20, height: 20 } }];
      if (id === 'one') elements.push({ type: 'image', assetId, box: { x: 0, y: 0, width: 10, height: 10 } });
      d.addSlide({ id, elements });
    }
    const layout = PPTKit.resolveLayout(d), seen = [];
    const isStaticImage = request => { seen.push({ slideId: request.slideId, slideIndex: request.slideIndex, mimeType: request.mimeType, frozen: Object.isFrozen(request) }); return request.dataUrl === png; };
    const accepted = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage, slideIndices: [0, 2] });
    const acceptedSeen = [...seen]; seen.length = 0;
    let rejected; try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage }); } catch (e) { rejected = { code: e.code, slideIndex: e.slideIndex }; }
    let invalid; try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage: () => Promise.reject(new Error('late')) }); } catch (e) { invalid = e.name; }
    let thrown; try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage: () => { throw new Error('failed'); } }); } catch (e) { thrown = e.code; }
    const fallback = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', slideIndices: [0] });
    return { count: accepted.slides.length, acceptedSeen, warnings: accepted.warnings.map(w => w.code), rejected, rejectedSeen: seen, invalid, thrown, fallbackWarnings: fallback.warnings.map(w => w.code) };
  }, { png: pixelPng, animated });
  expect(result.count).toBe(2);
  expect(result.acceptedSeen).toEqual([{ slideId: 'one', slideIndex: 0, mimeType: 'image/png', frozen: true }, { slideId: 'one', slideIndex: 0, mimeType: 'image/png', frozen: true }, { slideId: 'three', slideIndex: 2, mimeType: 'image/png', frozen: true }]);
  expect(result.warnings.filter(code => code === 'raster-image-static-attested')).toHaveLength(2);
  expect(result.rejected).toEqual({ code: 'raster-image-not-static', slideIndex: 1 });
  expect(result.rejectedSeen.map(item => item.slideId)).toEqual(['one', 'one', 'two']);
  expect(result.invalid).toBe('TypeError'); expect(result.thrown).toBe('raster-image-not-static'); expect(result.fallbackWarnings).toContain('raster-image-frame-unverified');
});

test('static image attestation abort consumes late rejection and permits retry', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async png => {
    const d = PPTKit.createPresentation({ size: { width: 20, height: 20 } }); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: png } });
    d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    const layout = PPTKit.resolveLayout(d), controller = new AbortController(), late = [];
    const listener = event => late.push(String(event.reason)); window.addEventListener('unhandledrejection', listener);
    let code;
    try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', signal: controller.signal, isStaticImage: () => { controller.abort(); return Promise.reject(new Error('late')); } }); }
    catch (e) { code = e.code; }
    await new Promise(resolve => setTimeout(resolve, 20)); window.removeEventListener('unhandledrejection', listener);
    const retry = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage: request => request.dataUrl === png });
    return { code, late, retry: retry.slides.length };
  }, pixelPng);
  expect(result).toEqual({ code: 'operation-aborted', late: [], retry: 1 });
});

test('static image policy and callback are captured before asynchronous asset resolution', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const result = await page.evaluate(async png => {
    const d = PPTKit.createPresentation({ size: { width: 20, height: 20 } }); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'path', value: 'virtual' } });
    d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 20, height: 20 } }] });
    let entered, release; const called = new Promise(resolve => { entered = resolve; });
    const options = { format: 'png', imageAnimationPolicy: 'require-static', isStaticImage: request => request.dataUrl === png, resolveAsset: () => { entered(); return new Promise(resolve => { release = resolve; }); } };
    const pending = PPTKit.renderLayoutToBrowserRaster(PPTKit.resolveLayout(d), options);
    await called; options.imageAnimationPolicy = 'allow'; options.isStaticImage = () => false; release(png);
    const result = await pending; return { count: result.slides.length, warnings: result.warnings.map(w => w.code) };
  }, pixelPng);
  expect(result.count).toBe(1); expect(result.warnings).toContain('raster-image-static-attested');
});

test('JPEG opacity inspection catches transparent decodes and enforces unique-source budget', async ({ page }, testInfo) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const input = Object.fromEntries(['exif-orientation6.jpg', 'jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg'].map(name => [name, readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url)).toString('base64')]));
  const result = await page.evaluate(async images => {
    const make = (base64, count = 1) => {
      const d = PPTKit.createPresentation({ size: { width: 6, height: 4 } }); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${base64}` } });
      for (let i = 0; i < count; i++) d.addSlide({ id: `s${i}`, elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 6, height: 4 }, fit: 'stretch' }] });
      return PPTKit.resolveLayout(d);
    };
    const run = async (layout, options) => {
      try {
        const result = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', backgroundMode: 'transparent', ...options });
        const blob = URL.createObjectURL(new Blob([result.slides[0].bytes], { type: 'image/png' }));
        try {
          const image = new Image(); image.src = blob; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 6; canvas.height = 4; const c = canvas.getContext('2d'); c.drawImage(image, 0, 0);
          const rgba = c.getImageData(0, 0, 6, 4).data; let transparent = 0; for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) transparent++;
          return { count: result.slides.length, transparent };
        } finally { URL.revokeObjectURL(blob); }
      } catch (e) { return { code: e.code, slideIndex: e.slideIndex }; }
    };
    const valid = make(images['exif-orientation6.jpg'], 2);
    const validResult = await run(valid, { jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 6 });
    const budget = await run(valid, { jpegDecodePolicy: 'require-opaque', maxJpegInspectionPixels: 5 });
    const sixteen = [];
    for (const name of ['jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg']) {
      const layout = make(images[name]);
      sixteen.push({ name, plain: await run(layout, {}), checked: await run(layout, { jpegDecodePolicy: 'require-opaque' }) });
    }
    return { validResult, budget, sixteen };
  }, input);
  expect(result.validResult).toEqual({ count: 2, transparent: 0 });
  expect(result.budget).toEqual({ code: 'raster-limit-exceeded', slideIndex: 0 });
  for (const sample of result.sixteen) {
    if (sample.checked.code) expect(sample.checked.code).toBe('raster-resource-unavailable');
    else expect(sample.checked).toEqual({ count: 1, transparent: 0 });
    if (testInfo.project.name === 'chromium') {
      expect(sample.plain).toEqual({ count: 1, transparent: 24 });
      expect(sample.checked.code).toBe('raster-resource-unavailable');
    }
  }
});

test('JPEG opacity inspection handles read failure and abort without poisoning retry', async ({ page }) => {
  for (const name of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: readFileSync(new URL(`../../${name}/dist/browser/global.js`, import.meta.url), 'utf8') });
  const input = readFileSync(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url)).toString('base64');
  const result = await page.evaluate(async base64 => {
    const d = PPTKit.createPresentation({ size: { width: 6, height: 4 } }); d.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${base64}` } });
    d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 6, height: 4 } }] });
    const layout = PPTKit.resolveLayout(d), original = CanvasRenderingContext2D.prototype.getImageData, controller = new AbortController();
    let failed, aborted;
    try {
      CanvasRenderingContext2D.prototype.getImageData = function () { throw new Error('pixel read failed'); };
      try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', jpegDecodePolicy: 'require-opaque' }); } catch (e) { failed = e.code; }
      CanvasRenderingContext2D.prototype.getImageData = function (...args) { const pixels = original.apply(this, args); controller.abort(); return pixels; };
      try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', jpegDecodePolicy: 'require-opaque', signal: controller.signal }); } catch (e) { aborted = e.code; }
    } finally { CanvasRenderingContext2D.prototype.getImageData = original; }
    const retry = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', jpegDecodePolicy: 'require-opaque' });
    return { failed, aborted, retry: retry.slides.length };
  }, input);
  expect(result).toEqual({ failed: 'raster-resource-unavailable', aborted: 'operation-aborted', retry: 1 });
});

test('embedded font affects isolated SVG image pixels without external requests', async ({ page }) => {
  const alias = 'PPTKit "Isolated" Font 91237';
  const d = createPresentation({ size: { width: 340, height: 80 }, theme: { fonts: { body: alias, heading: alias } } });
  d.addSlide({ id: 's', elements: [{ type: 'text', id: 'text', box: { x: 4, y: 4, width: 330, height: 60 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: 'iiii WWWW 12345', style: { fontFamily: alias, fontSize: 36 } }] }] }] });
  const plain = await renderPresentationToSvg(d);
  const embedded = await renderPresentationToSvg(d, { embeddedFonts: [{ family: alias, weight: 400, italic: false, format: 'woff2', base64: fixtureFont }] });
  const requests = []; await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
  const source = svg => `data:image/svg+xml;base64,${Buffer.from(svg).toString('base64')}`;
  const pixels = await page.evaluate(async ({ before, after }) => {
    async function draw(src) {
      const image = new Image(); image.src = src; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = 340; canvas.height = 80;
      const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
      return context.getImageData(0, 0, 340, 80).data;
    }
    const a = await draw(before), b = await draw(after); let changed = 0, dark = 0;
    for (let i = 0; i < a.length; i += 4) {
      if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2]) changed++;
      if (b[i] < 120 && b[i + 1] < 120 && b[i + 2] < 120 && b[i + 3] > 0) dark++;
    }
    return { changed, dark };
  }, { before: source(plain.slides[0].svg), after: source(embedded.slides[0].svg) });
  expect(pixels.changed).toBeGreaterThan(100);
  expect(pixels.dark).toBeGreaterThan(100);
  expect(requests).toEqual([]);
  await page.setContent(`<style>html,body{margin:0}</style>${embedded.slides[0].svg}<img id="isolated" src="${source(embedded.slides[0].svg)}">`);
  await page.evaluate(async () => { await document.fonts.ready; await document.querySelector('#isolated').decode(); });
  const session = await page.context().newCDPSession(page);
  await session.send('DOM.enable'); await session.send('CSS.enable');
  const root = await session.send('DOM.getDocument');
  const selected = await session.send('DOM.querySelector', { nodeId: root.root.nodeId, selector: 'svg tspan' });
  const loaded = (await session.send('CSS.getPlatformFontsForNode', { nodeId: selected.nodeId })).fonts;
  expect(loaded.some(face => face.isCustomFont && face.glyphCount > 0)).toBe(true);
  const inline = await page.locator('svg').screenshot(), isolated = await page.locator('#isolated').screenshot();
  expect(inline.equals(isolated)).toBe(true);
});

test('native table accessibility names preserve words split by visual wrapping', async ({ page }) => {
  const d = createPresentation();
  d.addSlide({ id: 's', elements: [
    { type: 'table', id: 'table', name: 'Report', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [50, 50], rows: [{ height: 40, cells: [
      { headerScope: 'column', content: 'Heading' }, { content: [{ runs: [{ text: 'Va' }, { text: 'lue' }] }] },
    ] }] },
    { type: 'table', id: 'decoration', accessibility: { decorative: true }, box: { x: 110, y: 0, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: 'Decorative' }] }] },
  ] });
  const rendered = await renderPresentationToSvg(d, { tableMode: 'native' });
  await page.setContent(rendered.slides[0].svg);
  await expect(page.getByRole('table', { name: 'Report' })).toHaveCount(1);
  await expect(page.getByRole('columnheader', { name: 'Heading', exact: true })).toHaveCount(1);
  await expect(page.getByRole('cell', { name: 'Value', exact: true })).toHaveCount(1);
  await expect(page.getByRole('cell', { name: 'Decorative', exact: true })).toHaveCount(0);
  await expect(page.getByRole('table', { name: 'Report' }).locator('[data-pptkit-cell="0:0"] text')).toHaveCount(2);
});

test('native quarter-turn table cells expose names and stay inside their clips', async ({ page }) => {
  const d = createPresentation();
  d.addSlide({ id: 's', elements: [{ type: 'table', id: 't', box: { x: 20, y: 20, width: 200, height: 160 }, columns: [100, 100], rows: [{ height: 160, cells: [
    { content: 'Clockwise', style: { textDirection: 'rotate90', margin: 8 } },
    { content: 'Counter', style: { textDirection: 'rotate270', margin: 8 } },
  ] }] }] });
  const result = await renderPresentationToSvg(d, { tableMode: 'native' });
  await page.setContent(result.slides[0].svg);
  for (const name of ['Clockwise', 'Counter']) {
    const cell = page.getByRole('cell', { name, exact: true });
    await expect(cell).toHaveCount(1);
    const geometry = await cell.evaluate(node => {
      const rect = node.querySelector('rect').getBoundingClientRect(), text = node.querySelector('text').getBoundingClientRect();
      return { cell: { x: rect.x, y: rect.y, right: rect.right, bottom: rect.bottom }, text: { x: text.x, y: text.y, right: text.right, bottom: text.bottom } };
    });
    expect(geometry.text.x).toBeGreaterThanOrEqual(geometry.cell.x);
    expect(geometry.text.y).toBeGreaterThanOrEqual(geometry.cell.y);
    expect(geometry.text.right).toBeLessThanOrEqual(geometry.cell.right);
    expect(geometry.text.bottom).toBeLessThanOrEqual(geometry.cell.bottom);
  }
});

for (const fixture of fixtureCatalog.filter(({ snapshot }) => snapshot !== undefined)) {
  test(`corpus: ${fixture.id}`, async ({ page }) => {
    const result = await renderPresentationToSvg(createFixture(fixture.id, createPresentation));
    expect(result.warnings.map(({ code }) => code).sort()).toEqual([...fixture.warnings.preview].sort());
    await page.setContent(`<style>@font-face{font-family:'PPTKit Fixture';src:url(data:font/woff2;base64,${fixtureFont}) format('woff2');font-style:normal;font-weight:400;font-display:block}html,body{margin:0}#frame{width:960px;height:540px;background:white}svg{display:block;width:960px;height:540px}</style><div id="frame">${result.slides[0].svg}</div>`);
    await page.evaluate(() => document.fonts.ready);
    await page.evaluate(async () => {
      const sources = [...document.querySelectorAll("svg image")].map((node) => node.getAttribute("href") ?? node.getAttributeNS("http://www.w3.org/1999/xlink", "href"));
      await Promise.all(sources.filter(Boolean).map(async (source) => { const image = new Image(); image.src = source; await image.decode(); }));
    });
    await expect(page.locator("#frame")).toHaveScreenshot(fixture.snapshot, { maxDiffPixelRatio: 0.01, threshold: 0.2 });
  });
}

function visualFixture() {
  const presentation = createPresentation({
    metadata: { title: "SVG visual regression" },
    theme: { colors: { accent1: "2457D6", accent2: "D94A66" }, fonts: { heading: "PPTKit Fixture", body: "PPTKit Fixture" } },
  });
  const asset = presentation.registerAsset({
    id: "visual-image",
    kind: "image",
    source: { type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='400' height='200'%3E%3Crect width='400' height='200' fill='%232457D6'/%3E%3Ccircle cx='200' cy='100' r='70' fill='%23F8FAFC'/%3E%3C/svg%3E" },
    mimeType: "image/svg+xml",
    width: 400,
    height: 200,
  });
  const slide = presentation.addSlide({ id: "visual", background: { type: "solid", color: "F8FAFC" } });
  slide.addElement({
    type: "text", id: "visual-text", content: [{ style: { bullet: { type: "bullet" }, lineSpacing: 1.2 }, runs: [
      { text: "Rich text ", style: { fontFamily: "PPTKit Fixture", fontSize: 24, bold: true, color: { theme: "accent1" } } },
      { text: "preview", style: { fontFamily: "PPTKit Fixture", fontSize: 24, italic: true, underline: true } },
    ] }], box: { x: 40, y: 35, width: 400, height: 90 }, frame: { verticalAlign: "middle" },
  });
  slide.addElement({ type: "shape", id: "visual-shape", shape: "roundRect", box: { x: 40, y: 145, width: 220, height: 120 }, transform: { rotation: -4 }, style: { fill: { type: "solid", color: "DBEAFE" }, stroke: { paint: { type: "solid", color: { theme: "accent1" } }, width: 2 } }, text: { content: "Editable shape", frame: { verticalAlign: "middle" } } });
  slide.addElement({ type: "connector", id: "visual-connector", start: { x: 275, y: 205 }, route: [{ x: 325, y: 170 }], end: { x: 390, y: 205 }, style: { paint: { type: "solid", color: { theme: "accent2" } }, width: 3, beginArrow: "oval", endArrow: "triangle" } });
  slide.addElement({ type: "image", id: "visual-crop", assetId: asset.id, fit: "crop", crop: { left: 0.2, right: 0.2, top: 0, bottom: 0 }, box: { x: 430, y: 35, width: 250, height: 150 } });
  slide.addElement({ type: "group", id: "visual-group", box: { x: 710, y: 35, width: 200, height: 150 }, coordinateSize: { width: 200, height: 150 }, transform: { rotation: 3 }, children: [
    { type: "shape", id: "visual-group-diamond", shape: "diamond", box: { x: 10, y: 20, width: 100, height: 100 }, style: { fill: { type: "solid", color: "FBCFE8" } } },
    { type: "text", id: "visual-group-label", content: "Group", box: { x: 105, y: 55, width: 85, height: 40 } },
  ] });
  slide.addElement({ type: "table", id: "visual-table", box: { x: 40, y: 310, width: 870, height: 170 }, columns: [290, 290, 290], rows: [
    { height: 55, cells: [{ content: "Merged heading", colSpan: 3, style: { fill: { type: "solid", color: "E0E7FF" } } }] },
    { height: 55, cells: [{ content: "Alpha" }, { content: "Beta" }, { content: "Gamma" }] },
    { height: 55, cells: [{ content: "Row span", rowSpan: 2 }, { content: "Delta" }, { content: "Epsilon" }] },
  ] });
  return presentation;
}

test("renders hybrid SVG fixtures in a real browser", async ({ page }) => {
  const result = await renderPresentationToSvg(visualFixture());
  expect(result.warnings).toEqual([]);
  await page.setContent(`<style>@font-face{font-family:'PPTKit Fixture';src:url(data:font/woff2;base64,${fixtureFont}) format('woff2');font-style:normal;font-weight:100 900;font-display:block}html,body{margin:0;background:#d1d5db}#frame{width:960px;height:540px;background:white}svg{display:block;width:960px;height:540px}</style><div id="frame">${result.slides[0].svg}</div>`);
  await page.evaluate(() => document.fonts.ready);

  await expect(page.locator("svg[data-pptkit-slide-id='visual']")).toHaveCount(1);
  await expect(page.locator("foreignObject[data-pptkit-element-id='visual-text']")).toHaveCount(1);
  await expect(page.locator("[data-pptkit-element-id='visual-table'] table")).toHaveCount(1);
  await expect(page.locator("#frame")).toHaveScreenshot("composite.png");

  const clips = [
    ["rich-text.png", { x: 25, y: 20, width: 420, height: 115 }],
    ["shape-connector.png", { x: 20, y: 125, width: 390, height: 165 }],
    ["image-crop.png", { x: 415, y: 20, width: 280, height: 180 }],
    ["nested-group.png", { x: 695, y: 20, width: 230, height: 180 }],
    ["table-spans.png", { x: 25, y: 295, width: 900, height: 200 }],
  ];
  for (const [name, clip] of clips) {
    expect(await page.screenshot({ clip, animations: "disabled" })).toMatchSnapshot(name);
  }
});

test("keeps long text inside its authored SVG bounds", async ({ page }) => {
  const presentation = createPresentation({
    theme: { fonts: { heading: "Arial", body: "Arial" } },
  });
  const slide = presentation.addSlide({ id: "text-units" });
  slide.addElement({
    type: "text",
    id: "long-text",
    content: [{
      style: { lineSpacing: 1.1 },
      runs: [{
        text: "Feature: bold, italic, underline, strike, color, size, and auto-fit",
        style: { fontFamily: "Arial", fontSize: 24 },
      }],
    }],
    box: { x: 48, y: 72, width: 624, height: 48 },
    frame: { margin: 0, autoFit: { mode: "shrink", fontScale: 0.8 } },
  });

  const result = await renderPresentationToSvg(presentation);
  await page.setContent(result.slides[0].svg);
  const textLines = page.locator("g[data-pptkit-element-id='long-text'] text[data-pptkit-text-line]");
  await expect(textLines).toHaveCount(1);
  await expect(textLines.locator("tspan")).toHaveAttribute("font-size", "19.2");
});

test("uses PowerPoint-oriented line wrapping for compact text boxes", async ({ page }) => {
  const presentation = createPresentation({
    theme: { fonts: { heading: "Helvetica Neue", body: "Helvetica Neue" } },
  });
  const slide = presentation.addSlide({ id: "font-overflow" });
  slide.addElement({
    type: "text",
    id: "discovery",
    content: [{
      style: { lineSpacing: 1 },
      runs: [{
        text: "Discovery becomes easier.",
        style: { fontFamily: "Helvetica Neue", fontSize: 21.75, bold: true },
      }],
    }],
    box: { x: 504, y: 243, width: 255, height: 39 },
    frame: { autoFit: { mode: "shrink", fontScale: 0.92595 } },
  });

  const result = await renderPresentationToSvg(presentation);
  await page.setContent(result.slides[0].svg);
  const lines = await page.locator("g[data-pptkit-element-id='discovery'] [data-pptkit-text-line]").allTextContents();
  expect(lines).toEqual(["Discovery becomes ", "easier."]);
  await expect(page.locator("foreignObject[data-pptkit-element-id='discovery']")).toHaveCount(0);
});

test("wrapped SVG text retains source characters without moving line glyphs", async ({ page }) => {
  const presentation = createPresentation();
  const slide = presentation.addSlide();
  const samples = [
    ["left", "ABC (words 123) XYZ", "left"],
    ["center", "ABC (more 123) XYZ", "center"],
    ["right", "Wide   words andmorewords", "right"],
    ["hard", "ABC (words 123) XYZ  \r\n  hello 123 XYZ\rLast words\nEnd", "right"],
    ["blank", "\r\n\r  \nABC words 123\n\n", "left"],
    ["bullet", "ABC words 123\r\n\nhello 456 XYZ", "left", true],
  ];
  for (const [index, [id, text, align, bullet]] of samples.entries()) slide.addElement({
    type: "text", id, box: { x: 30, y: 20 + index * 150, width: 190, height: 130 },
    frame: { margin: 0, wrap: true }, content: [{ style: { align, ...(bullet ? { bullet: { type: "bullet", character: "•" } } : {}) }, runs: [{ text, style: { fontFamily: "Arial", fontSize: 32 } }] }],
  });
  await page.setContent((await renderPresentationToSvg(presentation)).slides[0].svg);
  const results = await page.evaluate(ids => ids.map(id => {
    const group = document.querySelector(`[data-pptkit-element-id="${id}"]`);
    const lines = [...group.querySelectorAll("text[data-pptkit-text-line]")];
    const positions = node => [...node.querySelectorAll("tspan[font-family]")].flatMap(span =>
      Array.from({ length: span.getNumberOfChars() }, (_, index) => {
        const box = span.getExtentOfChar(index); return [[box.x, box.y, box.width, box.height]];
      }));
    const matches = lines.map(line => {
      const original = positions(line), clone = line.cloneNode(true);
      for (const node of clone.querySelectorAll('tspan:not([font-family])')) node.remove();
      group.appendChild(clone); const control = positions(clone); clone.remove();
      return JSON.stringify(original) === JSON.stringify(control);
    });
    return { text: group.textContent, lineCount: lines.length, matches };
  }), samples.map(([id]) => id));
  for (const [index, result] of results.entries()) {
    expect(result.lineCount).toBeGreaterThan(1);
    expect(result.text).toBe(`${samples[index][3] ? "• " : ""}${samples[index][1]}`);
    expect(result.matches.every(Boolean)).toBe(true);
  }
});

test("wrapped bidi keeps paragraph context across single and split runs", async ({ page }) => {
  const presentation = createPresentation();
  const slide = presentation.addSlide();
  const source = "ABC שלום - 123 XYZ";
  for (const [i, width] of [80, 120, 160].entries()) for (const [j, fragments] of [[source], ["ABC שלום - ", "123 XYZ"]].entries()) {
    slide.addElement({ type: "text", id: `w${width}-${j}`, box: { x: 20 + i * 210, y: 20 + j * 220, width, height: 200 },
      frame: { margin: 0, wrap: true }, content: [{ runs: fragments.map(text => ({ text, style: { fontFamily: "Arial", fontSize: 24 } })) }] });
  }
  slide.addElement({ type: "text", id: "nowrap", box: { x: 20, y: 470, width: 400, height: 40 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: source }] }] });
  slide.addElement({ type: "text", id: "controls", box: { x: 440, y: 470, width: 100, height: 40 }, frame: { margin: 0, wrap: true }, content: [{ runs: [{ text: "ABC \u202eDEF\u202c" }] }] });
  slide.addElement({ type: "text", id: "marks", box: { x: 550, y: 470, width: 100, height: 40 }, frame: { margin: 0, wrap: true }, content: [{ runs: [{ text: "ABC \u200fDEF" }] }] });
  for (const [i, text] of ["ABC ܫܠܡ - 123 XYZ", "ABC ދިވެހި - 123 XYZ", "ABC ߒߞߏ - 123 XYZ", "ABC 𞤀 - 123 XYZ", "ABC ١٢٣ XYZ"].entries()) {
    slide.addElement({ type: "text", id: `rtl-script-${i}`, box: { x: i * 100, y: 530, width: 80, height: 100 }, frame: { margin: 0, wrap: true }, content: [{ runs: [{ text }] }] });
  }
  await page.setContent((await renderPresentationToSvg(presentation)).slides[0].svg);
  const items = await page.evaluate(() => [80, 120, 160].map(width => [0, 1].map(index => {
    const group = document.querySelector(`[data-pptkit-element-id="w${width}-${index}"]`);
    const rect = group.getBoundingClientRect(), walker = document.createTreeWalker(group, NodeFilter.SHOW_TEXT), chars = []; let node;
    while ((node = walker.nextNode())) for (let offset = 0; offset < node.length; offset++) {
      const range = document.createRange(); range.setStart(node, offset); range.setEnd(node, offset + 1); const box = range.getBoundingClientRect();
      chars.push({ char: node.data[offset], x: Math.round((box.x - rect.x) * 100) / 100, y: Math.round((box.y - rect.y) * 100) / 100 });
    }
    return { tag: group.tagName, chars };
  })));
  for (const [single, split] of items) {
    expect(single.tag.toLowerCase()).toBe("foreignobject");
    expect(single.chars).toEqual(split.chars);
  }
  await expect(page.locator('g[data-pptkit-element-id="nowrap"] text')).toHaveCount(1);
  await expect(page.locator('foreignObject[data-pptkit-element-id="controls"]')).toHaveCount(1);
  await expect(page.locator('foreignObject[data-pptkit-element-id="marks"]')).toHaveCount(1);
  for (let i = 0; i < 5; i++) await expect(page.locator(`foreignObject[data-pptkit-element-id="rtl-script-${i}"]`)).toHaveCount(1);
});

test("explicit paragraph direction controls the base bidi order without forcing alignment", async ({ page }) => {
  const presentation = createPresentation(), slide = presentation.addSlide();
  const text = "שלום (ABC 123) סוף";
  for (const [i, direction] of ["ltr", "rtl"].entries()) slide.addElement({
    type: "text", id: direction, box: { x: 20, y: 20 + i * 100, width: 400, height: 80 }, frame: { margin: 0, wrap: true },
    content: [{ style: { direction, align: "left" }, runs: [{ text, style: { fontFamily: "Arial", fontSize: 24 } }] }],
  });
  slide.addElement({ type: "text", id: "explicit-ltr-native", box: { x: 20, y: 250, width: 300, height: 50 }, frame: { margin: 0, wrap: false }, content: [{ style: { direction: "ltr" }, runs: [{ text: "Plain text" }] }] });
  await page.setContent((await renderPresentationToSvg(presentation)).slides[0].svg);
  const actual = await page.evaluate(() => ["ltr", "rtl"].map(id => {
    const group = document.querySelector(`[data-pptkit-element-id="${id}"]`), paragraph = group.querySelector('div > div'), text = paragraph.querySelector('span[lang]').firstChild;
    const x = index => { const range = document.createRange(); range.setStart(text, index); range.setEnd(text, index + 1); return range.getBoundingClientRect().x; };
    return { tag: group.tagName.toLowerCase(), direction: getComputedStyle(paragraph).direction, alignment: getComputedStyle(paragraph).textAlign, first: x(0), last: x(text.length - 1) };
  }));
  expect(actual.map(item => [item.tag, item.direction, item.alignment])).toEqual([["foreignobject", "ltr", "left"], ["foreignobject", "rtl", "left"]]);
  expect(actual[0].first).toBeLessThan(actual[0].last);
  expect(actual[1].first).toBeGreaterThan(actual[1].last);
  await expect(page.locator('g[data-pptkit-element-id="explicit-ltr-native"] text')).toHaveAttribute('direction', 'ltr');
});

test("renders bar, line, and pie charts in a real browser", async ({ page }) => {
  const presentation = createPresentation({
    metadata: { title: "Chart visual regression" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A" } },
  });
  const slide = presentation.addSlide({ id: "chart-visual", background: { type: "solid", color: "F8FAFC" } });
  slide.addElement({
    type: "chart",
    id: "visual-bar",
    chartType: "bar",
    categories: ["Alpha", "Beta", "Gamma"],
    series: [
      { name: "A", values: [30, 50, 40], color: { theme: "accent1" } },
      { name: "B", values: [20, 35, 25], color: { theme: "accent2" } },
    ],
    title: "Bar",
    legend: { visible: true },
    box: { x: 40, y: 40, width: 280, height: 220 },
  });
  slide.addElement({
    type: "chart",
    id: "visual-line",
    chartType: "line",
    categories: ["Jan", "Feb", "Mar", "Apr"],
    series: [
      { name: "Growth", values: [10, 25, 20, 35], color: "22C55E" },
    ],
    title: "Line",
    legend: { visible: false },
    box: { x: 340, y: 40, width: 280, height: 220 },
  });
  slide.addElement({
    type: "chart",
    id: "visual-pie",
    chartType: "pie",
    categories: ["Red", "Green", "Blue"],
    series: [
      { name: "Slice", values: [45, 30, 25], color: "8B5CF6" },
    ],
    title: "Pie",
    legend: { visible: true },
    box: { x: 640, y: 40, width: 280, height: 220 },
  });

  const result = await renderPresentationToSvg(presentation);
  expect(result.warnings).toEqual([]);
  await page.setContent(`<style>html,body{margin:0;background:#d1d5db}#frame{width:960px;height:540px;background:white}svg{display:block;width:960px;height:540px}</style><div id="frame">${result.slides[0].svg}</div>`);

  await expect(page.locator("svg[data-pptkit-slide-id='chart-visual']")).toHaveCount(1);
  await expect(page.locator("[data-pptkit-element-id='visual-bar'] [data-chart-role='bar']")).toHaveCount(6);
  await expect(page.locator("[data-pptkit-element-id='visual-line'] polyline")).toHaveCount(1);
  await expect(page.locator("[data-pptkit-element-id='visual-pie'] path")).toHaveCount(3);
  await expect(page.locator("#frame")).toHaveScreenshot("charts.png");

  const clips = [
    ["chart-bar.png", { x: 25, y: 25, width: 310, height: 250 }],
    ["chart-line.png", { x: 325, y: 25, width: 310, height: 250 }],
    ["chart-pie.png", { x: 625, y: 25, width: 310, height: 250 }],
  ];
  for (const [name, clip] of clips) {
    expect(await page.screenshot({ clip, animations: "disabled" })).toMatchSnapshot(name);
  }
});


test("inline slides keep chart, image and marker references local and unique", async ({ page }) => {
  const deck = createPresentation();
  const asset = deck.registerAsset({ id: "image-asset", kind: "image", source: {
    type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20'%3E%3Crect width='20' height='20' fill='red'/%3E%3C/svg%3E" }, width: 20, height: 20 });
  for (const [index, slideId] of ["slide/one", "slide?one"].entries()) {
    const slide = deck.addSlide({ id: slideId });
    for (const [id, width, x] of [["chart/a", 240, 20], ["chart?a", 360, 300]]) {
      slide.addElement({ type: "chart", id: `${id}-${index}`, chartType: "bar", categories: ["A"],
        series: [{ name: "Data", values: [4] }], box: { x, y: 20, width, height: 180 } });
    }
    slide.addElement({ type: "image", id: `image/a-${index}`, assetId: asset.id, box: { x: 20, y: 240, width: 100, height: 100 } });
    slide.addElement({ type: "connector", id: `arrow/a-${index}`, start: { x: 150, y: 250 }, end: { x: 250, y: 250 },
      style: { paint: { type: "solid", color: "FF0000" }, beginArrow: "oval", endArrow: "triangle" } });
  }
  const first = await renderPresentationToSvg(deck);
  const second = await renderPresentationToSvg(deck);
  expect(first.slides.map(s => s.svg)).toEqual(second.slides.map(s => s.svg));
  await page.setContent(first.slides.map(s => s.svg).join(""));
  const result = await page.evaluate(() => {
    const nodes = [...document.querySelectorAll("svg defs [id]")];
    const ids = nodes.map(node => node.id);
    const refs = [...document.querySelectorAll("svg [clip-path], svg [marker-start], svg [marker-end]")];
    const badRefs = refs.flatMap(node => ["clip-path", "marker-start", "marker-end"].flatMap(attr => {
      const value = node.getAttribute(attr);
      if (!value) return [];
      const id = /^url\(#([^)]*)\)$/.exec(value)?.[1];
      const target = id && document.getElementById(id);
      return target && target.closest("svg") === node.closest("svg") ? [] : [value];
    }));
    const widths = [...document.querySelectorAll("svg [data-chart-layer=series]")].map(node => {
      const id = /^url\(#([^)]*)\)$/.exec(node.getAttribute("clip-path"))[1];
      return document.getElementById(id).firstElementChild.getBBox().width;
    });
    return { ids, badRefs, widths };
  });
  expect(result.ids).toHaveLength(10);
  expect(new Set(result.ids).size).toBe(result.ids.length);
  expect(result.badRefs).toEqual([]);
  expect(result.widths[0]).not.toBe(result.widths[1]);
  expect(result.widths.slice(0, 2)).toEqual(result.widths.slice(2));
});


test("namespaces isolate independently rendered copies in one document", async ({ page }) => {
  const deck = createPresentation();
  const slide = deck.addSlide({ id: "same-slide" });
  slide.addElement({ type: "chart", id: "same-chart", chartType: "bar", categories: ["A"],
    series: [{ name: "Data", values: [4] }], box: { x: 10, y: 10, width: 220, height: 160 } });
  const left = await renderPresentationToSvg(deck, { idNamespace: "copy / <&😀" });
  const right = await renderPresentationToSvg(deck, { idNamespace: "copy ? <&😀" });
  await page.setContent(left.slides[0].svg + right.slides[0].svg);
  const observed = await page.evaluate(() => {
    const svgs = [...document.querySelectorAll("svg")];
    const ids = [...document.querySelectorAll("[id]")].map(node => node.id);
    return { ids, refs: svgs.map(svg => {
      const title = document.getElementById(svg.getAttribute("aria-labelledby"));
      const series = svg.querySelector("[data-chart-layer=series]");
      const clip = /^url\(#([^)]*)\)$/.exec(series.getAttribute("clip-path"))[1];
      return title?.parentElement === svg && document.getElementById(clip)?.closest("svg") === svg;
    }) };
  });
  expect(new Set(observed.ids).size).toBe(observed.ids.length);
  expect(observed.refs).toEqual([true, true]);
});
