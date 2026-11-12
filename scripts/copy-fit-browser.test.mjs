import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';

const url = async file => `data:text/javascript;base64,${(await readFile(new URL(file, import.meta.url))).toString('base64')}`;

test('real loaded browser font detects overflow introduced by explicit geometry fit', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(); await page.route('**/*', route => route.abort());
  const imports = Object.fromEntries(await Promise.all([
    ['@pptkit/core', '../packages/core/dist/browser/index.js'],
    ['@pptkit/layout', '../packages/layout/dist/browser/index.js'],
    ['@pptkit/layout/dom', '../packages/layout/dist/browser/dom.js'],
  ].map(async ([name, file]) => [name, await url(file)])));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  const font = (await readFile(new URL('../packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2', import.meta.url))).toString('base64');
  const result = await page.evaluate(async fontBytes => {
    const { createPresentation } = await import('@pptkit/core');
    const { resolveLayout } = await import('@pptkit/layout');
    const { createBrowserTextMeasurement } = await import('@pptkit/layout/dom');
    const face = new FontFace('Fit Pinned', `url(data:font/woff2;base64,${fontBytes})`); await face.load(); document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: 'Fit Pinned', faces: [face], covers: ({ text, bold, italic }) => !bold && !italic && /^[\x00-\x7f]*$/.test(text) }] });
    const source = createPresentation({ size: { width: 400, height: 200 }, theme: { fonts: { body: 'Fit Pinned', heading: 'Fit Pinned' } } });
    const target = createPresentation({ size: { width: 200, height: 200 }, theme: { fonts: { body: 'Fit Pinned', heading: 'Fit Pinned' } } });
    source.addSlide({ id: 'original', elements: [{ id: 'text', type: 'text', box: { x: 20, y: 10, width: 160, height: 40 }, frame: { margin: 0, wrap: true, verticalAlign: 'top' }, content: [{ style: { lineSpacing: 1.2 }, runs: [{ text: 'Geometry fit', style: { fontSize: 16 } }] }] }] });
    const copied = target.copySlideFrom(source, 'original', { layoutId: null, theme: 'destination', geometry: 'fit-explicit-geometry', customData: 'preserve', assets: { deduplicate: 'content', external: 'reject' } });
    // Permit one point of font ascent beyond the CSS line box in this control.
    const options = { textMeasurement: adapter, diagnostics: { textOverflow: true, tolerance: 1.1 } };
    const before = resolveLayout(source, options).analysis, after = resolveLayout(target, options).analysis;
    const unavailable = resolveLayout(target, { diagnostics: { textOverflow: true } }).analysis;
    const value = { before, after, unavailable, copyId: copied.elementIds.text, fittedBox: copied.slide.elements[0].box };
    adapter.dispose(); value.hostsRemaining = document.querySelectorAll('[data-pptkit-measurement]').length;
    return value;
  }, font);
  assert.deepEqual(result.fittedBox, { x: 10, y: 55, width: 80, height: 20 });
  assert.equal(result.before.diagnostics.some(d => d.code === 'text-overflow'), false, JSON.stringify(result.before));
  assert.ok(result.after.diagnostics.some(d => d.code === 'text-overflow' && d.elementId === result.copyId));
  assert.ok(result.unavailable.diagnostics.some(d => d.code === 'text-measurement-unavailable' && d.elementId === result.copyId));
  assert.equal(result.before.measurements.length, 1); assert.equal(result.after.measurements.length, 1);
  assert.ok(result.after.measurements[0].bounds.height > 20);
  assert.equal(result.hostsRemaining, 0);
});
