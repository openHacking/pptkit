import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { resolveLayout } from '../../packages/layout/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';
import { renderLayoutToPdf } from '../../packages/pdf-renderer/dist/index.js';

// An experiment, not a PDF adapter: printing requires a Playwright-controlled browser.
const parent = resolve('.tmp/pdf-print-runs');
await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]);
await mkdir(out, { recursive: true });
const hash = value => createHash('sha256').update(value).digest('hex');
const deck = createPresentation({ size: { width: 400, height: 200 } });
deck.addSlide({ id: 'simple', elements: [{ type: 'text', box: { x: 10, y: 10, width: 300, height: 50 }, content: 'Alpha example' }] });
deck.addSlide({ id: 'mixed', elements: [{ type: 'text', box: { x: 10, y: 10, width: 300, height: 50 }, content: [{ runs: [{ text: 'Bold ', style: { bold: true } }, { text: 'linked', style: { italic: true }, action: { type: 'url', url: 'https://example.com/beta' } }, { text: ' next', action: { type: 'slide', slideId: 'simple' } }] }] }, { type: 'shape', shape: 'rect', box: { x: 10, y: 80, width: 100, height: 50 }, style: { fill: { type: 'solid', color: 'FF0000' } }, action: { type: 'slide', slideId: 'simple' } }] });
const rendered = await renderPresentationToSvg(deck, { slideIndices: [1, 0] });
const browser = await chromium.launch();
try {
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
  const css = '@page{size:400pt 200pt;margin:0}body{margin:0}section{width:400pt;height:200pt;break-after:page}svg{width:100%;height:100%}';
  const html = `<style>${css}</style>${rendered.slides.map(slide => `<section>${slide.svg}</section>`).join('')}`;
  await writeFile(`${out}/source.html`, html);
  await page.setContent(html);
  await page.evaluate(() => document.fonts.ready);
  const bytes = await page.pdf({ preferCSSPageSize: true, printBackground: true });
  await writeFile(`${out}/printed.pdf`, bytes);
  const linkedSvg = await renderPresentationToSvg(deck, { slideIndices: [1, 0], slideActionMode: 'document-links' });
  const linkedHtml = `<style>${css}</style>${linkedSvg.slides.map(slide => `<section>${slide.svg}</section>`).join('')}`;
  await writeFile(`${out}/linked-source.html`, linkedHtml);
  await page.setContent(linkedHtml);
  await page.evaluate(() => document.fonts.ready);
  const linked = await page.evaluate(() => document.querySelector('svg[data-pptkit-slide-id="simple"]').id);
  const linkedBytes = await page.pdf({ preferCSSPageSize: true, printBackground: true });
  await writeFile(`${out}/linked.pdf`, linkedBytes);
  const adapter = await renderLayoutToPdf(resolveLayout(deck), { browser, slideIndices: [1, 0] });
  await writeFile(`${out}/adapter.pdf`, adapter.bytes);
  const manifest = {
    producerSha256: hash(await readFile(new URL(import.meta.url))), adapterSourceSha256: hash(await readFile('packages/pdf-renderer/src/rendering/print.ts')),
    browser: browser.version(), pdfSha256: hash(bytes), linkedPdfSha256: hash(linkedBytes), adapterPdfSha256: hash(adapter.bytes), adapterPages: adapter.pages, adapterWarnings: adapter.warnings, linkedTargetId: linked, htmlSha256: hash(html), linkedHtmlSha256: hash(linkedHtml), linkedWarnings: linkedSvg.warnings,
    slideIds: rendered.slides.map(slide => slide.slideId),
    foreignObject: rendered.slides.map(slide => slide.svg.includes('foreignObject')),
    warnings: rendered.warnings, networkRequests: requests,
  };
  await writeFile(`${out}/evidence.json`, `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(JSON.stringify({ directory: out, ...manifest }, null, 2));
} finally { await browser.close(); }
