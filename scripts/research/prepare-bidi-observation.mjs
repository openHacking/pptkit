import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const directory = process.argv[2] ?? '.tmp/bidi-observation';
const wrap = process.argv[3] === '--wrap';
await mkdir(directory, { recursive: true });
const deck = createPresentation(), slide = deck.addSlide();
const samples = [
  ['hebrew-single', ['ABC (שלום 123) XYZ']],
  ['hebrew-runs', ['ABC (', 'של', 'ום', ' 123) XYZ']],
  ['arabic-single', ['ABC (مرحبا 123) XYZ']],
  ['arabic-runs', ['ABC (', 'مر', 'حبا', ' 123) XYZ']],
];
for (const [index, [id, fragments]] of samples.entries()) {
  slide.addElement({ type: 'text', id, box: wrap ? { x: 50, y: 20 + index * 130, width: 190, height: 120 } : { x: 50, y: 50 + index * 95, width: 850, height: 70 }, frame: { wrap, margin: 0 },
    content: [{ runs: fragments.map(text => ({ text, style: { fontFamily: 'Arial', fontSize: 32 } })) }] });
}
const [pptx, preview] = await Promise.all([generatePptx(deck), renderPresentationToSvg(deck)]);
await writeFile(`${directory}/source.pptx`, pptx.bytes);
await writeFile(`${directory}/preview.svg`, preview.slides[0].svg);
const browser = await chromium.launch({ headless: true });
let screenshot, requests = [];
try {
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.setContent(`<style>html,body{margin:0} svg{width:1920px;height:1080px;display:block}</style>${preview.slides[0].svg}`);
  await page.evaluate(() => document.fonts.ready);
  const positions = await page.evaluate(({ ids, measureLines }) => ids.map(id => {
    const element = document.querySelector(`[data-pptkit-element-id="${id}"]`);
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT); const chars = []; let node, index = 0;
    while ((node = walker.nextNode())) for (let offset = 0; offset < node.length; offset++) {
      const range = document.createRange(); range.setStart(node, offset); range.setEnd(node, offset + 1);
      const box = range.getBoundingClientRect(); chars.push({ index: index++, value: node.data[offset], x: box.x, width: box.width, ...(measureLines ? { y: box.y } : {}) });
    }
    return { id, chars };
  }), { ids: samples.map(([id]) => id), measureLines: wrap });
  await writeFile(`${directory}/positions.json`, `${JSON.stringify(positions, null, 2)}\n`);
  screenshot = await page.screenshot(); await writeFile(`${directory}/preview.png`, screenshot);
  console.log(JSON.stringify({ browser: browser.version(), requests, warnings: preview.warnings, samples, pptxSha256: createHash('sha256').update(pptx.bytes).digest('hex'), screenshotSha256: createHash('sha256').update(screenshot).digest('hex') }));
} finally { await browser.close(); }
