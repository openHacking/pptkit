import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const dir = '.tmp/native-jpeg-12';
const bytes = await readFile(`${dir}/crop.pptx`);
const session = await importPptx(bytes, { projection: 'content' });
try {
  const images = session.document.slides[0].elements.filter(element => element.type === 'image');
  assert.equal(images.length, 3); assert.equal(images[2].fit, 'crop');
  const preview = await renderPresentationToSvg(session.document);
  assert.equal(preview.slides.length, 1);
  const svg = preview.slides[0].svg;
  await writeFile(`${dir}/crop-preview.svg`, svg);
  const browser = await chromium.launch({ headless: true });
  let screenshot, decoded, control, requests = [];
  try {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
    await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
    await page.setContent(`<style>html,body{margin:0} svg{width:1920px;height:1080px;display:block}</style>${svg}`);
    decoded = await page.evaluate(async () => {
      const elements = [...document.querySelectorAll('svg image')];
      return Promise.all(elements.map(async element => {
        const url = element.getAttribute('href') || element.getAttributeNS('http://www.w3.org/1999/xlink', 'href');
        const image = new Image(); image.src = url;
        try {
          await image.decode();
          const canvas = document.createElement('canvas'); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
          const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
          return { ok: true, width: image.naturalWidth, height: image.naturalHeight, firstPixel: [...context.getImageData(0, 0, 1, 1).data] };
        }
        catch { return { ok: false, width: image.naturalWidth, height: image.naturalHeight }; }
      }));
    });
    control = await page.evaluate(async base64 => {
      const image = new Image(); image.src = `data:image/jpeg;base64,${base64}`; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
      const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
      return { width: image.naturalWidth, height: image.naturalHeight, firstPixel: [...context.getImageData(0, 0, 1, 1).data] };
    }, (await readFile('tests/fixtures/exif-orientation6.jpg')).toString('base64'));
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    screenshot = await page.screenshot();
    await writeFile(`${dir}/crop-preview.png`, screenshot);
  } finally { await browser.close(); }
  assert.deepEqual(requests, []);
  const sha = input => createHash('sha256').update(input).digest('hex');
  const report = { inputSha256: sha(bytes), svgSha256: sha(svg), screenshotSha256: sha(screenshot), decoded, control, requests, warnings: preview.warnings };
  await writeFile(`${dir}/crop-preview-report.json`, `${JSON.stringify(report, null, 2)}\n`);
  console.log(JSON.stringify(report));
} finally { session.dispose(); }
