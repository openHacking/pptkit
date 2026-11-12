import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const dir = '.tmp/native-jpeg-12';
const sha = value => createHash('sha256').update(value).digest('hex');
const entries = await Promise.all(['sequential', 'progressive'].map(async name => ({
  source: sha(await readFile(`tests/fixtures/jpeg-12bit-${name}.jpg`)),
  png: await readFile(`${dir}/decoded-${name}.png`),
})));
const input = await readFile(`${dir}/crop.pptx`), session = await importPptx(input, { projection: 'content' });
try {
  const resolved = [];
  const result = await renderPresentationToSvg(session.document, { resolveAsset: asset => {
    if (asset.source.type !== 'bytes') throw Error('Unexpected external source.');
    const entry = entries.find(item => item.source === sha(Buffer.from(asset.source.value, 'base64')));
    if (!entry) throw Error('Unexpected image bytes.');
    resolved.push({ assetId: asset.id, sourceSha256: entry.source, pngSha256: sha(entry.png) });
    return `data:image/png;base64,${entry.png.toString('base64')}`;
  } });
  const svg = result.slides[0].svg;
  await writeFile(`${dir}/resolved-preview.svg`, svg);
  const browser = await chromium.launch({ headless: true });
  let screenshot, decoded, requests = [];
  try {
    const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
    await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
    await page.setContent(`<style>html,body{margin:0} svg{width:1920px;height:1080px;display:block}</style>${svg}`);
    decoded = await page.evaluate(async () => Promise.all([...document.querySelectorAll('svg image')].map(async node => {
      const image = new Image(); image.src = node.getAttribute('href'); await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
      const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
      return { width: image.naturalWidth, height: image.naturalHeight, firstPixel: [...context.getImageData(0, 0, 1, 1).data] };
    })));
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    screenshot = await page.screenshot(); await writeFile(`${dir}/resolved-preview.png`, screenshot);
  } finally { await browser.close(); }
  assert.deepEqual(requests, []); assert.deepEqual(result.warnings, []);
  assert.equal(resolved.length, 2); assert.equal(decoded.length, 3); assert.ok(decoded.every(image => image.firstPixel[3] === 255));
  const report = { inputSha256: sha(input), svgSha256: sha(svg), screenshotSha256: sha(screenshot), resolved, decoded, requests, warnings: result.warnings };
  await writeFile(`${dir}/resolved-preview-report.json`, `${JSON.stringify(report, null, 2)}\n`);
  console.log(JSON.stringify(report));
} finally { session.dispose(); }
