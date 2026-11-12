import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { resolve } from 'node:path';
import { chromium, firefox, webkit } from '@playwright/test';

const parent = resolve('.tmp/jpeg-raster-opacity-runs'); await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]); await mkdir(out, { recursive: true });
const fixtures = ['exif-orientation6.jpg', 'jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg'];
const hash = value => createHash('sha256').update(value).digest('hex');
const images = Object.fromEntries(await Promise.all(fixtures.map(async name => [name, await readFile(`tests/fixtures/${name}`)])));
const observations = [];
for (const [engineName, engine] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await engine.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
    for (const pkg of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${pkg}/dist/browser/global.js`, 'utf8') });
    const results = await page.evaluate(async data => {
      const results = [];
      for (const [fixture, base64] of Object.entries(data)) {
        const d = PPTKit.createPresentation({ size: { width: 6, height: 4 } }); d.registerAsset({ id: 'image', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${base64}` } });
        d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'image', box: { x: 0, y: 0, width: 6, height: 4 }, fit: 'stretch' }] });
        const layout = PPTKit.resolveLayout(d);
        for (const policy of ['browser', 'require-opaque']) {
          try {
            const result = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', backgroundMode: 'transparent', jpegDecodePolicy: policy });
            const bytes = result.slides[0].bytes, url = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
            try { const image = new Image(); image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 6; canvas.height = 4; const ctx = canvas.getContext('2d'); ctx.drawImage(image, 0, 0); const rgba = ctx.getImageData(0, 0, 6, 4).data; let nonopaque = 0; for (let i = 3; i < rgba.length; i += 4) if (rgba[i] !== 255) nonopaque++; results.push({ fixture, policy, status: 'rendered', nonopaque, bytes: [...bytes] }); }
            finally { URL.revokeObjectURL(url); }
          } catch (error) { results.push({ fixture, policy, status: 'rejected', code: error.code ?? error.name }); }
        }
      }
      return results;
    }, Object.fromEntries(Object.entries(images).map(([name, bytes]) => [name, bytes.toString('base64')])));
    if (requests.length) throw new Error(`${engineName}: unexpected HTTP request`);
    for (const item of results) if (item.status === 'rendered') {
      const bytes = Buffer.from(item.bytes), path = `${engineName}-${fixtures.indexOf(item.fixture)}-${item.policy}.png`;
      await writeFile(`${out}/${path}`, bytes); item.png = path; item.pngSha256 = hash(bytes); delete item.bytes;
    }
    observations.push({ engine: engineName, version: browser.version(), networkRequests: requests, results });
  } finally { await browser.close(); }
}
const result = { producerSha256: hash(await readFile(new URL(import.meta.url))), fixtures: Object.fromEntries(Object.entries(images).map(([name, bytes]) => [name, hash(bytes)])), observations };
await writeFile(`${out}/evidence.json`, `${JSON.stringify(result, null, 2)}\n`);
console.log(JSON.stringify({ directory: out, ...result }, null, 2));
