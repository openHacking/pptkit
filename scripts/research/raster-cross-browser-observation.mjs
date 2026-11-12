// Reproducible bounded cross-engine raster and embedded-face observation.
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { resolve } from 'node:path';
import { chromium, firefox, webkit } from '@playwright/test';

const parent = resolve('.tmp/raster-cross-browser-runs');
await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]);
await mkdir(out, { recursive: true });
const fontBytes = await readFile('packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2');
const hash = value => createHash('sha256').update(value).digest('hex');
const observations = [];
for (const [name, engine] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await engine.launch();
  try {
    const page = await browser.newPage(), requests = [];
    await page.route('**/*', route => { const url = route.request().url(); if (/^https?:/i.test(url)) { requests.push(url); return route.abort(); } return route.continue(); });
    for (const pkg of ['core', 'layout', 'svg-renderer']) await page.addScriptTag({ content: await readFile(`packages/${pkg}/dist/browser/global.js`, 'utf8') });
    const sample = await page.evaluate(async base64 => {
      const family = 'Raster Cross Engine Unique', d = PPTKit.createPresentation({ size: { width: 340, height: 120 }, theme: { fonts: { body: family, heading: family } } });
      d.addSlide({ id: 's', elements: [
        { type: 'text', id: 't', box: { x: 4, y: 4, width: 330, height: 50 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: 'iiii WWWW 12345', style: { fontFamily: family, fontSize: 32 } }] }] },
        { type: 'shape', id: 'red', shape: 'rect', box: { x: 10, y: 70, width: 40, height: 40 }, style: { fill: { type: 'solid', color: 'FF0000' }, stroke: { paint: { type: 'none' } } } },
        { type: 'shape', id: 'blue', shape: 'rect', box: { x: 30, y: 70, width: 40, height: 40 }, opacity: 0.5, style: { fill: { type: 'solid', color: '0000FF' }, stroke: { paint: { type: 'none' } } } },
      ] });
      const layout = PPTKit.resolveLayout(d), font = { family, weight: 400, italic: false, format: 'woff2', base64 };
      const plain = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', backgroundMode: 'transparent' });
      const covered = await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', backgroundMode: 'transparent', fontPolicy: 'require-covered', embeddedFonts: [font], fontCoverage: request => /^[\x00-\x7f]*$/.test(request.text) });
      let missing; try { await PPTKit.renderLayoutToBrowserRaster(layout, { format: 'png', fontPolicy: 'require-embedded' }); } catch (e) { missing = e.code; }
      const decode = async bytes => {
        const url = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
        try { const image = new Image(); image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = 340; canvas.height = 120; const c = canvas.getContext('2d'); c.drawImage(image, 0, 0); return c.getImageData(0, 0, 340, 120).data; }
        finally { URL.revokeObjectURL(url); }
      };
      const a = await decode(plain.slides[0].bytes), b = await decode(covered.slides[0].bytes);
      let changed = 0, dark = 0;
      for (let i = 0; i < a.length; i += 4) {
        if (a[i] !== b[i] || a[i + 1] !== b[i + 1] || a[i + 2] !== b[i + 2] || a[i + 3] !== b[i + 3]) changed++;
        if (b[i] < 120 && b[i + 1] < 120 && b[i + 2] < 120 && b[i + 3] > 0) dark++;
      }
      const pixel = (x, y) => [...b.slice((y * 340 + x) * 4, (y * 340 + x) * 4 + 4)];
      return { plain: [...plain.slides[0].bytes], covered: [...covered.slides[0].bytes], changed, dark, overlay: pixel(40, 90), transparent: pixel(320, 110), missing, warnings: covered.warnings.map(w => w.code) };
    }, fontBytes.toString('base64'));
    if (requests.length || sample.missing !== 'raster-font-unavailable' || sample.changed < 100 || sample.dark < 100 || sample.transparent.join(',') !== '0,0,0,0' || Math.abs(sample.overlay[0] - 127) > 1 || sample.overlay.slice(1).join(',') !== '0,128,255') throw new Error(`Cross-browser specimen failed in ${name}: ${JSON.stringify({ requests, ...sample, plain: undefined, covered: undefined })}`);
    const plain = Buffer.from(sample.plain), covered = Buffer.from(sample.covered);
    await writeFile(`${out}/${name}-plain.png`, plain); await writeFile(`${out}/${name}-covered.png`, covered);
    observations.push({ engine: name, version: browser.version(), changedPixels: sample.changed, darkPixels: sample.dark, overlay: sample.overlay, transparent: sample.transparent, missingFont: sample.missing, warnings: sample.warnings, networkRequests: requests, plainSha256: hash(plain), coveredSha256: hash(covered) });
  } finally { await browser.close(); }
}
const result = { producerSha256: hash(await readFile(new URL(import.meta.url))), fontSha256: hash(fontBytes), observations };
await writeFile(`${out}/evidence.json`, `${JSON.stringify(result, null, 2)}\n`);
console.log(JSON.stringify({ directory: out, ...result }, null, 2));
