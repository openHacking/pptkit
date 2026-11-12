// Reproducible browser versus independent MuPDF SVG rendering probe.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../../packages/svg-renderer/dist/index.js';

const out = resolve(process.argv[2] ?? '.tmp/portable-svg-probe');
await mkdir(out, { recursive: true });
const p = createPresentation({ size: { width: 400, height: 260 }, theme: { fonts: { body: 'Arial', heading: 'Arial' } } });
p.registerAsset({ id: 'image', kind: 'image', source: { type: 'bytes', value: new Uint8Array(await readFile('tests/fixtures/exif-orientation6-normalized.png')) } });
p.addSlide({ id: 'specimen', elements: [
  { type: 'text', id: 'native', box: { x: 10, y: 10, width: 180, height: 30 }, content: 'Native text' },
  { type: 'text', id: 'rich', box: { x: 10, y: 50, width: 180, height: 30 }, content: [{ runs: [{ text: 'Rich ' }, { text: 'text', style: { bold: true } }] }] },
  { type: 'table', id: 'table', box: { x: 10, y: 95, width: 180, height: 45 }, columns: [90, 90], rows: [{ height: 45, cells: [{ content: 'Left', style: { fill: { type: 'solid', color: 'FFFF00' } } }, { content: 'Right', style: { fill: { type: 'solid', color: '00FFFF' } } }] }] },
  { type: 'image', id: 'crop', assetId: 'image', box: { x: 220, y: 10, width: 120, height: 120 }, fit: 'crop', crop: { left: 0, right: 0, top: 1 / 3, bottom: 1 / 3 } },
  { type: 'shape', id: 'red', shape: 'rect', box: { x: 220, y: 160, width: 100, height: 70 }, style: { fill: { type: 'solid', color: 'FF0000' }, stroke: { paint: { type: 'none' } } } },
  { type: 'shape', id: 'blue', shape: 'rect', box: { x: 270, y: 160, width: 100, height: 70 }, opacity: 0.5, style: { fill: { type: 'solid', color: '0000FF' }, stroke: { paint: { type: 'none' } } } },
] });
const result = await renderPresentationToSvg(p, { portabilityDiagnostics: 'warn' });
await writeFile(`${out}/specimen.svg`, result.slides[0].svg);
await writeFile(`${out}/warnings.json`, `${JSON.stringify(result.warnings, null, 2)}\n`);
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ viewport: { width: 400, height: 260 }, deviceScaleFactor: 1 });
  const requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.setContent(`<style>html,body{margin:0}</style>${result.slides[0].svg}`);
  await page.screenshot({ path: `${out}/chromium.png` });
  await writeFile(`${out}/browser-version.txt`, `${browser.version()}\n`);
  await writeFile(`${out}/requests.json`, `${JSON.stringify(requests)}\n`);
} finally { await browser.close(); }
console.log(out);
