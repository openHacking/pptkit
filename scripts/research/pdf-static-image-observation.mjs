import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { resolveLayout } from '../../packages/layout/dist/index.js';
import { renderLayoutToPdf } from '../../packages/pdf-renderer/dist/index.js';

const parent = resolve('.tmp/pdf-static-image-runs'); await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]); await mkdir(out, { recursive: true });
const hash = value => createHash('sha256').update(value).digest('hex');
const fixtures = ['transparent-replacement.png', 'two-frame-apng.png'];
const data = Object.fromEntries(await Promise.all(fixtures.map(async name => [name, await readFile(`tests/fixtures/${name}`)])));
const urls = Object.fromEntries(Object.entries(data).map(([name, bytes]) => [name, `data:image/png;base64,${bytes.toString('base64')}`]));
function deck(url, count) {
  const d = createPresentation({ size: { width: 40, height: 20 } });
  d.registerAsset({ id: 'asset', kind: 'image', source: { type: 'url', value: url } });
  for (let i = 0; i < count; i++) d.addSlide({ id: `s${i}`, elements: [{ type: 'image', assetId: 'asset', fit: 'stretch', box: { x: 0, y: 0, width: 40, height: 20 } }] });
  return resolveLayout(d);
}
const browser = await chromium.launch();
try {
  const calls = [];
  const still = await renderLayoutToPdf(deck(urls[fixtures[0]], 2), { browser, slideIndices: [1, 0], imageAnimationPolicy: 'require-static', isStaticImage: request => { calls.push({ slideId: request.slideId, slideIndex: request.slideIndex, mimeType: request.mimeType, dataSha256: hash(request.dataUrl) }); return request.dataUrl === urls[fixtures[0]]; } });
  await writeFile(`${out}/still.pdf`, still.bytes);
  let rejected;
  try { await renderLayoutToPdf(deck(urls[fixtures[1]], 1), { browser, imageAnimationPolicy: 'require-static', isStaticImage: request => request.dataUrl === urls[fixtures[0]] }); rejected = 'accepted'; }
  catch (error) { rejected = error.code ?? error.name; }
  const misattested = await renderLayoutToPdf(deck(urls[fixtures[1]], 1), { browser, imageAnimationPolicy: 'require-static', isStaticImage: () => true });
  await writeFile(`${out}/misattested.pdf`, misattested.bytes);
  const automatic = await renderLayoutToPdf(deck(urls[fixtures[0]], 2), { browser, slideIndices: [1, 0], imageAnimationPolicy: 'require-structurally-static' });
  await writeFile(`${out}/automatic.pdf`, automatic.bytes);
  let automaticRejected;
  try { await renderLayoutToPdf(deck(urls[fixtures[1]], 1), { browser, imageAnimationPolicy: 'require-structurally-static' }); automaticRejected = 'accepted'; }
  catch (error) { automaticRejected = error.code ?? error.name; }
  const stillGif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
  const animatedGif = Buffer.from('R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7', 'base64');
  await writeFile(`${out}/still.gif`, stillGif); await writeFile(`${out}/animated.gif`, animatedGif);
  const gif = await renderLayoutToPdf(deck(`data:image/gif;base64,${stillGif.toString('base64')}`, 1), { browser, imageAnimationPolicy: 'require-structurally-static' });
  await writeFile(`${out}/gif.pdf`, gif.bytes);
  let gifRejected;
  try { await renderLayoutToPdf(deck(`data:image/gif;base64,${animatedGif.toString('base64')}`, 1), { browser, imageAnimationPolicy: 'require-structurally-static' }); gifRejected = 'accepted'; }
  catch (error) { gifRejected = error.code ?? error.name; }
  const mediaSources = ['inspect-static.ts', 'capture-image-input.ts', 'inspect-png.ts', 'inspect-apng.ts', 'inspect-gif.ts', 'validate-gif-lzw.ts'];
  const manifest = { producerSha256: hash(await readFile(new URL(import.meta.url))), adapterSourceSha256: hash(await readFile('packages/pdf-renderer/src/rendering/print.ts')), imageDataSha256: hash(await readFile('packages/pdf-renderer/src/rendering/image-data.ts')), mediaSources: Object.fromEntries(await Promise.all(mediaSources.map(async name => [name, hash(await readFile(`packages/media-inspection/src/images/${name}`))]))), jpegSourceSha256: hash(await readFile('packages/media-inspection/src/jpeg/structure.ts')), browser: browser.version(), fixtures: Object.fromEntries(Object.entries(data).map(([name, bytes]) => [name, hash(bytes)])), gifFixtures: { still: hash(stillGif), animated: hash(animatedGif) }, calls, rejected, automaticRejected, gifRejected, still: { sha256: hash(still.bytes), pages: still.pages, warnings: still.warnings }, automatic: { sha256: hash(automatic.bytes), pages: automatic.pages, warnings: automatic.warnings }, gif: { sha256: hash(gif.bytes), pages: gif.pages, warnings: gif.warnings }, misattested: { sha256: hash(misattested.bytes), pages: misattested.pages, warnings: misattested.warnings } };
  await writeFile(`${out}/evidence.json`, `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(JSON.stringify({ directory: out, ...manifest }, null, 2));
} finally { await browser.close(); }
