import { createHash } from 'node:crypto';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { chromium } from '@playwright/test';
import { createPresentation } from '../../packages/core/dist/index.js';
import { resolveLayout } from '../../packages/layout/dist/index.js';
import { renderLayoutToPdf } from '../../packages/pdf-renderer/dist/index.js';

const parent = resolve('.tmp/pdf-jpeg-opacity-runs'); await mkdir(parent, { recursive: true });
const out = process.argv[2] === undefined ? await mkdtemp(`${parent}/run-`) : resolve(process.argv[2]); await mkdir(out, { recursive: true });
const hash = value => createHash('sha256').update(value).digest('hex');
const names = ['exif-orientation6.jpg', 'jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg'];
const fixtures = Object.fromEntries(await Promise.all(names.map(async name => [name, await readFile(`tests/fixtures/${name}`)])));
const browser = await chromium.launch();
try {
  const results = [];
  for (const [name, bytes] of Object.entries(fixtures)) {
    const d = createPresentation({ size: { width: 6, height: 4 } });
    d.registerAsset({ id: 'image', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${bytes.toString('base64')}` } });
    d.addSlide({ id: 's', elements: [{ type: 'image', assetId: 'image', box: { x: 0, y: 0, width: 6, height: 4 }, fit: 'stretch' }] });
    for (const [policy, options] of [['browser', {}], ['require-opaque', { jpegDecodePolicy: 'require-opaque' }], ['require-8-bit', { jpegPrecisionPolicy: 'require-8-bit' }], ['require-valid-8-bit', { jpegStructurePolicy: 'require-valid-8-bit' }]]) {
      try {
        const result = await renderLayoutToPdf(resolveLayout(d), { browser, ...options });
        const file = `${name}-${policy}.pdf`; await writeFile(`${out}/${file}`, result.bytes);
        results.push({ name, policy, status: 'rendered', file, sha256: hash(result.bytes), warnings: result.warnings });
      } catch (error) { results.push({ name, policy, status: 'rejected', code: error.code ?? error.name }); }
    }
  }
  const normal = fixtures['exif-orientation6.jpg']; let at = 2, duplicate;
  while (at < normal.length) {
    if (normal[at++] !== 0xff) throw new Error('Unexpected JPEG marker');
    while (normal[at] === 0xff) at++;
    const marker = normal[at++], length = normal.readUInt16BE(at), start = at - 2, end = at + length;
    if ([0xc0, 0xc1, 0xc2].includes(marker)) { duplicate = Buffer.concat([normal.subarray(0, end), normal.subarray(start, end), normal.subarray(end)]); break; }
    at = end;
  }
  if (!duplicate) throw new Error('Missing JPEG frame');
  await writeFile(`${out}/duplicate-frame.jpg`, duplicate);
  const malformed = createPresentation(); malformed.registerAsset({ id: 'a', kind: 'image', source: { type: 'url', value: `data:image/jpeg;base64,${duplicate.toString('base64')}` } });
  malformed.addSlide({ elements: [{ type: 'image', assetId: 'a', box: { x: 0, y: 0, width: 6, height: 4 } }] });
  let duplicateResult;
  try { await renderLayoutToPdf(resolveLayout(malformed), { browser, jpegStructurePolicy: 'require-valid-8-bit' }); duplicateResult = 'accepted'; }
  catch (error) { duplicateResult = error.code ?? error.name; }
  const manifest = { producerSha256: hash(await readFile(new URL(import.meta.url))), adapterSourceSha256: hash(await readFile('packages/pdf-renderer/src/rendering/print.ts')), jpegProbeSha256: hash(await readFile('packages/pdf-renderer/src/rendering/jpeg-prefix.ts')), imageDataSha256: hash(await readFile('packages/pdf-renderer/src/rendering/image-data.ts')), structureSourceSha256: hash(await readFile('packages/media-inspection/src/jpeg/structure.ts')), browser: browser.version(), fixtures: Object.fromEntries(Object.entries(fixtures).map(([name, bytes]) => [name, hash(bytes)])), results, duplicate: { sha256: hash(duplicate), result: duplicateResult } };
  await writeFile(`${out}/evidence.json`, `${JSON.stringify(manifest, null, 2)}\n`);
  console.log(JSON.stringify({ directory: out, ...manifest }, null, 2));
} finally { await browser.close(); }
