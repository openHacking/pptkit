import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { DEFAULT_SESSION_LIMITS } from '../packages/pptx-parser/dist/import/import-pptx.js';
import { preparePictureCandidate } from '../packages/pptx-parser/dist/import/prepare-picture-candidate.js';
import { imageFixture, encodeImages, picture, validGif as gif } from '../packages/pptx-parser/test/image-fixture.mjs';
import { readFile } from 'node:fs/promises';
import { validateJpegReplacement } from '../packages/pptx-parser/dist/import/validate-jpeg-replacement.js';

test('private picture candidate construction and validation match Node in Chromium', async t => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' }); t.after(() => session.dispose());
  const binding = session.report.pictureSources[0], limits = { ...DEFAULT_SESSION_LIMITS }, expected = await preparePictureCandidate(source, binding, gif, 'image/gif', {}, limits, async () => {});
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/import/prepare-picture-candidate.js', plugins: [nodeResolve({ browser: true }), commonjs()] }); t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'PictureCandidateTest' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(async ({ source, image, binding, limits }) => {
    const result = await PictureCandidateTest.preparePictureCandidate(new Uint8Array(source), binding, new Uint8Array(image), 'image/gif', {}, limits, async () => {});
    return { bytes: Array.from(result.bytes), part: result.part, relationshipId: result.relationshipId };
  }, { source: Array.from(source), image: Array.from(gif), binding, limits });
  assert.equal(actual.part, expected.part); assert.equal(actual.relationshipId, expected.relationshipId); assert.deepEqual(new Uint8Array(actual.bytes), expected.bytes); assert.deepEqual(requests, []);
});

test('JPEG component guards match Node in Chromium for baseline and progressive files', async t => {
  const baseline = await readFile('tests/fixtures/exif-orientation6.jpg');
  const progressive = await readFile('tests/fixtures/jpeg-progressive-components.jpg');
  const corrupt = Buffer.from(baseline), scan = corrupt.indexOf(Buffer.from([0xff, 0xda]));
  assert.ok(scan >= 0); corrupt[scan + 5] = 255;
  const missingTable = Buffer.from(baseline), frame = missingTable.indexOf(Buffer.from([0xff, 0xc0]));
  assert.ok(frame >= 0); missingTable[frame + 12] = 3;
  const dht = baseline.indexOf(Buffer.from([0xff, 0xc4])); assert.ok(dht >= 0);
  const noHuffman = Buffer.concat([baseline.subarray(0, dht), baseline.subarray(dht + 2 + baseline.readUInt16BE(dht + 2))]);
  const missingProgressiveDc = Buffer.from(progressive), firstScan = missingProgressiveDc.indexOf(Buffer.from([0xff, 0xda]));
  assert.ok(firstScan >= 0); missingProgressiveDc[firstScan + 6] |= 0x30;
  const badProgression = Buffer.from(progressive); badProgression[firstScan + 6 + 2 * badProgression[firstScan + 4]] = 1;
  const samples = [baseline, progressive, corrupt, missingTable, noHuffman, missingProgressiveDc, badProgression];
  const expected = await Promise.all(samples.map(bytes => validateJpegReplacement(bytes, 1e7, { work: 0 }, async () => {})));
  assert.deepEqual(expected, [true, true, false, false, false, false, false]);
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/import/validate-jpeg-replacement.js', plugins: [nodeResolve({ browser: true }), commonjs()] }); t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'JpegValidationTest' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(async samples => Promise.all(samples.map(bytes => JpegValidationTest.validateJpegReplacement(new Uint8Array(bytes), 1e7, { work: 0 }, async () => {}))), samples.map(bytes => Array.from(bytes)));
  assert.deepEqual(actual, expected);
});
