import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { inspectIsoBmff } from '../packages/media-inspection/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('ISO BMFF framing agrees in Node and Chromium without network access', async t => {
  const zip = await readFile('tests/fixtures/h264-mp4-powerpoint.pptx');
  const bytes = readGeneratedZip(zip).get('ppt/media/media1.mp4');
  const expected = inspectIsoBmff(bytes);
  const bundle = await rollup({ input: 'packages/media-inspection/dist/index.js' });
  t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'MediaInspection' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(data => {
    const bytes = Uint8Array.from(data), result = MediaInspection.inspectIsoBmff(bytes);
    let error;
    try { MediaInspection.inspectIsoBmff(bytes.subarray(0, bytes.length - 1)); } catch (e) { error = e.name; }
    return { result, error };
  }, [...bytes]);
  assert.deepEqual(actual, { result: expected, error: 'TypeError' });
  assert.deepEqual(requests, []);
});
