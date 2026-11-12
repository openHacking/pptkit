import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { normalizeJpegOrientationInBrowser } from '../dist/dom.js';
const source = () => readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg',import.meta.url));

test('browser adapter rejects unsupported environments, invalid limits and malformed input', async () => {
  const bytes=await source();
  await assert.rejects(normalizeJpegOrientationInBrowser(bytes),/browser document/);
  for (const key of ['maxInputBytes','maxPixels','maxOutputBytes']) for (const value of [0,-1,1.5,Infinity,null]) {
    await assert.rejects(normalizeJpegOrientationInBrowser(bytes,{[key]:value}),RangeError);
  }
  await assert.rejects(normalizeJpegOrientationInBrowser([]),TypeError);
  await assert.rejects(normalizeJpegOrientationInBrowser(Uint8Array.of(1,2,3)),/JPEG/);
  if (typeof SharedArrayBuffer !== 'undefined') await assert.rejects(normalizeJpegOrientationInBrowser(new Uint8Array(new SharedArrayBuffer(8))),/SharedArrayBuffer/);
  const controller=new AbortController(), reason=new Error('stop');controller.abort(reason);
  await assert.rejects(normalizeJpegOrientationInBrowser(bytes,{signal:controller.signal}),error=>error===reason);
});
