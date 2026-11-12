import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { inspectJpegIcc as inspect, extractJpegIccProfile, inspectIccProfile } from '../dist/index.js';
import { inspectJpegIcc as inspectNode, extractJpegIccProfile as extractNode, extractPngIccProfile } from '../dist/node.js';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const ident = Buffer.from('ICC_PROFILE\0', 'ascii');
const seg = (marker, payload) => Buffer.concat([Buffer.from([255, marker, (payload.length + 2) >> 8, (payload.length + 2) & 255]), payload]);
const chunk = (index, count, payload = Buffer.from([1, 2])) => seg(0xe2, Buffer.concat([ident, Buffer.from([index, count]), payload]));
const jpeg = (...segments) => Buffer.concat([Buffer.from([255,216]), ...segments, Buffer.from([255,217])]);

test('ICC inventory counts out-of-order chunks and reports missing chunks without profile claims', () => {
  const none = jpeg(seg(0xe2, Buffer.from('unrelated')));
  assert.deepEqual(inspect(none), { status: 'no-profile', stoppedAt: 'end', chunkCount: 0, receivedChunks: 0, missingChunks: [], profileBytes: 0 });
  const incomplete = jpeg(chunk(3, 3), chunk(1, 3, Buffer.from([7])));
  assert.deepEqual(inspect(incomplete), { status: 'incomplete', stoppedAt: 'end', chunkCount: 3, receivedChunks: 2, missingChunks: [2], profileBytes: 3 });
  const complete = jpeg(chunk(2, 2), chunk(1, 2, Buffer.from([9])));
  assert.deepEqual(inspect(complete), { status: 'complete', stoppedAt: 'end', chunkCount: 2, receivedChunks: 2, missingChunks: [], profileBytes: 3 });
  assert.deepEqual(inspectNode(complete), inspect(complete));
  const afterScan = Buffer.concat([Buffer.from([255,216]), seg(0xda, Buffer.alloc(0)), chunk(1,1), Buffer.from([255,217])]);
  assert.deepEqual(inspect(afterScan), { status: 'no-profile', stoppedAt: 'scan', chunkCount: 0, receivedChunks: 0, missingChunks: [], profileBytes: 0 });
});

test('ICC malformed chunks and JPEG framing reject, with independent budgets', () => {
  const malformed = [jpeg(chunk(0, 1)), jpeg(chunk(2, 1)), jpeg(chunk(1, 0)), jpeg(chunk(1, 1), chunk(1, 1)), jpeg(chunk(1, 2), chunk(2, 3)), jpeg(seg(0xe2, Buffer.from('ICC_PROFILE'))), jpeg(seg(0xe2, Buffer.from('ICC_PROFILEXzz'))), Buffer.from([255,216,255,226,0,8,1]), Buffer.from([255,216,255,0]), Buffer.from([255,216])];
  for (const bytes of malformed) assert.throws(() => inspect(bytes), /Invalid JPEG ICC/);
  const bytes = jpeg(chunk(1, 1));
  assert.throws(() => inspect(bytes, { maxBytes: bytes.length - 1 }), /maxBytes/);
  assert.throws(() => inspect(bytes, { maxProfileBytes: 1 }), /maxProfileBytes/);
  assert.throws(() => inspect(jpeg(chunk(1, 2), chunk(2, 2)), { maxSegments: 1 }), /maxSegments/);
  for (const key of ['maxBytes', 'maxSegments', 'maxProfileBytes']) for (const value of [0, -1, NaN, 1.5, Infinity]) assert.throws(() => inspect(bytes, { [key]: value }), RangeError);
  const controller = new AbortController(), reason = new Error('cancel'); controller.abort(reason);
  assert.throws(() => inspect(bytes, { signal: controller.signal }), error => error === reason);
  if (typeof SharedArrayBuffer !== 'undefined') assert.throws(() => inspect(new Uint8Array(new SharedArrayBuffer(bytes.length))), /SharedArrayBuffer/);
  class Hostile extends Uint8Array { get byteLength() { return 0; } get buffer() { return new ArrayBuffer(0); } *[Symbol.iterator]() { throw Error('unexpected iterator'); } }
  const hostile = new Hostile(bytes); assert.equal(inspect(hostile).status, 'complete');
  assert.throws(() => inspect(hostile, { maxBytes: bytes.length - 1 }), /maxBytes/);
});

test('ICC inspector rejects cross-realm shared backing storage', async () => {
  if (typeof SharedArrayBuffer === 'undefined') return;
  const vm = await import('node:vm'), foreign = vm.runInNewContext('new SharedArrayBuffer(32)');
  assert.throws(() => inspect(new Uint8Array(foreign)), /SharedArrayBuffer/);
});

test('JPEG ICC extraction reassembles shuffled chunks, bounds output and returns detached bytes', async () => {
  const source = await readFile(new URL('../../../tests/fixtures/jpeg-icc-srgb.jpg', import.meta.url));
  const expected = extractPngIccProfile(await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png', import.meta.url))).profileBytes;
  const actual = extractJpegIccProfile(source); assert.deepEqual(actual.profileBytes, expected);
  assert.deepEqual(extractNode(source), actual); assert.equal(inspectIccProfile(actual.profileBytes).tags.length, 11);
  const parts = [expected.subarray(0, 100), expected.subarray(100, 350), expected.subarray(350)];
  const shuffled = jpeg(chunk(3, 3, parts[2]), chunk(1, 3, parts[0]), chunk(2, 3, parts[1]));
  const reassembled = extractJpegIccProfile(shuffled, { maxProfileBytes: expected.length });
  assert.deepEqual(reassembled.profileBytes, expected); assert.equal(reassembled.inspection.receivedChunks, 3);
  reassembled.profileBytes[0] ^= 255; assert.deepEqual(extractJpegIccProfile(shuffled).profileBytes, expected);
  assert.throws(() => extractJpegIccProfile(shuffled, { maxProfileBytes: expected.length - 1 }), /maxProfileBytes/);
  assert.throws(() => extractJpegIccProfile(jpeg(chunk(1, 2))), /incomplete/);
  assert.throws(() => extractJpegIccProfile(jpeg(chunk(1, 1), chunk(1, 1))), /duplicate/);
  assert.equal(extractJpegIccProfile(jpeg()).profileBytes, undefined);
  const controller = new AbortController(), reason = new Error('cancel extract'); controller.abort(reason);
  assert.throws(() => extractJpegIccProfile(shuffled, { signal: controller.signal }), e => e === reason);
  let calls = 0; assert.throws(() => extractJpegIccProfile(shuffled, { get maxProfileBytes() { calls++; return 1; } }), /maxProfileBytes/); assert.equal(calls, 1);
});

test('real JPEG ICC chunks are inventoried; opt-in export warns and preserves exact bytes', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/jpeg-icc-srgb.jpg', import.meta.url));
  const info = inspect(bytes); assert.equal(info.status, 'complete'); assert.equal(info.chunkCount, 1); assert.ok(info.profileBytes > 128); assert.equal(info.stoppedAt, 'scan');
  const p = createPresentation(); p.registerAsset({ id: 'jpeg', kind: 'image', mimeType: 'image/jpeg', source: { type: 'bytes', value: bytes } });
  const slide = p.addSlide(); for (let i = 0; i < 2; i++) slide.addElement({ type: 'image', assetId: 'jpeg', box: { x: i * 30, y: 0, width: 20, height: 20 } });
  const normal = await generatePptx(p), diagnosed = await generatePptx(p, { jpegIccDiagnostics: 'warn' });
  assert.deepEqual(normal.warnings, []); assert.deepEqual(diagnosed.warnings.map(w => [w.code, w.assetId]), [['jpeg-icc-color-unverified', 'jpeg']]);
  assert.deepEqual(diagnosed.bytes, normal.bytes); assert.deepEqual(readGeneratedZip(diagnosed.bytes).get('ppt/media/jpeg.jpg'), bytes);
  await assert.rejects(generatePptx(p, { jpegIccDiagnostics: 'invalid' }), TypeError);
  for (const source of [jpeg(chunk(1,2)), jpeg(chunk(1,1), chunk(1,1))]) {
    const broken = createPresentation(); broken.registerAsset({ id: 'jpeg', kind: 'image', source: { type: 'bytes', value: source } });
    broken.addSlide().addElement({ type: 'image', assetId: 'jpeg', box: { x: 0, y: 0, width: 20, height: 20 } });
    const output = await generatePptx(broken, { jpegIccDiagnostics: 'warn' });
    assert.equal(output.warnings[0].code, 'jpeg-icc-inspection-incomplete');
    assert.deepEqual(readGeneratedZip(output.bytes).get('ppt/media/jpeg.jpg'), source);
  }
});
