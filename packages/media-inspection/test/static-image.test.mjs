import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { inspectGif, inspectStaticImage } from '../dist/index.js';

const fixture = name => readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url));
const stillGif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
const animatedGif = Buffer.from('R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7', 'base64');
function crc32(bytes) { let crc = 0xffffffff; for (const byte of bytes) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); } return (crc ^ 0xffffffff) >>> 0; }

test('static classifier distinguishes supported still and declared animation structures', () => {
  for (const [input, expected] of [
    [fixture('transparent-replacement.png'), { format: 'png', status: 'static' }],
    [fixture('two-frame-apng.png'), { format: 'png', status: 'animation-declared' }],
    [stillGif, { format: 'gif', status: 'static' }],
    [animatedGif, { format: 'gif', status: 'animation-declared' }],
    [fixture('exif-orientation6.jpg'), { format: 'jpeg', status: 'static' }],
    [Buffer.from('RIFFxxxxWEBP'), { format: 'unknown', status: 'unverified' }],
  ]) assert.deepEqual(inspectStaticImage(input), expected);
});

test('unknown PNG chunks and ambiguous single-frame GIF controls remain unverified', () => {
  const png = fixture('transparent-replacement.png'), type = Buffer.from('vpAg'), payload = Buffer.from([1]);
  const length = Buffer.alloc(4); length.writeUInt32BE(payload.length);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(Buffer.concat([type, payload])));
  const privateChunk = Buffer.concat([png.subarray(0, 33), length, type, payload, crc, png.subarray(33)]);
  assert.deepEqual(inspectStaticImage(privateChunk), { format: 'png', status: 'unverified' });
  const loop = Buffer.concat([Buffer.from([0x21, 0xff, 11]), Buffer.from('NETSCAPE2.0'), Buffer.from([3, 1, 0, 0, 0])]);
  assert.deepEqual(inspectStaticImage(Buffer.concat([stillGif.subarray(0, 19), loop, stillGif.subarray(19)])), { format: 'gif', status: 'unverified' });
  const unknown = Buffer.from([0x21, 0xce, 1, 0, 0]);
  assert.deepEqual(inspectStaticImage(Buffer.concat([stillGif.subarray(0, 19), unknown, stillGif.subarray(19)])), { format: 'gif', status: 'unverified' });
  const plainText = Buffer.concat([Buffer.from([0x21, 0x01, 12]), Buffer.alloc(12), Buffer.from([0])]);
  assert.deepEqual(inspectStaticImage(Buffer.concat([stillGif.subarray(0, 19), plainText, stillGif.subarray(19)])), { format: 'gif', status: 'unverified' });
  const interactive = Buffer.from(stillGif), gce = interactive.indexOf(Buffer.from([0x21, 0xf9, 4])); interactive[gce + 3] |= 2;
  assert.deepEqual(inspectStaticImage(interactive), { format: 'gif', status: 'unverified' });
  const disposal = Buffer.from(stillGif); disposal[gce + 3] |= 3 << 2;
  assert.deepEqual(inspectStaticImage(disposal), { format: 'gif', status: 'unverified' });
  const extraControl = Buffer.from([0x21, 0xf9, 4, 0, 0, 0, 0, 0]);
  const dangling = Buffer.concat([stillGif.subarray(0, -1), extraControl, stillGif.subarray(-1)]);
  assert.deepEqual(inspectStaticImage(dangling), { format: 'gif', status: 'unverified' });
  assert.deepEqual(inspectGif(dangling), inspectGif(stillGif));
  const reserved = Buffer.from(stillGif), imageDescriptor = reserved.indexOf(0x2c, 19); reserved[imageDescriptor + 9] |= 0x08;
  assert.deepEqual(inspectStaticImage(reserved), { format: 'gif', status: 'unverified' });
});

test('static classification rejects corrupted data, oversize inputs and cancellation', () => {
  const png = Buffer.from(fixture('transparent-replacement.png')); png[png.length - 1] ^= 1;
  assert.throws(() => inspectStaticImage(png), /CRC/);
  assert.throws(() => inspectStaticImage(stillGif.subarray(0, stillGif.length - 1)), /truncated/);
  assert.throws(() => inspectStaticImage(stillGif, { maxBytes: stillGif.length - 1 }), RangeError);
  const controller = new AbortController(); controller.abort(new Error('cancelled'));
  assert.throws(() => inspectStaticImage(stillGif, { signal: controller.signal }), /cancelled/);
});
