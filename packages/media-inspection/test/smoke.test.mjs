import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
import { validateJpegStructure } from '../dist/index.js';

const bytes = name => readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url));
function duplicateFrame(source) {
  let at = 2;
  while (at < source.length) {
    assert.equal(source[at++], 0xff);
    while (source[at] === 0xff) at++;
    const marker = source[at++], length = source.readUInt16BE(at), start = at - 2, end = at + length;
    if ([0xc0, 0xc1, 0xc2].includes(marker)) return Buffer.concat([source.subarray(0, end), source.subarray(start, end), source.subarray(end)]);
    at = end;
  }
  throw new Error('Fixture has no frame');
}

test('JPEG structure inspection reports precision without modifying input', () => {
  for (const [name, precision] of [['exif-orientation6.jpg', 8], ['jpeg-12bit-sequential.jpg', 12], ['jpeg-12bit-progressive.jpg', 12]]) {
    const input = bytes(name), before = Buffer.from(input);
    assert.deepEqual(validateJpegStructure(input), { precision });
    assert.deepEqual(input, before);
  }
  const source = bytes('exif-orientation6.jpg');
  const foreign = runInNewContext('Uint8Array.from(values)', { values: [...source] });
  assert.equal(validateJpegStructure(foreign).precision, 8);
});

test('JPEG structure rejects duplicate frames, truncation, trailing bytes and invalid types', () => {
  const source = bytes('exif-orientation6.jpg');
  assert.throws(() => validateJpegStructure(duplicateFrame(source)), /duplicate frame/);
  assert.throws(() => validateJpegStructure(source.subarray(0, source.length - 2)), /missing EOI/);
  assert.throws(() => validateJpegStructure(Buffer.concat([source, Buffer.from([0])])), /bytes after EOI/);
  assert.throws(() => validateJpegStructure(new DataView(new ArrayBuffer(4))), TypeError);
  const spoofed = new Uint8ClampedArray(4); Object.defineProperty(spoofed, Symbol.toStringTag, { value: 'Uint8Array' });
  assert.throws(() => validateJpegStructure(spoofed), TypeError);
  assert.throws(() => validateJpegStructure(new Uint8Array(64 * 1024 * 1024 + 1)), RangeError);
  const controller = new AbortController(); controller.abort(new Error('cancelled'));
  assert.throws(() => validateJpegStructure(source, controller.signal), /cancelled/);
});
