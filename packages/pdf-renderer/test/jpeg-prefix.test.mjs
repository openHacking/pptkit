import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { jpegDataUrlPrecision } from '../dist/rendering/jpeg-prefix.js';

const fixture = name => readFileSync(new URL(`../../../tests/fixtures/${name}`, import.meta.url));
const base64 = (bytes, mime = 'image/jpeg') => `data:${mime};base64,${bytes.toString('base64')}`;
const escaped = (bytes, mime = 'image/jpeg') => `data:${mime},${[...bytes].map(byte => `%${byte.toString(16).padStart(2, '0')}`).join('')}`;

test('bounded JPEG precision detects actual signatures independent of declared MIME', () => {
  const normal = fixture('exif-orientation6.jpg'), high = fixture('jpeg-12bit-sequential.jpg');
  assert.equal(jpegDataUrlPrecision(base64(normal)), 8);
  assert.equal(jpegDataUrlPrecision(base64(high)), 12);
  assert.equal(jpegDataUrlPrecision(base64(high, 'image/png')), 12);
  assert.equal(jpegDataUrlPrecision(escaped(high, 'image/png')), 12);
  assert.equal(jpegDataUrlPrecision(base64(fixture('transparent-replacement.png'), 'image/png')), 'not-jpeg');
  assert.equal(jpegDataUrlPrecision(base64(fixture('transparent-replacement.png'))), 'unavailable');
});

test('bounded probe fails closed on malformed or unavailable frame headers', () => {
  const normal = fixture('exif-orientation6.jpg');
  const delayed = Buffer.concat([normal.subarray(0, 2), Buffer.from([0xff, 0xe1, 0xff, 0xff]), Buffer.alloc(65533), normal.subarray(2)]);
  for (const value of [base64(delayed), 'data:image/jpeg;base64,/9j/2Q==', 'data:image/jpeg;base64,%%%%', 'data:image/jpeg,%FF%D8%ZZ', 'data:image/jpeg,é', 'data:image/jpeg,%FF%D8#tail']) {
    assert.equal(jpegDataUrlPrecision(value), 'unavailable');
  }
});

test('probe reports only the first frame header and does not validate later structure', () => {
  const contrived = Buffer.from([0xff, 0xd8, 0xff, 0xc0, 0, 8, 8, 0, 0, 0, 0, 0, 0xff, 0xc1, 0, 8, 12, 0, 0, 0, 0, 0]);
  assert.equal(jpegDataUrlPrecision(base64(contrived)), 8);
});
