// Optional independent pixel oracle; requires Python and Pillow, no browser or network.
import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { deflateSync } from 'node:zlib';
import { png } from '../packages/pptx-parser/test/image-fixture.mjs';
import { validatePngReplacement } from '../packages/pptx-parser/dist/import/validate-png-replacement.js';
import { indexedRows, filterRows } from '../tests/support/png-indexed-fixture.mjs';

function chunk(type, body) {
  const output = Buffer.alloc(12 + body.length); output.writeUInt32BE(body.length); output.write(type, 4); body.copy(output, 8);
  let crc = 0xffffffff; for (const byte of output.subarray(4, 8 + body.length)) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1; }
  output.writeUInt32BE((crc ^ 0xffffffff) >>> 0, 8 + body.length); return output;
}
function image(width, height, depth, interlaced, rows) {
  const header = Buffer.from(png.subarray(16, 29)); header.writeUInt32BE(width, 0); header.writeUInt32BE(height, 4); header[8] = depth; header[9] = 3; header[12] = interlaced ? 1 : 0;
  return Buffer.concat([png.subarray(0, 8), chunk('IHDR', header), chunk('PLTE', depth === 1 ? Buffer.from([0, 0, 0]) : Buffer.from([0, 0, 0, 255, 255, 255])), chunk('IDAT', deflateSync(rows)), chunk('IEND', Buffer.alloc(0))]);
}
function pillow(bytes) {
  const script = 'import sys,json,io; from PIL import Image; im=Image.open(io.BytesIO(sys.stdin.buffer.read())); im.load(); print(json.dumps({"mode":im.mode,"size":im.size,"pixels":list(im.getdata())}))';
  const result = spawnSync('python3', ['-c', script], { input: bytes, encoding: 'utf8', timeout: 10000, maxBuffer: 1024 * 1024 });
  assert.equal(result.status, 0, result.stderr); return JSON.parse(result.stdout);
}

test('Pillow independently decodes all filter types and indexed packed depths', async () => {
  for (const depth of [1, 2, 4, 8]) {
    const width = 11, height = 5, bytes = image(width, height, depth, false, filterRows(indexedRows(width, height, depth)));
    assert.equal(await validatePngReplacement(bytes, 100000, async () => {}), true);
    const result = pillow(bytes), expected = Array.from({ length: width * height }, (_, i) => depth === 1 ? 0 : (i % width + Math.floor(i / width)) % 2);
    assert.equal(result.mode, 'P'); assert.deepEqual(result.size, [width, height]); assert.deepEqual(result.pixels, expected);
  }
});

test('Pillow independently decodes indexed Adam7 pass placement and filter resets', async () => {
  const width = 11, height = 9, depth = 2, geometry = [[0,0,8,8],[4,0,8,8],[0,4,4,8],[2,0,4,4],[0,2,2,4],[1,0,2,2],[0,1,1,2]], pixels = Array(width * height).fill(-1), passes = [];
  for (const [startX, startY, dx, dy] of geometry) {
    const w = width <= startX ? 0 : Math.ceil((width - startX) / dx), h = height <= startY ? 0 : Math.ceil((height - startY) / dy);
    if (!w || !h) continue;
    passes.push(filterRows(indexedRows(w, h, depth)));
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) pixels[(startY + y * dy) * width + startX + x * dx] = (x + y) % 2;
  }
  assert.ok(pixels.every(value => value >= 0));
  const bytes = image(width, height, depth, true, Buffer.concat(passes));
  assert.equal(await validatePngReplacement(bytes, 100000, async () => {}), true);
  assert.deepEqual(pillow(bytes).pixels, pixels);
});

test('Pillow independently resolves grayscale, RGB and indexed transparency', async () => {
  for (const [color, depth, palette, trns, row, expected] of [
    [0, 8, null, [0, 17], [0, 17, 42], [0, 255]],
    [2, 8, null, [0, 1, 0, 2, 0, 3], [0, 1, 2, 3, 4, 5, 6], [0, 255]],
    [3, 1, [0, 0, 0, 255, 255, 255], [0, 127], [0, 0x40], [0, 127]],
  ]) {
    const header = Buffer.from(png.subarray(16, 29)); header.writeUInt32BE(2, 0); header.writeUInt32BE(1, 4); header[8] = depth; header[9] = color; header[12] = 0;
    const bytes = Buffer.concat([png.subarray(0, 8), chunk('IHDR', header), ...(palette ? [chunk('PLTE', Buffer.from(palette))] : []), chunk('tRNS', Buffer.from(trns)), chunk('IDAT', deflateSync(Buffer.from(row))), chunk('IEND', Buffer.alloc(0))]);
    assert.equal(await validatePngReplacement(bytes, 100000, async () => {}), true);
    const script = 'import sys,json,io; from PIL import Image; im=Image.open(io.BytesIO(sys.stdin.buffer.read())).convert("RGBA"); print(json.dumps([p[3] for p in im.getdata()]))';
    const result = spawnSync('python3', ['-c', script], { input: bytes, encoding: 'utf8', timeout: 10000 });
    assert.equal(result.status, 0, result.stderr); assert.deepEqual(JSON.parse(result.stdout), expected);
  }
});
