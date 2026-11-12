import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateSync, inflateSync } from 'node:zlib';
import { importPptx } from '../dist/index.js';
import { validatePngReplacement } from '../dist/import/validate-png-replacement.js';
import { encodeImages, png } from './image-fixture.mjs';
import { indexedRows, filterRows } from '../../../tests/support/png-indexed-fixture.mjs';

function chunk(type, body) {
  const name = Buffer.from(type), out = Buffer.alloc(12 + body.length); out.writeUInt32BE(body.length); name.copy(out, 4); body.copy(out, 8);
  let crc = 0xffffffff;
  for (const byte of Buffer.concat([name, body])) { crc ^= byte; for (let i = 0; i < 8; i++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1; }
  out.writeUInt32BE((crc ^ 0xffffffff) >>> 0, out.length - 4); return out;
}
function split(bytes) {
  const result = []; for (let offset = 8; offset < bytes.length;) { const length = bytes.readUInt32BE(offset), end = offset + 12 + length; result.push(bytes.subarray(offset, end)); offset = end; }
  return result;
}
const [header, data, end] = split(png), signature = png.subarray(0, 8), pack = (...parts) => Buffer.concat([signature, ...parts]);
test('PNG transparency chunks match color type, palette, sample depth and ordering', async () => {
  const make = (color, depth, transparency, options = {}) => {
    const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(1, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = depth; ihdr[9] = color;
    const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[color], palette = chunk('PLTE', Buffer.from([0, 0, 0, 255, 255, 255]));
    const trns = chunk('tRNS', transparency), pixels = chunk('IDAT', deflateSync(Buffer.alloc(1 + Math.ceil(channels * depth / 8))));
    const before = color === 3 ? [palette, trns] : [trns];
    if (options.beforePalette) before.splice(0, before.length, trns, palette);
    if (options.duplicate) before.push(trns);
    if (options.afterData) return pack(chunk('IHDR', ihdr), ...(color === 3 ? [palette] : []), pixels, trns, end);
    return pack(chunk('IHDR', ihdr), ...before, pixels, end);
  };
  for (const [color, depth, value] of [[0, 1, [0, 1]], [0, 16, [255, 255]], [2, 8, [0, 255, 0, 2, 0, 1]], [3, 1, [0, 255]]]) {
    assert.equal(await validatePngReplacement(make(color, depth, Buffer.from(value)), 100000, async () => {}), true, `color ${color} depth ${depth}`);
  }
  for (const image of [make(0, 1, Buffer.from([0, 2])), make(0, 8, Buffer.from([0])), make(2, 8, Buffer.from([1, 0, 0, 0, 0, 0])), make(2, 8, Buffer.alloc(5)), make(3, 1, Buffer.alloc(0)), make(3, 1, Buffer.alloc(3)), make(4, 8, Buffer.alloc(2)), make(6, 8, Buffer.alloc(6)), make(3, 1, Buffer.alloc(1), { beforePalette: true }), make(2, 8, Buffer.alloc(6), { beforePalette: true }), make(0, 8, Buffer.alloc(2), { duplicate: true }), make(0, 8, Buffer.alloc(2), { afterData: true })]) {
    assert.equal(await validatePngReplacement(image, 100000, async () => {}), false);
  }
});
test('PNG replacement inspection rejects corrupt CRC, truncation, ordering and unsupported critical chunks', async () => {
  assert.equal(await validatePngReplacement(png, 100000, async () => {}), true);
  assert.equal(await validatePngReplacement(pack(header, chunk('tEXt', Buffer.from('label\0value')), data, end), 100000, async () => {}), true);
  const corrupt = Buffer.from(png); corrupt[20] ^= 1;
  for (const bytes of [png.subarray(0, 33), corrupt, Buffer.concat([png, Buffer.from([0])]), pack(header, end), pack(header, header, data, end), pack(header, data, chunk('tEXt', Buffer.from('x\0y')), data, end), pack(header, chunk('ABCD', Buffer.alloc(0)), data, end), pack(header, data)]) assert.equal(await validatePngReplacement(bytes, 100000, async () => {}), false);
  await assert.rejects(validatePngReplacement(png, 1, async () => {}), error => error.code === 'save-limit-exceeded');
});

test('PNG validation rejects a CRC-correct but invalid compressed stream', async () => {
  const compressed = Buffer.from(data.subarray(8, -4)); compressed[0] = 0;
  const malformed = pack(header, chunk('IDAT', compressed), end);
  assert.equal(await validatePngReplacement(malformed, 100000, async () => {}), false);
  assert.throws(() => inflateSync(compressed));
});

test('PNG stream checks zlib checksum, exact scanline lengths, filters and split IDAT', async () => {
  const rows = Buffer.alloc(4 * (1 + 8 * 3));
  const valid = deflateSync(rows), middle = Math.floor(valid.length / 2);
  assert.equal(await validatePngReplacement(pack(header, chunk('IDAT', valid.subarray(0, middle)), chunk('IDAT', valid.subarray(middle)), end), 100000, async () => {}), true);
  const badChecksum = Buffer.from(valid); badChecksum[badChecksum.length - 1] ^= 1;
  const invalidFilter = Buffer.from(rows); invalidFilter[25] = 5;
  for (const encoded of [badChecksum, deflateSync(rows.subarray(0, rows.length - 1)), deflateSync(Buffer.concat([rows, Buffer.from([0])])), deflateSync(invalidFilter), Buffer.concat([valid.subarray(0, -4), Buffer.from([0]), valid.subarray(-4)])]) {
    assert.equal(await validatePngReplacement(pack(header, chunk('IDAT', encoded), end), 100000, async () => {}), false);
  }
});

test('PNG stream validation handles Adam7 pass geometry and bounds expanded work', async () => {
  const ihdr = Buffer.from(header.subarray(8, -4)); ihdr[12] = 1;
  const adam7 = [[0,0,8,8],[4,0,8,8],[0,4,4,8],[2,0,4,4],[0,2,2,4],[1,0,2,2],[0,1,1,2]];
  let size = 0;
  for (const [x,y,dx,dy] of adam7) { const w = 8 <= x ? 0 : Math.ceil((8-x)/dx), h = 4 <= y ? 0 : Math.ceil((4-y)/dy); if (w && h) size += h * (1 + w * 3); }
  const input = pack(chunk('IHDR', ihdr), chunk('IDAT', deflateSync(Buffer.alloc(size))), end);
  assert.equal(await validatePngReplacement(input, 100000, async () => {}), true);
  await assert.rejects(validatePngReplacement(input, 100000, async () => {}, { work: 0 }, size - 1), error => error.code === 'save-limit-exceeded');
  const huge = Buffer.from(ihdr); huge.writeUInt32BE(0x7fffffff, 0); huge.writeUInt32BE(0x7fffffff, 4);
  await assert.rejects(validatePngReplacement(pack(chunk('IHDR', huge), data, end), 100000, async () => {}, { work: 0 }, 1024), error => error.code === 'save-limit-exceeded');
});

test('PNG deflate references cannot exceed the zlib-advertised window', async () => {
  const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(1024, 0); ihdr.writeUInt32BE(1, 4); ihdr[12] = 0;
  const row = Buffer.alloc(1 + 1024 * 3); let state = 0x12345678;
  for (let i = 1; i <= 1536; i++) { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; row[i] = state & 255; }
  row.copy(row, 1537, 1, 1537);
  const encoded = deflateSync(row), limitedWindow = Buffer.from(encoded); limitedWindow[0] = 0x08; limitedWindow[1] = 0x1d;
  assert.equal(await validatePngReplacement(pack(chunk('IHDR', ihdr), chunk('IDAT', encoded), end), 100000, async () => {}), true);
  assert.equal(await validatePngReplacement(pack(chunk('IHDR', ihdr), chunk('IDAT', limitedWindow), end), 100000, async () => {}), false);
});

test('indexed PNG rejects palette indices outside the declared palette', async () => {
  const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(1, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = 8; ihdr[9] = 3; ihdr[12] = 0;
  const pixels = Buffer.from([0, 1]), image = pack(chunk('IHDR', ihdr), chunk('PLTE', Buffer.from([0, 0, 0])), chunk('IDAT', deflateSync(pixels)), end);
  assert.equal(await validatePngReplacement(image, 100000, async () => {}), false);
  assert.ok(pixels[1] >= 1, 'the sole palette has one entry and the stored index is outside it');
});

test('indexed PNG restores every filter and checks actual pixels at every packed depth', async () => {
  for (const depth of [1, 2, 4, 8]) {
    const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(11, 0); ihdr.writeUInt32BE(5, 4); ihdr[8] = depth; ihdr[9] = 3; ihdr[12] = 0;
    const palette = depth === 1 ? Buffer.from([0, 0, 0]) : Buffer.from([0, 0, 0, 255, 255, 255]);
    const create = rows => pack(chunk('IHDR', ihdr), chunk('PLTE', palette), chunk('IDAT', deflateSync(filterRows(rows))), end);
    assert.equal(await validatePngReplacement(create(indexedRows(11, 5, depth)), 100000, async () => {}), true, `valid depth ${depth}`);
    assert.equal(await validatePngReplacement(create(indexedRows(11, 5, depth, 4, depth === 1 ? 1 : 2)), 100000, async () => {}), false, `invalid depth ${depth}`);
  }
});

test('indexed Adam7 resets filter history per pass and rejects an invalid pass pixel', async () => {
  const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(11, 0); ihdr.writeUInt32BE(9, 4); ihdr[8] = 2; ihdr[9] = 3; ihdr[12] = 1;
  const geometry = [[0,0,8,8],[4,0,8,8],[0,4,4,8],[2,0,4,4],[0,2,2,4],[1,0,2,2],[0,1,1,2]];
  const create = invalid => {
    const passes = geometry.flatMap(([x,y,dx,dy], index) => { const w = 11 <= x ? 0 : Math.ceil((11-x)/dx), h = 9 <= y ? 0 : Math.ceil((9-y)/dy); return w && h ? [filterRows(indexedRows(w, h, 2, invalid && index === 5 ? h - 1 : -1))] : []; });
    return pack(chunk('IHDR', ihdr), chunk('PLTE', Buffer.from([0,0,0,255,255,255])), chunk('IDAT', deflateSync(Buffer.concat(passes))), end);
  };
  assert.equal(await validatePngReplacement(create(false), 100000, async () => {}), true);
  assert.equal(await validatePngReplacement(create(true), 100000, async () => {}), false);
});

test('scanline lengths honor PNG channel counts and packed or 16-bit samples', async () => {
  for (const [color, depth, channels] of [[0, 1, 1], [0, 16, 1], [2, 16, 3], [3, 1, 1], [4, 16, 2], [6, 16, 4]]) {
    const ihdr = Buffer.from(header.subarray(8, -4)); ihdr.writeUInt32BE(3, 0); ihdr.writeUInt32BE(2, 4); ihdr[8] = depth; ihdr[9] = color; ihdr[12] = 0;
    const rows = Buffer.alloc(2 * (1 + Math.ceil(3 * channels * depth / 8))), palette = color === 3 ? [chunk('PLTE', Buffer.from([0, 0, 0, 255, 255, 255]))] : [];
    const image = pack(chunk('IHDR', ihdr), ...palette, chunk('IDAT', deflateSync(rows)), end);
    assert.equal(await validatePngReplacement(image, 100000, async () => {}), true, `color ${color}, depth ${depth}`);
    const shortened = pack(chunk('IHDR', ihdr), ...palette, chunk('IDAT', deflateSync(rows.subarray(0, -1))), end);
    assert.equal(await validatePngReplacement(shortened, 100000, async () => {}), false, `short color ${color}, depth ${depth}`);
  }
});

test('public picture save blocks corrupt PNG and permits retry with valid bytes', async () => {
  const source = encodeImages(), original = Buffer.from(source), session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
  try {
    const slide = session.document.slides[0], image = slide.elements[0], broken = Buffer.from(png); broken[29] ^= 1;
    session.document.registerAsset({ id: 'new', kind: 'image', source: { type: 'bytes', value: broken }, mimeType: 'image/png', width: 8, height: 4 });
    slide.replaceElement(image.id, { ...image, assetId: 'new' });
    const rejected = await session.prepareSave({ mode: 'edit' }); assert.equal(rejected.status, 'blocked'); assert.equal(rejected.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const badStream = Buffer.from(data.subarray(8, -4)); badStream[0] = 0;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: pack(header, chunk('IDAT', badStream), end) }, mimeType: 'image/png', width: 8, height: 4 });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    const indexed = Buffer.from(header.subarray(8, -4)); indexed.writeUInt32BE(1, 0); indexed.writeUInt32BE(1, 4); indexed[8] = 8; indexed[9] = 3; indexed[12] = 0;
    const badPalette = pack(chunk('IHDR', indexed), chunk('PLTE', Buffer.from([0, 0, 0])), chunk('IDAT', deflateSync(Buffer.from([0, 1]))), end);
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: badPalette }, mimeType: 'image/png', width: 1, height: 1 });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready'); assert.deepEqual(source, original);
  } finally { session.dispose(); }
});

test('PNG inspection and subsequent package comparison share the save work budget', async () => {
  const usage = { work: 0 }; assert.equal(await validatePngReplacement(png, 100000, async () => {}, usage), true);
  const session = await importPptx(encodeImages(), { projection: 'content', pictureEditSaveProfile: 'replace-v1', sessionLimits: { maxSaveValidationWork: usage.work } });
  try {
    const slide = session.document.slides[0], image = slide.elements[0];
    session.document.registerAsset({ id: 'new', kind: 'image', source: { type: 'bytes', value: png }, mimeType: 'image/png', width: 8, height: 4 });
    slide.replaceElement(image.id, { ...image, assetId: 'new' });
    await assert.rejects(session.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded');
  } finally { session.dispose(); }
});
