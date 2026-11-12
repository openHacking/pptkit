import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { validateJpegReplacement } from '../dist/import/validate-jpeg-replacement.js';
import { importPptx } from '../dist/index.js';
import { encodeImages, jpegHeader } from './image-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const jpeg = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
const progressive = await readFile(new URL('../../../tests/fixtures/jpeg-progressive-components.jpg', import.meta.url));
// libjpeg-turbo cjpeg 3.1.3 encoded exif-orientation6 as RGB PPM, then -scans "0; 1; 2;".
// SHA-256 ba51b04c6b705518b5858aa889893e9e9ac513b1a3277442a5d6cd9afb3a7b66; Pillow independently decodes it.
const sequential = await readFile(new URL('../../../tests/fixtures/jpeg-sequential-multiscan.jpg', import.meta.url));
// libjpeg-turbo cjpeg 3.1.3: the same RGB PPM, -precision 12, with and without -progressive.
// SHA-256: sequential 4cbd30ea571652ace5a2450d0234dadf7067c3d91da59e750f7e1933a2115117;
// progressive 5b096a41566f899b48aa258689379198dc208fe7128df75655f7e8b5908b9966.
const sequential12 = await readFile(new URL('../../../tests/fixtures/jpeg-12bit-sequential.jpg', import.meta.url));
const progressive12 = await readFile(new URL('../../../tests/fixtures/jpeg-12bit-progressive.jpg', import.meta.url));
const inspect = (bytes, maxWork = 10000000) => validateJpegReplacement(bytes, maxWork, { work: 0 }, async () => {});

test('JPEG marker inspection rejects incomplete segments, scans and trailing bytes', async () => {
  assert.equal(await inspect(jpeg), true);
  assert.equal(await inspect(jpegHeader), false);
  for (const value of [jpeg.subarray(0, -1), jpeg.subarray(0, -2), Buffer.concat([jpeg, Buffer.from([0])]), Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 8, 1])]) assert.equal(await inspect(value), false);
  await assert.rejects(inspect(jpeg, 1), error => error.code === 'save-limit-exceeded');
});

test('JPEG rejects malformed quantization and Huffman table definitions', async () => {
  const dqt = jpeg.indexOf(Buffer.from([0xff, 0xdb])), dht = jpeg.indexOf(Buffer.from([0xff, 0xc4]));
  assert.ok(dqt >= 0 && dht >= 0);
  const change = (offset, value) => { const copy = Buffer.from(jpeg); copy[offset] = value; return copy; };
  for (const bad of [change(dqt + 4, 0x20), change(dqt + 4, 0x04), change(dqt + 5, 0), change(dqt + 3, 3), change(dht + 4, 0x20), change(dht + 4, 0x04), change(dht + 5, 3), change(dht + 3, 3)]) {
    assert.equal(await inspect(bad), false);
  }
});

test('JPEG scans reference distinct components declared by the frame', async () => {
  const frame = jpeg.indexOf(Buffer.from([0xff, 0xc0])), scan = jpeg.indexOf(Buffer.from([0xff, 0xda]));
  assert.ok(frame >= 0 && scan >= 0); assert.ok(jpeg[frame + 9] >= 2 && jpeg[scan + 4] >= 2);
  const duplicateFrame = Buffer.from(jpeg); duplicateFrame[frame + 13] = duplicateFrame[frame + 10];
  const unknownScan = Buffer.from(jpeg); unknownScan[scan + 5] = 255;
  const duplicateScan = Buffer.from(jpeg); duplicateScan[scan + 7] = duplicateScan[scan + 5];
  for (const candidate of [duplicateFrame, unknownScan, duplicateScan]) assert.equal(await inspect(candidate), false);
  assert.equal(await inspect(jpeg), true);
  assert.ok(progressive.indexOf(Buffer.from([0xff, 0xda])) !== progressive.lastIndexOf(Buffer.from([0xff, 0xda])));
  assert.equal(await inspect(progressive), true, 'components may recur across separate progressive scans');
});

test('sequential JPEG covers every frame component once across interleaved or separate scans', async () => {
  const positions = []; let at = -2;
  while ((at = sequential.indexOf(Buffer.from([0xff, 0xda]), at + 2)) !== -1) positions.push(at);
  assert.equal(positions.length, 3);
  assert.deepEqual(positions.map(p => sequential[p + 4]), [1, 1, 1]);
  const incomplete = Buffer.concat([sequential.subarray(0, positions[2]), Buffer.from([0xff, 0xd9])]);
  const repeated = Buffer.from(sequential); repeated[positions[1] + 5] = repeated[positions[0] + 5];
  for (const marker of [0xc0, 0xc1]) {
    const changeFrame = source => { const copy = Buffer.from(source); copy[copy.indexOf(Buffer.from([0xff, 0xc0])) + 1] = marker; return copy; };
    assert.equal(await inspect(changeFrame(sequential)), true);
    assert.equal(await inspect(changeFrame(incomplete)), false);
    assert.equal(await inspect(changeFrame(repeated)), false);
  }
  assert.equal(await inspect(jpeg), true);
  assert.equal(await inspect(progressive), true, 'progressive scans legitimately repeat components');
});

test('JPEG validates frame sampling and referenced quantization tables for common DCT frames', async () => {
  const frame = jpeg.indexOf(Buffer.from([0xff, 0xc0])); assert.ok(frame >= 0);
  const change = (offset, value) => { const copy = Buffer.from(jpeg); copy[offset] = value; return copy; };
  for (const invalid of [change(frame + 11, 0x01), change(frame + 11, 0x10), change(frame + 11, 0x51), change(frame + 11, 0x15), change(frame + 12, 4), change(frame + 12, 3)]) {
    assert.equal(await inspect(invalid), false);
  }
  assert.equal(await inspect(jpeg), true); assert.equal(await inspect(progressive), true);
  const dqt = jpeg.indexOf(Buffer.from([0xff, 0xdb])), dqtEnd = dqt + 2 + jpeg.readUInt16BE(dqt + 2);
  const frameEnd = frame + 2 + jpeg.readUInt16BE(frame + 2);
  assert.ok(dqtEnd < frame);
  const deferredTable = Buffer.concat([jpeg.subarray(0, dqt), jpeg.subarray(dqtEnd, frameEnd), jpeg.subarray(dqt, dqtEnd), jpeg.subarray(frameEnd)]);
  assert.equal(await inspect(deferredTable), true, 'quantization table may appear after the frame but before its scan');
  const tooLate = Buffer.concat([jpeg.subarray(0, dqt), jpeg.subarray(dqtEnd, -2), jpeg.subarray(dqt, dqtEnd), jpeg.subarray(-2)]);
  assert.equal(await inspect(tooLate), false, 'a table defined after its first scan is too late');
});

test('JPEG DCT sample precision matches baseline, extended sequential and progressive frames', async () => {
  for (const [bytes, marker, legal] of [[jpeg, 0xc0, [8]], [sequential12, 0xc1, [8, 12]], [progressive12, 0xc2, [8, 12]]]) {
    const frame = bytes.indexOf(Buffer.from([0xff, marker])); assert.ok(frame >= 0);
    assert.equal(await inspect(bytes), true);
    for (const precision of [0, 1, 7, 9, 11, 13, 16, ...(marker === 0xc0 ? [12] : [])]) {
      const candidate = Buffer.from(bytes); candidate[frame + 4] = precision;
      assert.equal(await inspect(candidate), false, `${marker.toString(16)} precision ${precision}`);
    }
    assert.ok(legal.includes(bytes[frame + 4]));
  }
});

test('baseline JPEG scans require declared DC and AC Huffman tables', async () => {
  const scan = jpeg.indexOf(Buffer.from([0xff, 0xda])), dht = jpeg.indexOf(Buffer.from([0xff, 0xc4]));
  assert.ok(scan > dht && dht >= 0); const dhtEnd = dht + 2 + jpeg.readUInt16BE(dht + 2);
  const missing = Buffer.concat([jpeg.subarray(0, dht), jpeg.subarray(dhtEnd)]);
  assert.equal(await inspect(missing), false);
  const badDc = Buffer.from(jpeg), badAc = Buffer.from(jpeg); badDc[scan + 6] = 0x20; badAc[scan + 6] = 0x02;
  assert.equal(await inspect(badDc), false); assert.equal(await inspect(badAc), false);
  const frame = jpeg.indexOf(Buffer.from([0xff, 0xc0])), frameEnd = frame + 2 + jpeg.readUInt16BE(frame + 2);
  assert.ok(dht >= frameEnd && dhtEnd < scan);
  const moved = Buffer.concat([jpeg.subarray(0, dht), jpeg.subarray(dhtEnd, scan), jpeg.subarray(dht, dhtEnd), jpeg.subarray(scan)]);
  assert.equal(await inspect(moved), true, 'tables may be declared immediately before the scan');
  assert.equal(await inspect(progressive), true);
});

test('extended and progressive scans require only Huffman tables used by that scan', async () => {
  const extended = Buffer.from(jpeg), frame = extended.indexOf(Buffer.from([0xff, 0xc0]));
  extended[frame + 1] = 0xc1;
  assert.equal(await inspect(extended), true);
  const extendedMissing = Buffer.from(extended), extendedScan = extendedMissing.indexOf(Buffer.from([0xff, 0xda]));
  extendedMissing[extendedScan + 6] = 0x30;
  assert.equal(await inspect(extendedMissing), false);
  const scans = []; let pos = 0;
  while ((pos = progressive.indexOf(Buffer.from([0xff, 0xda]), pos)) !== -1) {
    const count = progressive[pos + 4], fields = pos + 5 + 2 * count;
    scans.push({ pos, start: progressive[fields], high: progressive[fields + 2] >>> 4 }); pos += 2;
  }
  const initialDc = scans.find(scan => scan.start === 0 && scan.high === 0);
  const initialAc = scans.find(scan => scan.start > 0 && scan.high === 0);
  const refineDc = scans.find(scan => scan.start === 0 && scan.high > 0);
  const refineAc = scans.find(scan => scan.start > 0 && scan.high > 0);
  assert.ok(initialDc && initialAc && refineDc && refineAc);
  for (const [scan, mask] of [[initialDc, 0x30], [initialAc, 0x03], [refineAc, 0x03]]) {
    const candidate = Buffer.from(progressive); candidate[scan.pos + 6] |= mask;
    assert.equal(await inspect(candidate), false);
  }
  const unused = Buffer.from(progressive); unused[refineDc.pos + 6] = 0x33;
  assert.equal(await inspect(unused), true, 'DC refinement does not consume a Huffman table');
});

test('DCT scans reject invalid spectral and successive approximation parameters', async () => {
  const baselineScan = jpeg.indexOf(Buffer.from([0xff, 0xda]));
  const baselineFields = baselineScan + 5 + 2 * jpeg[baselineScan + 4];
  for (const [offset, value] of [[0, 1], [1, 62], [2, 1]]) {
    const invalid = Buffer.from(jpeg); invalid[baselineFields + offset] = value;
    assert.equal(await inspect(invalid), false);
  }
  const first = progressive.indexOf(Buffer.from([0xff, 0xda]));
  const second = progressive.indexOf(Buffer.from([0xff, 0xda]), first + 2);
  const firstFields = first + 5 + 2 * progressive[first + 4];
  const secondFields = second + 5 + 2 * progressive[second + 4];
  assert.equal(progressive[first + 4], 3); assert.equal(progressive[second + 4], 1);
  for (const [at, value] of [[firstFields + 1, 1], [firstFields, 1], [secondFields, 6], [secondFields + 1, 64], [secondFields + 2, 0x31], [secondFields + 2, 0x0e]]) {
    const invalid = Buffer.from(progressive); invalid[at] = value;
    assert.equal(await inspect(invalid), false);
  }
  assert.equal(await inspect(jpeg), true); assert.equal(await inspect(progressive), true);
});

test('progressive JPEG validates coefficient progression and initial DC coverage', async () => {
  const scans = []; let cursor = -2;
  while ((cursor = progressive.indexOf(Buffer.from([0xff, 0xda]), cursor + 2)) !== -1) {
    const fields = cursor + 5 + 2 * progressive[cursor + 4];
    scans.push({ at: cursor, fields, start: progressive[fields], end: progressive[fields + 1], approximation: progressive[fields + 2] });
  }
  const first = scans[0], initialAc = scans.find(s => s.start > 0 && s.approximation >>> 4 === 0);
  const refinement = scans.find(s => s.start > 0 && s.approximation >>> 4 > 0);
  assert.ok(first && initialAc && refinement); assert.equal(first.start, 0); assert.equal(progressive[first.at + 4], 3);
  const beforeInitial = Buffer.from(progressive); beforeInitial[first.fields + 2] = 0x21;
  const repeatedInitial = Buffer.from(progressive); repeatedInitial[refinement.fields + 2] &= 0x0f;
  const skippedRefinement = Buffer.from(progressive); skippedRefinement[refinement.fields + 2] = 0x32;
  const tooHigh = Buffer.from(progressive); tooHigh[refinement.fields + 2] = 0xed;
  const entropyStart = first.at + 2 + progressive.readUInt16BE(first.at + 2);
  let nextMarker = entropyStart;
  while (nextMarker < progressive.length - 1) {
    if (progressive[nextMarker] === 0xff && progressive[nextMarker + 1] !== 0 && !(progressive[nextMarker + 1] >= 0xd0 && progressive[nextMarker + 1] <= 0xd7)) break;
    nextMarker += progressive[nextMarker] === 0xff ? 2 : 1;
  }
  const duplicateDc = Buffer.concat([progressive.subarray(0, nextMarker), progressive.subarray(first.at, nextMarker), progressive.subarray(nextMarker)]);
  const noDc = Buffer.concat([progressive.subarray(0, first.at), progressive.subarray(nextMarker)]);
  const firstScanTwoComponents = Buffer.concat([progressive.subarray(0, first.at), Buffer.from([0xff, 0xda, 0, 10, 2]), progressive.subarray(first.at + 5, first.at + 9), progressive.subarray(first.fields, first.fields + 3), progressive.subarray(entropyStart, nextMarker), Buffer.from([0xff, 0xd9])]);
  for (const value of [beforeInitial, repeatedInitial, skippedRefinement, tooHigh, duplicateDc, noDc, firstScanTwoComponents]) assert.equal(await inspect(value), false);
  const dcOnly = Buffer.concat([progressive.subarray(0, nextMarker), Buffer.from([0xff, 0xd9])]);
  assert.equal(await inspect(dcOnly), true, 'untransmitted AC coefficients are permitted');
  assert.equal(await inspect(progressive), true);
  const usage = { work: 0 };
  assert.equal(await validateJpegReplacement(progressive, 10000000, usage, async () => {}), true);
  assert.equal(await inspect(progressive, usage.work), true);
  await assert.rejects(inspect(progressive, usage.work - 1), error => error.code === 'save-limit-exceeded');
});

test('JPEG inspection is cancellable within long entropy data', async () => {
  const scan = jpeg.indexOf(Buffer.from([0xff, 0xda])), length = jpeg.readUInt16BE(scan + 2), start = scan + 2 + length;
  const longScan = Buffer.concat([jpeg.subarray(0, start), Buffer.alloc(70000, 1), Buffer.from([0xff, 0xd9])]);
  const cancelled = new Error('cancelled'); let calls = 0;
  await assert.rejects(validateJpegReplacement(longScan, 1000000, { work: 0 }, async () => { if (++calls === 2) throw cancelled; }), error => error === cancelled);
});

test('picture save rejects a header-only JPEG and accepts a complete replacement after retry', async () => {
  const session = await importPptx(encodeImages(), { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
  try {
    const slide = session.document.slides[0], image = slide.elements[0];
    session.document.registerAsset({ id: 'new', kind: 'image', source: { type: 'bytes', value: jpegHeader }, mimeType: 'image/jpeg', width: 8, height: 4 });
    slide.replaceElement(image.id, { ...image, assetId: 'new' });
    const blocked = await session.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const { width, height } = await (await import('../dist/import/image-metadata.js')).rasterMetadata(jpeg, async () => {});
    const invalidTable = Buffer.from(jpeg); invalidTable[invalidTable.indexOf(Buffer.from([0xff, 0xdb])) + 5] = 0;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: invalidTable }, mimeType: 'image/jpeg', width, height });
    const invalid = await session.prepareSave({ mode: 'edit' }); assert.equal(invalid.status, 'blocked'); assert.equal(invalid.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const invalidPrecision = Buffer.from(jpeg); invalidPrecision[invalidPrecision.indexOf(Buffer.from([0xff, 0xc0])) + 4] = 12;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: invalidPrecision }, mimeType: 'image/jpeg', width, height });
    const badPrecision = await session.prepareSave({ mode: 'edit' }); assert.equal(badPrecision.status, 'blocked'); assert.equal(badPrecision.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: sequential12 }, mimeType: 'image/jpeg', width, height });
    const twelveReady = await session.prepareSave({ mode: 'edit' }); assert.equal(twelveReady.status, 'ready'); twelveReady.plan.dispose();
    const unknownComponent = Buffer.from(jpeg); unknownComponent[unknownComponent.indexOf(Buffer.from([0xff, 0xda])) + 5] = 255;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: unknownComponent }, mimeType: 'image/jpeg', width, height });
    const badScan = await session.prepareSave({ mode: 'edit' }); assert.equal(badScan.status, 'blocked'); assert.equal(badScan.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const missingTable = Buffer.from(jpeg); missingTable[missingTable.indexOf(Buffer.from([0xff, 0xc0])) + 12] = 3;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: missingTable }, mimeType: 'image/jpeg', width, height });
    const badTable = await session.prepareSave({ mode: 'edit' }); assert.equal(badTable.status, 'blocked'); assert.equal(badTable.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const missingHuffman = Buffer.from(jpeg), huffman = missingHuffman.indexOf(Buffer.from([0xff, 0xc4])), huffmanEnd = huffman + 2 + missingHuffman.readUInt16BE(huffman + 2);
    const withoutHuffman = Buffer.concat([missingHuffman.subarray(0, huffman), missingHuffman.subarray(huffmanEnd)]);
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: withoutHuffman }, mimeType: 'image/jpeg', width, height });
    const badHuffman = await session.prepareSave({ mode: 'edit' }); assert.equal(badHuffman.status, 'blocked'); assert.equal(badHuffman.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const invalidProgression = Buffer.from(progressive), firstScan = invalidProgression.indexOf(Buffer.from([0xff, 0xda]));
    invalidProgression[firstScan + 6 + 2 * invalidProgression[firstScan + 4]] = 1;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: invalidProgression }, mimeType: 'image/jpeg', width: 8, height: 8 });
    const badProgression = await session.prepareSave({ mode: 'edit' }); assert.equal(badProgression.status, 'blocked'); assert.equal(badProgression.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    const unordered = Buffer.from(progressive); unordered[firstScan + 7 + 2 * unordered[firstScan + 4]] = 0x21;
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: unordered }, mimeType: 'image/jpeg', width: 8, height: 8 });
    const badOrder = await session.prepareSave({ mode: 'edit' }); assert.equal(badOrder.status, 'blocked'); assert.equal(badOrder.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: progressive }, mimeType: 'image/jpeg', width: 8, height: 8 });
    const progressiveReady = await session.prepareSave({ mode: 'edit' }); assert.equal(progressiveReady.status, 'ready'); progressiveReady.plan.dispose();
    const lastScan = sequential.lastIndexOf(Buffer.from([0xff, 0xda]));
    const missingComponent = Buffer.concat([sequential.subarray(0, lastScan), Buffer.from([0xff, 0xd9])]);
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: missingComponent }, mimeType: 'image/jpeg', width, height });
    const badCoverage = await session.prepareSave({ mode: 'edit' }); assert.equal(badCoverage.status, 'blocked'); assert.equal(badCoverage.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: sequential }, mimeType: 'image/jpeg', width, height });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: jpeg }, mimeType: 'image/jpeg', width, height });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready');
  } finally { session.dispose(); }
});

test('sequential and 12-bit JPEG bytes survive three replacement and fresh reimport rounds', async () => {
  const original = encodeImages(), originalCopy = original.slice(), options = { projection: 'content', pictureEditSaveProfile: 'replace-v1' };
  let source = original;
  for (const [index, image] of [sequential, sequential12, progressive12].entries()) {
    const session = await importPptx(source, options);
    try {
      const slide = session.document.slides[0], target = slide.elements[0], id = `jpeg-round-${index}`;
      session.document.registerAsset({ id, kind: 'image', source: { type: 'bytes', value: image }, mimeType: 'image/jpeg', width: 3, height: 2 });
      slide.replaceElement(target.id, { ...target, assetId: id });
      const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
      source = (await session.save(prepared.plan)).bytes;
      assert.ok([...readGeneratedZip(source).values()].some(bytes => Buffer.from(bytes).equals(image)), `round ${index + 1}: JPEG bytes in output`);
    } finally { session.dispose(); }
    const reopened = await importPptx(source, options);
    try {
      assert.equal(reopened.document.slides[0].elements[0].type, 'image');
      const mediaPart = reopened.report.pictureSources[0].mediaPart;
      assert.ok(mediaPart); assert.deepEqual(readGeneratedZip(source).get(mediaPart.replace(/^\//, '')), image);
    }
    finally { reopened.dispose(); }
  }
  assert.deepEqual(original, originalCopy);
});
