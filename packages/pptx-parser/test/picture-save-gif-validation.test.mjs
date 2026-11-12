import test from 'node:test';
import assert from 'node:assert/strict';
import { validateGifReplacement } from '../dist/import/validate-gif-replacement.js';
import { importPptx } from '../dist/index.js';
import { encodeImages, gif, validGif } from './image-fixture.mjs';

const inspect = (bytes, maxWork = 100000, maxDecoded = 100000) => validateGifReplacement(bytes, maxWork, maxDecoded, { work: 0 }, async () => {});
const pictureOffset = validGif.indexOf(0x2c), gceOffset = validGif.indexOf(Buffer.from([0x21, 0xf9, 0x04]));
function codeStream(blocks) { return Buffer.concat([validGif.subarray(0, pictureOffset + 11), ...blocks.flatMap(block => [Buffer.from([block.length]), block]), Buffer.from([0, 0x3b])]); }

test('GIF validation accepts complete LZW streams, split sub-blocks, local palettes and repeated frames', async () => {
  assert.equal(await inspect(validGif), true);
  assert.equal(await inspect(codeStream([Buffer.from([0x44]), Buffer.from([0x01])])), true);
  const local = Buffer.concat([validGif.subarray(0, pictureOffset + 10), Buffer.from([0,0,0,255,255,255]), validGif.subarray(pictureOffset + 10)]); local[pictureOffset + 9] = 0x80;
  assert.equal(await inspect(local), true);
  const image = validGif.subarray(pictureOffset, validGif.length - 1), repeated = Buffer.concat([validGif.subarray(0, validGif.length - 1), image, Buffer.from([0x3b])]);
  assert.equal(await inspect(repeated), true);
  await assert.rejects(inspect(repeated, 100000, 1), error => error.code === 'save-limit-exceeded');
});

test('GIF inspection propagates cancellation during a large extension', async () => {
  const comment = Buffer.concat([Buffer.from([0x21, 0xfe]), ...Array.from({ length: 130 }, () => Buffer.concat([Buffer.from([255]), Buffer.alloc(255)])), Buffer.from([0])]);
  const input = Buffer.concat([validGif.subarray(0, pictureOffset), comment, validGif.subarray(pictureOffset)]);
  let calls = 0; const cancelled = new Error('cancelled');
  await assert.rejects(validateGifReplacement(input, 100000, 100000, { work: 0 }, async () => { if (++calls === 2) throw cancelled; }), error => error === cancelled);
  assert.equal(calls, 2);
});

test('GIF validation rejects missing end code, bad LZW, palette indices and malformed blocks', async () => {
  const badPixel = codeStream([Buffer.from([0x54, 0x01])]), badCode = codeStream([Buffer.from([0x74, 0x01])]), trailingCode = codeStream([Buffer.from([0x44, 0x01, 0x00])]);
  const badTransparent = Buffer.from(validGif); badTransparent[gceOffset + 6] = 2;
  const badDisposal = Buffer.from(validGif); badDisposal[gceOffset + 3] = 0x10;
  const badRectangle = Buffer.from(validGif); badRectangle.writeUInt16LE(2, pictureOffset + 5);
  const badMinimum = Buffer.from(validGif); badMinimum[pictureOffset + 10] = 1;
  for (const bytes of [gif, badPixel, badCode, trailingCode, badTransparent, badDisposal, badRectangle, badMinimum, validGif.subarray(0, -1), Buffer.concat([validGif, Buffer.from([0])])]) assert.equal(await inspect(bytes), false, bytes.toString('hex'));
  await assert.rejects(inspect(validGif, 1), error => error.code === 'save-limit-exceeded');
  await assert.rejects(inspect(validGif, 100000, 0), error => error.code === 'save-limit-exceeded');
});

test('picture save blocks an incomplete GIF and retries with a complete LZW stream', async () => {
  const source = encodeImages(), session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' });
  try {
    const slide = session.document.slides[0], image = slide.elements[0];
    session.document.registerAsset({ id: 'new', kind: 'image', source: { type: 'bytes', value: gif }, mimeType: 'image/gif', width: 1, height: 1 });
    slide.replaceElement(image.id, { ...image, assetId: 'new' });
    const blocked = await session.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.report.diagnostics[0].code, 'save-picture-candidate-unsupported');
    session.document.replaceAsset('new', { kind: 'image', source: { type: 'bytes', value: validGif }, mimeType: 'image/gif', width: 1, height: 1 });
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready');
  } finally { session.dispose(); }
});
