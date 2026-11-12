import test from 'node:test';
import assert from 'node:assert/strict';
import { appendImageContentType } from '../dist/preservation/append-image-content-type.js';

const ns = 'http://schemas.openxmlformats.org/package/2006/content-types';
const request = { part: '/ppt/media/new.png', mimeType: 'image/png' }, options = { maxOutputBytes: 4096 };
const be = text => { const bytes = Buffer.from(text, 'utf16le'); for (let i = 0; i < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1], bytes[i]]; return Buffer.concat([Buffer.from([0xfe, 0xff]), bytes]); };
const encodings = [text => Buffer.from(text), text => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)]), text => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]), be];

test('explicit image override retains original content types byte regions and encodings', async () => {
  const xml = `<c:Types xmlns:c="${ns}"><!-- </c:Types> --><c:Default Extension="png" ContentType="application/octet-stream"/></c:Types><!-- tail -->`;
  const expected = xml.replace('</c:Types><!-- tail -->', '<c:Override PartName="/ppt/media/new.png" ContentType="image/png"/></c:Types><!-- tail -->');
  for (const encode of encodings) { const bytes = encode(xml), original = Buffer.from(bytes); assert.deepEqual(Buffer.from(await appendImageContentType(bytes, request, options)), encode(expected)); assert.deepEqual(bytes, original); }
  const empty = `<Types xmlns="${ns}"/>`;
  assert.equal(Buffer.from(await appendImageContentType(Buffer.from(empty), request, options)).toString(), `<Types xmlns="${ns}"><Override PartName="/ppt/media/new.png" ContentType="image/png"/></Types>`);
});

test('image content type overrides reject source aliases, unsafe names and budgets', async () => {
  const duplicate = `<Types xmlns="${ns}"><Override PartName="/PPT/media/NEW.png" ContentType="image/jpeg"/></Types>`;
  const code = expected => error => error.code === expected;
  await assert.rejects(appendImageContentType(Buffer.from(duplicate), request, options), code('xml-edit-conflict'));
  const input = Buffer.from(`<Types xmlns="${ns}"/>`);
  for (const part of ['ppt/media/x.png', '/ppt/../x.png', '/ppt/a%20b.png', '/ppt/a&b.png', '/ppt/x.', '/ppt//x.png']) await assert.rejects(appendImageContentType(input, { ...request, part }, options), code('xml-edit-unsupported'));
  await assert.rejects(appendImageContentType(input, { ...request, mimeType: 'text/xml' }, options), TypeError);
  await assert.rejects(appendImageContentType(input, request, { maxOutputBytes: 2 }), code('xml-edit-limit-exceeded'));
  const controller = new AbortController(); controller.abort(); await assert.rejects(appendImageContentType(input, request, { ...options, signal: controller.signal }), code('operation-aborted'));
});

test('content type request and bytes are captured before asynchronous parsing', async () => {
  const input = Buffer.from(`<Types xmlns="${ns}"/>`), change = { ...request }, pending = appendImageContentType(input, change, options);
  input.fill(0); change.part = '/wrong.png';
  assert.match(Buffer.from(await pending).toString(), /PartName="\/ppt\/media\/new.png"/);
});
