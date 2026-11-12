import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { appendImageRelationship } from '../packages/pptx-parser/dist/preservation/append-image-relationship.js';
import { appendImageContentType } from '../packages/pptx-parser/dist/preservation/append-image-content-type.js';

test('browser image relationship insertion matches Node and independent DOM', async t => {
  const ns = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const source = new TextEncoder().encode(`<r:Relationships xmlns:r="${ns}"><!-- keep --></r:Relationships>`);
  const request = { id: 'rId2', target: '../media/new.png', dialect: 'strict' };
  const expected = await appendImageRelationship(source, request, { maxOutputBytes: 1024 });
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/preservation/append-image-relationship.js', plugins: [nodeResolve({ browser: true }), commonjs()] }); t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'ImageRelationshipTest' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(async ({ source, request, ns }) => {
    const input = new Uint8Array(source), pending = ImageRelationshipTest.appendImageRelationship(input, request, { maxOutputBytes: 1024 });
    input.fill(0); request.id = 'changed';
    const bytes = await pending, doc = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
    const node = doc.getElementsByTagNameNS(ns, 'Relationship')[0];
    return { bytes: Array.from(bytes), id: node.getAttribute('Id'), type: node.getAttribute('Type'), target: node.getAttribute('Target') };
  }, { source: Array.from(source), request, ns });
  assert.deepEqual(new Uint8Array(actual.bytes), expected);
  assert.deepEqual({ id: actual.id, type: actual.type, target: actual.target }, { id: 'rId2', type: 'http://purl.oclc.org/ooxml/officeDocument/relationships/image', target: '../media/new.png' });
});

test('browser image content type override matches Node and independent DOM', async t => {
  const ns = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const source = new TextEncoder().encode(`<Types xmlns="${ns}"><!-- keep --></Types>`), request = { part: '/ppt/media/new.png', mimeType: 'image/png' };
  const expected = await appendImageContentType(source, request, { maxOutputBytes: 1024 });
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/preservation/append-image-content-type.js', plugins: [nodeResolve({ browser: true }), commonjs()] }); t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'ImageContentTypeTest' });
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(); await page.addScriptTag({ content: output[0].code });
  const result = await page.evaluate(async ({ source, request, ns }) => {
    const input = new Uint8Array(source), pending = ImageContentTypeTest.appendImageContentType(input, request, { maxOutputBytes: 1024 }); input.fill(0); request.part = '/wrong.png';
    const bytes = await pending, doc = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml'), node = doc.getElementsByTagNameNS(ns, 'Override')[0];
    return { bytes: Array.from(bytes), part: node.getAttribute('PartName'), mime: node.getAttribute('ContentType') };
  }, { source: Array.from(source), request, ns });
  assert.deepEqual(new Uint8Array(result.bytes), expected); assert.equal(result.part, request.part); assert.equal(result.mime, request.mimeType);
});
