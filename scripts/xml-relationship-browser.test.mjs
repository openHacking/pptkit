import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { editXmlTextAttribute, replaceXmlNumericAttribute, replaceXmlRelationshipId } from '../packages/pptx-parser/dist/preservation/replace-xml-relationship-id.js';

test('browser relationship-ID lexical editing matches Node and independent DOM', async t => {
  const source = `<root xmlns:a="urn:drawing" xmlns:r="urn:relations"><!-- x --><a:blip r:embed='r&#73;d1'/><a:blip r:embed="rId1"/></root>`;
  const request = { elementIndex: 2, elementNamespace: 'urn:drawing', elementLocal: 'blip', attributeNamespace: 'urn:relations', attributeLocal: 'embed', expected: 'rId1', replacement: 'rId_new' };
  const bytes = new TextEncoder().encode(source), expected = await replaceXmlRelationshipId(bytes, request, { maxOutputBytes: 1024 });
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/preservation/replace-xml-relationship-id.js', plugins: [nodeResolve({ browser: true }), commonjs()] }); t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'XmlRelationshipTest' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(async ({ input, request }) => {
    const bytes = new Uint8Array(input), pending = XmlRelationshipTest.replaceXmlRelationshipId(bytes, request, { maxOutputBytes: 1024 });
    bytes.fill(0); request.replacement = 'wrong';
    const result = await pending, text = new TextDecoder().decode(result), dom = new DOMParser().parseFromString(text, 'application/xml');
    return { bytes: Array.from(result), values: Array.from(dom.getElementsByTagNameNS('urn:drawing', 'blip'), node => node.getAttributeNS('urn:relations', 'embed')) };
  }, { input: Array.from(bytes), request });
  assert.deepEqual(new Uint8Array(actual.bytes), expected);
  assert.deepEqual(actual.values, ['rId_new', 'rId1']);
  const numericSource = `<p:presentation xmlns:p="urn:p" firstSlideNum='+0&#49;'><p:other firstSlideNum="1"/></p:presentation>`;
  const numericRequest = { elementIndex: 1, elementNamespace: 'urn:p', elementLocal: 'presentation', attributeNamespace: '', attributeLocal: 'firstSlideNum', expected: '+01', replacement: '9999' };
  const numericBytes = new TextEncoder().encode(numericSource), numericExpected = await replaceXmlNumericAttribute(numericBytes, numericRequest, { maxOutputBytes: 1024 });
  const numericActual = await page.evaluate(async ({ input, request }) => {
    const bytes = await XmlRelationshipTest.replaceXmlNumericAttribute(new Uint8Array(input), request, { maxOutputBytes: 1024 });
    const dom = new DOMParser().parseFromString(new TextDecoder().decode(bytes), 'application/xml');
    return { bytes: Array.from(bytes), root: dom.documentElement.getAttribute('firstSlideNum'), child: dom.documentElement.firstElementChild.getAttribute('firstSlideNum') };
  }, { input: Array.from(numericBytes), request: numericRequest });
  assert.deepEqual(new Uint8Array(numericActual.bytes), numericExpected); assert.equal(numericActual.root, '9999'); assert.equal(numericActual.child, '1');
  const textSource = `<root xmlns:p="urn:p"><p:cNvPr descr='old'/><p:cNvPr descr="neighbor"/></root>`;
  const textRequest = { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'cNvPr', attributeNamespace: '', attributeLocal: 'descr', expected: 'old', replacement: `中😀 & < > " ' \t\n\r` };
  const textBytes = new TextEncoder().encode(textSource), textExpected = await editXmlTextAttribute(textBytes, textRequest, { maxOutputBytes: 1024 });
  const textActual = await page.evaluate(async ({ input, request }) => {
    const result = await XmlRelationshipTest.editXmlTextAttribute(new Uint8Array(input), request, { maxOutputBytes: 1024 });
    const dom = new DOMParser().parseFromString(new TextDecoder().decode(result), 'application/xml');
    return { bytes: Array.from(result), values: Array.from(dom.getElementsByTagNameNS('urn:p', 'cNvPr'), node => node.getAttribute('descr')) };
  }, { input: Array.from(textBytes), request: textRequest });
  assert.deepEqual(new Uint8Array(textActual.bytes), textExpected); assert.deepEqual(textActual.values, [textRequest.replacement, 'neighbor']);
});
