import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { reorderXmlSiblings } from '../dist/preservation/reorder-xml-siblings.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const options = { maxOutputBytes: 1_000_000, maxWorkUnits: 4_000_000, maxElements: 10 };
async function identities(bytes, namespaceUri = 'urn:test', local = 's') {
  const source = await captureXmlTextSource(bytes);
  try { const elements = source.elements.filter(item => item.namespaceUri === namespaceUri && item.local === local); return { parentElementIndex: elements[0].parentElementIndex, namespaceUri, local, elementIndexes: elements.map(item => item.elementIndex), requestedIndexes: elements.map(item => item.elementIndex).reverse() }; } finally { source.dispose(); }
}
const encode16 = (text, little) => { const bytes = new Uint8Array(2 + text.length * 2); bytes.set(little ? [255, 254] : [254, 255]); for (let i = 0; i < text.length; i++) { const unit = text.charCodeAt(i); bytes[2 + i * 2 + (little ? 0 : 1)] = unit & 255; bytes[2 + i * 2 + (little ? 1 : 0)] = unit >>> 8; } return bytes; };

test('reorders validated siblings while retaining lexical content, whitespace and encoding', async () => {
  for (const encoding of ['utf-8', 'utf-16le', 'utf-16be']) {
    const declaration = encoding === 'utf-8' ? 'UTF-8' : 'UTF-16', xml = `<?xml version="1.0" encoding="${declaration}"?><r xmlns="urn:test"><s a='1'>中文&amp;😀</s> \n <s><x/></s><z/></r>`;
    const bytes = encoding === 'utf-8' ? new TextEncoder().encode(xml) : encode16(xml, encoding === 'utf-16le');
    const request = await identities(bytes), result = await reorderXmlSiblings(bytes, request, options);
    const decoded = new TextDecoder(encoding).decode(result.bytes);
    assert.equal(decoded, xml.replace("<s a='1'>中文&amp;😀</s> \n <s><x/></s>", "<s><x/></s> \n <s a='1'>中文&amp;😀</s>"));
    assert.equal(result.bytes.length, bytes.length); assert.ok(result.workUnits > bytes.length);
    assert.deepEqual([...result.movedElementIndexes], [[request.elementIndexes[1], request.elementIndexes[0]], [request.elementIndexes[1] + 1, request.elementIndexes[0] + 1], [request.elementIndexes[0], request.elementIndexes[0] + 2]]);
  }
});

test('reorders the PowerPoint control by exact element bytes', async () => {
  const read = async name => readGeneratedZip(await readFile(new URL(`../../../tests/fixtures/${name}.pptx`, import.meta.url))).get('ppt/charts/chart1.xml');
  const source = await read('chart-label-source'), expected = await read('chart-order-powerpoint-move-only');
  const request = await identities(source, 'http://schemas.openxmlformats.org/drawingml/2006/chart', 'ser');
  assert.deepEqual((await reorderXmlSiblings(source, request, options)).bytes, Uint8Array.from(expected));
});

test('rejects unrelated or ambiguous siblings and enforces limits and cancellation', async () => {
  const bytes = new TextEncoder().encode('<r xmlns="urn:test"><s>A</s> <s>B</s></r>'), request = await identities(bytes);
  for (const bad of [{ ...request, parentElementIndex: 99 }, { ...request, namespaceUri: 'wrong' }, { ...request, elementIndexes: [...request.elementIndexes].reverse() }]) await assert.rejects(reorderXmlSiblings(bytes, bad, options), { code: 'xml-edit-conflict' });
  for (const bad of [{ ...request, requestedIndexes: [request.elementIndexes[0], request.elementIndexes[0]] }, { ...request, requestedIndexes: [1, 99] }]) await assert.rejects(reorderXmlSiblings(bytes, bad, options), TypeError);
  const separated = new TextEncoder().encode('<r xmlns="urn:test"><s>A</s><!-- belongs? --><s>B</s></r>');
  await assert.rejects(reorderXmlSiblings(separated, await identities(separated), options), { code: 'xml-edit-unsupported' });
  await assert.rejects(reorderXmlSiblings(bytes, request, { ...options, maxWorkUnits: 1 }), { code: 'xml-edit-limit-exceeded' });
  await assert.rejects(reorderXmlSiblings(bytes, request, { ...options, maxOutputBytes: bytes.length - 1 }), { code: 'xml-edit-limit-exceeded' });
  const controller = new AbortController(); controller.abort(); await assert.rejects(reorderXmlSiblings(bytes, request, { ...options, signal: controller.signal }), { code: 'operation-aborted' });
  const mutable = structuredClone(request), pending = reorderXmlSiblings(bytes, mutable, options); mutable.requestedIndexes.reverse();
  assert.match(new TextDecoder().decode((await pending).bytes), /<s>B<\/s> <s>A<\/s>/);
});
