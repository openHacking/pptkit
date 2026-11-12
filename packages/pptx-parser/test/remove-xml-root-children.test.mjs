import test from 'node:test';
import assert from 'node:assert/strict';
import { removeXmlRootChildren } from '../dist/preservation/remove-xml-root-children.js';

function encode(text, encoding) {
  if (encoding === 'utf-8') return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)]);
  const bytes = Buffer.from(text, 'utf16le');
  if (encoding === 'utf-16be') for (let i = 0; i < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1], bytes[i]];
  return Buffer.concat([Buffer.from(encoding === 'utf-16le' ? [0xff, 0xfe] : [0xfe, 0xff]), bytes]);
}
const root = { namespaceUri: 'urn:custom', local: 'Properties' };
const target = index => ({ elementIndex: index, namespaceUri: 'urn:custom', local: 'property' });
const options = { maxEdits: 5, maxOutputBytes: 10000 };

test('direct-child removal preserves every neighboring XML region and encoding', async () => {
  for (const encoding of ['utf-8', 'utf-16le', 'utf-16be']) {
    const decl = encoding === 'utf-8' ? 'UTF-8' : 'UTF-16';
    const prefix = `<?xml version="1.0" encoding="${decl}"?><c:Properties xmlns:c="urn:custom">`;
    const one = '<c:property id="1"><c:value>中文😀</c:value></c:property>';
    const two = '<c:property id="2"/>';
    const suffix = '<!--kept--><c:property id="3"><c:value>still</c:value></c:property></c:Properties>';
    const source = encode(prefix + one + '  ' + two + suffix, encoding), expected = encode(prefix + '  ' + suffix, encoding);
    const output = await removeXmlRootChildren(source, [target(4), target(2)], root, options);
    assert.deepEqual(Buffer.from(output), expected);
    assert.deepEqual(source, encode(prefix + one + '  ' + two + suffix, encoding));
  }
});

test('removal rejects duplicate, nested, wrong identity, limits and cancellation', async () => {
  const source = encode('<Properties xmlns="urn:custom"><property><value>x</value></property><property/></Properties>', 'utf-8');
  for (const removals of [[target(2), target(2)], [target(3)], [{ ...target(2), namespaceUri: 'urn:wrong' }], [{ ...target(2), local: 'other' }]]) {
    await assert.rejects(removeXmlRootChildren(source, removals, root, options), { code: 'xml-edit-conflict' });
  }
  await assert.rejects(removeXmlRootChildren(source, [target(2), target(4)], root, { ...options, maxEdits: 1 }), { code: 'xml-edit-limit-exceeded' });
  await assert.rejects(removeXmlRootChildren(source, [target(2)], root, { ...options, maxOutputBytes: 1 }), { code: 'xml-edit-limit-exceeded' });
  await assert.rejects(removeXmlRootChildren(source, [target(2)], { ...root, local: 'Wrong' }, options), { code: 'xml-edit-conflict' });
  const controller = new AbortController(); controller.abort();
  await assert.rejects(removeXmlRootChildren(source, [target(2)], root, { ...options, signal: controller.signal }), { code: 'operation-aborted' });
});


test('removal captures input and options before yielding and aborts long work', async () => {
  const text = '<Properties xmlns="urn:custom"><property/><keep>' + 'x'.repeat(80000) + '</keep></Properties>';
  const source = encode(text, 'utf-8'), expected = encode(text.replace('<property/>', ''), 'utf-8');
  const mutableOptions = { ...options, maxOutputBytes: 200000, limits: { maxInputBytes: 200000 } };
  const promise = removeXmlRootChildren(source, [target(2)], root, mutableOptions);
  source.fill(0); mutableOptions.maxOutputBytes = 1; mutableOptions.limits.maxInputBytes = 1;
  assert.deepEqual(Buffer.from(await promise), expected);
  const controller = new AbortController(), large = encode(text, 'utf-8');
  const pending = removeXmlRootChildren(large, [target(2)], root, { ...options, maxOutputBytes: 200000, signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, { code: 'operation-aborted' });
});
