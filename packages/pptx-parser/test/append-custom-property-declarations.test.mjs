import test from 'node:test';
import assert from 'node:assert/strict';
import { appendCustomPropertyDeclarations as append } from '../dist/preservation/append-custom-property-declarations.js';

const relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
const typesNs = 'http://schemas.openxmlformats.org/package/2006/content-types';
const trans = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const strict = 'http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties';
const mime = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const request = { part: '/docProps/custom.xml', relationshipId: 'rId9', dialect: 'transitional' };
const options = { maxOutputBytes: 8192 };
const rel = `<r:Relationships xmlns:r="${relNs}"><!-- marker --><r:Relationship Id="rId1" Type="urn:test" Target="a.xml"/></r:Relationships><!-- tail -->`;
const types = `<c:Types xmlns:c="${typesNs}"><c:Default Extension="xml" ContentType="application/xml"/></c:Types><!-- tail -->`;
const be = text => { const bytes = Buffer.from(text, 'utf16le'); for (let i = 0; i < bytes.length; i += 2) [bytes[i], bytes[i + 1]] = [bytes[i + 1], bytes[i]]; return Buffer.concat([Buffer.from([0xfe, 0xff]), bytes]); };
const encodings = [text => Buffer.from(text), text => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)]), text => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]), be];
const code = expected => error => error.code === expected;

test('declarations preserve original byte regions and encoding for both dialects', async () => {
  for (const dialect of ['transitional', 'strict']) for (const encode of encodings) {
    const inputRel = encode(rel), inputTypes = encode(types), beforeRel = Buffer.from(inputRel), beforeTypes = Buffer.from(inputTypes);
    const result = await append(inputRel, inputTypes, { ...request, dialect }, options);
    const expectedRel = rel.replace('</r:Relationships><!-- tail -->', `<r:Relationship Id="rId9" Type="${dialect === 'strict' ? strict : trans}" Target="docProps/custom.xml"/></r:Relationships><!-- tail -->`);
    const expectedTypes = types.replace('</c:Types><!-- tail -->', `<c:Override PartName="/docProps/custom.xml" ContentType="${mime}"/></c:Types><!-- tail -->`);
    assert.deepEqual(Buffer.from(result.relationships), encode(expectedRel));
    assert.deepEqual(Buffer.from(result.contentTypes), encode(expectedTypes));
    assert.deepEqual(inputRel, beforeRel); assert.deepEqual(inputTypes, beforeTypes);
  }
});

test('self-closing default-namespace roots become valid declaration containers', async () => {
  const result = await append(Buffer.from(`<Relationships xmlns="${relNs}"/>`), Buffer.from(`<Types xmlns="${typesNs}"/>`), request, options);
  assert.equal(Buffer.from(result.relationships).toString(), `<Relationships xmlns="${relNs}"><Relationship Id="rId9" Type="${trans}" Target="docProps/custom.xml"/></Relationships>`);
  assert.equal(Buffer.from(result.contentTypes).toString(), `<Types xmlns="${typesNs}"><Override PartName="/docProps/custom.xml" ContentType="${mime}"/></Types>`);
});

test('raw IDs, both custom relationship types, part aliases and MIME declarations conflict', async () => {
  const wrapRel = child => Buffer.from(`<Relationships xmlns="${relNs}">${child}</Relationships>`);
  const wrapTypes = child => Buffer.from(`<Types xmlns="${typesNs}">${child}</Types>`);
  for (const child of [`<Relationship Id="rId9" Type="urn:other" Target="x"/>`, ...[trans, strict].map(type => `<Relationship Id="other" Type="${type}" Target="x"/>`), `<Unknown><Relationship Id="rId9" Type="urn:other" Target="x"/></Unknown>`]) await assert.rejects(append(wrapRel(child), wrapTypes(''), request, options), code('xml-edit-conflict'));
  for (const child of [`<Override PartName="/DOCPROPS/CUSTOM.XML" ContentType="application/xml"/>`, `<Override PartName="/other.xml" ContentType="${mime.toUpperCase()}"/>`, `<Default Extension="xml" ContentType="${mime}"/>`]) await assert.rejects(append(wrapRel(''), wrapTypes(child), request, options), code('xml-edit-conflict'));
});

test('unsupported markup, malformed roots, unsafe names, budgets and abort fail closed', async () => {
  const r = Buffer.from(`<Relationships xmlns="${relNs}"/>`), t = Buffer.from(`<Types xmlns="${typesNs}"/>`);
  for (const part of ['docProps/custom.xml', '/docProps/../custom.xml', '/docProps/a%20b.xml', '/docProps/a&b.xml', '/docProps/a.', '/docProps//a.xml']) await assert.rejects(append(r, t, { ...request, part }, options), code('xml-edit-unsupported'));
  for (const relationshipId of ['1id', 'r id', 'r"id']) await assert.rejects(append(r, t, { ...request, relationshipId }, options), code('xml-edit-unsupported'));
  await assert.rejects(append(Buffer.from('<wrong/>'), t, request, options), code('xml-edit-conflict'));
  await assert.rejects(append(r, Buffer.from(`<Types xmlns="${typesNs}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>`), request, options), code('xml-edit-unsupported'));
  await assert.rejects(append(r, t, request, { maxOutputBytes: 2 }), code('xml-edit-limit-exceeded'));
  await assert.rejects(append(r, t, request, { ...options, limits: { maxInputBytes: 2 } }), code('xml-limit-exceeded'));
  const controller = new AbortController(); controller.abort(); await assert.rejects(append(r, t, request, { ...options, signal: controller.signal }), code('operation-aborted'));
});

test('request and buffers are captured before asynchronous parsing', async () => {
  const r = Buffer.from(rel), t = Buffer.from(types), mutable = { ...request }, pending = append(r, t, mutable, options);
  r.fill(0); t.fill(0); mutable.part = '/wrong.xml'; mutable.relationshipId = 'wrong';
  const result = await pending;
  assert.match(Buffer.from(result.relationships).toString(), /Id="rId9"/);
  assert.match(Buffer.from(result.contentTypes).toString(), /PartName="\/docProps\/custom.xml"/);
});
