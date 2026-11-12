import test from 'node:test';
import assert from 'node:assert/strict';
import { appendImageRelationship } from '../dist/preservation/append-image-relationship.js';
import { inspectOpcPackage } from '../dist/index.js';
import { imageFixture, encodeImages, png } from './image-fixture.mjs';

const ns = 'http://schemas.openxmlformats.org/package/2006/relationships';
const transitional = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const strict = 'http://purl.oclc.org/ooxml/officeDocument/relationships/image';
const request = { id: 'rIdNew', target: '../media/new.png', dialect: 'transitional' }, options = { maxOutputBytes: 4096 };
const child = (prefix, type = transitional) => `<${prefix}Relationship Id="rIdNew" Type="${type}" Target="../media/new.png"/>`;
const be = text => { const b = Buffer.from(text, 'utf16le'); for (let i = 0; i < b.length; i += 2) [b[i], b[i + 1]] = [b[i + 1], b[i]]; return Buffer.concat([Buffer.from([0xfe, 0xff]), b]); };
const encodings = [text => Buffer.from(text), text => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)]), text => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]), be];

test('image relationship insertion preserves lexical XML, MCE branches, prefixes and encoding', async () => {
  const xml = `<?xml version="1.0"?><r:Relationships xmlns:r="${ns}" xmlns:mc="urn:mc"><!-- </r:Relationships> --><mc:AlternateContent><mc:Choice><r:Relationship Id="hidden" Type="x" Target="x"/></mc:Choice><mc:Fallback/></mc:AlternateContent><?p x?></r:Relationships><!-- tail -->`;
  const wanted = xml.replace('</r:Relationships><!-- tail -->', `${child('r:')}</r:Relationships><!-- tail -->`);
  for (const encode of encodings) {
    const source = encode(xml), original = Buffer.from(source);
    assert.deepEqual(Buffer.from(await appendImageRelationship(source, request, options)), encode(wanted));
    assert.deepEqual(source, original);
  }
  const empty = `<Relationships xmlns="${ns}" />`;
  assert.equal(Buffer.from(await appendImageRelationship(Buffer.from(empty), { ...request, dialect: 'strict' }, options)).toString(), `<Relationships xmlns="${ns}" >${child('', strict)}</Relationships>`);
  const unicode = `<关:Relationships xmlns:关="${ns}"/>`, unicodeWanted = `<关:Relationships xmlns:关="${ns}">${child('关:')}</关:Relationships>`;
  const exact = Buffer.byteLength(unicodeWanted);
  assert.equal(Buffer.from(await appendImageRelationship(Buffer.from(unicode), request, { maxOutputBytes: exact })).toString(), unicodeWanted);
  await assert.rejects(appendImageRelationship(Buffer.from(unicode), request, { maxOutputBytes: exact - 1 }), error => error.code === 'xml-edit-limit-exceeded');
});

test('image relationship insertion rejects hidden ID collisions, unsafe fields and limits', async () => {
  const xml = `<Relationships xmlns="${ns}" xmlns:mc="urn:mc"><mc:AlternateContent><mc:Choice><Relationship Id="r&#73;dNew" Type="x" Target="x"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>`;
  const code = expected => error => error.code === expected;
  await assert.rejects(appendImageRelationship(Buffer.from(xml), request, options), code('xml-edit-conflict'));
  const source = Buffer.from(`<Relationships xmlns="${ns}"/>`);
  for (const change of [{ id: '0bad' }, { id: 'x"/>' }, { target: '/abs' }, { target: '../x/../../y' }, { target: 'https://example.com/x' }, { target: 'a%20b' }, { target: 'a?x' }]) await assert.rejects(appendImageRelationship(source, { ...request, ...change }, options), code('xml-edit-unsupported'));
  await assert.rejects(appendImageRelationship(Buffer.from('<bad/>'), request, options), code('xml-edit-conflict'));
  await assert.rejects(appendImageRelationship(source, request, { maxOutputBytes: 10 }), code('xml-edit-limit-exceeded'));
  const controller = new AbortController(); controller.abort();
  await assert.rejects(appendImageRelationship(source, request, { ...options, signal: controller.signal }), code('operation-aborted'));
  if (typeof SharedArrayBuffer !== 'undefined') await assert.rejects(appendImageRelationship(new Uint8Array(new SharedArrayBuffer(2)), request, options), /ordinary byte buffer/);
});

test('image relationship insertion snapshots fields and source before async work', async () => {
  const input = Buffer.from(`<Relationships xmlns="${ns}"/>`), mutable = { ...request };
  const pending = appendImageRelationship(input, mutable, options); input.fill(0); mutable.id = 'changed';
  assert.match(Buffer.from(await pending).toString(), /Id="rIdNew"/);
});

test('inserted image relationship resolves beside the existing source relationship in both dialects', async () => {
  for (const dialect of ['transitional', 'strict']) {
    const fixture = imageFixture(undefined, dialect === 'strict');
    const original = fixture.parts['deck/_rels/z.xml.rels'];
    fixture.parts['deck/_rels/z.xml.rels'] = await appendImageRelationship(Buffer.from(original), { ...request, dialect }, options);
    fixture.parts['media/new.png'] = png; fixture.mediaTypes['media/new.png'] = 'image/png';
    const inspection = await inspectOpcPackage(encodeImages(fixture));
    const relationships = inspection.relationships.filter(item => item.source === '/deck/z.xml');
    assert.equal(relationships.find(item => item.id === 'image').targetPart, '/media/raster.bin');
    const added = relationships.find(item => item.id === 'rIdNew');
    assert.equal(added.targetPart, '/media/new.png');
    assert.equal(added.type, dialect === 'strict' ? strict : transitional);
  }
});
