import test from 'node:test';
import assert from 'node:assert/strict';
import { appendImageContentType } from '../dist/preservation/append-image-content-type.js';
import { appendImageRelationship } from '../dist/preservation/append-image-relationship.js';
import { replaceXmlRelationshipId } from '../dist/preservation/replace-xml-relationship-id.js';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';
import { readArchive } from '../dist/archive/read-archive.js';
import { importPptx, inspectOpcPackage } from '../dist/index.js';
import { imageFixture, encodeImages, picture, png, validGif as gif } from './image-fixture.mjs';

test('private image candidate primitives isolate one shared picture across both OPC dialects', async () => {
  for (const strict of [false, true]) {
    const fixture = imageFixture(picture({ id: 2 }) + picture({ id: 3 }), strict), source = encodeImages(fixture);
    const session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
    try {
      const [first, second] = session.report.pictureSources;
      assert.equal(first.assetId, second.assetId);
      const archive = await readArchive(source, {}, true), entry = name => archive.parts.get(name);
      const slide = await replaceXmlRelationshipId(entry('deck/z.xml'), { elementIndex: first.sourceElementIndex, elementNamespace: fixture.A, elementLocal: 'blip', attributeNamespace: fixture.R, attributeLocal: 'embed', expected: 'image', replacement: 'rIdNew' }, { maxOutputBytes: 1024 * 1024 });
      const relationships = await appendImageRelationship(entry('deck/_rels/z.xml.rels'), { id: 'rIdNew', target: '../media/new.png', dialect: strict ? 'strict' : 'transitional' }, { maxOutputBytes: 1024 * 1024 });
      const types = await appendImageContentType(entry('[Content_Types].xml'), { part: '/media/new.png', mimeType: 'image/png' }, { maxOutputBytes: 1024 * 1024 });
      const candidate = await replaceArchiveEntries(source, [{ name: 'deck/z.xml', bytes: slide }, { name: 'deck/_rels/z.xml.rels', bytes: relationships }, { name: '[Content_Types].xml', bytes: types }], { maxOutputBytes: 1024 * 1024, additions: [{ name: 'media/new.png', bytes: png }] });
      const inspected = await inspectOpcPackage(candidate), refs = inspected.relationships.filter(item => item.source === '/deck/z.xml');
      assert.equal(refs.find(item => item.id === 'rIdNew').targetPart, '/media/new.png');
      assert.equal(refs.find(item => item.id === 'image').targetPart, '/media/raster.bin');
      assert.equal(inspected.parts.find(item => item.name === '/media/new.png').contentType, 'image/png');
      const reopened = await importPptx(candidate, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
      try {
        const [changed, untouched] = reopened.report.pictureSources;
        assert.equal(changed.relationshipId, 'rIdNew'); assert.equal(changed.mediaPart, '/media/new.png');
        assert.equal(untouched.relationshipId, 'image'); assert.equal(untouched.mediaPart, '/media/raster.bin');
        assert.notEqual(changed.assetId, untouched.assetId);
      } finally { reopened.dispose(); }
      const after = await readArchive(candidate, {}, true);
      for (const [name, bytes] of archive.parts) if (!['deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(after.parts.get(name), bytes, name);
    } finally { session.dispose(); }
  }
});

test('three private candidate rounds retain prior media and shared siblings while changing formats', async () => {
  for (const strict of [false, true]) {
    const fixture = imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict);
    let source = encodeImages(fixture), oldRelationship = 'image';
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(source, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
      try {
        const [target, sibling] = session.report.pictureSources, archive = await readArchive(source, {}, true);
        assert.equal(target.relationshipId, oldRelationship); assert.equal(sibling.mediaPart, '/media/raster.bin');
        const mimeType = round % 2 ? 'image/gif' : 'image/png', payload = round % 2 ? gif : png;
        const extension = round % 2 ? 'gif' : 'png', part = `/media/round-${round}.${extension}`, id = `rIdRound${round}`;
        const slide = await replaceXmlRelationshipId(archive.parts.get('deck/z.xml'), { elementIndex: target.sourceElementIndex, elementNamespace: fixture.A, elementLocal: 'blip', attributeNamespace: fixture.R, attributeLocal: 'embed', expected: oldRelationship, replacement: id }, { maxOutputBytes: 1024 * 1024 });
        const relationships = await appendImageRelationship(archive.parts.get('deck/_rels/z.xml.rels'), { id, target: `..${part}`, dialect: strict ? 'strict' : 'transitional' }, { maxOutputBytes: 1024 * 1024 });
        const types = await appendImageContentType(archive.parts.get('[Content_Types].xml'), { part, mimeType }, { maxOutputBytes: 1024 * 1024 });
        const candidate = await replaceArchiveEntries(source, [{ name: 'deck/z.xml', bytes: slide }, { name: 'deck/_rels/z.xml.rels', bytes: relationships }, { name: '[Content_Types].xml', bytes: types }], { maxOutputBytes: 1024 * 1024, additions: [{ name: part.slice(1), bytes: payload }] });
        const after = await readArchive(candidate, {}, true), inspection = await inspectOpcPackage(candidate);
        for (const [name, bytes] of archive.parts) if (!['deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(after.parts.get(name), bytes, `${round}: ${name}`);
        assert.equal(inspection.parts.find(item => item.name === part).contentType, mimeType);
        const reopened = await importPptx(candidate, { projection: 'content', pictureSourceProfile: 'embedded-v1' });
        try {
          const [changed, unchanged] = reopened.report.pictureSources;
          assert.equal(changed.mediaPart, part); assert.equal(unchanged.mediaPart, '/media/raster.bin');
          assert.equal(reopened.document.slides[0].elements[0].crop.left, 0.125);
          assert.equal(reopened.document.slides[0].elements[1].assetId, unchanged.assetId);
        } finally { reopened.dispose(); }
        source = candidate; oldRelationship = id;
      } finally { session.dispose(); }
    }
  }
});
