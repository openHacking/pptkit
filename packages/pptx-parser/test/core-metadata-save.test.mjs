import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { shape, body } from './content-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const options = { projection: 'content', metadataSaveProfile: 'core-strings-v1' };
const fixture = strict => { const f = metadataFixture(strict); f.add('core', '<d:title>Original &amp; title</d:title><d:creator>Alice</d:creator><d:subject>Subject</d:subject><d:description>Description</d:description><d:language>en-US</d:language><c:revision>7</c:revision><c:lastModifiedBy>Do not change</c:lastModifiedBy>'); return f; };

test('bound core strings save across dialects and three reopens with exact non-target bytes', async () => {
  for (const strict of [false, true]) {
    let source = encodeMetadata(fixture(strict));
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(source), s = await importPptx(source, options);
      try { const old = s.document.metadata, next = { ...old, title: `Title ${round} & 中文`, author: `Author ${round}`, subject: `Subject ${round}`, description: `Description ${round}`, language: `x-test-${round}` }; s.document.setMetadata(next);
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 5); assert.ok(p.plan.report.changes.every(change => change.kind === 'metadata' && change.source.part === '/inventory/core.data'));
        const output = (await s.save(p.plan)).bytes, after = readGeneratedZip(output); assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [name, payload] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), payload, name);
        const escape = value => value.replaceAll('&', '&amp;'); let expected = before.get('inventory/core.data').toString();
        for (const [field, local] of [['title','title'],['author','creator'],['subject','subject'],['description','description'],['language','language']]) expected = expected.replace(`<d:${local}>${escape(old[field])}</d:${local}>`, `<d:${local}>${escape(next[field])}</d:${local}>`);
        assert.equal(after.get('inventory/core.data').toString(), expected); assert.match(expected, /<c:revision>7<\/c:revision><c:lastModifiedBy>Do not change<\/c:lastModifiedBy>/);
        const reopened = await importPptx(output, options); try { for (const field of ['title','author','subject','description','language']) assert.equal(reopened.document.metadata[field], next[field]); assert.equal(reopened.document.metadata.revision, 7); } finally { reopened.dispose(); } source = output;
      } finally { s.dispose(); }
    }
  }
});

test('metadata save rejects absent or mixed-content properties and unrelated edits atomically', async () => {
  for (const body of ['<d:creator>Alice</d:creator>', '<d:title>Orig<!--hidden-->inal</d:title>']) {
    const f = metadataFixture(); f.add('core', body); const s = await importPptx(encodeMetadata(f), options);
    try { s.document.setMetadata({ ...s.document.metadata, title: 'Changed' }); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.equal(p.plan, undefined); } finally { s.dispose(); }
  }
  const source = encodeMetadata(fixture(false)), s = await importPptx(source, options);
  try { s.document.setMetadata({ ...s.document.metadata, title: 'Changed', revision: 8 }); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); s.document.setMetadata({ ...s.document.metadata, revision: 7 }); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); s.document.setMetadata({ ...s.document.metadata, title: 'Later' }); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' }); } finally { s.dispose(); }
  for (const extra of [{ projection: 'structure' }, { metadataSaveProfile: 'other' }, { pictureEditSaveProfile: 'replace-v1' }, { descriptionEditSaveProfile: 'batch-v3' }]) await assert.rejects(importPptx(source, { ...options, ...extra }), TypeError);
});

test('metadata preservation rejects signed package relationships and respects shared change limits', async () => {
  const f = fixture(false); f.parts['deck/_rels/z.xml.rels'] = f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const signed = await importPptx(encodeMetadata(f), options); try { signed.document.setMetadata({ ...signed.document.metadata, title: 'Changed' }); const p = await signed.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.ok(p.report.diagnostics.some(d => d.code === 'save-text-package-policy')); } finally { signed.dispose(); }
  const limited = await importPptx(encodeMetadata(fixture(false)), { ...options, sessionLimits: { maxSaveChanges: 1 } }); try { limited.document.setMetadata({ ...limited.document.metadata, title: 'Changed', author: 'Changed' }); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
});

test('metadata, visibility, text, object properties and starting number share one guarded candidate', async () => {
  const f = fixture(false); f.parts['deck/z.xml'] = f.p('sld', `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({ id: '2', txBox: true, text: body() })}</p:spTree></p:cSld>`, `xmlns:a="${f.A}" show="0"`);
  f.add('custom', custom(2, 'lpwstr', 'Custom original', 'name="Custom"'));
  const source = encodeMetadata(f), opts = { ...options, customPropertySaveProfile: 'scalar-and-name-v2', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', slideVisibilitySaveProfile: 'attribute-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  const s = await importPptx(source, opts);
  try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); s.document.setMetadata({ ...s.document.metadata, title: 'Mixed title', customProperties: [{ name: 'Renamed custom', type: 'string', value: 'Custom updated' }] }); slide.setHidden(false); s.document.setFirstSlideNumber(17); edit.name = 'Mixed name'; edit.accessibility.description = 'Mixed alt'; edit.content[0].runs[0].text += '!'; slide.replaceElement(old.id, edit);
    const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 8);
    const next = (await s.save(p.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(next); for (const [name, bytes] of before) if (!['inventory/core.data', 'inventory/custom.data', 'deck/main.data', 'deck/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
    const reopened = await importPptx(next, opts); try { assert.equal(reopened.document.metadata.title, 'Mixed title'); assert.equal(reopened.document.metadata.customProperties[0].value, 'Custom updated'); assert.equal(reopened.document.metadata.customProperties[0].name, 'Renamed custom'); assert.equal(reopened.document.slides[0].hidden, false); assert.equal(reopened.document.firstSlideNumber, 17); assert.equal(reopened.document.slides[0].elements[0].name, 'Mixed name'); assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'Mixed alt'); assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, edit.content[0].runs[0].text); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});

test('empty existing properties and CDATA values can be changed without changing other metadata', async () => {
  const f = metadataFixture(); f.add('core', '<d:title/><d:creator><![CDATA[A & B]]></d:creator><c:revision>5</c:revision>'); const source = encodeMetadata(f), s = await importPptx(source, options);
  try { assert.equal(s.document.metadata.title, ''); assert.equal(s.document.metadata.author, 'A & B'); s.document.setMetadata({ ...s.document.metadata, title: 'New title', author: 'New author' }); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); const next = (await s.save(p.plan)).bytes, xml = readGeneratedZip(next).get('inventory/core.data').toString(); assert.match(xml, /<d:title>New title<\/d:title>/); assert.match(xml, /<d:creator>New author<\/d:creator>/); assert.match(xml, /<c:revision>5<\/c:revision>/); } finally { s.dispose(); }
});

test('invalid metadata characters, cancellation and work limits do not produce a candidate', async () => {
  const source = encodeMetadata(fixture(false));
  const invalid = await importPptx(source, options); try { invalid.document.setMetadata({ ...invalid.document.metadata, title: 'bad\0value' }); const result = await invalid.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined); } finally { invalid.dispose(); }
  const limited = await importPptx(source, { ...options, sessionLimits: { maxSaveValidationWork: 100 } }); try { limited.document.setMetadata({ ...limited.document.metadata, title: 'Changed' }); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  const aborted = await importPptx(source, options); try { aborted.document.setMetadata({ ...aborted.document.metadata, title: 'Changed' }); const controller = new AbortController(); controller.abort(); await assert.rejects(aborted.prepareSave({ mode: 'edit', signal: controller.signal }), { code: 'operation-aborted' }); const retry = await aborted.prepareSave({ mode: 'edit' }); assert.equal(retry.status, 'ready', JSON.stringify(retry.report)); } finally { aborted.dispose(); }
});

test('core metadata source inventory distinguishes missing, mixed and bound fields and is opt-in', async () => {
  const f = metadataFixture(); f.add('core', '<d:title>Orig<!--opaque-->inal</d:title><d:creator>Alice</d:creator><d:subject/>'); const bytes = encodeMetadata(f), base = await importPptx(bytes, { projection: 'content' }), s = await importPptx(bytes, options);
  try { assert.equal(base.report.coreMetadataSources, undefined); const records = s.report.coreMetadataSources; assert.ok(Object.isFrozen(records)); assert.deepEqual(records.map(x => [x.field, x.status]), [['title', 'unmapped'], ['author', 'bound'], ['subject', 'bound'], ['description', 'unmapped'], ['language', 'unmapped']]); assert.ok(records.every(Object.isFrozen)); assert.equal(records[1].source.part, '/inventory/core.data'); assert.ok(records[1].source.elementIndex > 1); } finally { base.dispose(); s.dispose(); }
  const missing = await importPptx(encodeMetadata(metadataFixture()), options); try { assert.equal(missing.report.coreMetadataSources.length, 5); assert.ok(missing.report.coreMetadataSources.every(x => x.status === 'unmapped' && x.source === undefined)); } finally { missing.dispose(); }
});


test('core-string edits preserve mapped and opaque custom property source across three rounds', async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict);
    f.add('custom', custom(2, 'lpwstr', 'Original &amp; text', 'name="Mapped"') + custom(3, 'vector', '<v:lpstr>opaque</v:lpstr>', 'name="Opaque"'));
    let source = encodeMetadata(f);
    const originalCustom = readGeneratedZip(source).get('inventory/custom.data');
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(source, options);
      try {
        assert.deepEqual(session.document.metadata.customProperties, [{ name: 'Mapped', type: 'string', value: 'Original & text' }]);
        session.document.setMetadata({ ...session.document.metadata, title: `Round ${round}` });
        const prepared = await session.prepareSave({ mode: 'edit' });
        assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        source = (await session.save(prepared.plan)).bytes;
        assert.deepEqual(readGeneratedZip(source).get('inventory/custom.data'), originalCustom);
      } finally { session.dispose(); }
    }
  }
});
