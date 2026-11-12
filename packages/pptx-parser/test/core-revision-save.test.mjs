import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', metadataSaveProfile: 'core-values-v2' };
function fixture(revision, strict = false) { const f = metadataFixture(strict); f.add('core', `<d:title>Original</d:title>${revision === undefined ? '' : `<c:revision>${revision}</c:revision>`}<d:creator>Sentinel</d:creator>`); return encodeMetadata(f); }

test('explicit existing revision updates preserve all other parts over three rounds in both dialects', async () => {
  for (const strict of [false, true]) {
    let source = fixture('0', strict);
    for (const value of [1, 42, Number.MAX_SAFE_INTEGER]) {
      const before = readGeneratedZip(source), s = await importPptx(source, options);
      try {
        assert.equal(s.report.coreMetadataSources.find(x => x.field === 'revision').status, 'bound');
        s.document.setMetadata({ ...s.document.metadata, revision: value });
        const prepared = await s.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        source = (await s.save(prepared.plan)).bytes;
        const after = readGeneratedZip(source);
        for (const [name, data] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), data, name);
        assert.match(after.get('inventory/core.data').toString(), /<d:creator>Sentinel<\/d:creator>/);
      } finally { s.dispose(); }
      const reopened = await importPptx(source, options); try { assert.equal(reopened.document.metadata.revision, value); } finally { reopened.dispose(); }
    }
  }
});

test('missing, noncanonical and unsafe revisions do not acquire save authority', async () => {
  for (const sourceValue of [undefined, '', '01', '+1', '-1', '1.0', '9007199254740992']) {
    const s = await importPptx(fixture(sourceValue), options);
    try {
      assert.equal(s.report.coreMetadataSources.find(x => x.field === 'revision').status, 'unmapped', String(sourceValue));
      s.document.setMetadata({ ...s.document.metadata, revision: 3 });
      const prepared = await s.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'blocked', String(sourceValue));
    } finally { s.dispose(); }
  }
  const old = await importPptx(fixture('2'), { projection: 'content', metadataSaveProfile: 'core-strings-v1' });
  try { old.document.setMetadata({ ...old.document.metadata, revision: 3 }); assert.equal((await old.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { old.dispose(); }
});

test('revision combines with title and custom creation atomically', async () => {
  const s = await importPptx(fixture('2'), { ...options, customPropertySaveProfile: 'create-and-append-v5' });
  try {
    s.document.setMetadata({ ...s.document.metadata, title: 'Changed', revision: 8, customProperties: [{ name: 'A', type: 'string', value: 'v' }] });
    const p = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
    const reopened = await importPptx((await s.save(p.plan)).bytes, options);
    try { assert.equal(reopened.document.metadata.title, 'Changed'); assert.equal(reopened.document.metadata.revision, 8); assert.equal(reopened.document.metadata.customProperties[0].value, 'v'); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});

test('revision respects signature policy, cancellation, budget and stale plans', async () => {
  const f = metadataFixture(); f.add('core', '<c:revision>2</c:revision>');
  f.parts['deck/_rels/z.xml.rels'] = f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const signed = await importPptx(encodeMetadata(f), options);
  try { signed.document.setMetadata({ ...signed.document.metadata, revision: 3 }); const p = await signed.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.ok(p.report.diagnostics.some(d => d.code === 'save-text-package-policy')); } finally { signed.dispose(); }
  const s = await importPptx(fixture('2'), options);
  try {
    s.document.setMetadata({ ...s.document.metadata, revision: 3 }); const abort = new AbortController(); abort.abort();
    await assert.rejects(s.prepareSave({ mode: 'edit', signal: abort.signal }), { code: 'operation-aborted' });
    const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
    s.document.setMetadata({ ...s.document.metadata, revision: 4 }); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' });
  } finally { s.dispose(); }
  const limited = await importPptx(fixture('2'), { ...options, sessionLimits: { maxSaveValidationWork: 100 } });
  try { limited.document.setMetadata({ ...limited.document.metadata, revision: 3 }); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
});

test('negative zero is rejected as an unrepresentable revision edit', async () => {
  const s = await importPptx(fixture('1'), options);
  try { s.document.setMetadata({ ...s.document.metadata, revision: -0 }); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.equal(p.plan, undefined); }
  finally { s.dispose(); }
});
