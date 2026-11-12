import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, MC } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', metadataSaveProfile: 'extended-company-v4' };
function fixture(company, strict = false, attrs = '') { const f = metadataFixture(strict); f.add('core', '<d:title>Original</d:title><c:revision>7</c:revision><c:keywords>alpha</c:keywords>'); const part = f.add('extended', `${company}<Manager>Sentinel</Manager>`); if (attrs) f.parts[part] = f.properties('extended', `${company}<Manager>Sentinel</Manager>`, attrs); return encodeMetadata(f); }

test('source-bound company is saved across three rounds and preserves all other parts', async () => {
  for (const strict of [false, true]) {
    let source = fixture('<Company>Old</Company>', strict);
    for (const company of ['中文 & text', '', 'Final']) {
      const before = readGeneratedZip(source), s = await importPptx(source, options);
      try {
        assert.equal(s.report.coreMetadataSources.find(x => x.field === 'company').status, 'bound');
        s.document.setMetadata({ ...s.document.metadata, company }); const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        source = (await s.save(ready.plan)).bytes; const after = readGeneratedZip(source);
        for (const [name, bytes] of before) if (name !== 'inventory/extended.data') assert.deepEqual(after.get(name), bytes, name);
        assert.match(after.get('inventory/extended.data').toString(), /<Manager>Sentinel<\/Manager>/);
      } finally { s.dispose(); }
      const reopened = await importPptx(source, options); try { assert.equal(reopened.document.metadata.company, company); } finally { reopened.dispose(); }
    }
  }
});

test('missing, mixed or compatibility-processed company is unavailable', async () => {
  for (const [body, attrs] of [['', ''], ['<Company>old<!--opaque-->value</Company>', ''], ['<Company>old</Company>', `xmlns:mc="${MC}"`]]) {
    const s = await importPptx(fixture(body, false, attrs), options);
    try { assert.equal(s.report.coreMetadataSources.find(x => x.field === 'company').status, 'unmapped'); s.document.setMetadata({ ...s.document.metadata, company: 'New' }); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { s.dispose(); }
  }
  const old = await importPptx(fixture('<Company>Old</Company>'), { projection: 'content', metadataSaveProfile: 'core-keywords-v3' });
  try { old.document.setMetadata({ ...old.document.metadata, company: 'New' }); assert.equal((await old.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { old.dispose(); }
});

test('company combines atomically with core and new custom properties', async () => {
  const source = fixture('<Company>Old</Company>'), s = await importPptx(source, { ...options, customPropertySaveProfile: 'create-and-append-v5' });
  try {
    s.document.setMetadata({ ...s.document.metadata, company: 'New', title: 'Changed', revision: 8, keywords: ['changed'], customProperties: [{ name: 'A', type: 'string', value: 'v' }] });
    const ready = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    const reopened = await importPptx((await s.save(ready.plan)).bytes, options);
    try { assert.equal(reopened.document.metadata.company, 'New'); assert.equal(reopened.document.metadata.title, 'Changed'); assert.equal(reopened.document.metadata.revision, 8); assert.deepEqual(reopened.document.metadata.keywords, ['changed']); assert.equal(reopened.document.metadata.customProperties[0].value, 'v'); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});

test('company obeys package signature guards, cancellation, no-op and freshness', async () => {
  const f = metadataFixture(); f.add('extended', '<Company>Old</Company>');
  f.parts['deck/_rels/z.xml.rels'] = f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const signed = await importPptx(encodeMetadata(f), options);
  try { signed.document.setMetadata({ ...signed.document.metadata, company: 'New' }); const p = await signed.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.ok(p.report.diagnostics.some(d => d.code === 'save-text-package-policy')); } finally { signed.dispose(); }
  const source = fixture('<Company>Old</Company>'), s = await importPptx(source, options);
  try {
    const noOp = await s.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes), Buffer.from(source));
    s.document.setMetadata({ ...s.document.metadata, company: 'New' }); const abort = new AbortController(); abort.abort(); await assert.rejects(s.prepareSave({ mode: 'edit', signal: abort.signal }), { code: 'operation-aborted' });
    const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report)); s.document.setMetadata({ ...s.document.metadata, company: 'Changed' }); await assert.rejects(s.save(ready.plan), { code: 'save-snapshot-stale' });
  } finally { s.dispose(); }
});
