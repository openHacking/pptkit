import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
const options = { projection: 'content', metadataSaveProfile: 'core-keywords-v3' };
function fixture(keyword, strict = false) { const f = metadataFixture(strict); f.add('core', `<d:title>Original</d:title>${keyword}<c:revision>7</c:revision>`); return encodeMetadata(f); }

test('unsegmented aggregate keyword edits preserve other payloads over three rounds', async () => {
  for (const strict of [false, true]) {
    let source = fixture('<c:keywords>first, second</c:keywords>', strict);
    for (const value of ['中文 & text', '', 'last; value']) {
      const before = readGeneratedZip(source), s = await importPptx(source, options);
      try {
        assert.equal(s.report.coreMetadataSources.find(x => x.field === 'keywords').status, 'bound');
        s.document.setMetadata({ ...s.document.metadata, keywords: [value] });
        const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        source = (await s.save(ready.plan)).bytes;
        const after = readGeneratedZip(source); for (const [name, payload] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), payload, name);
        assert.match(after.get('inventory/core.data').toString(), /<c:revision>7<\/c:revision>/);
      } finally { s.dispose(); }
      const reopened = await importPptx(source, options); try { assert.deepEqual(reopened.document.metadata.keywords, [value]); } finally { reopened.dispose(); }
    }
  }
});

test('missing, language-tagged, segmented and mixed keywords do not bind', async () => {
  for (const value of ['', '<c:keywords xml:lang="en">alpha</c:keywords>', '<c:keywords><c:value>alpha</c:value></c:keywords>', '<c:keywords>alpha<!--opaque-->beta</c:keywords>', '<c:keywords>alpha<![CDATA[beta]]></c:keywords>']) {
    const s = await importPptx(fixture(value), options);
    try {
      assert.equal(s.report.coreMetadataSources.find(x => x.field === 'keywords').status, 'unmapped', value);
      s.document.setMetadata({ ...s.document.metadata, keywords: ['changed'] });
      assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked', value);
    } finally { s.dispose(); }
  }
});

test('keyword structure changes and old profiles remain blocked; revision can be combined', async () => {
  const source = fixture('<c:keywords>first</c:keywords>');
  const s = await importPptx(source, options);
  try {
    s.document.setMetadata({ ...s.document.metadata, keywords: ['one', 'two'] }); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked');
    s.document.setMetadata({ ...s.document.metadata, keywords: ['changed'], revision: 8 });
    const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    const reopened = await importPptx((await s.save(ready.plan)).bytes, options);
    try { assert.deepEqual(reopened.document.metadata.keywords, ['changed']); assert.equal(reopened.document.metadata.revision, 8); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
  const old = await importPptx(source, { projection: 'content', metadataSaveProfile: 'core-values-v2' });
  try { old.document.setMetadata({ ...old.document.metadata, keywords: ['changed'] }); assert.equal((await old.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { old.dispose(); }
});

test('single CDATA and empty keyword values are editable while no-op is exact', async () => {
  for (const keyword of ['<c:keywords><![CDATA[a & b]]></c:keywords>', '<c:keywords/>']) {
    const source = fixture(keyword), s = await importPptx(source, options);
    try {
      assert.equal(s.report.coreMetadataSources.find(x => x.field === 'keywords').status, 'bound');
      const noOp = await s.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes), Buffer.from(source));
      s.document.setMetadata({ ...s.document.metadata, keywords: ['<changed> &'] });
      const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
      const reopened = await importPptx((await s.save(ready.plan)).bytes, options);
      try { assert.deepEqual(reopened.document.metadata.keywords, ['<changed> &']); } finally { reopened.dispose(); }
    } finally { s.dispose(); }
  }
});
