import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';

for (const strict of [false, true]) test(`metadata projection distinguishes missing, empty and unrepresentable values (strict=${strict})`, async () => {
  const absent = await importPptx(encodeMetadata(metadataFixture(strict)), { projection: 'content' });
  try { assert.deepEqual(absent.document.metadata, {}); assert.deepEqual(absent.report.inspection.metadata, {}); } finally { absent.dispose(); }
  const f = metadataFixture(strict);
  f.add('core', '<d:title/><d:creator>A &amp; B</d:creator><c:keywords xml:lang="zh">前<c:value xml:lang="en">word</c:value>后</c:keywords><c:revision>9007199254740992</c:revision><t:created xsi:type="t:W3CDTF">2020-01-01T00:00:00.0001Z</t:created><t:modified xsi:type="t:W3CDTF">2020-01-01T08:00:00+08:00</t:modified><c:category>source only</c:category>');
  f.add('extended', '<Company/><Manager>source only</Manager>');
  f.add('custom', custom(2, 'lpwstr', '', 'name="Empty"') + custom(3, 'null', '', 'name="Null"') + custom(4, 'i8', '2147483648', 'name="Wide"') + custom(5, 'lpwstr', 'source', 'name="Linked" linkTarget="elsewhere"'));
  const bytes = encodeMetadata(f), session = await importPptx(bytes, { projection: 'content' });
  try {
    assert.deepEqual(session.document.metadata, { title: '', author: 'A & B', keywords: ['前word后'], modifiedAt: '2020-01-01T00:00:00.000Z', company: '', customProperties: [{ name: 'Empty', type: 'string', value: '' }] });
    const findings = session.report.findings;
    for (const feature of ['metadata.keyword-segments', 'metadata.revision', 'metadata.core.created', 'metadata.core.category', 'metadata.extended.Manager']) assert.ok(findings.some(item => item.feature === feature && item.tier === 'unsupported'), feature);
    assert.equal(findings.filter(item => item.feature === 'metadata.custom-property' && item.tier === 'unsupported').length, 3);
    assert.equal(session.report.inspection.metadata.core.properties.find(item => item.name === 'created').value, '2020-01-01T00:00:00.0001Z');
    assert.equal(session.report.inspection.metadata.core.properties.find(item => item.name === 'revision').value, '9007199254740992');
    const plan = await session.prepareSave({mode:'no-op'}); assert.equal(plan.status, 'ready');
    assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes));
  } finally { session.dispose(); }
});

test('empty revision and date are distinct from missing strings', async () => {
  const f = metadataFixture(); f.add('core', '<d:subject/><c:keywords/><c:revision/><t:created xsi:type="t:W3CDTF"/>');
  const session = await importPptx(encodeMetadata(f), { projection: 'structure' });
  try {
    assert.deepEqual(session.document.metadata, { subject: '', keywords: [''] });
    assert.deepEqual(session.report.inspection.metadata.core.properties.map(p => [p.name, p.value]), [['subject', ''], ['keywords', ''], ['revision', ''], ['created', '']]);
    for (const feature of ['metadata.revision', 'metadata.core.created']) assert.ok(session.report.findings.some(item => item.feature === feature && item.tier === 'unsupported'));
  } finally { session.dispose(); }
});
