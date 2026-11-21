import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { importPptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const root = new URL('../../../', import.meta.url);
const path = name => fileURLToPath(new URL(name, root));
const options = { projection: 'content', customPropertySaveProfile: 'scalar-v1' };
const hash = bytes => createHash('sha256').update(bytes).digest('hex');

test('currency variants remain source-only and survive unrelated edits', async () => {
  const evidence = JSON.parse(await readFile(path('docs/testing/custom-currency-libreoffice-evidence.json'), 'utf8'));
  let bytes = await readFile(path('tests/fixtures/custom-currency-source.pptx'));
  assert.equal(hash(bytes), evidence.source_sha256);
  for (let round = 0; round < 3; round++) {
    const before = readGeneratedZip(bytes), session = await importPptx(bytes, options);
    try {
      const currencies = session.report.inspection.metadata.custom.properties.filter(item => item.valueType === 'cy');
      assert.equal(currencies.length, 3);
      assert.ok(currencies.every(item => item.value.status === 'uninterpreted'));
      assert.deepEqual(session.document.metadata.customProperties.map(item => item.name), ['Control']);
      const no = await session.prepareSave({ mode: 'no-op' }); assert.equal(no.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes), Buffer.from(bytes));
      const property = session.document.metadata.customProperties[0];
      session.document.setMetadata({ ...session.document.metadata, customProperties: [{ ...property, value: round === 0 ? 'after' : `edit-${round}` }] });
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      bytes = (await session.save(plan.plan)).bytes;
      if (round === 0) assert.equal(hash(bytes), evidence.pptkit_edited_sha256);
      const after = readGeneratedZip(bytes);
      for (const [name, payload] of before) if (name !== 'docProps/custom.xml') assert.deepEqual(after.get(name), payload, name);
      const xml = after.get('docProps/custom.xml').toString();
      for (const value of ['1.2500', '0.1000', '-12.5000']) assert.ok(xml.includes(`<vt:cy>${value}</vt:cy>`));
    } finally { session.dispose(); }
  }
  for (const kind of ['control', 'edited']) {
    const records = evidence.records.filter(item => item.kind === kind);
    assert.deepEqual(records.map(item => item.round), [0, 1, 2, 3]);
    assert.equal(records[0].currency_elements.length, 3);
    assert.ok(records.slice(1).every(item => item.currency_elements.length === 0));
    assert.ok(records.every(item => item.properties.some(p => p.name === 'Control' && p.value === (kind === 'control' ? 'before' : 'after'))));
  }
});
