import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { verifyTextSourceAncestry } from '../dist/import/prepare-text-candidate.js';
const options = { projection: 'content', customPropertySaveProfile: 'scalar-v1' };
function fixture(strict) {
  const f = metadataFixture(strict);
  f.add('custom', custom(2, 'lpwstr', 'old &amp; text', 'name="Text"') + custom(3, 'bool', '0', 'name="Flag"') + custom(4, 'i4', '-10', 'name="Count"') + custom(5, 'r8', '1.25', 'name="Real"') + custom(6, 'filetime', '2020-01-02T03:04:05Z', 'name="Date"') + custom(7, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
  return encodeMetadata(f);
}
test('bound custom scalars save atomically for both dialects across three reopens', async () => {
  for (const strict of [false, true]) {
    let source = fixture(strict);
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(source), session = await importPptx(source, options);
      try {
        const values = session.document.metadata.customProperties.map(p => ({ ...p, value: p.name === 'Text' ? `round ${round} & _x0041_` : p.name === 'Flag' ? round % 2 === 0 : p.name === 'Count' ? round : p.name === 'Real' ? round + 0.5 : `2020-01-0${round + 3}T03:04:05.000Z` }));
        session.document.setMetadata({ ...session.document.metadata, customProperties: values });
        const prepared = await session.prepareSave({ mode: 'edit' });
        assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        const saved = await session.save(prepared.plan), after = readGeneratedZip(saved.bytes);
        for (const name of before.keys()) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), before.get(name), name);
        assert.match(after.get('inventory/custom.data').toString(), /<v:vector><v:lpwstr>opaque<\/v:lpwstr><\/v:vector>/);
        const reopened = await importPptx(saved.bytes, options);
        try { assert.deepEqual(reopened.document.metadata.customProperties, values); } finally { reopened.dispose(); }
        source = saved.bytes;
      } finally { session.dispose(); }
    }
  }
});

test('unbound values, topology changes and overlapping escapes block the entire edit', async () => {
  const f = metadataFixture(); f.add('custom', custom(2, 'lpwstr', 'a<!--opaque-->b', 'name="A"') + custom(3, 'lpwstr', 'ok', 'name="B"'));
  const source = encodeMetadata(f);
  for (const mutation of [
    values => values.map(p => p.name === 'A' ? { ...p, value: 'changed' } : p),
    values => values.map(p => p.name === 'B' ? { ...p, name: 'renamed' } : p),
    values => [...values].reverse(),
    values => values.map(p => p.name === 'B' ? { ...p, value: '_x005F_x0041_' } : p),
  ]) {
    const session = await importPptx(source, options);
    try { session.document.setMetadata({ ...session.document.metadata, customProperties: mutation(session.document.metadata.customProperties) });
      const prepared = await session.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'blocked'); assert.equal(prepared.plan, undefined);
    } finally { session.dispose(); }
  }
});


test('custom value ancestry is checked against independently captured raw XML', async () => {
  for (const strict of [false, true]) {
    const source = fixture(strict), session = await importPptx(source, options);
    try {
      const binding = session.report.customPropertySources[0], part = binding.source.part;
      const lexical = await captureXmlTextSource(readGeneratedZip(source).get('inventory/custom.data'), { partName: part });
      const value = lexical.elements[binding.valueElementIndex - 1], root = lexical.elements[0];
      const edit = { elementIndex: binding.valueElementIndex, namespaceUri: value.namespaceUri, local: value.local, expectedText: value.text, text: 'new', rawCustomPropertyElementIndex: binding.source.elementIndex, rawCustomRootNamespace: root.namespaceUri };
      await verifyTextSourceAncestry(lexical.elements, [edit], part, async () => {});
      for (const bad of [{ ...edit, rawCustomPropertyElementIndex: binding.source.elementIndex + 1 }, { ...edit, rawCustomRootNamespace: 'urn:wrong' }, { ...edit, elementIndex: 1 }]) {
        await assert.rejects(verifyTextSourceAncestry(lexical.elements, [bad], part, async () => {}), { code: 'save-candidate-invalid' });
      }
    } finally { session.dispose(); }
  }
});


test('custom edits obey cancellation, work limits and stale-plan checks', async () => {
  const source = fixture(false), session = await importPptx(source, options);
  try {
    const change = value => session.document.setMetadata({ ...session.document.metadata, customProperties: session.document.metadata.customProperties.map(p => p.name === 'Text' ? { ...p, value } : p) });
    change('first');
    const controller = new AbortController(); controller.abort();
    await assert.rejects(session.prepareSave({ mode: 'edit', signal: controller.signal }), { code: 'operation-aborted' });
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
    change('second'); await assert.rejects(session.save(plan.plan), { code: 'save-snapshot-stale' });
    const retry = await session.prepareSave({ mode: 'edit' }); assert.equal(retry.status, 'ready');
  } finally { session.dispose(); }
  const limited = await importPptx(source, { ...options, sessionLimits: { maxSaveChanges: 1 } });
  try {
    limited.document.setMetadata({ ...limited.document.metadata, customProperties: limited.document.metadata.customProperties.map(p => p.name === 'Text' ? { ...p, value: 'changed' } : p.name === 'Flag' ? { ...p, value: true } : p) });
    await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' });
  } finally { limited.dispose(); }
});


test('custom scalar profile rejects incompatible projections and drawing-property profiles', async () => {
  const bytes = fixture(false);
  for (const extra of [
    { projection: 'structure' }, { pictureSourceProfile: 'embedded-v1' }, { masterArtwork: 'pictures-v1' },
    { descriptionEditSaveProfile: 'existing-v1' }, { drawingNameSaveProfile: 'existing-v1' }, { decorativeSaveProfile: 'existing-v1' },
  ]) await assert.rejects(importPptx(bytes, { ...options, ...extra }), TypeError);
});


test('custom property names and values save together for both dialects through three reopens', async () => {
  for (const strict of [false, true]) {
    let source = fixture(strict);
    for (let round = 0; round < 3; round++) {
      const opts = { projection: 'content', customPropertySaveProfile: 'scalar-and-name-v2' };
      const s = await importPptx(source, opts);
      try {
        const entries = s.document.metadata.customProperties.map(p => p.type === 'string' ? { ...p, name: `Name ${round}\t_x0041_`, value: `Value ${round}\r & 中文` } : p);
        s.document.setMetadata({ ...s.document.metadata, customProperties: entries });
        const prepared = await s.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
        const output = (await s.save(prepared.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
        for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
        assert.match(after.get('inventory/custom.data').toString(), /<v:vector><v:lpwstr>opaque<\/v:lpwstr><\/v:vector>/);
        const reopened = await importPptx(output, opts);
        try { assert.deepEqual(reopened.document.metadata.customProperties, entries); } finally { reopened.dispose(); }
        source = output;
      } finally { s.dispose(); }
    }
  }
});

test('rename remains unsupported in v1 and collision with an opaque property yields no candidate', async () => {
  const bytes = fixture(false), first = await importPptx(bytes, options);
  try {
    first.document.setMetadata({ ...first.document.metadata, customProperties: first.document.metadata.customProperties.map(p => p.name === 'Text' ? { ...p, name: 'Changed' } : p) });
    assert.equal((await first.prepareSave({ mode: 'edit' })).status, 'blocked');
  } finally { first.dispose(); }
  const s = await importPptx(bytes, { projection: 'content', customPropertySaveProfile: 'scalar-and-name-v2' });
  try {
    s.document.setMetadata({ ...s.document.metadata, customProperties: s.document.metadata.customProperties.map(p => p.name === 'Text' ? { ...p, name: 'Opaque' } : p) });
    const result = await s.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined);
  } finally { s.dispose(); }
});


test('name-only edits bind original attributes and reject overlapping names atomically', async () => {
  const source = fixture(false), opts = { projection: 'content', customPropertySaveProfile: 'scalar-and-name-v2' };
  const s = await importPptx(source, opts);
  try {
    const update = name => s.document.setMetadata({ ...s.document.metadata, customProperties: s.document.metadata.customProperties.map(p => p.name === 'Text' ? { ...p, name } : p) });
    update('_x005F_x0041_'); const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
    s.document.setMetadata({ ...s.document.metadata, customProperties: s.document.metadata.customProperties.map(p => p.name === '_x005F_x0041_' ? { ...p, name: 'Text' } : p) });
    update('Name & \"中文'); const prepared = await s.prepareSave({ mode: 'edit' }); assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
    assert.equal(prepared.plan.report.changes.length, 1);
    const output = (await s.save(prepared.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
    const reopened = await importPptx(output, opts);
    try { assert.equal(reopened.document.metadata.customProperties[0].name, 'Name & \"中文'); assert.equal(reopened.document.metadata.customProperties[0].value, 'old & text'); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});


test('rename and value edit handle empty and CDATA scalar forms', async () => {
  for (const body of ['<v:lpwstr/>', '<v:lpwstr><![CDATA[old & value]]></v:lpwstr>']) {
    const f = metadataFixture(); f.add('custom', `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Old">${body}</property>`);
    const s = await importPptx(encodeMetadata(f), { projection: 'content', customPropertySaveProfile: 'scalar-and-name-v2' });
    try {
      s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'New', type: 'string', value: body.includes('CDATA') ? '' : 'filled' }] });
      const plan = await s.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      const reopened = await importPptx((await s.save(plan.plan)).bytes, { projection: 'content' });
      try { assert.deepEqual(reopened.document.metadata.customProperties, s.document.metadata.customProperties); } finally { reopened.dispose(); }
    } finally { s.dispose(); }
  }
});


test('rename rejects semantic collisions with opaque encoded names and unknown source name semantics', async () => {
  for (const strict of [false, true]) {
    for (const [opaqueName, proposed] of [['_x005F_x0041_', '_x0041_'], ['_x005F_x005F_x0041_', 'Different']]) {
      const f = metadataFixture(strict);
      f.add('custom', custom(2, 'lpwstr', 'value', 'name="Editable"') + custom(3, 'vector', '<v:lpwstr>opaque</v:lpwstr>', `name="${opaqueName}"`));
      const bytes = encodeMetadata(f), opts = { projection: 'content', customPropertySaveProfile: 'scalar-and-name-v2' }, s = await importPptx(bytes, opts);
      try {
        assert.equal(s.report.customPropertySources[0].nameEditable, opaqueName !== '_x005F_x005F_x0041_');
        s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: proposed, type: 'string', value: 'changed' }] });
        const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
        s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'Editable', type: 'string', value: 'changed' }] });
        const valueOnly = await s.prepareSave({ mode: 'edit' }); assert.equal(valueOnly.status, 'ready', JSON.stringify(valueOnly.report));
        const output = (await s.save(valueOnly.plan)).bytes;
        assert.match(readGeneratedZip(output).get('inventory/custom.data').toString(), new RegExp(`name="${opaqueName}"`));
      } finally { s.dispose(); }
    }
  }
});
