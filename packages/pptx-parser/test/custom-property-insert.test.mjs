import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom, GUID, MC } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { shape, body } from './content-fixture.mjs';
const options = { projection: 'content', customPropertySaveProfile: 'append-and-remove-v4' };
const id = n => `${GUID}:${n}`;
function fixture(strict, empty = false) {
  const f = metadataFixture(strict), part = f.add('custom', empty ? '' : custom(2, 'lpwstr', 'old', 'name="A"') + custom(3, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"') + custom(4, 'bool', 'false', 'name="B"'));
  if (empty) f.parts[part] = f.parts[part].replace('></Properties>', '/>');
  return encodeMetadata(f);
}
test('explicit appends avoid existing and removed IDs and preserve opaque siblings in both dialects', async () => {
  for (const strict of [false, true]) {
    const bytes = fixture(strict), before = readGeneratedZip(bytes), s = await importPptx(bytes, options);
    try {
      assert.equal(s.report.customPropertyInsertion.status, 'available');
      const entries = [{ name: 'Renamed', type: 'boolean', value: true }, { name: 'New _x0041_', type: 'string', value: '中文\r & _x0041_' }, { name: 'Real', type: 'real', value: -0 }, { name: 'When', type: 'date', value: '2020-01-02T03:04:05.000Z' }];
      s.document.setMetadata({ ...s.document.metadata, customProperties: entries });
      const prepared = await s.prepareSave({ mode: 'edit', customPropertyRemovals: [id(2)], customPropertyAdditions: [1, 2, 3] });
      assert.equal(prepared.status, 'ready', JSON.stringify(prepared.report));
      const output = (await s.save(prepared.plan)).bytes, after = readGeneratedZip(output);
      for (const [name, value] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), value, name);
      const reopened = await importPptx(output, options);
      try {
        assert.deepEqual(reopened.document.metadata.customProperties, entries);
        assert.deepEqual(reopened.report.inspection.metadata.custom.properties.map(p => p.id), [3, 4, 5, 6, 7]);
        assert.equal(reopened.report.inspection.metadata.custom.properties[0].value.status, 'uninterpreted');
      } finally { reopened.dispose(); }
    } finally { s.dispose(); }
  }
});

test('an existing empty self-closing part accepts the first explicit insertion', async () => {
  for (const strict of [false, true]) {
    const s = await importPptx(fixture(strict, true), options);
    try {
      s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'First', type: 'integer', value: 42 }] });
      const p = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
      const next = await importPptx((await s.save(p.plan)).bytes, options);
      try { assert.equal(next.report.inspection.metadata.custom.properties[0].id, 2); assert.equal(next.document.metadata.customProperties[0].value, 42); } finally { next.dispose(); }
    } finally { s.dispose(); }
  }
});

test('missing part, wrong index, unknown names and incompatible profiles do not create candidates', async () => {
  const absent = await importPptx(encodeMetadata(metadataFixture()), options);
  try { assert.equal(absent.report.customPropertyInsertion.status, 'unavailable'); absent.document.setMetadata({ ...absent.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'v' }] }); assert.equal((await absent.prepareSave({ mode: 'edit', customPropertyAdditions: [0] })).status, 'blocked'); } finally { absent.dispose(); }
  const s = await importPptx(fixture(false), options);
  try {
    s.document.setMetadata({ ...s.document.metadata, customProperties: [...s.document.metadata.customProperties, { name: 'Opaque', type: 'string', value: 'v' }] });
    assert.equal((await s.prepareSave({ mode: 'edit', customPropertyAdditions: [2] })).status, 'blocked');
    s.document.setMetadata({ ...s.document.metadata, customProperties: [...s.document.metadata.customProperties.slice(0, 2), { name: 'New', type: 'string', value: 'v' }] });
    assert.equal((await s.prepareSave({ mode: 'edit', customPropertyAdditions: [1] })).status, 'blocked');
    await assert.rejects(s.prepareSave({ mode: 'edit', customPropertyAdditions: [2, 2] }), TypeError);
  } finally { s.dispose(); }
  const old = await importPptx(fixture(false), { projection: 'content', customPropertySaveProfile: 'remove-v3' });
  try { await assert.rejects(old.prepareSave({ mode: 'edit', customPropertyAdditions: [2] }), TypeError); } finally { old.dispose(); }
});


test('appends retain allocated identities across three reopens', async () => {
  for (const strict of [false, true]) {
    let source = fixture(strict, true);
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(source, options);
      try {
        const existing = s.document.metadata.customProperties;
        s.document.setMetadata({ ...s.document.metadata, customProperties: [...existing, { name: `Name ${round}`, type: 'integer', value: round }] });
        const p = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [existing.length] }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
        source = (await s.save(p.plan)).bytes;
      } finally { s.dispose(); }
      const reopened = await importPptx(source, options);
      try { assert.deepEqual(reopened.report.inspection.metadata.custom.properties.map(p => p.id), Array.from({ length: round + 1 }, (_, index) => index + 2)); }
      finally { reopened.dispose(); }
    }
  }
});

test('unknown source names and compatibility processing disable insertion without changing the source', async () => {
  for (const type of ['unknown-name', 'compatibility']) {
    const f = metadataFixture(), body = custom(2, 'vector', '<v:lpwstr>opaque</v:lpwstr>', `name="${type === 'unknown-name' ? '_x005F_x005F_x0041_' : 'Opaque'}"`), part = f.add('custom', body);
    if (type === 'compatibility') f.parts[part] = f.properties('custom', body, `xmlns:mc="${MC}"`);
    const s = await importPptx(encodeMetadata(f), options);
    try {
      assert.equal(s.report.customPropertyInsertion.status, 'unavailable');
      s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'v' }] });
      const result = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined);
    } finally { s.dispose(); }
  }
});

test('insertion rejects invalid XML, overlapping escape sequences and constrained budgets', async () => {
  for (const value of ['bad\u0000text', '_x005F_x0041_']) {
    const s = await importPptx(fixture(false, true), options);
    try {
      s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'New', type: 'string', value }] });
      const result = await s.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined);
    } finally { s.dispose(); }
  }
  const limited = await importPptx(fixture(false, true), { ...options, sessionLimits: { maxSaveAnalysisWork: 3 } });
  try { limited.document.setMetadata({ ...limited.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'v' }] }); await assert.rejects(limited.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
});


test('custom insertion and deletion combine atomically with core, visibility and text edits', async () => {
  const f = metadataFixture(); f.add('core', '<d:title>Original</d:title>'); f.add('custom', custom(2, 'lpwstr', 'old', 'name="A"'));
  f.parts['deck/z.xml'] = f.p('sld', `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({ id: '2', txBox: true, text: body() })}</p:spTree></p:cSld>`, `xmlns:a="${f.A}" show="0"`);
  const source = encodeMetadata(f), config = { ...options, metadataSaveProfile: 'core-strings-v1', slideVisibilitySaveProfile: 'attribute-v1', editSaveProfile: 'text-runs-v3' }, s = await importPptx(source, config);
  try {
    const slide = s.document.slides[0], original = slide.elements[0], changed = structuredClone(original);
    changed.content[0].runs[0].text += '!'; slide.replaceElement(original.id, changed); slide.setHidden(false);
    s.document.setMetadata({ ...s.document.metadata, title: 'Changed', customProperties: [{ name: 'New', type: 'string', value: 'New value' }] });
    const request = { mode: 'edit', customPropertyRemovals: [id(2)], customPropertyAdditions: [0] };
    const invalid = structuredClone(changed); invalid.box.x++; slide.replaceElement(original.id, invalid);
    const blocked = await s.prepareSave(request); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
    slide.replaceElement(original.id, changed);
    const ready = await s.prepareSave(request); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    const output = (await s.save(ready.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const [name, bytes] of before) if (!['inventory/core.data', 'inventory/custom.data', 'deck/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
    const reopened = await importPptx(output, config);
    try { assert.equal(reopened.document.metadata.title, 'Changed'); assert.equal(reopened.document.slides[0].hidden, false); assert.deepEqual(reopened.document.metadata.customProperties, [{ name: 'New', type: 'string', value: 'New value' }]); assert.equal(reopened.report.inspection.metadata.custom.properties[0].id, 3); }
    finally { reopened.dispose(); }
  } finally { s.dispose(); }
});


test('insertion captures requested indexes and enforces cancellation, freshness and count limits', async () => {
  const source = fixture(false, true), s = await importPptx(source, options);
  try {
    s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'v' }] });
    const abort = new AbortController(); abort.abort();
    await assert.rejects(s.prepareSave({ mode: 'edit', customPropertyAdditions: [0], signal: abort.signal }), { code: 'operation-aborted' });
    await assert.rejects(s.prepareSave({ mode: 'edit', customPropertyAdditions: new Array(1) }), TypeError);
    const indexes = [0], pending = s.prepareSave({ mode: 'edit', customPropertyAdditions: indexes }); indexes[0] = 1;
    const ready = await pending; assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
    s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'New', type: 'string', value: 'changed' }] });
    await assert.rejects(s.save(ready.plan), { code: 'save-snapshot-stale' });
  } finally { s.dispose(); }
  const limited = await importPptx(source, { ...options, sessionLimits: { maxSaveChanges: 1 } });
  try {
    limited.document.setMetadata({ ...limited.document.metadata, customProperties: [{ name: 'A', type: 'string', value: 'a' }, { name: 'B', type: 'string', value: 'b' }] });
    await assert.rejects(limited.prepareSave({ mode: 'edit', customPropertyAdditions: [0, 1] }), { code: 'save-limit-exceeded' });
  } finally { limited.dispose(); }
});


test('explicit removal and reappend is a source change even when the normalized document is unchanged', async () => {
  for (const strict of [false, true]) {
    const f=metadataFixture(strict); f.add('custom',custom(2,'lpwstr','same','name="A"')+custom(3,'vector','<v:lpwstr>opaque</v:lpwstr>','name="Opaque"'));
    const source=encodeMetadata(f), before=readGeneratedZip(source), s=await importPptx(source,options);
    try {
      const original=structuredClone(s.document.metadata.customProperties);
      const plan=await s.prepareSave({mode:'edit',customPropertyRemovals:[id(2)],customPropertyAdditions:[0]});
      assert.equal(plan.status,'ready',JSON.stringify(plan.report));
      assert.equal(plan.plan.report.changes.length,2);
      const output=(await s.save(plan.plan)).bytes, after=readGeneratedZip(output);
      for(const [name,bytes] of before) if(name!=='inventory/custom.data') assert.deepEqual(after.get(name),bytes,name);
      const reopened=await importPptx(output,options);
      try {assert.deepEqual(reopened.document.metadata.customProperties,original);assert.deepEqual(reopened.report.inspection.metadata.custom.properties.map(p=>p.id),[3,4]);assert.equal(reopened.report.inspection.metadata.custom.properties[0].value.status,'uninterpreted');} finally {reopened.dispose();}
      const onlyRemoval=await s.prepareSave({mode:'edit',customPropertyRemovals:[id(2)]}); assert.equal(onlyRemoval.status,'blocked');
      const onlyAddition=await s.prepareSave({mode:'edit',customPropertyAdditions:[0]}); assert.equal(onlyAddition.status,'blocked');
    } finally {s.dispose();}
  }
});


test('same-value structural replacements honor source guards, cancellation and plan freshness', async () => {
  const f=metadataFixture(); f.add('custom',custom(2,'lpwstr','same','name="A"'));
  const source=encodeMetadata(f), request={mode:'edit',customPropertyRemovals:[id(2)],customPropertyAdditions:[0]}, s=await importPptx(source,options);
  try {
    const abort=new AbortController(); abort.abort(); await assert.rejects(s.prepareSave({...request,signal:abort.signal}),{code:'operation-aborted'});
    const p=await s.prepareSave(request); assert.equal(p.status,'ready',JSON.stringify(p.report));
    s.document.setMetadata({...s.document.metadata,customProperties:[{name:'A',type:'string',value:'changed'}]});
    await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});
  } finally {s.dispose();}
  const signed=metadataFixture(); signed.add('custom',custom(2,'lpwstr','same','name="A"'));
  signed.parts['deck/_rels/z.xml.rels']=signed.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const guarded=await importPptx(encodeMetadata(signed),options); try {const p=await guarded.prepareSave(request);assert.equal(p.status,'blocked');assert.ok(p.report.diagnostics.some(d=>d.code==='save-text-package-policy'));} finally {guarded.dispose();}
});
