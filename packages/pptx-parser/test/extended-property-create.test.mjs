import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata } from './metadata-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const config = { projection: 'content', metadataSaveProfile: 'extended-create-v12', customPropertySaveProfile: 'create-and-append-v5' };

test('creates extended properties in both dialects and edits across three reopens', async () => {
  for (const strict of [false, true]) {
    let bytes = encodeMetadata(metadataFixture(strict));
    for (const [round, value] of ['', '中文 & Co', 'final'].entries()) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, config);
      try {
        assert.equal(session.report.coreMetadataSources.find(item => item.field === 'company').status, round ? 'bound' : 'insertable');
        session.document.setMetadata({ ...session.document.metadata, company: value });
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
        const after = readGeneratedZip(bytes);
        for (const [name, payload] of before) if (!['_rels/.rels', '[Content_Types].xml', 'pptkit-extended-properties-1.xml'].includes(name)) assert.deepEqual(after.get(name), payload, name);
        assert.ok(after.has('pptkit-extended-properties-1.xml'));
      } finally { session.dispose(); }
      const reopened = await importPptx(bytes, config); try { assert.equal(reopened.document.metadata.company, value); } finally { reopened.dispose(); }
    }
  }
});

test('creates core, extended and custom properties atomically', async () => {
  for (const strict of [false, true]) {
    const session = await importPptx(encodeMetadata(metadataFixture(strict)), config);
    try {
      session.document.setMetadata({ ...session.document.metadata, title: 'Title', company: 'Company', customProperties: [{ name: 'Count', type: 'integer', value: 42 }] });
      const plan = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      const bytes = (await session.save(plan.plan)).bytes, zip = readGeneratedZip(bytes);
      for (const kind of ['core', 'extended', 'custom']) assert.ok(zip.has(`pptkit-${kind}-properties-1.xml`));
      const reopened = await importPptx(bytes, config); try { assert.equal(reopened.document.metadata.title, 'Title'); assert.equal(reopened.document.metadata.company, 'Company'); assert.equal(reopened.document.metadata.customProperties[0].value, 42); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test('creation respects prior profiles, malformed values, stale plans and source conflicts', async () => {
  const source = encodeMetadata(metadataFixture()), old = await importPptx(source, { projection: 'content', metadataSaveProfile: 'extended-company-insert-v11' });
  try { old.document.setMetadata({ ...old.document.metadata, company: 'A' }); assert.equal((await old.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { old.dispose(); }
  const session = await importPptx(source, config);
  try {
    session.document.setMetadata({ ...session.document.metadata, company: 'bad\u0000' }); const invalid = await session.prepareSave({ mode: 'edit' }); assert.equal(invalid.status, 'blocked'); assert.equal(invalid.plan, undefined);
    session.document.setMetadata({ ...session.document.metadata, company: 'A' }); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
    session.document.setMetadata({ ...session.document.metadata, company: 'B' }); await assert.rejects(session.save(plan.plan), { code: 'save-snapshot-stale' });
  } finally { session.dispose(); }
  const f = metadataFixture(); f.parts['orphan.data'] = '<unused/>'; f.metadataTypes['orphan.data'] = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
  await assert.rejects(importPptx(encodeMetadata(f), config), { code: 'import-invalid-package' });
});

test('creation preserves no-op bytes and rejects cancellation and exhausted budgets', async () => {
  const source = encodeMetadata(metadataFixture()), session = await importPptx(source, config);
  try {
    const no = await session.prepareSave({ mode: 'no-op' }); assert.equal(no.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes), Buffer.from(source));
    session.document.setMetadata({ ...session.document.metadata, company: 'A' }); const controller = new AbortController(); controller.abort();
    await assert.rejects(session.prepareSave({ mode: 'edit', signal: controller.signal }), { code: 'operation-aborted' });
  } finally { session.dispose(); }
  const limited = await importPptx(source, { ...config, sessionLimits: { maxSaveAnalysisWork: 1 } });
  try { limited.document.setMetadata({ ...limited.document.metadata, company: 'A' }); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
});

test('invalid combined creation produces no partial save plan', async () => {
  const session = await importPptx(encodeMetadata(metadataFixture()), config);
  try {
    session.document.setMetadata({ ...session.document.metadata, title: 'Title', company: 'bad\u0000', customProperties: [{ name: 'Count', type: 'integer', value: 42 }] });
    const plan = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] });
    assert.equal(plan.status, 'blocked'); assert.equal(plan.plan, undefined);
  } finally { session.dispose(); }
});

test('extended creation avoids case-insensitive part and relationship ID collisions', async () => {
  for (const strict of [false, true]) {
    const f=metadataFixture(strict);f.parts['PPTKIT-EXTENDED-PROPERTIES-1.XML']='<sentinel/>';f.metadataTypes['PPTKIT-EXTENDED-PROPERTIES-1.XML']='application/xml';f.rootRelations.push({Id:'pptkitExtendedProperties2',Type:'urn:test',Target:'deck/main.data'});
    const session=await importPptx(encodeMetadata(f),config);
    try {
      assert.equal(session.report.coreMetadataSources.find(x=>x.field==='company').source.part,'/pptkit-extended-properties-3.xml');
      session.document.setMetadata({...session.document.metadata,company:'Safe'});const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));
      const bytes=(await session.save(plan.plan)).bytes,zip=readGeneratedZip(bytes);assert.equal(zip.get('PPTKIT-EXTENDED-PROPERTIES-1.XML').toString(),'<sentinel/>');assert.ok(zip.has('pptkit-extended-properties-3.xml'));
      const reopened=await importPptx(bytes,config);try{assert.equal(reopened.document.metadata.company,'Safe');assert.equal(reopened.report.inspection.metadata.extended.part,'/pptkit-extended-properties-3.xml');}finally{reopened.dispose();}
    } finally {session.dispose();}
  }
});

test('extended creation blocks an inactive relationship that could become live', async () => {
  const {sourceAlternates}=await import('./closure-fixture.mjs');
  for (const strict of [false,true]) {
    const f=metadataFixture(strict);f.parts['deck/_rels/z.xml.rels']=sourceAlternates('<Relationship Id="hidden" Type="urn:asset" Target="../pptkit-extended-properties-1.xml"/>',`<Relationship Id="layout" Type="${f.R}/slideLayout" Target="layout.xml"/>`);
    const session=await importPptx(encodeMetadata(f),config);try{assert.equal(session.report.coreMetadataSources.find(x=>x.field==='company').status,'unmapped');session.document.setMetadata({...session.document.metadata,company:'Unsafe'});const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'blocked');assert.equal(plan.plan,undefined);assert.ok(plan.report.diagnostics.some(x=>x.documentPath==='/metadata' && /metadata change/.test(x.message)));}finally{session.dispose();}
  }
});

test('extended creation refuses an inactive signature relationship without mutating the source', async () => {
  const f=metadataFixture();f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>','<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const source=encodeMetadata(f),session=await importPptx(source,config);try{const baseline=session.document.metadata;assert.equal(session.report.coreMetadataSources.find(x=>x.field==='company').status,'unmapped');session.document.setMetadata({...baseline,company:'Unsafe'});const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'blocked');assert.equal(plan.plan,undefined);session.document.setMetadata(baseline);const no=await session.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes),Buffer.from(source));}finally{session.dispose();}
});

test('content import refuses an active signature relationship before extended creation', async () => {
  const f=metadataFixture();f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('</Relationships>','<Relationship Id="signature" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="layout.xml"/></Relationships>');
  await assert.rejects(importPptx(encodeMetadata(f),config),{code:'import-unsupported'});
});
