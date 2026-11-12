import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { contentFixture, shape, body } from './content-fixture.mjs';
import { encodeStructure } from './structure-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const opts = { projection: 'content', slideVisibilitySaveProfile: 'attribute-v1' };
test('slide visibility preserves present/absent source attributes across both dialects and three rounds', async () => {
  for (const strict of [false, true]) {
    let source = encodeStructure(contentFixture(undefined, strict));
    for (const [round, values] of [[false, true], [true, false], [false, true]].entries()) {
      const before = readGeneratedZip(source), s = await importPptx(source, opts);
      try {
        assert.deepEqual(s.document.slides.map(slide => slide.hidden), round === 0 ? [true, false] : round === 1 ? [false, true] : [true, false]);
        s.document.slides.forEach((slide, i) => slide.setHidden(values[i]));
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 2);
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next); assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [name, data] of before) {
          if (name === 'deck/z.xml') assert.equal(after.get(name).toString(), data.toString().replace(`show="${values[0] ? '1' : '0'}"`, `show="${values[0] ? '0' : '1'}"`));
          else if (name === 'deck/a.xml') { const xml = data.toString(); assert.equal(after.get(name).toString(), round === 0 ? xml.replace('><p:cSld', ' show="0"><p:cSld') : xml.replace(`show="${values[1] ? '1' : '0'}"`, `show="${values[1] ? '0' : '1'}"`)); }
          else assert.deepEqual(after.get(name), data, name);
        }
        const reopened = await importPptx(next, opts); try { assert.deepEqual(reopened.document.slides.map(slide => slide.hidden), values); } finally { reopened.dispose(); } source = next;
      } finally { s.dispose(); }
    }
  }
});

test('slide visibility joins text, description, name and numbering in one candidate', async () => {
  const f = contentFixture(shape({ id: '2', txBox: true, text: body() })), source = encodeStructure(f);
  const options = { ...opts, editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  const s = await importPptx(source, options);
  try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); slide.setHidden(false); s.document.setFirstSlideNumber(17); edit.name = 'Changed'; edit.accessibility.description = 'New alt'; edit.content[0].runs[0].text += '!'; slide.replaceElement(old.id, edit);
    const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 5);
    const next = (await s.save(p.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(next);
    for (const [name, data] of before) if (!['deck/z.xml', 'deck/main.data'].includes(name)) assert.deepEqual(after.get(name), data, name);
    const reopened = await importPptx(next, options); try { assert.equal(reopened.document.slides[0].hidden, false); assert.equal(reopened.document.firstSlideNumber, 17); assert.equal(reopened.document.slides[0].elements[0].name, 'Changed'); assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'New alt'); assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, edit.content[0].runs[0].text); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});

test('slide visibility rejects unsupported changes, limits, stale plans and signed packages', async () => {
  const source = encodeStructure(contentFixture());
  const s = await importPptx(source, opts);
  try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); slide.setHidden(false); edit.box.x++; slide.replaceElement(old.id, edit); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); edit.box.x--; slide.replaceElement(old.id, edit); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); slide.setHidden(true); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' }); } finally { s.dispose(); }
  const limited = await importPptx(source, { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try { limited.document.slides[0].setHidden(false); limited.document.slides[1].setHidden(true); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  const fixture = contentFixture(); fixture.parts['deck/_rels/z.xml.rels'] = fixture.parts['deck/_rels/z.xml.rels'].replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ').replace('</Relationships>', '<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>');
  const signed = await importPptx(encodeStructure(fixture), opts); try { signed.document.slides[0].setHidden(false); const p = await signed.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'blocked'); assert.ok(p.report.diagnostics.some(d => d.code === 'save-text-package-policy')); } finally { signed.dispose(); }
  for (const extra of [{ projection: 'structure' }, { slideVisibilitySaveProfile: 'other' }, { pictureEditSaveProfile: 'replace-v1' }, { descriptionEditSaveProfile: 'batch-v3' }]) await assert.rejects(importPptx(source, { ...opts, ...extra }), TypeError);
});

test('shared source slide visibility views must agree even when only one view changes', async () => {
  const { analyzeTextEdits } = await import('../dist/import/analyze-text-edits.js'); const { DEFAULT_SESSION_LIMITS } = await import('../dist/index.js'); const { normalizePresentation } = await import('@pptkit/core');
  const f = contentFixture(); f.parts['deck/z.xml'] = f.parts['deck/z.xml'].replace('show="0"', 'show="1"'); const s = await importPptx(encodeStructure(f), opts);
  try { const baseline = normalizePresentation(s.document), first = s.document.slides[0], second = s.document.slides[1]; first.setHidden(true);
    const binding = slide => ({ slideId: slide.id, part: '/deck/shared.xml', elementIndex: 1, namespace: f.P, expected: '1', baseline: false });
    const result = await analyzeTextEdits(baseline, normalizePresentation(s.document), [], [], DEFAULT_SESSION_LIMITS, async () => {}, [], undefined, [binding(first), binding(second)]);
    assert.ok(result.diagnostics.some(d => d.code === 'save-source-shared-conflict'));
  } finally { s.dispose(); }
});

test('visibility source inventory is opt-in, detached and identifies source roots without granting save authority', async () => {
  const source = encodeStructure(contentFixture()); const base = await importPptx(source, { projection: 'content' }), session = await importPptx(source, opts);
  try {
    assert.equal(base.report.slideVisibilitySources, undefined); const records = session.report.slideVisibilitySources; assert.equal(records.length, 2); assert.ok(Object.isFrozen(records)); assert.ok(records.every(record => Object.isFrozen(record) && record.status === 'bound' && record.source.elementIndex === 1));
    assert.deepEqual(records.map(record => record.slideId), session.document.slides.map(slide => slide.id)); assert.deepEqual(records.map(record => record.source.part), ['/deck/z.xml', '/deck/a.xml']);
    const slide = session.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.box.x++; slide.replaceElement(old.id, edit); slide.setHidden(false); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
  } finally { base.dispose(); session.dispose(); }
});
