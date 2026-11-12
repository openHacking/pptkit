import test from 'node:test';
import assert from 'node:assert/strict';
import { structureContext } from '../dist/presentation/structure-context.js';
import { bindDrawingDescription } from '../dist/import/drawing-description-binding.js';
import { projectContent } from '../dist/import/project-content.js';
import { projectStructure } from '../dist/import/project-structure.js';
import { readStructureFromContext } from '../dist/presentation/read-structure.js';
import { DEFAULT_SESSION_LIMITS } from '../dist/import/import-pptx.js';
import { analyzeDescriptionEdit } from '../dist/import/analyze-description-edit.js';
import { prepareTextCandidate } from '../dist/import/prepare-text-candidate.js';
import { normalizePresentation } from '@pptkit/core';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';
import { importPptx } from '../dist/index.js';
import { bytes, shape, body, contentFixture } from './content-fixture.mjs';
import { encodeStructure } from './structure-fixture.mjs';
import { table, tableFixture } from './table-fixture.mjs';
import { notesFixture } from './notes-fixture.mjs';

const child = (node, namespace, local) => node.children.find(item => item.kind === 'element' && item.namespaceUri === namespace && item.local === local);

async function fixture(strict = false, first = shape({ id: '2' }), second = shape({ id: '3' })) {
  const context = await structureContext(bytes(first + second, strict));
  const root = await context.root('/deck/z.xml', 'sld');
  const tree = child(child(root, context.p, 'cSld'), context.p, 'spTree');
  const shapes = tree.children.filter(item => item.kind === 'element' && item.local === 'sp');
  const nonvisual = shapes.map(item => child(item, context.p, 'nvSpPr'));
  const properties = nonvisual.map(item => child(item, context.p, 'cNvPr'));
  const bind = (owner = shapes[0], nv = nonvisual[0], cnv = properties[0], projected = 'Description', id = '2') => bindDrawingDescription(context, '/deck/z.xml', 's', 'element', id, owner, nv, cnv, projected);
  return { context, shapes, nonvisual, properties, bind };
}

test('description source binding requires exact raw owner and attribute in both dialects', async () => {
  for (const strict of [false, true]) {
    const f = await fixture(strict);
    const result = await f.bind();
    assert.equal(result.status, 'bound'); assert.equal(result.expected, 'Description');
    assert.equal(result.sourceObjectId, '2'); assert.ok(result.ownerIndex < result.nonvisualIndex && result.nonvisualIndex < result.elementIndex);
    assert.equal((await f.bind(f.shapes[1], f.nonvisual[0], f.properties[0])).status, 'unmapped');
    assert.equal((await f.bind(f.shapes[0], f.nonvisual[1], f.properties[0])).status, 'unmapped');
    assert.equal((await f.bind(f.shapes[0], f.nonvisual[0], f.properties[1])).status, 'unmapped');
    assert.equal((await f.bind(undefined, undefined, undefined, 'different')).status, 'unmapped');
    assert.equal((await f.bind(undefined, undefined, undefined, 'Description', '3')).status, 'unmapped');
  }
});

test('absent, empty and lexical drawing identities are distinguished', async () => {
  const absent = await fixture(false, shape({ id: '2' }).replace(' descr="Description"', ''));
  assert.equal((await absent.bind()).status, 'unmapped');
  const empty = await fixture(false, shape({ id: '2' }).replace('descr="Description"', 'descr=""'));
  assert.equal((await empty.bind(undefined, undefined, undefined, '')).status, 'bound');
  const alternate = await fixture(false, shape({ id: '+002' }));
  assert.equal((await alternate.bind()).status, 'bound');
  const minusZero = await fixture(false, shape({ id: '-0' }));
  assert.equal((await minusZero.bind(undefined, undefined, undefined, 'Description', '0')).status, 'unmapped');
});

test('internal content projection captures mapped source descriptions without granting save permission', async () => {
  for (const strict of [false, true]) {
    const first = shape({ id: '2' }), second = shape({ id: '3' }).replace(' descr="Description"', '');
    const context = await structureContext(bytes(first + second, strict)), checkpoint = async () => {};
    const inspection = await readStructureFromContext(context);
    const projection = await projectStructure(inspection, DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint);
    const bindings = [];
    const report = await projectContent(projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, ...Array(20).fill(undefined), bindings);
    assert.deepEqual(bindings.map(item => [item.sourceObjectId, item.status, item.expected]), [['2', 'bound', 'Description'], ['3', 'unmapped', undefined]]);
    assert.equal(report.findings.find(item => item.feature === 'drawing.shape.mapped-fields' && item.operation === 'edit-save').tier, 'unsupported');
    assert.equal(projection.document.slides[0].elements.length, 2);
  }
});

test('effective MCE property under a raw wrapper is not bound as a direct owner', async () => {
  const original = shape({ id: '2' });
  const properties = '<p:cNvPr id="2" name="Object 2" descr="Description" />';
  const wrapped = original.replace(properties, `<mc:AlternateContent><mc:Choice Requires="x"><p:cNvPr id="22" name="Choice" descr="Choice"/></mc:Choice><mc:Fallback>${properties}</mc:Fallback></mc:AlternateContent>`);
  assert.notEqual(wrapped, original);
  const source = contentFixture(wrapped); source.parts['deck/z.xml'] = source.parts['deck/z.xml'].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown" mc:Ignorable="x" ');
  const context = await structureContext(encodeStructure(source)), root = await context.root('/deck/z.xml', 'sld');
  const tree = child(child(root, context.p, 'cSld'), context.p, 'spTree'), owner = child(tree, context.p, 'sp');
  const nv = child(owner, context.p, 'nvSpPr'), cnv = child(nv, context.p, 'cNvPr');
  const result = await bindDrawingDescription(context, '/deck/z.xml', 's', 'object', '2', owner, nv, cnv, 'Description');
  assert.equal(result.status, 'unmapped'); assert.match(result.reason, /ancestry/);
  const session = await importPptx(encodeStructure(source), { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const slide = session.document.slides[0], element = slide.elements[0];
  assert.equal(session.report.findings.find(item => item.feature === 'drawing.description-existing').tier, 'unsupported');
  slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: 'changed' } });
  assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); session.dispose();
  const v2 = await importPptx(encodeStructure(source), { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' });
  const v2Slide = v2.document.slides[0], v2Element = v2Slide.elements[0];
  assert.equal(v2.report.findings.find(item => item.feature === 'drawing.description-attribute').tier, 'unsupported');
  v2Slide.replaceElement(v2Element.id, { ...v2Element, accessibility: { ...v2Element.accessibility, description: 'changed' } });
  assert.equal((await v2.prepareSave({ mode: 'edit' })).status, 'blocked'); v2.dispose();
  const v4 = await importPptx(encodeStructure(source), { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v4' });
  const v4Slide = v4.document.slides[0], v4Element = v4Slide.elements[0];
  v4Slide.replaceElement(v4Element.id, { ...v4Element, accessibility: { ...v4Element.accessibility, description: 'changed' } });
  assert.equal((await v4.prepareSave({ mode: 'edit' })).status, 'blocked'); v4.dispose();
});

test('group and nested child retain separate raw owners and bindings', async () => {
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group" descr="Group description"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp>${shape({ id: '2' }).slice('<p:sp>'.length, -'</p:sp>'.length)}</p:sp></p:grpSp>`;
  const context = await structureContext(bytes(group)), root = await context.root('/deck/z.xml', 'sld');
  const tree = child(child(root, context.p, 'cSld'), context.p, 'spTree'), owner = child(tree, context.p, 'grpSp');
  const nv = child(owner, context.p, 'nvGrpSpPr'), cnv = child(nv, context.p, 'cNvPr');
  const nested = child(owner, context.p, 'sp'), nestedNv = child(nested, context.p, 'nvSpPr'), nestedCnv = child(nestedNv, context.p, 'cNvPr');
  const a = await bindDrawingDescription(context, '/deck/z.xml', 's', 'group', '4', owner, nv, cnv, 'Group description');
  const b = await bindDrawingDescription(context, '/deck/z.xml', 's', 'child', '2', nested, nestedNv, nestedCnv, 'Description');
  assert.equal(a.status, 'bound'); assert.equal(b.status, 'bound'); assert.ok(a.elementIndex < b.ownerIndex);
  assert.equal((await bindDrawingDescription(context, '/deck/z.xml', 's', 'child', '2', owner, nestedNv, nestedCnv, 'Description')).status, 'unmapped');
});

test('description snapshot analyzer admits one bound existing value and blocks unrelated changes', async () => {
  const context = await structureContext(bytes(shape({ id: '2' }) + shape({ id: '3' }))), checkpoint = async () => {};
  const projection = await projectStructure(await readStructureFromContext(context), DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), bindings = [];
  await projectContent(projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, ...Array(20).fill(undefined), bindings);
  const document = projection.document, slide = document.slides[0], baseline = normalizePresentation(document), first = slide.elements[0], second = slide.elements[1];
  slide.replaceElement(first.id, { ...first, accessibility: { ...first.accessibility, description: '新的 & description 😀' } });
  const candidate = normalizePresentation(document), ready = await analyzeDescriptionEdit(baseline, candidate, bindings, DEFAULT_SESSION_LIMITS, checkpoint);
  assert.equal(ready.diagnostics.length, 0); assert.equal(ready.descriptions[0].value, '新的 & description 😀'); assert.equal(ready.changes[0].source.sourceObjectId, '2');
  const mutated = structuredClone(candidate); mutated.slides[0].elements[0].box.x += 1;
  assert.equal((await analyzeDescriptionEdit(baseline, mutated, bindings, DEFAULT_SESSION_LIMITS, checkpoint)).diagnostics[0].code, 'save-description-edit-unsupported');
  slide.replaceElement(second.id, { ...second, accessibility: { ...second.accessibility, description: 'second edit' } });
  assert.equal((await analyzeDescriptionEdit(baseline, normalizePresentation(document), bindings, DEFAULT_SESSION_LIMITS, checkpoint)).diagnostics[0].code, 'save-description-edit-unsupported');
  await assert.rejects(analyzeDescriptionEdit(baseline, candidate, bindings, { ...DEFAULT_SESSION_LIMITS, maxSaveAnalysisWork: 1 }, checkpoint), error => error.code === 'save-limit-exceeded');
});

test('private candidate round-trips one bound description and preserves unrelated payloads for three rounds', async () => {
  for (const strict of [false, true]) {
    let source = bytes(shape({ id: '2' }) + shape({ id: '3' }), strict), expected = 'Description';
    for (const value of ['A & B', '中文😀', 'tab\tquote"']) {
      const context = await structureContext(source), checkpoint = async () => {};
      const projection = await projectStructure(await readStructureFromContext(context), DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), bindings = [];
      await projectContent(projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, ...Array(20).fill(undefined), bindings);
      const slide = projection.document.slides[0], element = slide.elements[0], baseline = normalizePresentation(projection.document);
      slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: value } });
      const candidate = normalizePresentation(projection.document), analysis = await analyzeDescriptionEdit(baseline, candidate, bindings, DEFAULT_SESSION_LIMITS, checkpoint);
      assert.deepEqual(analysis.diagnostics, []); assert.equal(analysis.descriptions[0].binding.expected, expected);
      const result = await prepareTextCandidate(source, candidate, { ...analysis, parts: new Map(), chartValues: [] }, {}, DEFAULT_SESSION_LIMITS, checkpoint);
      assert.ok('bytes' in result, JSON.stringify(result));
      const before = readGeneratedZip(source), after = readGeneratedZip(result.bytes);
      assert.deepEqual([...before.keys()], [...after.keys()]);
      for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
      assert.equal((after.get('deck/z.xml').toString().match(/descr=/g) ?? []).length, 2);
      source = result.bytes; expected = value;
    }
  }
});

test('explicit description preservation profile saves one source attribute for three rounds', async () => {
  for (const strict of [false, true]) {
    let source = bytes(shape({ id: '2' }) + shape({ id: '3' }), strict), old = 'Description', lexical = 'Description';
    for (const [value, encoded] of [['A & B', 'A &amp; B'], ['中文😀', '中文😀'], ['quote " tab\t', 'quote &quot; tab&#x9;']]) {
      const session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
      const slide = session.document.slides[0], element = slide.elements[0];
      const finding = session.report.findings.find(item => item.feature === 'drawing.description-existing' && item.elementId === element.id);
      assert.equal(finding.tier, 'editable');
      slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: value } });
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      assert.equal(plan.plan.report.assurance, 'scoped-preservation'); assert.equal(plan.plan.report.changes[0].source.sourceObjectId, '2');
      const output = (await session.save(plan.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
      for (const [name, payload] of before) {
        if (name === 'deck/z.xml') assert.equal(after.get(name).toString(), payload.toString().replace(`descr="${lexical}"`, `descr="${encoded}"`));
        else assert.deepEqual(after.get(name), payload, name);
      }
      session.dispose(); source = output; old = value; lexical = encoded;
      const reopened = await importPptx(source, { projection: 'content' }); assert.equal(reopened.document.slides[0].elements[0].accessibility.description, old); reopened.dispose();
    }
  }
});

test('description profile rejects unbound, mixed, duplicate and stale changes', async () => {
  const source = bytes(shape({ id: '2' }) + shape({ id: '3' }));
  const open = () => importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const s = await open(), slide = s.document.slides[0], first = slide.elements[0], second = slide.elements[1];
  const noop = await s.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready'); assert.deepEqual(Buffer.from((await s.save(noop.plan)).bytes), Buffer.from(source));
  slide.replaceElement(first.id, { ...first, accessibility: { ...first.accessibility, description: 'changed' } });
  const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready');
  slide.replaceElement(second.id, { ...second, accessibility: { ...second.accessibility, description: 'second' } });
  await assert.rejects(s.save(ready.plan), error => error.code === 'save-snapshot-stale');
  assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); s.dispose();
  const mixed = await open(), page = mixed.document.slides[0], object = page.elements[0];
  page.replaceElement(object.id, { ...object, box: { ...object.box, x: object.box.x + 1 }, accessibility: { ...object.accessibility, description: 'changed' } });
  assert.equal((await mixed.prepareSave({ mode: 'edit' })).status, 'blocked'); mixed.dispose();
  const unbound = await importPptx(bytes(shape({ id: '2' }).replace(' descr="Description"', '')), { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const unboundSlide = unbound.document.slides[0], unboundObject = unboundSlide.elements[0];
  unboundSlide.replaceElement(unboundObject.id, { ...unboundObject, accessibility: { description: 'new' } });
  assert.equal((await unbound.prepareSave({ mode: 'edit' })).status, 'blocked'); unbound.dispose();
  await assert.rejects(importPptx(source, { projection: 'structure', descriptionEditSaveProfile: 'existing-v1' }), TypeError);
  await assert.rejects(importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1', editSaveProfile: 'text-runs-v1' }), TypeError);
  const ordinary = await importPptx(source, { projection: 'content' }), ordinarySlide = ordinary.document.slides[0], ordinaryObject = ordinarySlide.elements[0];
  ordinarySlide.replaceElement(ordinaryObject.id, { ...ordinaryObject, accessibility: { ...ordinaryObject.accessibility, description: 'not admitted' } });
  assert.equal((await ordinary.prepareSave({ mode: 'edit' })).status, 'blocked'); ordinary.dispose();
  const v1Removal = await open(), v1Slide = v1Removal.document.slides[0], v1Object = structuredClone(v1Slide.elements[0]);
  delete v1Object.accessibility.description; v1Slide.replaceElement(v1Object.id, v1Object);
  assert.equal((await v1Removal.prepareSave({ mode: 'edit' })).status, 'blocked'); v1Removal.dispose();
});

test('description profile binds and saves nested group owner without touching siblings', async () => {
  const childShape = shape({ id: '2' }).replace('descr="Description"', 'descr="Child alt"');
  const groupXml = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group" descr="Group alt"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/><a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${childShape}</p:grpSp>`;
  const source = bytes(groupXml), session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const slide = session.document.slides[0], root = structuredClone(slide.elements[0]);
  root.children[0].accessibility.description = 'Changed child'; slide.replaceElement(root.id, root);
  const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  const saved = await session.save(plan.plan), before = readGeneratedZip(source), after = readGeneratedZip(saved.bytes);
  for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
  assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace('descr="Child alt"', 'descr="Changed child"'));
  const reopened = await importPptx(saved.bytes, { projection: 'content' });
  assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'Group alt');
  assert.equal(reopened.document.slides[0].elements[0].children[0].accessibility.description, 'Changed child');
  reopened.dispose(); session.dispose();
  const groupSession = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const groupSlide = groupSession.document.slides[0], group = groupSlide.elements[0];
  groupSlide.replaceElement(group.id, { ...group, accessibility: { ...group.accessibility, description: 'Changed group' } });
  const groupPlan = await groupSession.prepareSave({ mode: 'edit' }); assert.equal(groupPlan.status, 'ready', JSON.stringify(groupPlan.report));
  const changedXml = readGeneratedZip((await groupSession.save(groupPlan.plan)).bytes).get('deck/z.xml').toString();
  assert.equal(changedXml, before.get('deck/z.xml').toString().replace('descr="Group alt"', 'descr="Changed group"'));
  groupSession.dispose();
});

test('description preservation respects inactive closure and package signature guards', async () => {
  for (const type of ['urn:resource', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject']) {
    const fixture = contentFixture();
    fixture.parts['deck/_rels/z.xml.rels'] = (fixture.parts['deck/_rels/z.xml.rels'] ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
      .replace('<Relationships ', '<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace('</Relationships>', `<mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="hidden" Type="${type}" Target="missing.bin"/></mc:Choice><mc:Fallback/></mc:AlternateContent></Relationships>`);
    const session = await importPptx(encodeStructure(fixture), { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
    const slide = session.document.slides[0], element = slide.elements[0];
    slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: 'changed' } });
    const plan = await session.prepareSave({ mode: 'edit' });
    assert.equal(plan.status, 'blocked'); assert.equal(plan.report.diagnostics[0].code, type === 'urn:resource' ? 'save-source-closure-incomplete' : 'save-text-package-policy');
    session.dispose();
  }
});

test('description edit preserves unknown nonvisual extension and neighboring lexical bytes', async () => {
  const original = shape({ id: '2' });
  const decorated = original.replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', '<p:cNvPr id="2" name="Object 2" descr="Description" custom="unchanged"><a:extLst><a:ext uri="urn:future"><x:unknown xmlns:x="urn:future" value="&amp;"/></a:ext></a:extLst></p:cNvPr>');
  assert.notEqual(decorated, original);
  const source = bytes(decorated), session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' });
  const slide = session.document.slides[0], element = slide.elements[0];
  slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: 'New & safe' } });
  const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  const before = readGeneratedZip(source), after = readGeneratedZip((await session.save(plan.plan)).bytes);
  assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace('descr="Description"', 'descr="New &amp; safe"'));
  for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
  session.dispose();
});

test('attribute-v2 distinguishes absent and empty descriptions over insert, edit and remove in both dialects', async () => {
  for (const strict of [false, true]) {
    let source = bytes(shape({ id: '2' }).replace(' descr="Description"', '') + shape({ id: '3' }), strict);
    const values = ['', '中文 & 😀', undefined];
    for (const value of values) {
      const session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' });
      const slide = session.document.slides[0], element = slide.elements[0], next = structuredClone(element);
      assert.equal(session.report.findings.find(item => item.feature === 'drawing.description-attribute' && item.elementId === element.id).tier, 'editable');
      if (value === undefined) delete next.accessibility.description; else next.accessibility = { ...next.accessibility, description: value };
      slide.replaceElement(element.id, next);
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      const before = readGeneratedZip(source), output = (await session.save(plan.plan)).bytes, after = readGeneratedZip(output);
      for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
      const xml = after.get('deck/z.xml').toString();
      const priorXml = before.get('deck/z.xml').toString();
      const expectedXml = value === '' ? priorXml.replace('name="Object 2" />', 'name="Object 2"  descr=""/>') : value === undefined ? priorXml.replace('descr="中文 &amp; 😀"', '') : priorXml.replace('descr=""', 'descr="中文 &amp; 😀"');
      assert.notEqual(expectedXml, priorXml); assert.equal(xml, expectedXml);
      if (value === undefined) assert.equal((xml.match(/descr=/g) ?? []).length, 1);
      else assert.match(xml, value === '' ? /descr=""/ : /descr="中文 &amp; 😀"/);
      const reopened = await importPptx(output, { projection: 'content' });
      assert.equal(reopened.document.slides[0].elements[0].accessibility?.description, value);
      assert.equal(reopened.document.slides[0].elements[1].accessibility.description, 'Description');
      reopened.dispose(); session.dispose(); source = output;
    }
  }
});

test('attribute-v2 rejects unrelated edits and profile combinations', async () => {
  const source = bytes(shape({ id: '2' }).replace(' descr="Description"', ''));
  const session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' });
  const slide = session.document.slides[0], element = slide.elements[0];
  slide.replaceElement(element.id, { ...element, box: { ...element.box, x: element.box.x + 1 }, accessibility: { description: 'new' } });
  assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); session.dispose();
  await assert.rejects(importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'attribute-v2', firstSlideNumberSaveProfile: 'attribute-v2' }), TypeError);
  const two = await importPptx(bytes(shape({ id: '2' }).replace(' descr="Description"', '') + shape({ id: '3' }).replace(' descr="Description"', '')), { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' });
  const twoSlide = two.document.slides[0];
  for (const element of [...twoSlide.elements]) twoSlide.replaceElement(element.id, { ...element, accessibility: { description: 'new' } });
  assert.equal((await two.prepareSave({ mode: 'edit' })).status, 'blocked'); two.dispose();
});

test('description profile presence-state matrix preserves v1 compatibility and v2 semantics', async () => {
  const states = [undefined, '', 'value'];
  for (const profile of ['existing-v1', 'attribute-v2']) for (const before of states) for (const after of states) {
    const markup = shape({ id: '2' }).replace(' descr="Description"', before === undefined ? '' : ` descr="${before}"`);
    const session = await importPptx(bytes(markup), { projection: 'content', descriptionEditSaveProfile: profile });
    const slide = session.document.slides[0], element = slide.elements[0], next = structuredClone(element);
    if (after === undefined) { if (next.accessibility) delete next.accessibility.description; }
    else next.accessibility = { ...next.accessibility, description: after };
    slide.replaceElement(element.id, next);
    const plan = await session.prepareSave({ mode: 'edit' });
    const admitted = before !== after && (profile === 'attribute-v2' || before !== undefined && after !== undefined);
    assert.equal(plan.status, admitted ? 'ready' : 'blocked', `${profile}: ${String(before)} -> ${String(after)}: ${JSON.stringify(plan.report)}`);
    session.dispose();
  }
});

test('batch-v3 saves mixed description operations across two objects and two slide parts over three rounds', async () => {
  for (const strict of [false, true]) {
    const fixture = contentFixture(shape({ id: '2' }) + shape({ id: '3' }).replace(' descr="Description"', ''), strict);
    fixture.parts['deck/a.xml'] = fixture.p('sld', `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({ id: '2' })}${shape({ id: '3' })}</p:spTree></p:cSld>`, `xmlns:a="${fixture.A}"`);
    let source = encodeStructure(fixture);
    let previous = ['Description', undefined, 'Description', 'Description'];
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'batch-v3' });
      const slides = session.document.slides;
      const values = [`first ${round}`, `added ${round}`, undefined, `fourth ${round}`];
      if (round > 1) values[2] = `restored ${round}`;
      const flattened = slides.flatMap(slide => slide.elements.map(element => ({ slide, element })));
      for (let i = 0; i < flattened.length; i++) {
        const { slide, element } = flattened[i], next = structuredClone(element);
        if (values[i] === undefined) delete next.accessibility.description; else next.accessibility = { ...next.accessibility, description: values[i] };
        slide.replaceElement(element.id, next);
      }
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      assert.equal(plan.plan.report.changes.length, 4);
      const before = readGeneratedZip(source), result = (await session.save(plan.plan)).bytes, after = readGeneratedZip(result);
      assert.deepEqual([...before.keys()], [...after.keys()]);
      for (const [name, payload] of before) if (!['deck/z.xml', 'deck/a.xml'].includes(name)) assert.deepEqual(after.get(name), payload, name);
      for (const [part, base] of [['deck/z.xml', 0], ['deck/a.xml', 2]]) {
        let expected = before.get(part).toString();
        for (let offset = 0; offset < 2; offset++) {
          const index = base + offset, marker = `name="Object ${offset + 2}"`, start = expected.indexOf(marker), end = expected.indexOf('/>', start);
          assert.ok(start >= 0 && end > start);
          const segment = expected.slice(start, end);
          let updated;
          if (previous[index] === undefined) updated = `${segment} descr="${values[index]}"`;
          else if (values[index] === undefined) updated = segment.replace(`descr="${previous[index]}"`, '');
          else updated = segment.replace(`descr="${previous[index]}"`, `descr="${values[index]}"`);
          assert.notEqual(segment, updated); expected = expected.slice(0, start) + updated + expected.slice(end);
        }
        assert.equal(after.get(part).toString(), expected);
      }
      const reopened = await importPptx(result, { projection: 'content' });
      assert.deepEqual(reopened.document.slides.flatMap(slide => slide.elements.map(element => element.accessibility?.description)), values);
      reopened.dispose(); session.dispose(); source = result; previous = values;
    }
  }
});

test('batch-v3 rejects mixed changes, stale plans and exhausted budgets atomically', async () => {
  const source = bytes(shape({ id: '2' }) + shape({ id: '3' }));
  const open = extra => importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'batch-v3', ...extra });
  const editBoth = session => {
    const slide = session.document.slides[0];
    for (const element of [...slide.elements]) slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: `changed ${element.id}` } });
  };
  const limited = await open({ sessionLimits: { maxSaveChanges: 1 } }); editBoth(limited);
  await assert.rejects(limited.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded'); limited.dispose();
  const lowWork = await open({ sessionLimits: { maxSaveValidationWork: 100 } }); editBoth(lowWork);
  await assert.rejects(lowWork.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded'); lowWork.dispose();
  const lowBytes = await open({ sessionLimits: { maxSaveValidationBytes: 100 } }); editBoth(lowBytes);
  await assert.rejects(lowBytes.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded'); lowBytes.dispose();
  const session = await open(); editBoth(session);
  const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
  const slide = session.document.slides[0], first = slide.elements[0];
  slide.replaceElement(first.id, { ...first, box: { ...first.box, x: first.box.x + 1 } });
  await assert.rejects(session.save(plan.plan), error => error.code === 'save-snapshot-stale');
  assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
  session.dispose();
});

test('batch description analysis detects disagreeing views of one raw source attribute', async () => {
  const context = await structureContext(bytes(shape({ id: '2' }) + shape({ id: '3' }))), checkpoint = async () => {};
  const projection = await projectStructure(await readStructureFromContext(context), DEFAULT_SESSION_LIMITS.maxReportFindings, checkpoint), bindings = [];
  await projectContent(projection.document, projection.report, context, DEFAULT_SESSION_LIMITS, checkpoint, {}, ...Array(20).fill(undefined), bindings, true);
  const baseline = normalizePresentation(projection.document), slide = projection.document.slides[0], first = slide.elements[0];
  slide.replaceElement(first.id, { ...first, accessibility: { ...first.accessibility, description: 'changed' } });
  const alias = { ...bindings[1], part: bindings[0].part, elementIndex: bindings[0].elementIndex };
  const result = await analyzeDescriptionEdit(baseline, normalizePresentation(projection.document), [bindings[0], alias], DEFAULT_SESSION_LIMITS, checkpoint, true);
  assert.equal(result.diagnostics[0].code, 'save-source-shared-conflict'); assert.deepEqual(result.descriptions, []);
});

test('mixed-v4 atomically saves text and descriptions in one source part for three rounds', async () => {
  for (const strict of [false, true]) {
    let source = bytes(shape({ id: '2', txBox: true, text: body() }) + shape({ id: '3' }), strict), oldText = ' A &amp; B ', oldFirst = 'Description', oldSecond = 'Description';
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v4' });
      const slide = session.document.slides[0], first = structuredClone(slide.elements[0]), second = structuredClone(slide.elements[1]);
      first.content[0].runs[0].text = `Text ${round}`; first.accessibility.description = `First ${round}`; second.accessibility.description = `Second ${round}`;
      slide.replaceElement(first.id, first); slide.replaceElement(second.id, second);
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      assert.equal(plan.plan.report.changes.length, 3);
      const before = readGeneratedZip(source), output = (await session.save(plan.plan)).bytes, after = readGeneratedZip(output);
      for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
      const expected = before.get('deck/z.xml').toString().replace(`>${oldText}</a:t>`, `>Text ${round}</a:t>`).replace(`name="Object 2" descr="${oldFirst}"`, `name="Object 2" descr="First ${round}"`).replace(`name="Object 3" descr="${oldSecond}"`, `name="Object 3" descr="Second ${round}"`);
      assert.equal(after.get('deck/z.xml').toString(), expected);
      const reopened = await importPptx(output, { projection: 'content' });
      assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, `Text ${round}`);
      assert.deepEqual(reopened.document.slides[0].elements.map(element => element.accessibility.description), [`First ${round}`, `Second ${round}`]);
      reopened.dispose(); session.dispose(); source = output; oldText = `Text ${round}`; oldFirst = `First ${round}`; oldSecond = `Second ${round}`;
    }
  }
});

test('mixed-v4 combines table cell and notes text with the table description', async () => {
  for (const strict of [false, true]) {
    const fixture = notesFixture(undefined, strict), slideSource = tableFixture(table(), strict);
    fixture.parts['deck/z.xml'] = slideSource.parts['deck/z.xml'];
    let source = encodeStructure(fixture);
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v4' });
      const slide = session.document.slides[0], original = slide.elements[0], changed = structuredClone(original), notes = structuredClone(slide.notes);
      changed.accessibility.description = `Table alt ${round}`; changed.rows[0].cells[0].content[0].runs[0].text = `Cell ${round}`;
      notes[0].runs[0].text = `Notes ${round}`; slide.replaceElement(original.id, changed); slide.setNotes(notes);
      const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
      assert.equal(plan.plan.report.changes.length, 3);
      const before = readGeneratedZip(source), output = (await session.save(plan.plan)).bytes, after = readGeneratedZip(output);
      for (const [name, payload] of before) if (!['deck/z.xml', 'deck/notes/z.xml'].includes(name)) assert.deepEqual(after.get(name), payload, name);
      const oldAlt = round === 1 ? 'Table description' : `Table alt ${round - 1}`, oldCell = round === 1 ? 'R1C1' : `Cell ${round - 1}`, oldNotes = round === 1 ? ' A &amp; B ' : `Notes ${round - 1}`;
      assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace(`descr="${oldAlt}"`, `descr="Table alt ${round}"`).replace(`>${oldCell}</a:t>`, `>Cell ${round}</a:t>`));
      assert.equal(after.get('deck/notes/z.xml').toString(), before.get('deck/notes/z.xml').toString().replace(`>${oldNotes}</a:t>`, `>Notes ${round}</a:t>`));
      const reopened = await importPptx(output, { projection: 'content' });
      assert.equal(reopened.document.slides[0].elements[0].accessibility.description, `Table alt ${round}`);
      assert.equal(reopened.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].text, `Cell ${round}`);
      assert.equal(reopened.document.slides[0].notes[0].runs[0].text, `Notes ${round}`);
      reopened.dispose(); session.dispose(); source = output;
    }
  }
});

test('mixed-v4 admits either domain alone and blocks unsupported mixed edits and shared budgets', async () => {
  const source = bytes(shape({ id: '2', txBox: true, text: body() }));
  const open = extra => importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v4', ...extra });
  for (const domain of ['text', 'description']) {
    const session = await open(), slide = session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
    if (domain === 'text') next.content[0].runs[0].text = 'Only text'; else next.accessibility.description = 'Only description';
    slide.replaceElement(original.id, next);
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, 1);
    session.dispose();
  }
  const mixed = await open(), slide = mixed.document.slides[0], element = structuredClone(slide.elements[0]);
  element.content[0].runs[0].text = 'changed'; element.accessibility.description = 'changed'; element.box.x += 1; slide.replaceElement(element.id, element);
  assert.equal((await mixed.prepareSave({ mode: 'edit' })).status, 'blocked'); mixed.dispose();
  const limited = await open({ sessionLimits: { maxSaveChanges: 1 } }), limitedSlide = limited.document.slides[0], limitedElement = structuredClone(limitedSlide.elements[0]);
  limitedElement.content[0].runs[0].text = 'changed'; limitedElement.accessibility.description = 'changed'; limitedSlide.replaceElement(limitedElement.id, limitedElement);
  await assert.rejects(limited.prepareSave({ mode: 'edit' }), error => error.code === 'save-limit-exceeded'); limited.dispose();
  const stale = await open(), staleSlide = stale.document.slides[0], staleElement = structuredClone(staleSlide.elements[0]);
  staleElement.content[0].runs[0].text = 'text'; staleElement.accessibility.description = 'alt'; staleSlide.replaceElement(staleElement.id, staleElement);
  const stalePlan = await stale.prepareSave({ mode: 'edit' }); assert.equal(stalePlan.status, 'ready');
  staleElement.accessibility.description = 'later'; staleSlide.replaceElement(staleElement.id, staleElement);
  await assert.rejects(stale.save(stalePlan.plan), error => error.code === 'save-snapshot-stale'); stale.dispose();
  await assert.rejects(importPptx(source, { projection: 'content', descriptionEditSaveProfile: 'mixed-v4' }), TypeError);
  await assert.rejects(importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v2', descriptionEditSaveProfile: 'mixed-v4' }), TypeError);
});

test('directly bound drawing names preserve all unrelated source bytes across dialects and three rounds', async () => {
  const opts = { projection: 'content', drawingNameSaveProfile: 'existing-v1' };
  for (const strict of [false, true]) {
    let source = encodeStructure(contentFixture(shape({ id: '2' }) + shape({ id: '3' }), strict));
    for (const [round, prefix] of ['New & 中文', '', 'Final 😀'].entries()) {
      const s = await importPptx(source, opts), before = readGeneratedZip(source);
      try {
        const bindings = s.report.findings.filter(f => f.feature === 'drawing.name-existing'); assert.equal(bindings.length, 2); assert.ok(bindings.every(f => f.tier === 'editable'));
        const slide = s.document.slides[0], oldNames = slide.elements.map(e => e.name);
        for (const [i, old] of [...slide.elements].entries()) { const edited = structuredClone(old); edited.name = `${prefix}${i}`; slide.replaceElement(old.id, edited); }
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 2); assert.ok(p.plan.report.changes.every(c => c.documentPath.endsWith('/name')));
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next);
        for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
        const escape = value => value.replaceAll('&', '&amp;');
        let expected = before.get('deck/z.xml').toString(); for (const [i, old] of oldNames.entries()) expected = expected.replace(`name="${escape(old)}"`, `name="${escape(prefix + i)}"`);
        assert.equal(after.get('deck/z.xml').toString(), expected);
        const reopened = await importPptx(next, opts); try { assert.deepEqual(reopened.document.slides[0].elements.map(e => e.name), [`${prefix}0`, `${prefix}1`]); } finally { reopened.dispose(); }
        source = next; void round;
      } finally { s.dispose(); }
    }
  }
});

test('drawing name save rejects missing source names, unrelated changes and incompatible profiles', async () => {
  const opts = { projection: 'content', drawingNameSaveProfile: 'existing-v1' };
  const missing = encodeStructure(contentFixture(shape({ id: '2' }).replace(' name="Object 2"', '')));
  const s = await importPptx(missing, opts);
  try { assert.equal(s.report.findings.find(f => f.feature === 'drawing.name-existing').tier, 'unsupported'); const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'Inserted'; slide.replaceElement(old.id, edit); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { s.dispose(); }
  const source = bytes(shape({ id: '2' })), good = await importPptx(source, opts);
  try { const slide = good.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'Changed'; edit.box.x++; slide.replaceElement(old.id, edit); assert.equal((await good.prepareSave({ mode: 'edit' })).status, 'blocked'); edit.box.x--; slide.replaceElement(old.id, edit); const p = await good.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); edit.name = 'Later'; slide.replaceElement(old.id, edit); await assert.rejects(good.save(p.plan), { code: 'save-snapshot-stale' }); } finally { good.dispose(); }
  for (const extra of [{ projection: 'structure' }, { descriptionEditSaveProfile: 'batch-v3' }, { editSaveProfile: 'text-runs-v3' }, { chartEditSaveProfile: 'existing-values-v1' }, { firstSlideNumberSaveProfile: 'attribute-v2' }, { drawingNameSaveProfile: 'other' }]) await assert.rejects(importPptx(source, { ...opts, ...extra }), TypeError);
});

test('drawing name edits enforce shared limits and reject a name on MCE-promoted ancestry', async () => {
  const opts = { projection: 'content', drawingNameSaveProfile: 'existing-v1' }, source = bytes(shape({ id: '2' }) + shape({ id: '3' }));
  const limited = await importPptx(source, { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try { const slide = limited.document.slides[0]; for (const old of [...slide.elements]) { const edit = structuredClone(old); edit.name += '!'; slide.replaceElement(old.id, edit); } await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  const f = contentFixture(shape({ id: '2' })); const xml = f.parts['deck/z.xml'], cnv = '<p:cNvPr id="2" name="Object 2" descr="Description" />'; assert.ok(xml.includes(cnv));
  f.parts['deck/z.xml'] = xml.replace(cnv, `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:x="urn:unsupported" Requires="x"><p:cNvPr id="22" name="Other"/></mc:Choice><mc:Fallback>${cnv}</mc:Fallback></mc:AlternateContent>`);
  const wrapped = await importPptx(encodeStructure(f), opts);
  try { const finding = wrapped.report.findings.find(item => item.feature === 'drawing.name-existing'); assert.equal(finding.tier, 'unsupported'); const slide = wrapped.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'Unsafe'; slide.replaceElement(old.id, edit); const result = await wrapped.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined); } finally { wrapped.dispose(); }
});

test('combined name and description edits isolate unchanged sibling attributes and share change limits', async () => {
  const source = bytes(shape({ id: '2', txBox: true, text: body() }));
  const opts = { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1' };
  for (const field of ['name', 'description']) {
    const s = await importPptx(source, opts);
    try { const slide = s.document.slides[0], original = slide.elements[0], edit = structuredClone(original); if (field === 'name') edit.name = 'New name'; else edit.accessibility.description = 'New alt'; slide.replaceElement(original.id, edit); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 1); const next = await importPptx((await s.save(p.plan)).bytes, opts); try { assert.equal(next.document.slides[0].elements[0].name, field === 'name' ? 'New name' : original.name); assert.equal(next.document.slides[0].elements[0].accessibility.description, field === 'description' ? 'New alt' : original.accessibility.description); } finally { next.dispose(); } } finally { s.dispose(); }
  }
  const limited = await importPptx(source, { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try { const slide = limited.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'New'; edit.accessibility.description = 'Alt'; slide.replaceElement(old.id, edit); await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  const s = await importPptx(source, opts);
  try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'New'; edit.accessibility.description = 'Alt'; slide.replaceElement(old.id, edit); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); edit.name = 'Changed again'; slide.replaceElement(old.id, edit); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' }); } finally { s.dispose(); }
});

test('mixed profile admits bound descriptions while rejecting an unbound missing name', async () => {
  const opts = { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1' };
  const source = bytes(shape({ id: '2' }).replace(' name="Object 2"', '')), s = await importPptx(source, opts);
  try {
    const nameFinding = s.report.findings.find(f => f.feature === 'drawing.name-existing'), descriptionFinding = s.report.findings.find(f => f.feature === 'drawing.description-attribute'); assert.equal(nameFinding.tier, 'unsupported'); assert.equal(descriptionFinding.tier, 'editable');
    const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.accessibility.description = 'New alt'; slide.replaceElement(old.id, edit);
    const ready = await s.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report)); assert.equal(ready.plan.report.changes.length, 1);
    edit.name = 'Forbidden'; slide.replaceElement(old.id, edit); const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
  } finally { s.dispose(); }
});

test('existing decorative values preserve unknown extensions across dialects and three rounds', async () => {
  const uri = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}', extension = `<a:extLst><a:ext uri="urn:future"><x:keep xmlns:x="urn:future" value="&amp;"/></a:ext><a:ext uri="${uri}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="true"/></a:ext></a:extLst>`;
  const original = shape({ id: '2' }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${extension}</p:cNvPr>`), opts = { projection: 'content', decorativeSaveProfile: 'existing-v1' };
  for (const strict of [false, true]) {
    let source = bytes(original, strict), old = 'true';
    for (const value of [false, true, false]) {
      const before = readGeneratedZip(source), s = await importPptx(source, opts);
      try {
        const binding = s.report.findings.find(f => f.feature === 'drawing.decorative-existing'); assert.equal(binding.tier, 'editable');
        const slide = s.document.slides[0], element = slide.elements[0], edit = structuredClone(element); edit.accessibility.decorative = value; slide.replaceElement(element.id, edit);
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 1); assert.ok(p.plan.report.changes[0].documentPath.endsWith('/accessibility/decorative'));
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next), lexical = value ? '1' : '0';
        for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name); else assert.equal(after.get(name).toString(), payload.toString().replace(`val="${old}"`, `val="${lexical}"`));
        const reopened = await importPptx(next, opts); try { assert.equal(reopened.document.slides[0].elements[0].accessibility.decorative, value); } finally { reopened.dispose(); }
        source = next; old = lexical;
      } finally { s.dispose(); }
    }
  }
});

test('decorative, name, description and text edits share one guarded mixed candidate', async () => {
  const extension = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="0"/></a:ext></a:extLst>';
  const original = shape({ id: '2', txBox: true, text: body() }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${extension}</p:cNvPr>`);
  const opts = { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', decorativeSaveProfile: 'existing-v1' };
  for (const strict of [false, true]) {
    let source = bytes(original, strict);
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(source), s = await importPptx(source, opts);
      try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = `Name ${round}`; edit.accessibility.description = `Alt ${round}`; edit.accessibility.decorative = round % 2 === 0; edit.content[0].runs[0].text += '!'; slide.replaceElement(old.id, edit);
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 4);
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next); for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
        const reopened = await importPptx(next, opts); try { const e = reopened.document.slides[0].elements[0]; assert.equal(e.name, edit.name); assert.equal(e.accessibility.description, edit.accessibility.description); assert.equal(e.accessibility.decorative, edit.accessibility.decorative); assert.equal(e.content[0].runs[0].text, edit.content[0].runs[0].text); } finally { reopened.dispose(); } source = next;
      } finally { s.dispose(); }
    }
  }
});

test('managed edit batches compose with guarded source saves and preserve unrelated parts for three rounds', async () => {
  const extension = '<a:extLst><a:ext uri="urn:future"><x:keep xmlns:x="urn:future" value="&amp;"/></a:ext><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="0"/></a:ext></a:extLst>';
  const original = shape({ id: '2', txBox: true, text: body() }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${extension}</p:cNvPr>`);
  const opts = { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', decorativeSaveProfile: 'existing-v1' };
  for (const strict of [false, true]) {
    let source = bytes(original, strict);
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(source), session = await importPptx(source, opts);
      try {
        const doc = session.document, slide = doc.slides[0], old = slide.elements[0];
        const target = { slideId: slide.id, elementId: old.id }, text = old.content[0].runs[0].text;
        const values = { name: `Batch ${round}`, description: `描述 & ${round}`, decorative: !old.accessibility.decorative, text: `${text}!` };
        const request = { schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
          { type: 'element-name', ...target, expected: old.name, value: values.name },
          { type: 'element-accessibility', ...target, expected: { description: old.accessibility.description, decorative: old.accessibility.decorative }, value: { description: values.description, decorative: values.decorative } },
          { type: 'text-run', ...target, paragraphIndex: 0, runIndex: 0, expected: text, value: values.text },
        ] };
        const inverse = doc.applyEditOperation(JSON.parse(JSON.stringify(request))).inverse;
        const stale = await session.prepareSave({ mode: 'edit' }); assert.equal(stale.status, 'ready', JSON.stringify(stale.report));
        const redo = doc.applyEditOperation(inverse).inverse;
        await assert.rejects(session.save(stale.plan), error => error.code === 'save-snapshot-stale');
        doc.applyEditOperation(redo);
        const ready = await session.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        assert.equal(ready.plan.report.changes.length, 4);
        const output = (await session.save(ready.plan)).bytes, after = readGeneratedZip(output);
        for (const [part, payload] of before) {
          if (part !== 'deck/z.xml') assert.deepEqual(after.get(part), payload, part);
          else assert.match(after.get(part).toString(), /<x:keep xmlns:x="urn:future" value="&amp;"\/>/);
        }
        const reopened = await importPptx(output, opts);
        try { const e = reopened.document.slides[0].elements[0]; assert.equal(e.name, values.name); assert.equal(e.accessibility.description, values.description); assert.equal(e.accessibility.decorative, values.decorative); assert.equal(e.content[0].runs[0].text, values.text); }
        finally { reopened.dispose(); }
        source = output;
      } finally { session.dispose(); }
    }
  }
});

test('managed accessibility edits do not bypass source binding and undo permits byte-exact no-op', async () => {
  const source = bytes(shape({ id: '2' })), opts = { projection: 'content', decorativeSaveProfile: 'existing-v1' };
  const session = await importPptx(source, opts);
  try {
    const doc = session.document, slide = doc.slides[0], element = slide.elements[0];
    const inverse = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [{
      type: 'element-accessibility', slideId: slide.id, elementId: element.id,
      expected: { description: 'Description', decorative: null }, value: { description: 'Description', decorative: true },
    }] }).inverse;
    const blocked = await session.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked');
    doc.applyEditOperation(inverse);
    const noop = await session.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready');
    assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(source));
  } finally { session.dispose(); }
});

test('decorative saves reject absent, ambiguous and unrelated changes and enforce budgets', async () => {
  const opts = { projection: 'content', decorativeSaveProfile: 'existing-v1' }, bare = bytes(shape({ id: '2' })), missing = await importPptx(bare, opts);
  try { assert.equal(missing.report.findings.find(f => f.feature === 'drawing.decorative-existing').tier, 'unsupported'); const slide = missing.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.accessibility = { ...edit.accessibility, decorative: true }; slide.replaceElement(old.id, edit); assert.equal((await missing.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { missing.dispose(); }
  const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
  const sh = shape({ id: '2' }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${ext}</p:cNvPr>`), source = bytes(sh);
  const s = await importPptx(source, opts); try { const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.accessibility.decorative = false; edit.box.x++; slide.replaceElement(old.id, edit); const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); edit.box.x--; slide.replaceElement(old.id, edit); const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); edit.name = 'Changed'; slide.replaceElement(old.id, edit); await assert.rejects(s.save(p.plan), { code: 'save-snapshot-stale' }); } finally { s.dispose(); }
  const duplicated = bytes(sh.replace(ext, ext.replace('</a:extLst>', ext.slice('<a:extLst>'.length, -'</a:extLst>'.length) + '</a:extLst>'))), d = await importPptx(duplicated, opts);
  try { assert.equal(d.report.findings.find(f => f.feature === 'drawing.decorative-existing').tier, 'unsupported'); } finally { d.dispose(); }
  const two = bytes(sh + sh.replaceAll('id="2"', 'id="3"').replaceAll('Object 2', 'Object 3')), limited = await importPptx(two, { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try { const slide = limited.document.slides[0]; for (const old of [...slide.elements]) { const edit = structuredClone(old); edit.accessibility.decorative = false; slide.replaceElement(old.id, edit); } await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' }); } finally { limited.dispose(); }
  for (const extra of [{ projection: 'structure' }, { drawingNameSaveProfile: 'existing-v1' }, { editSaveProfile: 'text-runs-v3' }, { decorativeSaveProfile: 'other' }]) await assert.rejects(importPptx(source, { ...opts, ...extra }), TypeError);
});

test('decorative source binding rejects MCE-promoted values and extension payloads', async () => {
  const uri = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}', target = '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>';
  const wrapped = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:x="urn:unknown" Requires="x"><x:unknown/></mc:Choice><mc:Fallback>${target}</mc:Fallback></mc:AlternateContent>`;
  for (const content of [wrapped, `${target}<x:foreign xmlns:x="urn:foreign"/>`]) {
    const extension = `<a:extLst><a:ext uri="${uri}">${content}</a:ext></a:extLst>`;
    const source = bytes(shape({ id: '2' }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${extension}</p:cNvPr>`));
    const s = await importPptx(source, { projection: 'content', decorativeSaveProfile: 'existing-v1' });
    try { assert.equal(s.report.findings.find(f => f.feature === 'drawing.decorative-existing').tier, 'unsupported'); const slide = s.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.accessibility.decorative = !edit.accessibility.decorative; slide.replaceElement(old.id, edit); assert.equal((await s.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { s.dispose(); }
  }
});

test('large real slide property batches fit a bounded transaction budget and preserve every other package part', async () => {
  const count = 200, input = bytes(Array.from({ length: count }, (_, i) => shape({ id: String(i + 2) })).join(''));
  const before = readGeneratedZip(input), s = await importPptx(input, { projection: 'content', descriptionEditSaveProfile: 'batch-v3', sessionLimits: { maxSaveValidationWork: 4_000_000 } });
  try {
    const slide = s.document.slides[0]; assert.equal(slide.elements.length, count);
    for (const [i, original] of [...slide.elements].entries()) { const next = structuredClone(original); next.accessibility.description = `Batch ${i} & 中文`; slide.replaceElement(original.id, next); }
    const plan = await s.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal(plan.plan.report.changes.length, count);
    const output = (await s.save(plan.plan)).bytes, after = readGeneratedZip(output);
    for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
    let expected = before.get('deck/z.xml').toString(); for (let i = 0; i < count; i++) expected = expected.replace(`name="Object ${i + 2}" descr="Description"`, `name="Object ${i + 2}" descr="Batch ${i} &amp; 中文"`);
    assert.equal(after.get('deck/z.xml').toString(), expected);
    const reopened = await importPptx(output, { projection: 'content' }); try { assert.deepEqual(reopened.document.slides[0].elements.map(e => e.accessibility.description), Array.from({ length: count }, (_, i) => `Batch ${i} & 中文`)); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});
