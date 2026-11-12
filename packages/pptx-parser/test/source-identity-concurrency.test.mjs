import assert from 'node:assert/strict';
import test from 'node:test';
import { structureContext } from '../dist/presentation/structure-context.js';
import { bindTextRuns } from '../dist/import/text-run-bindings.js';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { verifyTextSourceAncestry } from '../dist/import/prepare-text-candidate.js';
import { paragraphFixture } from './paragraph-target-fixture.mjs';
import { encodeStructure, structureFixture } from './structure-fixture.mjs';

test('concurrent source identity lookups share only a complete raw preorder index', async () => {
  for (const strict of [false, true]) {
    const context = await structureContext(encodeStructure(paragraphFixture({ strict, txBox: true })));
    const part = '/deck/z.xml', root = context.rawXml.get(part);
    assert.ok(root);
    const ordered = [], parents = new Map(), stack = [root];
    while (stack.length) {
      const node = stack.pop(); ordered.push(node);
      for (let i = node.children.length - 1; i >= 0; i--) if (node.children[i].kind === 'element') { parents.set(node.children[i], node); stack.push(node.children[i]); }
    }
    assert.ok(ordered.length > 5);
    const result = await Promise.all(ordered.map(node => context.sourceElement(part, node)));
    for (let i = 0; i < ordered.length; i++) {
      assert.equal(result[i].node, ordered[i]);
      assert.equal(result[i].elementIndex, i + 1);
      assert.equal(result[i].parent, parents.get(ordered[i]));
    }
    await assert.rejects(context.sourceElement('/missing.xml', root));
    assert.equal((await context.sourceElement(part, ordered.at(-1))).elementIndex, ordered.length);
  }
});

test('concurrent MCE projections retain one tree and its exact raw origins', async () => {
  for (const strict of [false, true]) {
    const fixture = structureFixture(strict), part = '/deck/z.xml';
    fixture.parts['deck/z.xml'] = fixture.parts['deck/z.xml']
      .replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace('<p:cSld name="First"/>', '<mc:AlternateContent><mc:Choice Requires="x"><p:cSld name="Hidden"/></mc:Choice><mc:Fallback><p:cSld name="First"/></mc:Fallback></mc:AlternateContent>');
    const context = await structureContext(encodeStructure(fixture));
    const roots = await Promise.all(Array.from({ length: 8 }, () => context.root(part, 'sld')));
    assert.ok(roots.every(root => root === roots[0]));
    const effective = roots[0].children.find(node => node.kind === 'element' && node.local === 'cSld');
    assert.ok(effective);
    const source = await context.sourceElement(part, effective);
    assert.equal(source.node.local, 'cSld');
    assert.equal(source.parent.local, 'Fallback');
    assert.ok(source.node.attributes.some(attribute => attribute.local === 'name' && attribute.value === 'First'));
    assert.equal(context.reports.filter(report => report.part === part).length, 1);
  }
});

test('text binding records only direct raw run and paragraph ancestry', async () => {
  for (const strict of [false, true]) {
    const fixture = paragraphFixture({ strict, txBox: true }), part = '/deck/z.xml';
    const context = await structureContext(encodeStructure(fixture)), root = context.rawXml.get(part);
    const stack = [root], atoms = [];
    while (stack.length) {
      const node = stack.pop();
      if (node.namespaceUri === context.a && node.local === 't') atoms.push(node);
      for (const child of node.children) if (child.kind === 'element') stack.push(child);
    }
    const atom = atoms.find(node => node.children.some(child => child.kind === 'text' && child.value === 'First'));
    const binding = (await bindTextRuns([{ paragraphIndex: 0, runIndex: 0, text: 'First', node: atom }], { slideId: 's', elementId: 'e' }, { part }, context, async () => {}))[0];
    const raw = await context.sourceElement(part, atom), run = await context.sourceElement(part, raw.parent), paragraph = await context.sourceElement(part, run.parent);
    assert.equal(binding.status, 'bound');
    assert.equal(binding.rawRunElementIndex, run.elementIndex);
    assert.equal(binding.rawParagraphElementIndex, paragraph.elementIndex);
    assert.equal(run.node.local, 'r'); assert.equal(paragraph.node.local, 'p');
    const lexical = await captureXmlTextSource(new TextEncoder().encode(fixture.parts['deck/z.xml']));
    const edit = { elementIndex: binding.sourceElementIndex, namespaceUri: context.a, local: 't', expectedText: 'First', text: 'Changed', rawRunElementIndex: run.elementIndex, rawParagraphElementIndex: paragraph.elementIndex };
    await verifyTextSourceAncestry(lexical.elements, [edit], part, async () => {});
    const split = await lexical.preflightRunSplit({ elementIndex: edit.elementIndex, namespaceUri: context.a, expectedText: 'First', offset: 2 });
    assert.deepEqual(split, { left: 'Fi', right: 'rst', runElementIndex: run.elementIndex, paragraphElementIndex: paragraph.elementIndex });
    for (const invalid of [{ ...edit, rawRunElementIndex: paragraph.elementIndex }, { ...edit, rawParagraphElementIndex: run.elementIndex }, { ...edit, namespaceUri: 'urn:wrong' }, { ...edit, rawRunElementIndex: undefined }]) {
      await assert.rejects(verifyTextSourceAncestry(lexical.elements, [invalid], part, async () => {}), { code: 'save-candidate-invalid' });
    }
    await assert.rejects(verifyTextSourceAncestry(lexical.elements, [edit], part, async () => { throw new Error('cancelled'); }), /cancelled/);
    lexical.dispose();
  }
});

test('effective fallback run does not invent a direct raw paragraph parent', async () => {
  for (const strict of [false, true]) {
    const fixture = paragraphFixture({ strict, txBox: true }), part = '/deck/z.xml';
    fixture.parts['deck/z.xml'] = fixture.parts['deck/z.xml']
      .replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace('<a:r><a:t>First</a:t></a:r>', '<mc:AlternateContent><mc:Choice Requires="x"><a:r><a:t>Hidden</a:t></a:r></mc:Choice><mc:Fallback><a:r><a:t>First</a:t></a:r></mc:Fallback></mc:AlternateContent>');
    const context = await structureContext(encodeStructure(fixture)), root = await context.root(part, 'sld'), stack = [root];
    let atom;
    while (stack.length) {
      const node = stack.pop();
      if (node.namespaceUri === context.a && node.local === 't' && node.children.some(child => child.kind === 'text' && child.value === 'First')) atom = node;
      for (const child of node.children) if (child.kind === 'element') stack.push(child);
    }
    assert.ok(atom);
    const binding = (await bindTextRuns([{ paragraphIndex: 0, runIndex: 0, text: 'First', node: atom }], { slideId: 's', elementId: 'e' }, { part }, context, async () => {}))[0];
    assert.equal(binding.status, 'bound'); assert.ok(binding.rawRunElementIndex);
    assert.equal(binding.rawParagraphElementIndex, undefined);
    const lexical = await captureXmlTextSource(new TextEncoder().encode(fixture.parts['deck/z.xml']));
    await verifyTextSourceAncestry(lexical.elements, [{ elementIndex: binding.sourceElementIndex, namespaceUri: context.a, local: 't', expectedText: 'First', text: 'Changed', rawRunElementIndex: binding.rawRunElementIndex }], part, async () => {});
    await assert.rejects(lexical.preflightRunSplit({ elementIndex: binding.sourceElementIndex, namespaceUri: context.a, expectedText: 'First', offset: 2 }), { code: 'xml-edit-unsupported' });
    lexical.dispose();
  }
});
