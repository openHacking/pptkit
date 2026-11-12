import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('browser concurrent MCE projections preserve the selected raw source identity', async t => {
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/presentation/structure-context.js', plugins: [nodeResolve({ browser: true }), commonjs()] });
  t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'SourceIdentityTest' });
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); await page.addScriptTag({ content: output[0].code });
  for (const strict of [false, true]) {
    const fixture = structureFixture(strict);
    fixture.parts['deck/z.xml'] = fixture.parts['deck/z.xml']
      .replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
      .replace('<p:cSld name="First"/>', '<mc:AlternateContent><mc:Choice Requires="x"><p:cSld name="Hidden"/></mc:Choice><mc:Fallback><p:cSld name="First"/></mc:Fallback></mc:AlternateContent>');
    const result = await page.evaluate(async bytes => {
      const context = await SourceIdentityTest.structureContext(new Uint8Array(bytes));
      const part = '/deck/z.xml';
      const roots = await Promise.all(Array.from({ length: 8 }, () => context.root(part, 'sld')));
      const selected = roots[0].children.find(node => node.kind === 'element' && node.local === 'cSld');
      const [a, b] = await Promise.all([context.sourceElement(part, selected), context.sourceElement(part, selected)]);
      return { same: roots.every(root => root === roots[0]), index: a.elementIndex, sameSource: a.node === b.node, name: a.node.attributes.find(attribute => attribute.local === 'name').value, reports: context.reports.filter(report => report.part === part).length };
    }, Array.from(encodeStructure(fixture)));
    assert.deepEqual(result, { same: true, index: 6, sameSource: true, name: 'First', reports: 1 });
  }
});
