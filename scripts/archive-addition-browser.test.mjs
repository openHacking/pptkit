import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { replaceArchiveEntries } from '../packages/pptx-parser/dist/preservation/replace-archive-entries.js';
import { inspectZipArchive } from '../packages/pptx-parser/dist/index.js';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('private ZIP addition has Node and browser parity without exposing a public API', async t => {
  const source = encodeStructure(structureFixture());
  const addition = { name: 'ppt/media/independent.png', bytes: new Uint8Array([0, 1, 127, 128, 255]) };
  const expected = await replaceArchiveEntries(source, [], { maxOutputBytes: 1024 * 1024, additions: [addition] });
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/preservation/replace-archive-entries.js', plugins: [nodeResolve({ browser: true }), commonjs()] });
  t.after(() => bundle.close());
  const { output } = await bundle.generate({ format: 'iife', name: 'ArchiveAdditionTest' });
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  await page.addScriptTag({ content: output[0].code });
  const actual = await page.evaluate(async ({ source, addition }) => {
    const input = new Uint8Array(source), payload = new Uint8Array(addition.bytes);
    const pending = ArchiveAdditionTest.replaceArchiveEntries(input, [], { maxOutputBytes: 1024 * 1024, additions: [{ name: addition.name, bytes: payload }] });
    input.fill(0); payload.fill(0);
    const bytes = await pending;
    let conflict;
    try { await ArchiveAdditionTest.replaceArchiveEntries(bytes, [], { maxOutputBytes: 1024 * 1024, additions: [{ name: addition.name.toUpperCase(), bytes: new Uint8Array() }] }); }
    catch (error) { conflict = error.code; }
    return { bytes: Array.from(bytes), conflict };
  }, { source: Array.from(source), addition: { name: addition.name, bytes: Array.from(addition.bytes) } });
  assert.deepEqual(new Uint8Array(actual.bytes), expected);
  assert.equal(actual.conflict, 'archive-duplicate-path');
  const inspection = await inspectZipArchive(new Uint8Array(actual.bytes));
  assert.ok(inspection.entries.some(entry => entry.name === addition.name));
});
