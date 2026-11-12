import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { inspectZipArchive } from '../packages/pptx-parser/dist/index.js';
import { zip } from '../packages/pptx-parser/test/zip-fixture.mjs';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('ZIP directory type decisions agree across Node and browser public exports', async t => {
  const fixture = (name, host, mode, directoryFlag = false) => {
    const { bytes, directory } = zip([{ name, method: 0 }]);
    bytes.writeUInt16LE((host << 8) | 20, directory + 4);
    bytes.writeUInt32LE(((mode << 16) | (directoryFlag ? 16 : 0)) >>> 0, directory + 38);
    return bytes;
  };
  const cases = [fixture('folder/', 3, 0x8000), fixture('file', 3, 0x4000), fixture('file', 3, 0x8000, true), fixture('folder/', 3, 0x4000, true), fixture('file', 3, 0x8000), fixture('file', 0, 0x4000), fixture('link/', 3, 0xa000), fixture('file', 19, 0xa000), fixture('folder/', 19, 0x8000), fixture('file', 19, 0x4000), fixture('folder/', 19, 0x4000), fixture('file', 19, 0x8000)];
  const expected = await Promise.all(cases.map(async bytes => { try { return { directory: (await inspectZipArchive(bytes)).entries[0].directory }; } catch (error) { return { code: error.code }; } }));
  assert.deepEqual(expected, [{ code: 'archive-path-invalid' }, { code: 'archive-path-invalid' }, { code: 'archive-path-invalid' }, { directory: true }, { directory: false }, { directory: false }, { code: 'archive-unsupported' }, { code: 'archive-unsupported' }, { code: 'archive-path-invalid' }, { code: 'archive-path-invalid' }, { directory: true }, { directory: false }]);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const requests = []; await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  const esm = await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8');
  const actual = await page.evaluate(async ({ esm, cases }) => {
    const module = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(esm)))}`);
    return Promise.all([module, globalThis.PPTKit].map(api => Promise.all(cases.map(async bytes => { try { return { directory: (await api.inspectZipArchive(new Uint8Array(bytes))).entries[0].directory }; } catch (error) { return { code: error.code }; } }))));
  }, { esm, cases: cases.map(bytes => Array.from(bytes)) });
  assert.deepEqual(actual, [expected, expected]); assert.deepEqual(requests, []);
});

test('OS X ZIP type handling reaches browser PPTX import and byte-exact no-op save', async t => {
  const original = Buffer.from(encodeStructure(structureFixture())), end = original.length - 22, central = original.readUInt32LE(end + 16);
  const make = mode => { const bytes = Buffer.from(original); bytes.writeUInt16LE((19 << 8) | 20, central + 4); bytes.writeUInt32LE((mode << 16) >>> 0, central + 38); return Array.from(bytes); };
  const samples = [make(0x8000), make(0xa000), make(0x4000)];
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  const actual = await page.evaluate(async samples => Promise.all(['structure', 'content'].map(async projection => Promise.all(samples.map(async data => {
    try {
      const session = await PPTKit.importPptx(new Uint8Array(data), { projection });
      try { const plan = await session.prepareSave({ mode: 'no-op' }); const output = await session.save(plan.plan); return { status: plan.status, identical: output.bytes.length === data.length && output.bytes.every((value, index) => value === data[index]) }; }
      finally { session.dispose(); }
    } catch (error) { return { code: error.code }; }
  })))), samples);
  assert.deepEqual(actual, Array.from({ length: 2 }, () => [{ status: 'ready', identical: true }, { code: 'import-unsupported' }, { code: 'import-invalid-package' }]));
});
