import assert from 'node:assert/strict';
import test from 'node:test';
import { rollup } from 'rollup';
import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import { chromium } from '@playwright/test';
import { captureXmlTextSource } from '../packages/pptx-parser/dist/preservation/xml-text-source.js';

test('private plain-run split matches Node and Chromium without network access', async t => {
  const run = async capture => {
    const xml = '<a:p xmlns:a="urn:a"><a:r><a:rPr b="1" sz="2400"/><a:t>Ae\u0301👨‍👩‍👧‍👦Z</a:t></a:r></a:p>';
    const source = await capture(new TextEncoder().encode(xml));
    try {
      const request = { elementIndex: 4, namespaceUri: 'urn:a', expectedText: 'Ae\u0301👨‍👩‍👧‍👦Z', offset: 3 };
      const result = await source.preflightRunSplit(request);
      const candidate = new TextDecoder().decode((await source.splitSimpleRun(request)).bytes);
      const failures = [];
      for (const item of [{ ...request, offset: 2 }, { ...request, offset: 4 }, { ...request, expectedText: 'different' }]) {
        try { await source.preflightRunSplit(item); failures.push('accepted'); } catch (error) { failures.push(error.code); }
      }
      return { result, candidate, failures, unchanged: source.readElement({ elementIndex: 1, namespaceUri: 'urn:a', local: 'p' }) === xml };
    } finally { source.dispose(); }
  };
  const expected = await run(captureXmlTextSource);
  assert.deepEqual(expected.failures, ['split-invalid', 'split-invalid', 'xml-edit-conflict']);
  assert.equal(expected.unchanged, true);
  assert.equal(expected.candidate, '<a:p xmlns:a="urn:a"><a:r><a:rPr b="1" sz="2400"/><a:t>Ae\u0301</a:t></a:r><a:r><a:rPr b="1" sz="2400"/><a:t>👨‍👩‍👧‍👦Z</a:t></a:r></a:p>');
  const bundle = await rollup({ input: 'packages/pptx-parser/dist/preservation/xml-text-source.js', plugins: [nodeResolve({ browser: true }), commonjs()] });
  t.after(() => bundle.close());
  const output = await bundle.generate({ format: 'es' });
  const url = `data:text/javascript;base64,${Buffer.from(output.output[0].code).toString('base64')}`;
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const actual = await page.evaluate(async ({ url, source }) => {
    const module = await import(url), run = (0, eval)(`(${source})`);
    return run(module.captureXmlTextSource);
  }, { url, source: run.toString() });
  assert.deepEqual(actual, expected);
  assert.deepEqual(requests, []);
});
