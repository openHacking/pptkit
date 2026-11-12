import assert from 'node:assert/strict';
import { readFile, mkdtemp, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
import { inspectPptx, importPptx } from '../packages/pptx-parser/dist/index.js';
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const evidenceFile = path.resolve('docs/testing/master-artwork-native-evidence.json'), evidenceBytes = await readFile(evidenceFile), evidence = JSON.parse(evidenceBytes);
const directory = await mkdtemp(path.join(evidence.directory, 'parser-'));
const inputs = ['input.pptx','artwork-round1.pptx','artwork-round2.pptx','artwork-round3.pptx'];
const browser = await chromium.launch(), page = await browser.newPage(), requests = [], records = [];
await page.route('**/*', request => { requests.push(request.request().url()); return request.abort(); });
try {
  const url = text => 'data:text/javascript;base64,' + Buffer.from(text).toString('base64');
  const core = url(await readFile('packages/core/dist/browser/index.js','utf8'));
  const parser = url((await readFile('packages/pptx-parser/dist/browser/index.js','utf8')).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const name of ['core','pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`,'utf8') });
  for (const name of inputs) {
    const file = path.join(evidence.directory, name), bytes = await readFile(file), binding = evidence.artifacts.find(a => a.file === file);
    assert.ok(binding); assert.equal(hash(bytes), binding.sha256);
    const inspection = await inspectPptx(bytes), session = await importPptx(bytes,{ projection:'structure' });
    assert.equal(inspection.slides.length,6);
    assert.equal(inspection.slideMasters.length,3); assert.equal(inspection.layouts.length,4);
    let expected;
    try {
      const plan = await session.prepareSave({ mode:'no-op' }); assert.equal(plan.status,'ready');
      expected = { inspection, report:session.report, bytes:[...(await session.save(plan.plan)).bytes] }; assert.deepEqual(expected.bytes,[...bytes]);
    } finally { session.dispose(); }
    for (const format of ['esm','global']) {
      const actual = await page.evaluate(async ({ bytes, parser }) => {
        const api = parser ? await import(parser) : PPTKit, source = Uint8Array.from(bytes);
        const inspection = await api.inspectPptx(source), session = await api.importPptx(source,{ projection:'structure' });
        try {
          const plan = await session.prepareSave({ mode:'no-op' }); if (plan.status !== 'ready') throw Error('Native no-op blocked');
          return { inspection, report:session.report, bytes:[...(await session.save(plan.plan)).bytes] };
        } finally { session.dispose(); }
      },{ bytes:[...bytes], parser:format === 'esm' ? parser : null });
      assert.deepEqual(actual,expected);
    }
    const snapshot = JSON.stringify({ inspection, report:expected.report },null,2)+'\n', snapshotFile = path.join(directory,name+'.json');
    await writeFile(snapshotFile,snapshot);
    records.push({ file, sha256:hash(bytes), snapshotFile, snapshotSha256:hash(snapshot), fullNodeEsmGlobalEqual:true, exactNoOp:true });
  }
  assert.deepEqual(requests,[]);
} finally { await browser.close(); }
await writeFile('docs/testing/master-artwork-native-parser-evidence.json',JSON.stringify({ schemaVersion:1, nativeEvidence:evidenceFile, nativeEvidenceSha256:hash(evidenceBytes), records, total:{ actualInputs:records.length, consumerNoOps:records.length*3 }, scope:'Four actual source/reopened master-artwork files, full structural inspections/reports in Node/ESM/global and 12 exact source-only no-ops. No complete master/artwork Core import, visibility projection or source style/structure write authority.' },null,2)+'\n');
console.log(JSON.stringify({ inputs:records.length, noOps:records.length*3 }));
