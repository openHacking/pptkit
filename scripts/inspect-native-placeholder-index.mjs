import assert from 'node:assert/strict';
import { readFile, mkdtemp, writeFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { inspectPptx, importPptx } from '../packages/pptx-parser/dist/index.js';
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const evidenceFile = path.resolve('docs/testing/placeholder-index-native-evidence.json'), evidenceBytes = await readFile(evidenceFile), evidence = JSON.parse(evidenceBytes);
const directory = await mkdtemp(path.join(evidence.directory, 'parser-'));
const inputs = ['before.pptx','input.pptx','index-round1.pptx','index-round2.pptx','index-round3.pptx'];
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
    const inspection = await inspectPptx(bytes), session = await importPptx(bytes,{ projection:'content' });
    assert.equal(inspection.slides.length,3);
    assert.equal(inspection.slideMasters.length,3); assert.equal(inspection.layouts.length,3);
    let expected;
    try {
      const plan = await session.prepareSave({ mode:'no-op' }); assert.equal(plan.status,'ready');
      expected = { inspection, report:session.report, normalized:normalizePresentation(session.document), bytes:[...(await session.save(plan.plan)).bytes] }; assert.deepEqual(expected.bytes,[...bytes]);
      const local=expected.report.placeholderInheritance.placeholders.filter(p=>p.level==='slide');assert.equal(local.length,12);
      assert.equal(local.filter(p=>p.index===0).length,3);assert.equal(local.filter(p=>p.index===4294967295).length,3);
      for(const placeholder of local){assert.equal(placeholder.parent.rule,'index');assert.equal(placeholder.parent.status,'matched');assert.equal(placeholder.parent.key,placeholder.index);if(placeholder.index===0)assert.equal(placeholder.explicitIndex,!name.startsWith('index-round'));}
    } finally { session.dispose(); }
    for (const format of ['esm','global']) {
      const actual = await page.evaluate(async ({ bytes, parser, core }) => {
        const api = parser ? Object.assign({},await import(parser),await import(core)) : PPTKit, source = Uint8Array.from(bytes);
        const inspection = await api.inspectPptx(source), session = await api.importPptx(source,{ projection:'content' });
        try {
          const plan = await session.prepareSave({ mode:'no-op' }); if (plan.status !== 'ready') throw Error('Native no-op blocked');
          return { inspection, report:session.report, normalized:api.normalizePresentation(session.document), bytes:[...(await session.save(plan.plan)).bytes] };
        } finally { session.dispose(); }
      },{ bytes:[...bytes], parser:format === 'esm' ? parser : null, core });
      assert.deepEqual(actual,expected);
    }
    const snapshot = JSON.stringify({ inspection, report:expected.report, normalized:expected.normalized },null,2)+'\n', snapshotFile = path.join(directory,name+'.json');
    await writeFile(snapshotFile,snapshot);
    records.push({ file, sha256:hash(bytes), snapshotFile, snapshotSha256:hash(snapshot), fullNodeEsmGlobalEqual:true, exactNoOp:true });
  }
  assert.deepEqual(requests,[]);
} finally { await browser.close(); }
const dependencyFiles=['scripts/inspect-native-placeholder-index.mjs','scripts/build-browser-bundle.mjs','pnpm-lock.yaml','packages/core/package.json','packages/pptx-parser/package.json'];
async function dependenciesIn(directory){for(const entry of await readdir(directory,{withFileTypes:true})){const file=path.join(directory,entry.name);if(entry.isDirectory())await dependenciesIn(file);else if(file.endsWith('.ts')||file.endsWith('.js'))dependencyFiles.push(file);}}
for(const name of ['core','pptx-parser'])for(const folder of ['src','dist'])await dependenciesIn('packages/'+name+'/'+folder);
const dependencies=await Promise.all(dependencyFiles.sort().map(async file=>({file,sha256:hash(await readFile(file))})));
await writeFile('docs/testing/placeholder-index-native-parser-evidence.json',JSON.stringify({ schemaVersion:1, nativeEvidence:evidenceFile, nativeEvidenceSha256:hash(evidenceBytes), records, dependencies, total:{ actualInputs:records.length, consumerNoOps:records.length*3 }, scope:'Five actual before/after/reopened explicit-index files, full content projections/inspections/reports in Node/ESM/global and 15 exact source-only no-ops. Source and compiled runtime dependencies are bound; changed Parser/Core code requires a new consumer inspection. Native omitted-zero/retained-max index provenance and unique parent matches are checked. No complete master/artwork Core import, inherited-artwork visibility model or source template/index structure-write authority. Slide order/hidden-state projection retains its existing mapped profile.' },null,2)+'\n');
console.log(JSON.stringify({ inputs:records.length, noOps:records.length*3 }));
