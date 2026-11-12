import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
import { inspectPptx, importPptx } from '../packages/pptx-parser/dist/index.js';
const hash = b => createHash('sha256').update(b).digest('hex');
const nativeFile = path.resolve('docs/testing/multi-master-native-evidence.json'), nativeBytes = await readFile(nativeFile), native = JSON.parse(nativeBytes);
const inputs = [{ file:path.join(native.directory,'fixed.pptx'), sha256:native.fixedSha256 }, ...native.rounds];
const browser = await chromium.launch(), page = await browser.newPage(), records = [], requests = [];
await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
try {
  const core = 'data:text/javascript;base64,' + Buffer.from(await readFile('packages/core/dist/browser/index.js','utf8')).toString('base64');
  const parser = 'data:text/javascript;base64,' + Buffer.from((await readFile('packages/pptx-parser/dist/browser/index.js','utf8')).replaceAll('"@pptkit/core"',JSON.stringify(core))).toString('base64');
  for (const name of ['core','pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`,'utf8') });
  const output = path.join(native.directory,'parser-inspections'); await mkdir(output,{recursive:true});
  for (const input of inputs) {
    const bytes = await readFile(input.file); assert.equal(hash(bytes),input.sha256);
    const inspection = await inspectPptx(bytes); assert.equal(inspection.slideMasters.length,4); assert.equal(inspection.layouts.length,5);
    const session = await importPptx(bytes,{projection:'structure'});
    let expected;
    try { const plan=await session.prepareSave({mode:'no-op'}); assert.equal(plan.status,'ready');expected={inspection,report:session.report,bytes:[...(await session.save(plan.plan)).bytes]}; assert.deepEqual(expected.bytes,[...bytes]); } finally {session.dispose();}
    for (const format of ['esm','global']) {
      const actual = await page.evaluate(async ({bytes,parser})=>{
        const api=parser?await import(parser):PPTKit;const source=Uint8Array.from(bytes),inspection=await api.inspectPptx(source),session=await api.importPptx(source,{projection:'structure'});
        try { const plan=await session.prepareSave({mode:'no-op'});if(plan.status!=='ready')throw Error('Native master no-op blocked');return {inspection,report:session.report,bytes:[...(await session.save(plan.plan)).bytes]}; }finally{session.dispose();}
      },{bytes:[...bytes],parser:format==='esm'?parser:null});assert.deepEqual(actual,expected);
    }
    const snapshot=JSON.stringify({inspection:expected.inspection,report:expected.report},null,2)+'\n', snapshotFile=path.join(output,path.basename(input.file)+'.json');await writeFile(snapshotFile,snapshot);
    records.push({file:input.file,sha256:hash(bytes),snapshotFile,snapshotSha256:hash(snapshot),nodeEsmGlobalReportsAndNoOpBytesEqual:true});
  }
  assert.deepEqual(requests,[]);
} finally {await browser.close();}
await writeFile('docs/testing/multi-master-native-parser-evidence.json',JSON.stringify({schemaVersion:1,nativeEvidence:nativeFile,nativeEvidenceSha256:hash(nativeBytes),records,total:{actualInputs:records.length,consumerNoOps:records.length*3},scope:'Current structural inspection/full report and exact no-op across captured source/native inputs. No authored registry import, source master editing or new native app operation.'},null,2)+'\n');console.log(JSON.stringify({inputs:records.length,noOps:records.length*3}));
