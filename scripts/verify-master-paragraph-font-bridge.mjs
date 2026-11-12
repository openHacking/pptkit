import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {createPresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {masterAdvancedDocument,editMasterHeading} from '../tests/fixtures/master-advanced.mjs';
import {verifyAdvancedTableParagraphFontBridge} from './research/advanced-table-font-bridge.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');
const file='docs/testing/master-paragraph-font-bridge-evidence.json',manifest=JSON.parse(await readFile(file));
const nativeFile='docs/testing/master-paragraph-font-native-evidence.json',nativeBytes=await readFile(nativeFile),native=JSON.parse(nativeBytes);
assert.equal(native.verified,true);assert.equal(native.records.length,2);
for(const artifact of [...native.artifacts,...native.negativeFixtures]){const bytes=await readFile(artifact.file);assert.equal(bytes.length,artifact.bytes);assert.equal(sha(bytes),artifact.sha256);}
for(const item of [...native.dependencies,...native.baselineEvidence])assert.equal(sha(await readFile(item.file)),item.sha256);
const d=masterAdvancedDocument(createPresentation);
for(const record of native.records){
  if(record.mode==='after')editMasterHeading(d);
  const out=await generatePptx(d,{connectorMode:'native'}),old=await readFile(record.historicalFile);
  assert.equal(sha(old),record.historicalSha256);assert.equal(sha(out.bytes),record.sha256);assert.deepEqual(Buffer.from(out.bytes),await readFile(record.file));
  assert.equal(record.fullPageRecordsExact,true);assert.equal(record.pages.length,4);assert.equal(sha(await readFile(record.referencePdf)),record.referencePdfSha256);assert.equal(sha(await readFile(record.pdf)),record.pdfSha256);
  assert.deepEqual(verifyAdvancedTableParagraphFontBridge(old,out.bytes),{changedParts:record.changedParts,tables:record.tables,paragraphs:record.paragraphs,delta:record.delta});
}
const old=await readFile(native.records[0].historicalFile);assert.equal(native.negativeFixtures.length,4);
for(const fixture of native.negativeFixtures){const bytes=await readFile(fixture.file);assert.throws(()=>verifyAdvancedTableParagraphFontBridge(old,bytes),{name:'AssertionError'});}
const deps=['scripts/prepare-master-paragraph-font-bridge.mjs','scripts/verify-master-paragraph-font-bridge.mjs','scripts/research/advanced-table-font-bridge.mjs','packages/pptx-exporter/src/ooxml/elements.ts','tests/fixtures/master-advanced.mjs'];
manifest.verified=true;manifest.nativeEvidence={file:nativeFile,sha256:sha(nativeBytes)};manifest.negativeRejections=native.negativeFixtures.map(({name,file,sha256})=>({name,file,sha256}));manifest.dependencies=await Promise.all(deps.map(async file=>({file,sha256:sha(await readFile(file))})));manifest.nativeApplication='Actual PowerPoint 16.104 (25121423), two current source first-open prints/eight page records exactly equal to the historical before/edited native controls; historical reopen rounds retain their original identities.';
await writeFile(file,JSON.stringify(manifest,null,2)+'\n');console.log(JSON.stringify({verified:true,exactNativePages:8,byteBridges:2,actualZipNegatives:4}));
