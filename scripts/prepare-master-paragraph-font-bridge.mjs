import assert from 'node:assert/strict';
import {readFile,writeFile,mkdtemp} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {createPresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {masterAdvancedDocument,editMasterHeading} from '../tests/fixtures/master-advanced.mjs';
import {verifyAdvancedTableParagraphFontBridge} from './research/advanced-table-font-bridge.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');
const baseline=JSON.parse(await readFile('docs/testing/master-advanced-evidence.json'));
const shared=JSON.parse(await readFile('docs/testing/master-shared-evidence.json'));
const directory=path.resolve(await mkdtemp('.tmp/master-advanced/font-mark-bridge-')),d=masterAdvancedDocument(createPresentation),records=[];
for(const[mode,historicalFile,historicalSha256,referencePdf]of [
  ['before',path.join(baseline.nativeDirectory,'input.pptx'),baseline.sourceSha256,path.join(baseline.nativeDirectory,'first.pdf')],
  ['after',path.join(shared.nativeDirectory,'fresh-core-edit.pptx'),shared.freshEditedSha256,path.join(shared.nativeDirectory,'fresh-core-edit.pdf')],
]){
  if(mode==='after')editMasterHeading(d);
  const output=await generatePptx(d,{connectorMode:'native'}),historical=await readFile(historicalFile);assert.equal(sha(historical),historicalSha256);
  const delta=verifyAdvancedTableParagraphFontBridge(historical,output.bytes),file=path.join(directory,mode+'.pptx');await writeFile(file,output.bytes);
  records.push({mode,file,sha256:sha(output.bytes),historicalFile,historicalSha256,referencePdf,referencePdfSha256:sha(await readFile(referencePdf)),...delta});
}
const manifest={schemaVersion:1,directory,records,nativeApplication:'pending actual first-open printing; no previous reopen round is relabeled'};
await writeFile('docs/testing/master-paragraph-font-bridge-evidence.json',JSON.stringify(manifest,null,2)+'\n');console.log(JSON.stringify(manifest));
