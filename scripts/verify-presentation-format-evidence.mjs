import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {createPresentation,normalizePresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {masterArtworkDocument} from '../tests/fixtures/master-artwork.mjs';
import {presentationFormatDocuments} from '../tests/fixtures/presentation-format.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');
const file='docs/testing/presentation-format-evidence.json',e=JSON.parse(await readFile(file));
const nativeFile='docs/testing/presentation-format-native-evidence.json',nativeBytes=await readFile(nativeFile),n=JSON.parse(nativeBytes);
const parserFile='docs/testing/presentation-format-native-parser-evidence.json',parserBytes=await readFile(parserFile),p=JSON.parse(parserBytes);
assert.equal(n.verified,true);assert.equal(n.sourceSha256,e.sourceSha256);assert.equal(n.historicalBaselineSha256,e.oldSourceSha256);
const d=masterArtworkDocument(createPresentation),out=await generatePptx(d);assert.deepEqual(out.warnings,[]);assert.equal(sha(out.bytes),e.sourceSha256);
assert.deepEqual(JSON.parse(await readFile(path.join(e.directory,'normalized.json'))),normalizePresentation(d));
const oldBytes=await readFile(e.oldSourceFile);assert.equal(sha(oldBytes),e.oldSourceSha256);
const before=readGeneratedZip(oldBytes),after=readGeneratedZip(out.bytes);assert.deepEqual([...after.keys()].sort(),[...before.keys()].sort());
for(const[name,body]of before)if(name!=='docProps/app.xml')assert.deepEqual(after.get(name),body,name);
assert.equal(after.get('docProps/app.xml').toString(),before.get('docProps/app.xml').toString().replace('<PresentationFormat>On-screen Show (16:9)</PresentationFormat>','<PresentationFormat>On-screen Show (4:3)</PresentationFormat>'));
for(const s of presentationFormatDocuments(createPresentation)){
  const record=e.scenes.find(r=>r.id===s.document.id);assert.ok(record);assert.equal(record.expected,s.label);assert.equal(record.cx,s.cx);assert.equal(record.cy,s.cy);
  const actual=await generatePptx(s.document);assert.deepEqual(actual.warnings,[]);assert.equal(sha(actual.bytes),record.sha256);assert.deepEqual(Buffer.from(actual.bytes),await readFile(path.join(e.directory,s.document.id+'.pptx')));
}
for(const a of [...e.artifacts,...n.artifacts]){const bytes=await readFile(a.file);assert.equal(bytes.length,a.bytes,a.file);assert.equal(sha(bytes),a.sha256,a.file);}
for(const a of n.dependencies)assert.equal(sha(await readFile(a.file)),a.sha256,a.file);
assert.equal(n.firstPages.length,6);assert.equal(n.rounds.length,3);assert.equal(n.rounds.reduce((sum,r)=>sum+r.pages.length,0),18);assert.equal(n.negatives.length,4);
assert.equal(p.nativeEvidenceSha256,sha(nativeBytes));assert.deepEqual(p.total,{actualInputs:4,consumerNoOps:12});
for(const r of p.records){assert.equal(sha(await readFile(r.file)),r.sha256);assert.equal(sha(await readFile(r.snapshotFile)),r.snapshotSha256);assert.equal(r.fullNodeEsmGlobalEqual,true);assert.equal(r.exactNoOp,true);}
const dependencies=['packages/pptx-exporter/src/ooxml/package-parts.ts','packages/pptx-exporter/src/packaging/build-package.ts','packages/pptx-exporter/test/presentation-format.test.mjs','tests/fixtures/presentation-format.mjs','tests/fixtures/master-artwork-format-baseline.json','tests/fixtures/master-artwork.mjs','scripts/presentation-format-parity.test.mjs','scripts/prepare-presentation-format.mjs','scripts/inspect-native-presentation-format.mjs','scripts/verify-presentation-format-evidence.mjs','package.json'];
e.dependencies=await Promise.all(dependencies.map(async file=>({file,sha256:sha(await readFile(file))})));
e.nativeApplication='Actual PowerPoint 16.104 (25121423); metadata-corrected 4:3 source, six first-open and eighteen reopened pages';
e.nativeEvidence={file:nativeFile,sha256:sha(nativeBytes)};e.parserEvidence={file:parserFile,sha256:sha(parserBytes)};e.finalizedAt=new Date().toISOString();
await writeFile(file,JSON.stringify(e,null,2)+'\n');
console.log(JSON.stringify({artifactHashChecks:e.artifacts.length+n.artifacts.length,implementationDependencies:e.dependencies.length,oracleDependencies:n.dependencies.length,scenes:e.scenes.length,nativeRounds:3,reopenedPages:18,zipNegatives:4,consumerNoOps:12,sourceSha256:e.sourceSha256}));
