import assert from 'node:assert/strict';
import {readFile,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import path from 'node:path';
import {createPresentation,normalizePresentation,resolvePresentationInheritance} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {masterAdvancedDocument} from '../tests/fixtures/master-advanced.mjs';
import {verifyAdvancedTableParagraphFontBridge} from './research/advanced-table-font-bridge.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');
const file='docs/testing/master-advanced-evidence.json',e=JSON.parse(await readFile(file));
const nativeFile='docs/testing/master-cell-properties-native-evidence.json',nativeBytes=await readFile(nativeFile),n=JSON.parse(nativeBytes);
const parserFile='docs/testing/master-cell-properties-native-parser-evidence.json',parserBytes=await readFile(parserFile),p=JSON.parse(parserBytes);
const d=masterAdvancedDocument(createPresentation),l=resolveLayout(d),out=await generatePptx(d,{connectorMode:'native'}),svg=await renderLayoutToSvg(l);
const bridgeFile='docs/testing/master-paragraph-font-bridge-evidence.json',bridgeBytes=await readFile(bridgeFile),bridge=JSON.parse(bridgeBytes);assert.equal(bridge.verified,true);
const bridgeNativeBytes=await readFile(bridge.nativeEvidence.file),bridgeNative=JSON.parse(bridgeNativeBytes);assert.equal(sha(bridgeNativeBytes),bridge.nativeEvidence.sha256);assert.equal(bridgeNative.verified,true);
for(const a of [...bridgeNative.artifacts,...bridgeNative.negativeFixtures]){const b=await readFile(a.file);assert.equal(sha(b),a.sha256);assert.equal(b.length,a.bytes);}
for(const a of [...bridge.dependencies,...bridgeNative.dependencies,...bridgeNative.baselineEvidence])assert.equal(sha(await readFile(a.file)),a.sha256);
const current=bridgeNative.records.find(r=>r.mode==='before');assert.ok(current);assert.equal(current.fullPageRecordsExact,true);assert.equal(current.historicalSha256,e.sourceSha256);assert.equal(sha(out.bytes),current.sha256);assert.deepEqual(Buffer.from(out.bytes),await readFile(current.file));
assert.equal(n.verified,true);assert.equal(n.sourceSha256,e.sourceSha256);const historical=await readFile(path.join(e.nativeDirectory,'input.pptx'));assert.equal(sha(historical),e.sourceSha256);verifyAdvancedTableParagraphFontBridge(historical,out.bytes);
for(const[name,value]of [['normalized.json',normalizePresentation(d)],['origins.json',resolvePresentationInheritance(d).origins],['layout.json',l],['warnings.json',out.warnings]])assert.deepEqual(JSON.parse(await readFile(path.join(e.directory,name))),JSON.parse(JSON.stringify(value)));
for(const[i,s]of svg.slides.entries())assert.equal(await readFile(path.join(e.directory,'page-'+i+'.svg'),'utf8'),s.svg);
for(const a of [...e.artifacts,...n.artifacts]){const bytes=await readFile(a.file);assert.equal(bytes.length,a.bytes,a.file);assert.equal(sha(bytes),a.sha256,a.file);}
for(const a of n.dependencies)assert.equal(sha(await readFile(a.file)),a.sha256,a.file);
assert.equal(n.firstPages.length,4);assert.equal(n.rounds.length,3);assert.equal(n.rounds.reduce((sum,r)=>sum+r.pages.length,0),12);assert.equal(n.negatives.length,6);
assert.equal(p.nativeEvidenceSha256,sha(nativeBytes));assert.deepEqual(p.total,{actualInputs:4,consumerNoOps:12});
for(const r of p.records){assert.equal(sha(await readFile(r.file)),r.sha256);assert.equal(sha(await readFile(r.snapshotFile)),r.snapshotSha256);assert.equal(r.fullNodeEsmGlobalEqual,true);assert.equal(r.exactNoOp,true);}
const dependencies=['packages/pptx-exporter/src/ooxml/elements.ts','packages/pptx-exporter/test/table-cell-properties.test.mjs','tests/fixtures/master-advanced.mjs','scripts/prepare-master-advanced.mjs','scripts/inspect-native-master-cell-properties.mjs','scripts/verify-master-cell-properties-evidence.mjs','scripts/table-style-parity.test.mjs','scripts/table-origins-parity.test.mjs','scripts/inspect-table-style.mjs','scripts/research/advanced-table-font-bridge.mjs'];
e.currentProducer={sourceSha256:sha(out.bytes),historicalSourceSha256:e.sourceSha256,bridgeEvidence:{file:bridgeFile,sha256:sha(bridgeBytes)},dependencies:await Promise.all(dependencies.map(async file=>({file,sha256:sha(await readFile(file))})))};
e.nativeApplication='Actual PowerPoint 16.104 (25121423); historical corrected-cell three-round controls retained. Current producer has a separate exact six-paragraph font delta and actual four-page first-open equivalence bridge.';e.scope=n.scope;
e.nativeEvidence={file:nativeFile,sha256:sha(nativeBytes)};e.parserEvidence={file:parserFile,sha256:sha(parserBytes)};e.finalizedAt=new Date().toISOString();await writeFile(file,JSON.stringify(e,null,2)+'\n');
console.log(JSON.stringify({artifactHashChecks:e.artifacts.length+n.artifacts.length,implementationDependencies:e.currentProducer.dependencies.length,oracleDependencies:n.dependencies.length,historicalNativeRounds:3,reopenedPages:12,zipNegatives:6,consumerNoOps:12,historicalSourceSha256:e.sourceSha256,currentSourceSha256:e.currentProducer.sourceSha256,currentNativeFirstPrintExact:true}));
