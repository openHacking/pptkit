import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {readFile,writeFile,readdir} from 'node:fs/promises';
import path from 'node:path';
import {createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {placeholderIndexDocument} from '../tests/fixtures/placeholder-index.mjs';
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const manifestFile='docs/testing/placeholder-index-evidence.json', manifest=JSON.parse(await readFile(manifestFile,'utf8'));
const regenerated=new Map(), save=(name,value)=>regenerated.set(name,typeof value==='string'||value instanceof Uint8Array?Buffer.from(value):Buffer.from(JSON.stringify(value,null,2)+'\n'));
const d=placeholderIndexDocument(createPresentation), before=await generatePptx(d), oldLayout=resolveLayout(d);
save('before.pptx',before.bytes);save('before-normalized.json',normalizePresentation(d));save('before-layout.json',oldLayout);
for(const[i,slide]of(await renderLayoutToSvg(oldLayout)).slides.entries())save('before-'+i+'.svg',slide.svg);
const preserve=prepareLayoutRebind(d,[{slideId:'page-0',layoutId:'target'}],{localAppearance:'preserve'});d.applyLayoutRebind(preserve);
const adopt=prepareLayoutRebind(d,[{slideId:'page-1',layoutId:'target'}],{localAppearance:'adopt',reset:['geometry','text','background']});d.applyLayoutRebind(adopt);
const after=await generatePptx(d), layout=resolveLayout(d);
save('after.pptx',after.bytes);save('after-normalized.json',normalizePresentation(d));save('after-layout.json',layout);save('after-origins.json',resolvePresentationInheritance(d).origins);save('plans.json',{preserve,adopt});save('warnings.json',{before:before.warnings,after:after.warnings});
for(const[i,slide]of(await renderLayoutToSvg(layout)).slides.entries())save('after-'+i+'.svg',slide.svg);
assert.equal(regenerated.size,manifest.artifacts.length);
for(const artifact of manifest.artifacts){const bytes=await readFile(artifact.file);assert.equal(bytes.length,artifact.bytes);assert.equal(sha(bytes),artifact.sha256);assert.deepEqual(bytes,regenerated.get(path.basename(artifact.file)));}
assert.equal(sha(regenerated.get('before.pptx')),manifest.beforeSha256);assert.equal(sha(regenerated.get('after.pptx')),manifest.sourceSha256);
const links=[];
for(const file of ['docs/testing/placeholder-index-native-evidence.json','docs/testing/placeholder-index-native-parser-evidence.json']){
  const bytes=await readFile(file), proof=JSON.parse(bytes);links.push({file,sha256:sha(bytes)});
  for(const artifact of [...(proof.artifacts??[]),...(proof.negatives??[])]){const body=await readFile(artifact.file);assert.equal(sha(body),artifact.sha256);if(artifact.bytes!==undefined)assert.equal(body.length,artifact.bytes);}
  for(const dependency of proof.dependencies??[])assert.equal(sha(await readFile(dependency.file)),dependency.sha256);
  if(proof.records){assert.equal(proof.records.length,5);assert.equal(proof.total.consumerNoOps,15);assert.equal(sha(await readFile(proof.nativeEvidence)),proof.nativeEvidenceSha256);for(const record of proof.records){assert.equal(sha(await readFile(record.file)),record.sha256);assert.equal(sha(await readFile(record.snapshotFile)),record.snapshotSha256);assert.equal(record.exactNoOp,true);assert.equal(record.fullNodeEsmGlobalEqual,true);}}
  else{assert.equal(proof.verified,true);assert.equal(proof.nativeIndexGateAccepted,true);assert.equal(proof.completeTemplateRebindAccepted,false);assert.equal(proof.fullFirstSaveEquivalence,false);assert.equal(proof.sourceSha256,manifest.sourceSha256);assert.equal(proof.beforeSha256,manifest.beforeSha256);assert.equal(proof.rounds.length,3);assert.equal(proof.firstPages.length,3);for(const round of proof.rounds)assert.equal(sha(await readFile(round.file)),round.sha256);assert.equal(proof.negatives.filter(n=>n.file.endsWith('.pptx')).length,23);assert.equal(proof.negatives.filter(n=>n.file.endsWith('.pdf')).length,4);}
}
const dependencyFiles=[];
async function sourceFiles(directory){for(const entry of await readdir(directory,{withFileTypes:true})){const file=path.join(directory,entry.name);if(entry.isDirectory())await sourceFiles(file);else if(file.endsWith('.ts'))dependencyFiles.push(file);}}
for(const name of ['core','layout','pptx-exporter','svg-renderer','pptx-parser'])await sourceFiles('packages/'+name+'/src');
dependencyFiles.push('tests/fixtures/placeholder-index.mjs','scripts/prepare-placeholder-index.mjs','scripts/verify-placeholder-index-evidence.mjs','scripts/inspect-native-placeholder-index.mjs','scripts/layout-rebind-parity.test.mjs','scripts/build-browser-bundle.mjs','package.json');
manifest.currentProducerDependencies=await Promise.all(dependencyFiles.sort().map(async file=>({file,sha256:sha(await readFile(file))})));
const contact=path.join(manifest.directory,'native-contact.png'),contactBytes=await readFile(contact);
if(manifest.visualArtifacts)for(const artifact of manifest.visualArtifacts){const bytes=await readFile(artifact.file);assert.equal(sha(bytes),artifact.sha256);assert.equal(bytes.length,artifact.bytes);}
else manifest.visualArtifacts=[{file:contact,bytes:contactBytes.length,sha256:sha(contactBytes)}];
manifest.evidence=links;manifest.nativeApplication='Named explicit-zero/max/mixed implicit allocation control accepted over three actual native reopen rounds, 23 corrupt ZIPs/four corrupt PDFs, complete content projection/inspection/report comparisons and fifteen exact source-only no-ops. General native template import/reset/cross-document/source editing remains open.';
await writeFile(manifestFile,JSON.stringify(manifest,null,2)+'\n');
console.log(JSON.stringify({preparedArtifacts:regenerated.size,producerDependencies:dependencyFiles.length,linkedProofs:links.length,nativeIndexGateAccepted:true,completeTemplateRebindAccepted:false}));
