import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdir,mkdtemp,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {placeholderIndexDocument} from '../tests/fixtures/placeholder-index.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');await mkdir('.tmp/placeholder-index',{recursive:true});
const directory=path.resolve(await mkdtemp('.tmp/placeholder-index/run-')),nativeDirectory=path.resolve(await mkdtemp('.tmp/placeholder-index/native-')),artifacts=[];
async function save(name,value){const bytes=typeof value==='string'||value instanceof Uint8Array?value:JSON.stringify(value,null,2)+'\n',file=path.join(directory,name);await writeFile(file,bytes);artifacts.push({file,bytes:Buffer.byteLength(bytes),sha256:sha(bytes)});}
const d=placeholderIndexDocument(createPresentation),before=await generatePptx(d);await save('before.pptx',before.bytes);await save('before-normalized.json',normalizePresentation(d));const oldLayout=resolveLayout(d);await save('before-layout.json',oldLayout);for(const[i,s]of(await renderLayoutToSvg(oldLayout)).slides.entries())await save('before-'+i+'.svg',s.svg);
const preserve=prepareLayoutRebind(d,[{slideId:'page-0',layoutId:'target'}],{localAppearance:'preserve'});assert.equal(preserve.status,'ready');d.applyLayoutRebind(preserve);
const adopt=prepareLayoutRebind(d,[{slideId:'page-1',layoutId:'target'}],{localAppearance:'adopt',reset:['geometry','text','background']});assert.equal(adopt.status,'ready');d.applyLayoutRebind(adopt);
const after=await generatePptx(d),l=resolveLayout(d);await save('after.pptx',after.bytes);await save('after-normalized.json',normalizePresentation(d));await save('after-layout.json',l);await save('after-origins.json',resolvePresentationInheritance(d).origins);await save('plans.json',{preserve,adopt});await save('warnings.json',{before:before.warnings,after:after.warnings});for(const[i,s]of(await renderLayoutToSvg(l)).slides.entries())await save('after-'+i+'.svg',s.svg);
await writeFile(path.join(nativeDirectory,'input.pptx'),after.bytes);await writeFile(path.join(nativeDirectory,'before.pptx'),before.bytes);const manifest={schemaVersion:1,recordedAt:new Date().toISOString(),directory,nativeDirectory,beforeSha256:sha(before.bytes),sourceSha256:sha(after.bytes),artifacts,nativeApplication:'pending actual explicit-zero/max-index opening and three native reopen rounds'};await writeFile('docs/testing/placeholder-index-evidence.json',JSON.stringify(manifest,null,2)+'\n');console.log(JSON.stringify({directory,nativeDirectory,beforeSha256:manifest.beforeSha256,sourceSha256:manifest.sourceSha256,warnings:{before:before.warnings.length,after:after.warnings.length}}));
