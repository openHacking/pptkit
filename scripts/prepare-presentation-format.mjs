import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdir,mkdtemp,readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {createPresentation,normalizePresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {masterArtworkDocument} from '../tests/fixtures/master-artwork.mjs';
import {presentationFormatDocuments} from '../tests/fixtures/presentation-format.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');await mkdir('.tmp/presentation-format',{recursive:true});
const directory=path.resolve(await mkdtemp('.tmp/presentation-format/run-')),nativeDirectory=path.resolve(await mkdtemp('.tmp/presentation-format/native-'));
const artifacts=[];async function save(name,value){const file=path.join(directory,name);await writeFile(file,typeof value==='string'||value instanceof Uint8Array?value:JSON.stringify(value,null,2)+'\n');const bytes=await readFile(file);artifacts.push({file,bytes:bytes.length,sha256:sha(bytes)});}
const originalManifest=JSON.parse(await readFile('docs/testing/master-artwork-evidence.json','utf8')),originalFile=path.join(originalManifest.nativeDirectory,'input.pptx'),original=await readFile(originalFile),d=masterArtworkDocument(createPresentation),out=await generatePptx(d);
assert.equal(sha(original),'7773cbaa76a74d1bea8d714dc60ece07a2423b94b01b72aee441e061d535940b');assert.deepEqual(out.warnings,[]);
const before=readGeneratedZip(original),after=readGeneratedZip(out.bytes);assert.deepEqual([...after.keys()].sort(),[...before.keys()].sort());const changed=[...before].filter(([name,bytes])=>!bytes.equals(after.get(name))).map(([name])=>name);assert.deepEqual(changed,['docProps/app.xml']);assert.equal(after.get(changed[0]).toString(),before.get(changed[0]).toString().replace('<PresentationFormat>On-screen Show (16:9)</PresentationFormat>','<PresentationFormat>On-screen Show (4:3)</PresentationFormat>'));
await save('source.pptx',out.bytes);await save('normalized.json',normalizePresentation(d));
const scenes=[];for(const s of presentationFormatDocuments(createPresentation)){const p=await generatePptx(s.document);assert.deepEqual(p.warnings,[]);await save(s.document.id+'.pptx',p.bytes);scenes.push({id:s.document.id,expected:s.label,cx:s.cx,cy:s.cy,sha256:sha(p.bytes)});}
await writeFile(path.join(nativeDirectory,'input.pptx'),out.bytes);await writeFile(path.join(nativeDirectory,'input-binding.json'),JSON.stringify({sourceFile:path.join(directory,'source.pptx'),sha256:sha(out.bytes)},null,2)+'\n');
const manifest={schemaVersion:1,recordedAt:new Date().toISOString(),directory,nativeDirectory,sourceSha256:sha(out.bytes),oldSourceFile:originalFile,oldSourceSha256:sha(original),changedParts:changed,delta:'Only PresentationFormat 16:9→4:3 in docProps/app.xml; all other named ZIP payloads are exact.',scenes,artifacts,nativeApplication:'pending actual metadata-corrected first-open and reopen controls'};await writeFile('docs/testing/presentation-format-evidence.json',JSON.stringify(manifest,null,2)+'\n');console.log(JSON.stringify({directory,nativeDirectory,sourceSha256:manifest.sourceSha256,changedParts:changed,scenes:scenes.length}));
