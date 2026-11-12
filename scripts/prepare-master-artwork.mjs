import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdir,mkdtemp,readFile,writeFile} from 'node:fs/promises';
import path from 'node:path';
import {chromium} from '@playwright/test';
import {createPresentation,normalizePresentation,resolvePresentationInheritance} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {masterArtworkDocument} from '../tests/fixtures/master-artwork.mjs';
import {textOriginsDocument} from '../tests/fixtures/text-origins.mjs';
const sha=b=>createHash('sha256').update(b).digest('hex');
await mkdir('.tmp/master-artwork',{recursive:true});
const directory=path.resolve(await mkdtemp('.tmp/master-artwork/run-')),nativeDirectory=path.resolve(await mkdtemp('.tmp/master-artwork/native-'));
const d=masterArtworkDocument(createPresentation),r=resolvePresentationInheritance(d),n=normalizePresentation(d),layout=resolveLayout(d),svg=await renderLayoutToSvg(layout),out=await generatePptx(d);
assert.deepEqual(r.presentation,n);assert.deepEqual(out.warnings,[]);
const artifacts=[];async function save(name,value){const file=path.join(directory,name);await writeFile(file,typeof value==='string'||value instanceof Uint8Array?value:JSON.stringify(value,null,2)+'\n');const bytes=await readFile(file);artifacts.push({file,bytes:bytes.length,sha256:sha(bytes)});}
await save('source.pptx',out.bytes);await save('normalized.json',n);await save('origins.json',r.origins);await save('layout.json',layout);await save('fixture.mjs',await readFile('tests/fixtures/master-artwork.mjs','utf8'));
for(const[i,s]of svg.slides.entries())await save(`page-${i+1}.svg`,s.svg);
const old=textOriginsDocument(createPresentation),g=old.slides[0].elements.find(e=>e.id==='group'),{placeholderKey:_binding,...ordinary}=g;old.slides[0].replaceElement(g.id,ordinary);
const legacy=(await generatePptx(old)).bytes;assert.equal(sha(legacy),'ef3d886dcf09d80e23b58cbbafb67cbb26be84fbb2a5c71797c2eb0a3219cd2d');await save('legacy-native-source.pptx',legacy);
const browser=await chromium.launch();let browserVersion;
try{browserVersion=browser.version();const p=await browser.newPage({viewport:{width:1000,height:1140},deviceScaleFactor:1});await p.route('**/*',r=>r.abort());await p.setContent('<style>body{margin:8px;display:grid;grid-template-columns:480px 480px;gap:8px;background:#ddd}svg{width:480px;height:360px;background:white}</style>'+svg.slides.map(s=>s.svg).join(''));await p.evaluate(()=>document.fonts.ready);const file=path.join(directory,'browser-preview.png');await p.screenshot({path:file});const b=await readFile(file);artifacts.push({file,bytes:b.length,sha256:sha(b)});}finally{await browser.close();}
await writeFile(path.join(nativeDirectory,'input.pptx'),out.bytes);await writeFile(path.join(nativeDirectory,'input-binding.json'),JSON.stringify({sourceFile:path.join(directory,'source.pptx'),sha256:sha(out.bytes)},null,2)+'\n');
const manifest={schemaVersion:1,recordedAt:new Date().toISOString(),directory,nativeDirectory,sourceSha256:sha(out.bytes),artifacts,originRecords:r.origins.length,browserVersion,warnings:out.warnings,expectedPages:[{master:'red',layout:'red-on',masterShapes:true,layoutShapes:true},{master:'red',layout:'red-on',masterShapes:false,layoutShapes:false},{master:'red',layout:'red-off',masterShapes:false,layoutShapes:true},{master:'blue',layout:'blue-on',masterShapes:true,layoutShapes:true},{master:'red',layout:'red-off',masterShapes:false,layoutShapes:true},{master:'red',layout:'red-on',masterShapes:true,layoutShapes:true}],nativeApplication:'pending actual first-open and reopen verification'};
await writeFile('docs/testing/master-artwork-evidence.json',JSON.stringify(manifest,null,2)+'\n');console.log(JSON.stringify({directory,nativeDirectory,sha256:manifest.sourceSha256,artifacts:artifacts.length}));
