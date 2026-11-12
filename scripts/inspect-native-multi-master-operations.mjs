import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { chromium } from '@playwright/test';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
const hash=b=>createHash('sha256').update(b).digest('hex');
const evidenceFile=path.resolve('docs/testing/multi-master-operations-evidence.json'),evidenceBytes=await readFile(evidenceFile),evidence=JSON.parse(evidenceBytes);
const options={projection:'content',textFontProfile:'script-families-v3',notesTextProfile:'master-style-v1',editSaveProfile:'text-runs-v3'};
async function exercise(api,bytes,options){
 const s=await api.importPptx(Uint8Array.from(bytes),options);
 try{
  const document=api.normalizePresentation(s.document),report=s.report,first=s.document.slides[0],original=structuredClone(first.elements[0]);
  const noOp=await s.prepareSave({mode:'no-op'});if(noOp.status!=='ready')throw Error('No-op blocked');
  const unchanged=[...(await s.save(noOp.plan)).bytes];
  const candidate=structuredClone(original);candidate.content[0].runs.at(-1).text+='X';
  const mixed=structuredClone(candidate);mixed.content[0].runs.at(-1).style.fontFamily='Forbidden';
  first.replaceElement(original.id,mixed);if((await s.prepareSave({mode:'edit'})).status!=='blocked')throw Error('Mixed style write admitted');
  first.replaceElement(original.id,original);const restored=await s.prepareSave({mode:'no-op'});if(restored.status!=='ready')throw Error('Restored no-op blocked');
  if(!unchanged.every((v,i)=>v===bytes[i]))throw Error('No-op changed original bytes');
  const restore=[...(await s.save(restored.plan)).bytes];if(restore.length!==bytes.length||restore.some((v,i)=>v!==bytes[i]))throw Error('Restored bytes changed');
  first.replaceElement(original.id,candidate);const plan=await s.prepareSave({mode:'edit'});if(plan.status!=='ready')throw Error('Bound scalar write blocked '+JSON.stringify(plan));
  const saved=await s.save(plan.plan),reopened=await api.importPptx(saved.bytes,options);
  try{const expected=api.normalizePresentation(s.document),actual=api.normalizePresentation(reopened.document);if(JSON.stringify(actual)!==JSON.stringify(expected))throw Error('Scalar candidate snapshot differs');}
  finally{reopened.dispose();}
  return{document,report,noOpBytes:unchanged,candidateBytes:[...saved.bytes],lastSourceText:original.content[0].runs.at(-1).text};
 }finally{s.dispose();}
}
const browser=await chromium.launch(),page=await browser.newPage(),requests=[],records=[];
await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});
try{
 const moduleUrl=text=>'data:text/javascript;base64,'+Buffer.from(text).toString('base64');
 const core=moduleUrl(await readFile('packages/core/dist/browser/index.js','utf8'));
 const parser=moduleUrl((await readFile('packages/pptx-parser/dist/browser/index.js','utf8')).replaceAll('"@pptkit/core"',JSON.stringify(core)));
 for(const name of ['core','pptx-parser'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const directory=path.join(evidence.directory,'operation-parser-inspections');await mkdir(directory,{recursive:true});
 for(const stage of evidence.stages)for(const input of stage.rounds){
  const bytes=await readFile(input.file);assert.equal(hash(bytes),input.sha256);
  const expected=await exercise({importPptx,normalizePresentation},[...bytes],options);
  const colors=stage.stage==='theme'?['4472C4','0044CC','CC2200','112233']:['0044CC','0044CC','CC2200','112233'];
  const families=stage.stage==='theme'?['等线','Times New Roman','Georgia','Arial']:['Times New Roman','Times New Roman','Georgia','Arial'];
  for(const [i,slide]of expected.document.slides.entries()){
   assert.equal(slide.elements.length,1);const e=slide.elements[0];assert.deepEqual(e.box,{x:10,y:20,width:300,height:60});
   assert.equal(e.content[0].runs.map(r=>r.text).join(''),'Owned '+i+(i===0&&stage.stage!=='layout'?' EDIT':''));
   assert.ok(e.content[0].runs.every(r=>r.style.fontFamily===families[i]&&r.style.color===colors[i]));
  }
  assert.deepEqual(expected.noOpBytes,[...bytes]);
  const source=readGeneratedZip(bytes),candidate=readGeneratedZip(Uint8Array.from(expected.candidateBytes));assert.deepEqual([...source.keys()],[...candidate.keys()]);
  const part=input.domains[0].slide,from='<a:t>'+expected.lastSourceText+'</a:t>',to='<a:t>'+expected.lastSourceText+'X</a:t>';
  for(const [name,body]of source){if(name===part){assert.equal(body.toString().split(from).length,2);assert.equal(candidate.get(name).toString(),body.toString().replace(from,to));}else assert.deepEqual(candidate.get(name),body);}
  for(const format of ['esm','global']){
   const actual=await page.evaluate(async({bytes,options,core,parser,exercise})=>{
    const api=parser?{...await import(core),...await import(parser)}:PPTKit;
    return await Function('api','bytes','options',`return (${exercise})(api,bytes,options)`)(api,bytes,options);
   },{bytes:[...bytes],options,core,parser:format==='esm'?parser:null,exercise:exercise.toString()});assert.deepEqual(actual,expected);
  }
  const snapshot=JSON.stringify({document:expected.document,report:expected.report},null,2)+'\n',snapshotFile=path.join(directory,path.basename(input.file)+'.json'),candidateFile=path.join(directory,path.basename(input.file));
  await writeFile(snapshotFile,snapshot);await writeFile(candidateFile,Uint8Array.from(expected.candidateBytes));
  records.push({stage:stage.stage,file:input.file,sha256:hash(bytes),snapshotFile,snapshotSha256:hash(snapshot),candidateFile,candidateSha256:hash(Uint8Array.from(expected.candidateBytes)),fullNodeEsmGlobalEqual:true,exactNoOpAndRestore:true,mixedStyleBlocked:true,exactOutsideTargetCharacters:true,candidateNativeReopen:'not performed'});
 }
 assert.deepEqual(requests,[]);
}finally{await browser.close();}
const result={schemaVersion:1,nativeEvidence:evidenceFile,nativeEvidenceSha256:hash(evidenceBytes),options,records,total:{nativeInputs:records.length,consumerNoOps:records.length*3,consumerMixedStyleRejections:records.length*3,consumerScalarCandidates:records.length*3,actualCandidateFiles:records.length},scope:'Current captured-profile content, inherited geometry/owner theme values, exact no-op/restoration and existing scalar source writer only. Headless candidates are not native reopened. No source layout/theme/style write authority.'};
await writeFile('docs/testing/multi-master-operations-parser-evidence.json',JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify(result.total));
