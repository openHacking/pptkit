// Independent single-master override and diagnosed non-conforming source-preservation evidence.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { notesMasterIsolationSourceBytes, notesOverrideIsolationSourceBytes } from "../tests/fixtures/notes-master-isolation.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if(process.argv.length!==2)throw Error("Usage: node scripts/inspect-notes-master-isolation.mjs");
const invalid=await notesMasterIsolationSourceBytes({listedSecond:true}),groups=[],options={projection:"content",notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3"},hash=b=>createHash("sha256").update(b).digest("hex");
await mkdir(".tmp/notes-master-isolation",{recursive:true});const base=await mkdtemp(path.resolve(".tmp/notes-master-isolation/run-"));
for(const f of [await notesMasterIsolationSourceBytes(),await notesOverrideIsolationSourceBytes()]){
const directory=path.join(base,f.nonConforming?"unlisted":"override"),snapshots=[],hashes=[];await mkdir(directory);
await writeFile(path.join(directory,"controls.json"),JSON.stringify(f.controls,null,2)+"\n");await writeFile(path.join(directory,"listed-two-rejected.pptx"),invalid.bytes);
let rejection;try{const s=await importPptx(invalid.bytes,options);s.dispose();}catch(e){rejection={code:e.code,part:e.source?.part,message:e.message};}assert.equal(rejection?.code,"import-invalid-package");assert.equal(rejection.part,"/ppt/presentation.xml");
let bytes=f.bytes;
for(let stage=0;stage<=3;stage++){
  const file=`round-${stage}.pptx`;await writeFile(path.join(directory,file),bytes);hashes.push({file,sha256:hash(bytes)});const s=await importPptx(bytes,options);
  try{
    const document=normalizePresentation(s.document),report=s.report;assert.equal(report.inspection.findings.filter(p=>p.code==="unlisted-master").length,f.nonConforming?15:0);const noOp=await s.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),Buffer.from(bytes));snapshots.push({file,document,report});if(stage===3)break;
    const slide=s.document.slides[1],old=structuredClone(slide.notes),wrong=structuredClone(old);wrong[0].runs[0].style.fontFamily="Forbidden";slide.setNotes(wrong);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");slide.setNotes(old);
    const variant=stage%2===0?"A":"B",targets=new Set();for(const [i,c]of f.controls.entries())if(c.variant===variant){const slide=s.document.slides[i],notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);targets.add(c.notesPart);}
    const ready=await s.prepareSave({mode:"edit"});assert.equal(ready.status,"ready");const saved=(await s.save(ready.plan)).bytes,before=readGeneratedZip(bytes),after=readGeneratedZip(saved);assert.deepEqual([...after.keys()],[...before.keys()]);
    for(const [name,body]of before)if(!targets.has(name))assert.deepEqual(after.get(name),body);
    const candidate=await importPptx(saved,options);try{assert.deepEqual(normalizePresentation(candidate.document),normalizePresentation(s.document));assert.deepEqual(candidate.report.inspection.findings,report.inspection.findings);}finally{candidate.dispose();}bytes=saved;
  }finally{s.dispose();}
}
await writeFile(path.join(directory,"snapshots.json"),JSON.stringify(snapshots)+"\n");
const independent=JSON.parse((await promisify(execFile)("python3",["scripts/research/oracles/notes-master-isolation.py",directory],{timeout:120000,maxBuffer:8*1024*1024})).stdout);
const evidence={classification:f.nonConforming?"non-conforming-multi-notes-master-adversarial-source":"single-master-notes-font-override-source",sourceAssembly:f.sourceAssembly,hashes,listedTwoRejection:{...rejection,sha256:hash(invalid.bytes)},independent,total:{sourcePages:30,states:4,nodeExactNoOps:4,scalarRounds:3,styleWriteRejections:3,unlistedFindingsPerState:f.nonConforming?15:0,actualCorruptionRejections:independent.rejectedCount},nativeApplication:"not tested",fullSchemaConformance:f.nonConforming?false:"not validated",scope:"Located source projection/preservation and cardinality controls; no native typography or full schema certificate."};
await writeFile(path.join(directory,"evidence.json"),JSON.stringify(evidence,null,2)+"\n");groups.push(evidence);
}
const manifest={schemaVersion:1,recordedAt:new Date().toISOString(),artifactsDirectory:path.relative(process.cwd(),base),standardReference:"https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document#notes-master-part",options,groups,total:{sourcePages:60,states:8,nodeExactNoOps:8,scalarRounds:6,styleWriteRejections:6,actualCorruptionRejections:groups.reduce((n,g)=>n+g.total.actualCorruptionRejections,0)},nativeApplication:"not tested",scope:"Single-master override isolation and diagnosed non-conforming multi-master tolerance are separate controls."};
await writeFile("docs/testing/notes-master-isolation-evidence.json",JSON.stringify(manifest,null,2)+"\n");console.log(JSON.stringify({directory:base,...manifest.total}));
