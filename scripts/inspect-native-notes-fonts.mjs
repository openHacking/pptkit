import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

if(process.argv.length!==3)throw Error("Usage: node scripts/inspect-native-notes-fonts.mjs <native-notes-run-directory>");
const directory=path.resolve(process.argv[2]),exec=promisify(execFile),hash=b=>createHash("sha256").update(b).digest("hex"),groups=[],ea=new Set(["Hans","Hant","Jpan","Hang"]);
const options={projection:"content",notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3"};
const browser=await chromium.launch();
try{
  for(const variant of ["A","B"]){
    const folder=path.join(directory,variant),expected=JSON.parse(await readFile(path.join(folder,"expected.json"),"utf8"));let negatives;
    for(const file of ["native-notes-fonts.py","native-notes-fonts-checks.py"]){const output=(await exec("python3",[path.resolve("scripts/research/oracles",file),folder],{timeout:120000,maxBuffer:4*1024*1024})).stdout;if(file.endsWith("-checks.py"))negatives=JSON.parse(output);else console.log(output.trim());}
    const independent=JSON.parse(await readFile(path.join(folder,"notes-inspection.json"),"utf8")),projections=[];
    for(const name of ["notes-fonts","native-round1","native-round2","native-round3"]){
      const bytes=await readFile(path.join(folder,name+".pptx")),s=await importPptx(bytes,options);let snapshot,report;
      try{
        snapshot=normalizePresentation(s.document);report=s.report;const controls=[];
        for(const [i,c]of expected.controls.entries()){
          const notes=snapshot.slides[i].notes;assert.equal(notes.length,1);assert.equal(notes[0].runs.length,1);const run=notes[0].runs[0];assert.equal(run.text,c.text);assert.equal(run.style.language,c.language);assert.equal(run.style.fontSize,40);
          const map=c.id.startsWith("local-")?undefined:Object.fromEntries(Object.entries(expected.map).filter(([script])=>!(["empty-ea-tc","bad-ea-tc"].includes(c.id)&&ea.has(script))));assert.deepEqual(run.style.scriptFontFamilies,map);
          if(c.id.startsWith("local-")){assert.equal(run.style.fontFamily,"Arial");assert.equal(run.style.eastAsianFontFamily,"Arial Unicode MS");assert.equal(run.style.complexScriptFontFamily,"Arial Unicode MS");}
          const selected=resolveTextFontSegments(c.text,run.style,snapshot.theme);assert.deepEqual(selected,expected.projections[i].selected);
          controls.push({id:c.id,notes,selected});
        }
        const page=await browser.newPage();
        try{
          const checked=await page.evaluate(({parts,properties,theme})=>{
            const docs=Object.fromEntries(Object.entries(parts).map(([n,x])=>["/"+n,new DOMParser().parseFromString(x,"application/xml")]));const select=s=>{const doc=docs[s.part];if(!doc)throw Error("Invented source part");return doc.evaluate(s.path,doc,p=>doc.documentElement.lookupNamespaceURI(p),XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;};let count=0;
            for(const p of properties.filter(p=>p.target.part.includes("/notesSlides/"))){
              if(!select(p.target)||!select(p.source))throw Error("Missing actual notes source/target");for(const d of p.resolvedFrom??[])if(!select(d))throw Error("Invented notes dependency");
              if(p.status!=="mapped")continue;
              if(p.property.startsWith("run.scriptFontFamilies.")){const n=select(p.source);if(p.source.part!==theme||n.value!==p.value||n.ownerElement.getAttribute("script")!==p.property.split(".").at(-1))throw Error("Wrong independent actual theme rule");count++;}
              if(p.property==="run.fontSize"){if(p.level!=="notes-master-placeholder"||p.value!==40||select(p.source).value!=="4000")throw Error("Wrong notes type-match font size");const types=p.resolvedFrom.map(select);if(types.length!==2||types.some(n=>n.value!=="body")||types.map(n=>n.ownerElement.getAttribute("idx")).sort().join(",")!=="7,99")throw Error("Type-vs-index provenance wrong");count++;}
            }
            return count;
          },{parts:Object.fromEntries([...readGeneratedZip(bytes)].filter(([n])=>n.endsWith(".xml")||n.endsWith(".rels")).map(([n,b])=>[n,b.toString()])),properties:report.textStyleInheritance.properties,theme:"/"+(name==="notes-fonts"?expected.themePart:"ppt/theme/theme2.xml")});assert.ok(checked>0);
        }finally{await page.close();}
        const noOp=await s.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),bytes);const result={name,sha256:hash(bytes),controls,exactNoOp:true,independentXPathValidated:true,scalarCandidates:[]};projections.push(result);
        if(name!=="notes-fonts"){
          const first=s.document.slides[0],original=structuredClone(first.notes),wrong=structuredClone(original);wrong[0].runs[0].style.fontFamily="Forbidden";first.setNotes(wrong);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");first.setNotes(original);result.fontWriteBlockedAndRestored=true;let previous=bytes;
          for(let round=1;round<=3;round++){
            for(const slide of s.document.slides){const notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);}
            const plan=await s.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=(await s.save(plan.plan)).bytes,before=readGeneratedZip(previous),after=readGeneratedZip(saved);assert.deepEqual([...after.keys()],[...before.keys()]);
            for(const [part,body]of before){const match=/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/.exec(part);if(match){const from=`<a:t>${expected.controls[Number(match[1])-1].text}${"X".repeat(round-1)}</a:t>`,to=from.replace("</a:t>","X</a:t>");assert.equal(body.toString().split(from).length,2);assert.equal(after.get(part).toString(),body.toString().replace(from,to));}else assert.deepEqual(after.get(part),body);}
            const imported=await importPptx(saved,options);try{assert.deepEqual(normalizePresentation(imported.document),normalizePresentation(s.document));}finally{imported.dispose();}result.scalarCandidates.push({round,sha256:hash(saved),nonCharacterBytesExact:true,reprojectionEqual:true,nativeReopen:"not tested"});previous=saved;
          }
        }
      }finally{s.dispose();}
      if(name==="native-round3"){
        const page=await browser.newPage(),requests=[];await page.route("**/*",r=>{requests.push(r.request().url());return r.abort();});
        try{
          for(const module of ["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${module}/dist/browser/global.js`,"utf8")});
          const actual=await page.evaluate(async({bytes,options})=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),options);try{const document=PPTKit.normalizePresentation(s.document),report=s.report,plan=await s.prepareSave({mode:"no-op"});if(plan.status!=="ready")throw Error("Browser no-op blocked");return{document,report,bytes:[...(await s.save(plan.plan)).bytes]};}finally{s.dispose();}},{bytes:[...bytes],options});
          assert.deepEqual(actual.document,snapshot);assert.deepEqual(actual.report,report);assert.deepEqual(actual.bytes,[...bytes]);assert.deepEqual(requests,[]);projections.at(-1).browserParityAndExactNoOp=true;
        }finally{await page.close();}
      }
    }
    for(const font of expected.fonts)assert.equal(hash(await readFile(font.file)),font.sha256);
    groups.push({variant,sourceFontEnvironment:expected.fonts,independent,negatives,projections});
  }
}finally{await browser.close();}
const app=JSON.parse((await exec("python3",["-c","import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))","/Applications/Microsoft PowerPoint.app/Contents/Info.plist"])).stdout);
const evidence={schemaVersion:1,recordedAt:new Date().toISOString(),application:{name:"Microsoft PowerPoint for Mac",...app},groups,total:{controls:30,matches:groups.reduce((n,g)=>n+g.independent.pdf.matches,0),differences:groups.reduce((n,g)=>n+g.independent.pdf.differences,0),nativeRoundsPerGroup:3,nativeNotesPDFPageRecords:90,sourceNotesPDFPageRecords:30,nodeExactNoOps:8,browserExactNoOps:2,sourceFontWriteRejections:6,scalarCandidates:18,rejectedControls:groups.reduce((n,g)=>n+g.negatives.rejectedCount,0)},scope:"Named independent notes font/no-edit baseline; source assembler patched notes/master/theme/labels. No native text/style/theme edits, complete notes-page imagery/fields, Windows or bidi/mark conformance. Scalar candidates not native reopened."};
await writeFile("docs/testing/native-notes-font-evidence.json",JSON.stringify(evidence,null,2)+"\n");console.log(JSON.stringify(evidence.total));
