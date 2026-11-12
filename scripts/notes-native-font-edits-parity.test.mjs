import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { nativeNotesFontEditsSourceBytes } from "../tests/fixtures/native-notes-font-edits.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options={projection:"content",notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3"};
const ea=new Set(["Hans","Hant","Jpan","Hang"]),moduleUrl=s=>`data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
async function rounds(f){
  let bytes=f.bytes;const results=[];
  for(let round=0;round<3;round++){
    const s=await importPptx(bytes,options);
    try{
      const document=normalizePresentation(s.document),report=s.report;
      for(const [i,c]of f.controls.entries()){
        const runs=document.slides[i].notes[0].runs,target=f.targets.includes(i+1);assert.equal(runs.length,target?2:1);assert.equal(runs[0].text,c.text);assert.equal(runs[0].style.language,c.language);
        const map=c.id.startsWith("local-")?undefined:Object.fromEntries(Object.entries(f.map).filter(([script])=>!(["empty-ea-tc","bad-ea-tc"].includes(c.id)&&ea.has(script))));
        for(const run of runs){assert.deepEqual(run.style.scriptFontFamilies,map);assert.equal(run.style.fontSize,40);}
        assert.equal(runs[0].style.bold??false,false);assert.equal(runs[0].style.italic??false,false);
        if(target){assert.equal(runs[1].text,"Ω"+"X".repeat(round));assert.equal(runs[1].style.language,"en-US");assert.equal(runs[1].style.bold??false,f.stage==="bold");assert.equal(runs[1].style.italic??false,f.stage==="italic");}
      }
      const noOp=await s.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),Buffer.from(bytes));
      const slide=s.document.slides[f.targets[0]-1],old=structuredClone(slide.notes),wrong=structuredClone(old);wrong[0].runs[1].style.bold=f.stage!=="bold";slide.setNotes(wrong);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");slide.setNotes(old);
      for(const i of f.targets){const slide=s.document.slides[i-1],notes=structuredClone(slide.notes);notes[0].runs[1].text+="X";slide.setNotes(notes);}
      const plan=await s.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=(await s.save(plan.plan)).bytes,before=readGeneratedZip(bytes),after=readGeneratedZip(saved);assert.deepEqual([...after.keys()],[...before.keys()]);
      for(const [name,body]of before){if(f.targets.some(i=>name===`ppt/notesSlides/notesSlide${i}.xml`)){const from=`<a:t>Ω${"X".repeat(round)}</a:t>`,to=`<a:t>Ω${"X".repeat(round+1)}</a:t>`;assert.equal(body.toString().split(from).length,2);assert.equal(after.get(name).toString(),body.toString().replace(from,to));}else assert.deepEqual(after.get(name),body);}
      const candidate=await importPptx(saved,options);try{assert.deepEqual(normalizePresentation(candidate.document),normalizePresentation(s.document));}finally{candidate.dispose();}
      results.push({document,report,bytes:[...saved]});bytes=saved;
    }finally{s.dispose();}
  }
  return results;
}

test("two-run notes and local bold/pure-italic/regular source controls preserve maps/owners across three saves and public consumers",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());const core=moduleUrl(await readFile("packages/core/dist/browser/index.js","utf8")),parser=moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js","utf8")).replaceAll('"@pptkit/core"',JSON.stringify(core)));
  for(const variant of ["A","B"])for(const stage of ["append","bold","italic","regular"]){
    const f=await nativeNotesFontEditsSourceBytes(variant,stage),expected=await rounds(f),parts=readGeneratedZip(f.bytes),page=await browser.newPage(),requests=[];await page.route("**/*",r=>{requests.push(r.request().url());return r.abort();});
    const checked=await page.evaluate(({parts,properties})=>{
      const docs=Object.fromEntries(Object.entries(parts).map(([n,x])=>["/"+n,new DOMParser().parseFromString(x,"application/xml")]));const select=s=>{const doc=docs[s.part];return doc.evaluate(s.path,doc,p=>doc.documentElement.lookupNamespaceURI(p),XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;};let count=0;
      for(const p of properties.filter(p=>p.target.part.includes("/notesSlides/"))){if(!select(p.target)||!select(p.source))throw Error("Missing actual notes source/target");for(const d of p.resolvedFrom??[])if(!select(d))throw Error("Missing notes dependency");if(p.status==="mapped"&&p.property.startsWith("run.scriptFontFamilies.")){if(select(p.source).value!==p.value)throw Error("Wrong source rule");count++;}if(p.status==="mapped"&&["run.bold","run.italic"].includes(p.property)){if((select(p.source).value==="1")!==p.value)throw Error("Wrong local style value");count++;}}
      return count;
    },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith(".xml")||n.endsWith(".rels")).map(([n,b])=>[n,b.toString()])),properties:expected[0].report.textStyleInheritance.properties});assert.ok(checked>0);
    for(const format of ["esm","global"]){
      if(format==="global")for(const name of ["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
      const actual=await page.evaluate(async({bytes,targets,stage,options,core,parser})=>{
        const c=core?await import(core):PPTKit,p=parser?await import(parser):PPTKit,results=[];
        for(let round=0;round<3;round++){const s=await p.importPptx(Uint8Array.from(bytes),options);try{
          const document=c.normalizePresentation(s.document),report=s.report,noOp=await s.prepareSave({mode:"no-op"});if(noOp.status!=="ready")throw Error("No-op blocked");const saved=[...(await s.save(noOp.plan)).bytes];if(saved.length!==bytes.length||saved.some((b,i)=>b!==bytes[i]))throw Error("No-op differs");
          const slide=s.document.slides[targets[0]-1],old=structuredClone(slide.notes),wrong=structuredClone(old);wrong[0].runs[1].style.bold=stage!=="bold";slide.setNotes(wrong);if((await s.prepareSave({mode:"edit"})).status!=="blocked")throw Error("Style writer admitted");slide.setNotes(old);
          for(const i of targets){const slide=s.document.slides[i-1],notes=structuredClone(slide.notes);notes[0].runs[1].text+="X";slide.setNotes(notes);}const plan=await s.prepareSave({mode:"edit"});if(plan.status!=="ready")throw Error("Scalar blocked");bytes=[...(await s.save(plan.plan)).bytes];results.push({document,report,bytes});
        }finally{s.dispose();}}
        return results;
      },{bytes:[...f.bytes],targets:f.targets,stage,options,core:format==="esm"?core:null,parser:format==="esm"?parser:null});assert.deepEqual(actual,expected);
    }
    assert.deepEqual(requests,[]);await page.close();
  }
});
