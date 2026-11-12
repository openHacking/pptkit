import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { notesMasterIsolationSourceBytes, notesOverrideIsolationSourceBytes } from "../tests/fixtures/notes-master-isolation.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const options={projection:"content",notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3"};
const ea=new Set(["Hans","Hant","Jpan","Hang"]),moduleUrl=s=>`data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
const encode=parts=>zip([...parts].map(([name,body])=>({name,body}))).bytes;
const mapFor=c=>c.id.startsWith("local-")?undefined:Object.fromEntries(Object.entries(c.map).filter(([script])=>!(["empty-ea-tc","bad-ea-tc"].includes(c.id)&&ea.has(script))));

async function sourceRounds(f){
  let bytes=f.bytes;const results=[];
  for(let round=0;round<3;round++){
    const s=await importPptx(bytes,options);
    try{
      const document=normalizePresentation(s.document),report=s.report;assert.equal(document.slides.length,30);assert.equal(report.inspection.notesMasters.length,1);
      const unlisted=report.inspection.findings.filter(p=>p.code==="unlisted-master");assert.equal(unlisted.length,f.nonConforming?15:0);assert.deepEqual(unlisted.map(p=>p.part).sort(),f.controls.filter(c=>f.nonConforming&&c.variant==="B").map(c=>"/"+c.notesPart).sort());
      for(const [i,c]of f.controls.entries()){
        const run=document.slides[i].notes[0].runs[0];assert.equal(run.text,c.text+"X".repeat(c.variant==="A"?Math.ceil(round/2):Math.floor(round/2)));assert.equal(run.style.fontSize,40);assert.equal(run.style.language,c.language);assert.deepEqual(run.style.scriptFontFamilies,mapFor(c));assert.equal(run.style.fontFamily,c.id.startsWith("local-")?"Arial":c.map.Latn);assert.equal(run.style.eastAsianFontFamily,c.id.startsWith("local-")?"Arial Unicode MS":undefined);assert.equal(run.style.complexScriptFontFamily,c.id.startsWith("local-")?"Arial Unicode MS":undefined);
        const properties=report.textStyleInheritance.properties.filter(p=>p.target.part==="/"+c.notesPart);for(const p of properties.filter(p=>p.property.startsWith("run.scriptFontFamilies.")&&p.status==="mapped")){assert.equal(p.source.part,"/"+c.themePart);assert.equal(p.value,mapFor(c)[p.property.split(".").at(-1)]);}
        const size=properties.find(p=>p.property==="run.fontSize");assert.equal(size.source.part,"/"+c.masterPart);assert.equal(size.level,"notes-master-placeholder");
      }
      const noOp=await s.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),Buffer.from(bytes));
      const first=s.document.slides[1],old=structuredClone(first.notes),wrong=structuredClone(old);wrong[0].runs[0].style.fontFamily="Forbidden";first.setNotes(wrong);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");first.setNotes(old);
      const variant=round%2===0?"A":"B",changed=new Set();for(const [i,c]of f.controls.entries())if(c.variant===variant){const slide=s.document.slides[i],notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);changed.add(c.notesPart);}
      const plan=await s.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=(await s.save(plan.plan)).bytes,before=readGeneratedZip(bytes),after=readGeneratedZip(saved);assert.deepEqual([...after.keys()],[...before.keys()]);
      for(const [part,body]of before){if(changed.has(part)){const c=f.controls.find(c=>c.notesPart===part),count=variant==="A"?Math.ceil(round/2):Math.floor(round/2),from=`<a:t>${c.text}${"X".repeat(count)}</a:t>`,to=from.replace("</a:t>","X</a:t>");assert.equal(body.toString().split(from).length,2);assert.equal(after.get(part).toString(),body.toString().replace(from,to));}else assert.deepEqual(after.get(part),body);}
      const candidate=await importPptx(saved,options);try{assert.deepEqual(normalizePresentation(candidate.document),normalizePresentation(s.document));assert.deepEqual(candidate.report.inspection.findings,report.inspection.findings);}finally{candidate.dispose();}
      results.push({document,report,bytes:[...saved]});bytes=saved;
    }finally{s.dispose();}
  }
  return results;
}

test("single-master font overrides and diagnosed malformed master controls isolate interleaved A/B/A source saves across consumers",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());
  for(const f of [await notesMasterIsolationSourceBytes(),await notesOverrideIsolationSourceBytes(),await notesOverrideIsolationSourceBytes({distinctParts:true})]){
  const expected=await sourceRounds(f),parts=readGeneratedZip(f.bytes),page=await browser.newPage(),requests=[];await page.route("**/*",r=>{requests.push(r.request().url());return r.abort();});
  const independent=await page.evaluate(({parts,properties})=>{
    const P="http://schemas.openxmlformats.org/presentationml/2006/main",R="http://schemas.openxmlformats.org/officeDocument/2006/relationships",REL="http://schemas.openxmlformats.org/package/2006/relationships",A="http://schemas.openxmlformats.org/drawingml/2006/main";
    const docs=Object.fromEntries(Object.entries(parts).map(([n,x])=>["/"+n,new DOMParser().parseFromString(x,"application/xml")]));for(const doc of Object.values(docs))if(doc.getElementsByTagName("parsererror").length)throw Error("Invalid XML");
    const relation=(owner,kind,id,optional=false)=>{const slash=owner.lastIndexOf("/"),part=owner.slice(0,slash+1)+"_rels/"+owner.slice(slash+1)+".rels",nodes=[...docs[part].getElementsByTagNameNS(REL,"Relationship")].filter(n=>n.getAttribute("Type").endsWith("/"+kind)&&(id===undefined||n.getAttribute("Id")===id));if(optional&&nodes.length===0)return undefined;if(nodes.length!==1||nodes[0].getAttribute("TargetMode")==="External")throw Error("Ambiguous/external edge");return new URL(nodes[0].getAttribute("Target"),"https://fixture.invalid"+owner).pathname;};
    const select=s=>{const doc=docs[s.part];if(!doc)throw Error("Invented source part");return doc.evaluate(s.path,doc,p=>doc.documentElement.lookupNamespaceURI(p),XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;};
    const rows=[];for(const slide of docs["/ppt/presentation.xml"].getElementsByTagNameNS(P,"sldId")){const part=relation("/ppt/presentation.xml","slide",slide.getAttributeNS(R,"id")),notes=relation(part,"notesSlide"),master=relation(notes,"notesMaster"),theme=relation(notes,"themeOverride",undefined,true)??relation(master,"theme");if(relation(notes,"slide")!==part)throw Error("Wrong notes backlink");const font=docs[theme].getElementsByTagNameNS(A,"minorFont")[0];rows.push({part,notes,master,theme,map:Object.fromEntries([...font.children].filter(n=>n.localName==="font").map(n=>[n.getAttribute("script"),n.getAttribute("typeface")]))});}
    let count=0;for(const p of properties.filter(p=>p.target.part.includes("/notesSlides/"))){if(!select(p.source)||!select(p.target))throw Error("Missing independent provenance");for(const d of p.resolvedFrom??[])if(!select(d))throw Error("Missing source dependency");if(p.status==="mapped"&&p.property.startsWith("run.scriptFontFamilies.")){const row=rows.find(r=>r.notes===p.target.part);if(p.source.part!==row.theme||select(p.source).value!==p.value)throw Error("Borrowed notes font source");count++;}}
    return{rows,count,listed:docs["/ppt/presentation.xml"].getElementsByTagNameNS(P,"notesMasterId").length};
  },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith(".xml")||n.endsWith(".rels")).map(([n,b])=>[n,b.toString()])),properties:expected[0].report.textStyleInheritance.properties});
  assert.equal(independent.listed,1);assert.ok(independent.count>0);for(const [i,c]of f.controls.entries()){const row=independent.rows[i];assert.equal(row.part,`/ppt/slides/slide${c.sourceSlide}.xml`);assert.equal(row.notes,"/"+c.notesPart);assert.equal(row.master,"/"+c.masterPart);assert.equal(row.theme,"/"+c.themePart);assert.deepEqual(row.map,c.map);}
  const core=moduleUrl(await readFile("packages/core/dist/browser/index.js","utf8")),parser=moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js","utf8")).replaceAll('"@pptkit/core"',JSON.stringify(core)));
  const invalid=await notesMasterIsolationSourceBytes({listedSecond:true});
  for(const format of ["esm","global"]){
    if(format==="global")for(const name of ["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
    const actual=await page.evaluate(async({bytes,invalid,controls,options,core,parser})=>{
      const c=core?await import(core):PPTKit,p=parser?await import(parser):PPTKit,results=[];let rejected;
      try{const unexpected=await p.importPptx(Uint8Array.from(invalid),options);unexpected.dispose();}catch(e){rejected={code:e.code,part:e.source?.part};}if(rejected?.code!=="import-invalid-package")throw Error("Two listed notes masters admitted");
      for(let round=0;round<3;round++){const s=await p.importPptx(Uint8Array.from(bytes),options);try{
        const document=c.normalizePresentation(s.document),report=s.report,noOp=await s.prepareSave({mode:"no-op"});if(noOp.status!=="ready")throw Error("No-op blocked");const noOpBytes=[...(await s.save(noOp.plan)).bytes];if(noOpBytes.length!==bytes.length||noOpBytes.some((v,i)=>v!==bytes[i]))throw Error("No-op differs");
        const first=s.document.slides[1],old=structuredClone(first.notes),wrong=structuredClone(old);wrong[0].runs[0].style.fontFamily="Forbidden";first.setNotes(wrong);if((await s.prepareSave({mode:"edit"})).status!=="blocked")throw Error("Style writer admitted");first.setNotes(old);
        const variant=round%2===0?"A":"B";for(const [i,control]of controls.entries())if(control.variant===variant){const slide=s.document.slides[i],notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);}const plan=await s.prepareSave({mode:"edit"});if(plan.status!=="ready")throw Error("Scalar blocked");bytes=[...(await s.save(plan.plan)).bytes];results.push({document,report,bytes});
      }finally{s.dispose();}}
      return{results,rejected};
    },{bytes:[...f.bytes],invalid:[...invalid.bytes],controls:f.controls,options,core:format==="esm"?core:null,parser:format==="esm"?parser:null});assert.deepEqual(actual.results,expected);assert.deepEqual(actual.rejected,{code:"import-invalid-package",part:"/ppt/presentation.xml"});
  }
  await assert.rejects(importPptx(invalid.bytes,options),e=>e.code==="import-invalid-package"&&e.source.part==="/ppt/presentation.xml");assert.deepEqual(requests,[]);await page.close();
  }
});

test("missing/bad actual B source stays bounded to B; declaration-list fallback and type-by-index are not invented",async()=>{
  const f=await notesMasterIsolationSourceBytes(),parts=readGeneratedZip(f.bytes);
  for(const kind of ["theme","type","master"]){
    const changed=new Map(parts);
    if(kind==="theme"){const name="ppt/notesMasters/_rels/notesMaster2.xml.rels";changed.set(name,Buffer.from(changed.get(name).toString().replace(/<Relationship\b[^>]*Type="[^"]*\/theme"[^>]*\/>/,"")));}
    if(kind==="type"){const name="ppt/notesMasters/notesMaster2.xml";changed.set(name,Buffer.from(changed.get(name).toString().replace('type="body" idx="99"','type="title" idx="7"')));}
    if(kind==="master"){const name="ppt/notesSlides/_rels/notesSlide16.xml.rels";changed.set(name,Buffer.from(changed.get(name).toString().replace(/<Relationship\b[^>]*Type="[^"]*\/notesMaster"[^>]*\/>/,"")));}
    const s=await importPptx(encode(changed),options);
    try{for(const [i,c]of f.controls.entries()){
      const run=s.document.slides[i].notes[0].runs[0];if(c.variant==="A"){assert.deepEqual(run.style.scriptFontFamilies,mapFor(c));assert.equal(run.style.fontSize,40);}
      else if(kind==="theme"&&!c.id.startsWith("local-")){assert.equal(run.style.fontFamily,undefined);assert.equal(run.style.scriptFontFamilies,undefined);}
      else if(kind==="type")assert.equal(run.style.fontSize,18);
      else if(kind==="master"&&c.sourceSlide===16){assert.equal(run.style.fontFamily,undefined);assert.equal(run.style.scriptFontFamilies,undefined);assert.equal(run.style.fontSize,undefined);}
      else{assert.equal(run.style.fontSize,40);assert.deepEqual(run.style.scriptFontFamilies,mapFor(c));}
    }}finally{s.dispose();}
  }
});


test("a single shared notes master keeps per-notes override replacement local and falls back only along its actual chain",async()=>{
  const f=await notesOverrideIsolationSourceBytes(),parts=readGeneratedZip(f.bytes);
  for(const variant of ["missing-override","missing-Hant"]){
    const changed=new Map(parts);
    if(variant==="missing-override"){const name="ppt/notesSlides/_rels/notesSlide16.xml.rels";changed.set(name,Buffer.from(changed.get(name).toString().replace(/<Relationship\b[^>]*Type="[^"]*\/themeOverride"[^>]*\/>/,"")));}
    else{const name="ppt/theme/notes-font-override-B.xml",xml=changed.get(name).toString();let count=0;changed.set(name,Buffer.from(xml.replace(/<a:font script="Hant"[^>]*\/>/g,()=>{count++;return"";})));assert.equal(count,2);}
    const s=await importPptx(encode(changed),options);
    try{
      assert.equal(s.report.inspection.findings.filter(p=>p.code==="unlisted-master").length,0);
      for(const [i,c]of f.controls.entries()){
        const run=s.document.slides[i].notes[0].runs[0];assert.equal(run.style.fontSize,40);
        let expected=mapFor(c);if(variant==="missing-override"&&c.sourceSlide===16)expected=mapFor(f.controls[0]);else if(variant==="missing-Hant"&&c.variant==="B"&&expected){expected=structuredClone(expected);delete expected.Hant;}
        assert.deepEqual(run.style.scriptFontFamilies,expected);assert.equal(run.style.fontFamily,c.id.startsWith("local-")?"Arial":variant==="missing-override"&&c.sourceSlide===16?f.controls[0].map.Latn:c.map.Latn);
      }
    }finally{s.dispose();}
  }
});
