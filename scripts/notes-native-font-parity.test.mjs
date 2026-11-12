import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { nativeNotesFontBytes } from "../tests/fixtures/native-notes-fonts.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const options={projection:"content",notesTextProfile:"master-style-v1",textFontProfile:"script-families-v3",editSaveProfile:"text-runs-v3"};
const ea=new Set(["Hans","Hant","Jpan","Hang"]),moduleUrl=s=>`data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
async function rounds(fixture){
  let bytes=fixture.bytes;const results=[];
  for(let round=0;round<3;round++){
    const s=await importPptx(bytes,options);
    try{
      const document=normalizePresentation(s.document),report=s.report;
      for(const [i,c]of fixture.controls.entries()){
        const runs=document.slides[i].notes[0].runs;assert.equal(runs.length,1);const run=runs[0];assert.equal(run.text,c.text+"X".repeat(round));assert.equal(run.style.language,c.language);assert.equal(run.style.fontSize,40);
        const map=c.id.startsWith("local-")?undefined:Object.fromEntries(Object.entries(fixture.map).filter(([script])=>!(["empty-ea-tc","bad-ea-tc"].includes(c.id)&&ea.has(script))));
        assert.deepEqual(run.style.scriptFontFamilies,map);
        if(c.id.startsWith("local-")){assert.equal(run.style.eastAsianFontFamily,"Arial Unicode MS");assert.equal(run.style.complexScriptFontFamily,"Arial Unicode MS");}
        else if(!c.id.includes("ea-tc"))assert.equal(resolveTextFontSegments(c.text,run.style,document.theme)[0].family,fixture.map[c.script]);
        const size=report.textStyleInheritance.properties.find(p=>p.target.part===`/ppt/notesSlides/notesSlide${i+1}.xml`&&p.property==="run.fontSize");assert.equal(size.level,"notes-master-placeholder");assert.equal(size.value,40);
        const scriptRules=report.textStyleInheritance.properties.filter(p=>p.target.part===`/ppt/notesSlides/notesSlide${i+1}.xml`&&p.property.startsWith("run.scriptFontFamilies.")&&p.status==="mapped");
        for(const r of scriptRules){assert.equal(r.source.part,"/"+fixture.themePart);assert.equal(r.value,map[r.property.split(".").at(-1)]);}
      }
      const noOp=await s.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual(Buffer.from((await s.save(noOp.plan)).bytes),Buffer.from(bytes));
      const original=structuredClone(s.document.slides[0].notes),wrong=structuredClone(original);wrong[0].runs[0].style.fontFamily="Forbidden";s.document.slides[0].setNotes(wrong);assert.equal((await s.prepareSave({mode:"edit"})).status,"blocked");s.document.slides[0].setNotes(original);
      for(const slide of s.document.slides){const notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);}
      const plan=await s.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const saved=(await s.save(plan.plan)).bytes,before=readGeneratedZip(bytes),after=readGeneratedZip(saved);assert.deepEqual([...after.keys()],[...before.keys()]);
      for(const [name,body]of before){
        const match=/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/.exec(name);
        if(match){const from=`<a:t>${fixture.controls[Number(match[1])-1].text}${"X".repeat(round)}</a:t>`,to=from.replace("</a:t>","X</a:t>");assert.equal(body.toString().split(from).length,2);assert.equal(after.get(name).toString(),body.toString().replace(from,to));}
        else assert.deepEqual(after.get(name),body);
      }
      const r=await importPptx(saved,options);try{assert.deepEqual(normalizePresentation(r.document),normalizePresentation(s.document));}finally{r.dispose();}
      results.push({document,report,bytes:[...saved]});bytes=saved;
    }finally{s.dispose();}
  }
  return results;
}

test("real-family notes source controls isolate actual theme, type matching, empty/bad/local slots and three saves in Node/ESM/global",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());const core=moduleUrl(await readFile("packages/core/dist/browser/index.js","utf8")),parser=moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js","utf8")).replaceAll('"@pptkit/core"',JSON.stringify(core)));
  for(const variant of ["A","B"]){
    const fixture=await nativeNotesFontBytes(variant),expected=await rounds(fixture),parts=readGeneratedZip(fixture.bytes);
    const page=await browser.newPage(),requests=[];await page.route("**/*",r=>{requests.push(r.request().url());return r.abort();});
    const independent=await page.evaluate(({parts,properties})=>{
      const A="http://schemas.openxmlformats.org/drawingml/2006/main",P="http://schemas.openxmlformats.org/presentationml/2006/main",R="http://schemas.openxmlformats.org/package/2006/relationships",docs=Object.fromEntries(Object.entries(parts).map(([n,x])=>["/"+n,new DOMParser().parseFromString(x,"application/xml")]));
      const select=s=>{const doc=docs[s.part];return doc.evaluate(s.path,doc,p=>doc.documentElement.lookupNamespaceURI(p),XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;};
      const master=docs["/ppt/notesMasters/notesMaster1.xml"],local=docs["/ppt/notesSlides/notesSlide1.xml"],type=d=>[...d.getElementsByTagNameNS(P,"ph")].find(n=>n.getAttribute("type")==="body");
      if(type(master).getAttribute("idx")!=="99"||type(local).getAttribute("idx")!=="7")throw Error("Did not distinguish type from idx");
      const image = doc => [...doc.getElementsByTagNameNS(P,"sp")].filter(s => [...s.getElementsByTagNameNS(P,"ph")].some(p => p.getAttribute("type") === "sldImg"));
      const masterImages = image(master);
      if(masterImages.length !== 1)throw Error("Missing/ambiguous master slide image");
      const geometry = masterImages[0].getElementsByTagNameNS(A,"prstGeom");
      if(geometry.length !== 1 || geometry[0].getAttribute("prst") !== "rect" || geometry[0].getElementsByTagNameNS(A,"avLst").length !== 1)throw Error("Master slide image needs explicit rect geometry for native printing");
      const relation = (owner, kind) => {
        const slash = owner.lastIndexOf("/"), rels = docs[owner.slice(0,slash+1)+"_rels/"+owner.slice(slash+1)+".rels"];
        const edges = [...rels.getElementsByTagNameNS(R,"Relationship")].filter(r => r.getAttribute("Type").endsWith("/"+kind));
        if(edges.length !== 1 || edges[0].getAttribute("TargetMode") === "External")throw Error("Missing/ambiguous image relationship");
        return new URL(edges[0].getAttribute("Target"),"https://fixture.invalid"+owner).pathname;
      };
      const noteParts = Object.keys(docs).filter(n => /^\/ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n));
      if(noteParts.length !== 15)throw Error("Notes page count changed");
      for(const note of noteParts){
        const images = image(docs[note]);
        if(images.length !== 1 || images[0].getElementsByTagNameNS(P,"ph")[0].getAttribute("idx") !== "2" || images[0].getElementsByTagNameNS(A,"prstGeom").length)throw Error("Local slide image inheritance changed");
        if(relation(note,"notesMaster") !== "/ppt/notesMasters/notesMaster1.xml" || relation(relation(note,"slide"),"notesSlide") !== note)throw Error("Wrong notes/slide correspondence");
      }
      const rel=[...docs["/ppt/notesMasters/_rels/notesMaster1.xml.rels"].getElementsByTagNameNS(R,"Relationship")].find(n=>n.getAttribute("Type").endsWith("/theme"));const theme=new URL(rel.getAttribute("Target"),"https://fixture.invalid/ppt/notesMasters/notesMaster1.xml").pathname;
      let count=0;for(const p of properties.filter(p=>p.target.part.includes("/notesSlides/"))){if(!select(p.target)||!select(p.source))throw Error("Missing notes provenance");for(const d of p.resolvedFrom??[])if(!select(d))throw Error("Missing notes dependency");if(p.status==="mapped"&&p.property.startsWith("run.scriptFontFamilies.")){const n=select(p.source);if(n.value!==p.value||p.source.part!==theme||n.ownerElement.getAttribute("script")!==p.property.split(".").at(-1))throw Error("Wrong actual notes font source");count++;}}
      return{theme,count,themeName:docs[theme].documentElement.getAttribute("name")};
    },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith(".xml")||n.endsWith(".rels")).map(([n,b])=>[n,b.toString()])),properties:expected[0].report.textStyleInheritance.properties});
    assert.equal(independent.theme,"/"+fixture.themePart);assert.ok(independent.count>0);assert.equal(independent.themeName,`PPTKit Notes ${variant}`);
    for(const format of ["esm","global"]){
      if(format==="global")for(const name of ["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
      const actual=await page.evaluate(async({bytes,options,core,parser})=>{
        const c=core?await import(core):PPTKit,p=parser?await import(parser):PPTKit,results=[];
        for(let round=0;round<3;round++){const s=await p.importPptx(Uint8Array.from(bytes),options);try{
          const document=c.normalizePresentation(s.document),report=s.report,noOp=await s.prepareSave({mode:"no-op"});if(noOp.status!=="ready")throw Error("No-op blocked");const saved=[...(await s.save(noOp.plan)).bytes];if(saved.length!==bytes.length||saved.some((b,i)=>b!==bytes[i]))throw Error("No-op differs");
          const old=structuredClone(s.document.slides[0].notes),wrong=structuredClone(old);wrong[0].runs[0].style.fontFamily="Forbidden";s.document.slides[0].setNotes(wrong);if((await s.prepareSave({mode:"edit"})).status!=="blocked")throw Error("Style writer admitted");s.document.slides[0].setNotes(old);
          for(const slide of s.document.slides){const notes=structuredClone(slide.notes);notes[0].runs[0].text+="X";slide.setNotes(notes);}const plan=await s.prepareSave({mode:"edit"});if(plan.status!=="ready")throw Error("Scalar blocked");bytes=[...(await s.save(plan.plan)).bytes];results.push({document,report,bytes});
        }finally{s.dispose();}}
        return results;
      },{bytes:[...fixture.bytes],options,core:format==="esm"?core:null,parser:format==="esm"?parser:null});assert.deepEqual(actual,expected);
    }
    assert.deepEqual(requests,[]);await page.close();
    // Realistic source negative: an equal idx cannot rescue the wrong placeholder type.
    const changed=new Map(parts),name="ppt/notesMasters/notesMaster1.xml";changed.set(name,Buffer.from(parts.get(name).toString().replace('type="body" idx="99"','type="title" idx="7"')));const wrong=await importPptx(zip([...changed].map(([name,body])=>({name,body}))).bytes,options);
    try{assert.equal(wrong.document.slides[0].notes[0].runs[0].style.fontSize,18);}finally{wrong.dispose();}
    for(const missing of ["theme","master"]){
      const broken=new Map(parts),part=missing==="theme"?"ppt/notesMasters/_rels/notesMaster1.xml.rels":"ppt/notesSlides/_rels/notesSlide1.xml.rels",kind=missing==="theme"?"theme":"notesMaster";
      const input=broken.get(part).toString(),pattern=new RegExp(`<Relationship\\b[^>]*Type="[^"]*/${kind}"[^>]*/>`);assert.match(input,pattern);broken.set(part,Buffer.from(input.replace(pattern,"")));
      const blocked=await importPptx(zip([...broken].map(([name,body])=>({name,body}))).bytes,options);
      try{
        const run=blocked.document.slides[0].notes[0].runs[0];assert.equal(run.style.fontFamily,undefined);assert.equal(run.style.scriptFontFamilies,undefined);
        const properties=blocked.report.textStyleInheritance.properties.filter(p=>p.target.part==="/ppt/notesSlides/notesSlide1.xml"&&p.property==="run.fontFamily");if(missing==="theme")assert.ok(properties.some(p=>p.status==="unsupported"));else assert.equal(properties.length,0);assert.ok(properties.every(p=>!(p.resolvedFrom??[]).some(d=>d.part==="/ppt/theme/theme1.xml")));
        if(missing==="master")assert.deepEqual(blocked.document.slides[1].notes[0].runs[0].style.scriptFontFamilies,fixture.map);
      }finally{blocked.dispose();}
    }

  }
});
