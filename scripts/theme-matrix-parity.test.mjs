import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { matrixFixture } from "../packages/pptx-parser/test/theme-matrix-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const moduleUrl=s=>`data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
async function rounds(bytes){
  const results=[];for(let round=0;round<3;round++){
    const s=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});try{
      const document=normalizePresentation(s.document),report=s.report.tableStyleInheritance,slide=s.document.slides[0],old=slide.elements[0],changed=structuredClone(old);changed.rows[0].cells[0].content[0].runs[0].text=`Matrix ${round}`;slide.replaceElement(old.id,changed);const plan=await s.prepareSave({mode:"edit"});assert.equal(plan.status,"ready");const next=await s.save(plan.plan);results.push({document,report,saved:[...next.bytes]});bytes=next.bytes;
    }finally{s.dispose();}
  }return results;
}
test("Node/ESM/global matrix values, source dependencies and three source saves agree with independent DOM slot/color checks",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());const core=moduleUrl(await readFile("packages/core/dist/browser/index.js","utf8")),parser=moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js","utf8")).replaceAll('"@pptkit/core"',JSON.stringify(core)));
  for(const strict of[false,true]){
    const f=matrixFixture({strict}),bytes=encodeStructure(f),expected=await rounds(bytes);
    for(const format of["esm","global"]){
      const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort();});if(format==="global")for(const name of["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
      const result=await page.evaluate(async({core,parser,bytes,parts,namespace})=>{
        const c=core?await import(core):PPTKit,p=parser?await import(parser):PPTKit,results=[];
        for(let round=0;round<3;round++){
          const s=await p.importPptx(Uint8Array.from(bytes),{projection:"content",editSaveProfile:"text-runs-v3"}),document=c.normalizePresentation(s.document),report=s.report.tableStyleInheritance,slide=s.document.slides[0],old=slide.elements[0],changed=structuredClone(old);changed.rows[0].cells[0].content[0].runs[0].text=`Matrix ${round}`;slide.replaceElement(old.id,changed);const ready=await s.prepareSave({mode:"edit"});if(ready.status!=="ready")throw new Error("not ready");const saved=await s.save(ready.plan);results.push({document,report,saved:[...saved.bytes]});bytes=[...saved.bytes];s.dispose();
        }
        const docs=Object.fromEntries(Object.entries(parts).filter(([,text])=>typeof text==="string"&&text.startsWith("<")).map(([part,xml])=>[`/${part}`,new DOMParser().parseFromString(xml,"application/xml")])),theme=docs["/deck/theme.xml"],defs=docs["/deck/tableStyles.xml"];
        const palette=Object.fromEntries([...theme.getElementsByTagNameNS(namespace,"clrScheme")[0].children].map(n=>[n.localName,n.firstElementChild.getAttribute("val")]));let checked=0;
        for(const info of results[0].report.properties){
          if(!info.property.startsWith("cell.")||info.status!=="mapped")continue;
          const index=defs.evaluate(info.source.path,defs,prefix=>prefix==="a"?namespace:null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;
          if(index?.nodeType!==2||index.name!=="idx")throw new Error("not a source matrix index");
          const ref=index.ownerElement,kind=ref.localName==="fillRef"?"fillStyleLst":"lnStyleLst",slot=theme.getElementsByTagNameNS(namespace,kind)[0].children[Number(index.value)-1],ph=slot.getElementsByTagNameNS(namespace,"schemeClr")[0];
          if(ph?.getAttribute("val")!=="phClr")throw new Error("fixture slot must bind phClr");const token=ref.firstElementChild.getAttribute("val"),paint=info.property==="cell.fill"?info.value:info.value.paint;if(paint.color!==palette[token])throw new Error("independent placeholder color mismatch");
          if(info.property!=="cell.fill"&&info.value.width!==Number(slot.getAttribute("w"))/12700)throw new Error("independent line width mismatch");
          if(!info.resolvedFrom.some(d=>d.part==="/deck/theme.xml"&&d.path.includes(kind)))throw new Error("missing matrix source dependency");checked++;
        }
        return {results,checked};
      },{core:format==="esm"?core:null,parser:format==="esm"?parser:null,bytes:[...bytes],parts:f.parts,namespace:f.A});
      assert.deepEqual(result.results,expected);assert.equal(result.checked,20);assert.deepEqual(requests,[]);await page.close();
    }
    const a=readGeneratedZip(bytes),b=readGeneratedZip(Uint8Array.from(expected.at(-1).saved));for(const[name,value]of a)if(name!=="deck/z.xml")assert.deepEqual(b.get(name),value);
  }
});

test("browser matrix bounds/unsupported slots share Node errors, captured budgets and cancellation retry",async t=>{
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of["core","pptx-parser"])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,"utf8")});
  const bytes=encodeStructure(matrixFixture());const result=await page.evaluate(async data=>{
    const bytes=Uint8Array.from(data),options={projection:"content",sessionLimits:{maxTableStyleWork:2000000}},pending=PPTKit.importPptx(bytes,options);bytes.fill(0);options.sessionLimits.maxTableStyleWork=1;const captured=await pending,value=captured.report.tableStyleInheritance.properties.find(p=>p.property==="cell.fill").value.color;captured.dispose();
    let limited,aborted;try{await PPTKit.importPptx(Uint8Array.from(data),{projection:"content",sessionLimits:{maxTableStyleDependencies:1}});}catch(e){limited=e.code;}
    const c=new AbortController(),operation=PPTKit.importPptx(Uint8Array.from(data),{projection:"content",signal:c.signal});queueMicrotask(()=>c.abort());try{await operation;}catch(e){aborted=e.code;}
    const retry=await PPTKit.importPptx(Uint8Array.from(data),{projection:"content"});retry.dispose();return {value,limited,aborted};
  },[...bytes]);assert.deepEqual(result,{value:"3579BD",limited:"import-limit-exceeded",aborted:"operation-aborted"});
});
