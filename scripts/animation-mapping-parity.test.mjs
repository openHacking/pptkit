import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { animationSpecimens, nativeTimingSpecimen } from "./research/native-timing-fixture.mjs";
import { chromium } from "@playwright/test";

async function source() {
 const d=createPresentation();
 for(const s of animationSpecimens)d.addSlide().addElement({type:"shape",shape:"rect",name:s.name,text:{content:s.name},box:{x:10,y:10,width:100,height:40}});
 const parts=readGeneratedZip((await generatePptx(d)).bytes);
 for(const [i,s]of animationSpecimens.entries()){const name=`ppt/slides/slide${i+1}.xml`;parts.set(name,Buffer.from(parts.get(name).toString().replace("</p:sld>",nativeTimingSpecimen(s.name,2)+"</p:sld>")));}
 return zip([...parts].map(([name,body])=>({name,body}))).bytes;
}

test("independent DOM validates behavior category, duration, unique time IDs and resolved target names",async t=>{
 const parts=readGeneratedZip(await source()),browser=await chromium.launch({headless:true});t.after(()=>browser.close());const page=await browser.newPage();
 const xml=animationSpecimens.map((_,i)=>parts.get(`ppt/slides/slide${i+1}.xml`).toString());
 const rows=await page.evaluate(xml=>xml.map(text=>{
  const d=new DOMParser().parseFromString(text,"application/xml"),ns="http://schemas.openxmlformats.org/presentationml/2006/main";
  if(d.querySelector("parsererror"))throw new Error("Bad XML");
  const nodes=[...d.getElementsByTagNameNS(ns,"cTn")],ids=nodes.map(n=>n.getAttribute("id"));if(new Set(ids).size!==ids.length)throw new Error("Duplicate timing ID");
  const b=[...d.getElementsByTagNameNS(ns,"cBhvr")][0],target=b.getElementsByTagNameNS(ns,"spTgt")[0].getAttribute("spid");
  const drawing=[...d.getElementsByTagNameNS(ns,"cNvPr")].find(n=>n.getAttribute("id")===target);
  return {tag:b.parentElement.localName,attributes:Object.fromEntries([...b.parentElement.attributes].map(a=>[a.name,a.value])),targetName:drawing?.getAttribute("name"),duration:b.getElementsByTagNameNS(ns,"cTn")[0].getAttribute("dur"),root:nodes[0].getAttribute("nodeType")};
 }),xml);
 assert.deepEqual(rows,animationSpecimens.map(s=>({tag:s.tag,attributes:s.attributes,targetName:s.name,duration:"500",root:"tmRoot"})));
});

test("source-only timing survives three unrelated shape-text saves with exact timing markup and targets",async()=>{
 let bytes=new Uint8Array(await source());
 for(let round=0;round<3;round++){
  const before=readGeneratedZip(bytes),session=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});
  try{
   const noOp=await session.prepareSave({mode:"no-op"});assert.equal(noOp.status,"ready");assert.deepEqual((await session.save(noOp.plan)).bytes,bytes);
   for(const slide of session.document.slides){const element=structuredClone(slide.elements[0]);element.text.content[0].runs[0].text=`Edit ${round}`;slide.removeElement(element.id);slide.insertElement(0,element);}
   const p=await session.prepareSave({mode:"edit"});assert.equal(p.status,"ready");bytes=(await session.save(p.plan)).bytes;const after=readGeneratedZip(bytes);
   const timing=x=>x.toString().match(/<p:timing>[\s\S]*?<\/p:timing>/)[0];
   for(let i=1;i<=4;i++){const name=`ppt/slides/slide${i}.xml`;assert.equal(timing(after.get(name)),timing(before.get(name)));}
   for(const[name,body]of before)if(!/^ppt\/slides\/slide\d+\.xml$/.test(name))assert.deepEqual(after.get(name),body,name);
  }finally{session.dispose();}
 }
});
