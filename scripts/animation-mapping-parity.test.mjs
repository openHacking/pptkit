import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, inspectPptxTiming } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { animationSpecimens, nativeTimingSpecimen } from "./research/native-timing-fixture.mjs";
import { chromium } from "@playwright/test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

async function source() {
 const d=createPresentation();
 for(const s of animationSpecimens)d.addSlide().addElement({type:"shape",shape:"rect",name:s.name,text:{content:s.name},box:{x:10,y:10,width:100,height:40}});
 const parts=readGeneratedZip((await generatePptx(d)).bytes);
 for(const [i,s]of animationSpecimens.entries()){const name=`ppt/slides/slide${i+1}.xml`;parts.set(name,Buffer.from(parts.get(name).toString().replace("</p:sld>",nativeTimingSpecimen(s.name,2)+"</p:sld>")));}
 return zip([...parts].map(([name,body])=>({name,body}))).bytes;
}

test("native-edit XML oracle rejects corruptions without private PowerPoint artifacts",async()=>{
 const original=await source(),parts=readGeneratedZip(original);
 for(const [index,specimen] of animationSpecimens.entries()){
  const part=`ppt/slides/slide${index+1}.xml`,xml=parts.get(part).toString();
  const text=`<a:t>${specimen.name}</a:t>`,duration='dur="500"';
  assert.ok(xml.includes(text)&&xml.includes(duration));
  parts.set(part,Buffer.from(xml.replace(text,`<a:t>${specimen.name} edited</a:t>`).replace(duration,'dur="2000"')));
 }
 const edited=zip([...parts].map(([name,body])=>({name,body}))).bytes;
 const dir=await mkdtemp(join(tmpdir(),"pptkit-animation-oracle-"));
 try{
  const before=join(dir,"source.pptx"),after=join(dir,"edited.pptx");
  await Promise.all([writeFile(before,original),writeFile(after,edited)]);
  const {stdout}=await promisify(execFile)("python3",["scripts/research/oracles/test-animation-powerpoint-edit.py",before,after]);
  assert.match(stdout,/optimized=False: valid accepted; 6 mutations rejected/);
  assert.match(stdout,/optimized=True: valid accepted; 6 mutations rejected/);
 }finally{await rm(dir,{recursive:true,force:true});}
});

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

test("opaque timing and inactive text-range targets survive unrelated edits but block their own target",async()=>{
 const parts=readGeneratedZip(await source()),part="ppt/slides/slide1.xml";
 const original=parts.get(part).toString().replace("<p:sld ",'<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future-timing" mc:Ignorable="x" ')
  .replace("</p:timing>",'<x:unknownBehavior x:token="preserve"><x:payload>opaque</x:payload></x:unknownBehavior><mc:AlternateContent><mc:Choice Requires="x"><p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></mc:Choice><mc:Fallback><x:marker/></mc:Fallback></mc:AlternateContent></p:timing>');
 parts.set(part,Buffer.from(original));const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes;
 const session=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});
 try{
  const first=session.document.slides[0],second=session.document.slides[1];
  const replace=(slide,text)=>{const element=structuredClone(slide.elements[0]);element.text.content[0].runs[0].text=text;slide.replaceElement(element.id,element);};
  const baseline=first.elements[0].text.content[0].runs[0].text;
  replace(second,"unrelated");replace(first,"unsafe");
  const blocked=await session.prepareSave({mode:"edit"});assert.equal(blocked.status,"blocked");
  assert.ok(blocked.report.diagnostics.some(d=>d.code==="save-text-range-reference"));
  replace(first,baseline);
  const retry=await session.prepareSave({mode:"edit"});assert.equal(retry.status,"ready");
  const output=readGeneratedZip((await session.save(retry.plan)).bytes);
  assert.equal(output.get(part).toString(),original);
  assert.ok(output.get("ppt/slides/slide2.xml").toString().includes("unrelated"));
  for(const [name,body]of parts)if(name!=="ppt/slides/slide2.xml")assert.deepEqual(output.get(name),body,name);
 }finally{session.dispose();}
});

test("combined timing targets distinguish a same-slide bystander from the guarded owner and a dangling reference",async()=>{
 const deck=createPresentation(),slide=deck.addSlide();
 for(const name of ["target","bystander"])slide.addElement({type:"shape",shape:"rect",name,text:{content:name},box:{x:10,y:10,width:100,height:40}});
 const parts=readGeneratedZip((await generatePptx(deck)).bytes),part="ppt/slides/slide1.xml";
 const extra='<p:cond evt="onClick"><p:tgtEl><p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></p:tgtEl></p:cond><p:cond evt="onClick"><p:tgtEl><p:spTgt spid="999"><p:txEl/></p:spTgt></p:tgtEl></p:cond><x:future x:mode="opaque"/>';
 const timing=nativeTimingSpecimen("fade-in",2).replace("</p:timing>",extra+"</p:timing>");
 const xml=parts.get(part).toString().replace("<p:sld ",'<p:sld xmlns:x="urn:future-timing" ').replace("</p:sld>",timing+"</p:sld>");
 parts.set(part,Buffer.from(xml));const bytes=zip([...parts].map(([name,body])=>({name,body}))).bytes;
 const inventory=await inspectPptxTiming(bytes),refs=inventory.slides[0].nodes.filter(n=>n.drawingReference).map(n=>[n.drawingReference.rawId,n.drawingReference.status]);
 assert.equal(refs.filter(([id,status])=>id==="2"&&status==="located").length,2,JSON.stringify(refs));
 assert.equal(refs.filter(([id,status])=>id==="999"&&status==="missing").length,1,JSON.stringify(refs));
 assert.ok(inventory.findings.some(f=>f.code==="unknown-timing-node"));
 const session=await importPptx(bytes,{projection:"content",editSaveProfile:"text-runs-v3"});
 try{
  const owner=session.document.slides[0],replace=(index,text)=>{const e=structuredClone(owner.elements[index]);e.text.content[0].runs[0].text=text;owner.replaceElement(e.id,e);};
  replace(0,"changed target");replace(1,"changed bystander");
  const blocked=await session.prepareSave({mode:"edit"});assert.equal(blocked.status,"blocked");
  assert.ok(blocked.report.diagnostics.some(d=>d.code==="save-text-range-reference"));
  replace(0,"target");const ready=await session.prepareSave({mode:"edit"});assert.equal(ready.status,"ready");
  const output=readGeneratedZip((await session.save(ready.plan)).bytes),saved=output.get(part).toString();
  assert.equal(saved.match(/<p:timing>[\s\S]*?<\/p:timing>/)[0],timing);
  assert.ok(xml.includes(">bystander</a:t>"));assert.equal(saved,xml.replace(">bystander</a:t>",">changed bystander</a:t>"));
  for(const [name,body]of parts)if(name!==part)assert.deepEqual(output.get(name),body,name);
 }finally{session.dispose();}
});
