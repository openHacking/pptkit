import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync, strFromU8 } from "fflate";
import { createPresentation, planImageReplacement } from "@pptkit/core";
import { resolveLayout, planVisibleImageReplacement, planSharedVisibleImageReplacement } from "@pptkit/layout";
import { generatePptx } from "../dist/index.js";

test("replacement cover recomputes layout from new dimensions and preserves the other image",async()=>{
  const png=await readFile(new URL("../../../tests/fixtures/exif-orientation6-normalized.png",import.meta.url));
  const jpeg=await readFile(new URL("../../../tests/fixtures/exif-orientation6.jpg",import.meta.url));
  const exif=jpeg.indexOf(Buffer.from("Exif\0\0"));assert.ok(exif>0);jpeg[exif+6+19]=1;
  const deck=createPresentation();
  const old=deck.registerAsset({id:"old",kind:"image",source:{type:"bytes",value:png},mimeType:"image/png",width:2,height:3});
  const next=deck.registerAsset({id:"next",kind:"image",source:{type:"bytes",value:jpeg},mimeType:"image/jpeg",width:3,height:2});
  const slide=deck.addSlide({elements:[{type:"image",id:"one",assetId:old.id,fit:"crop",crop:{left:.1,right:.2},box:{x:0,y:0,width:100,height:100}},
    {type:"image",id:"two",assetId:old.id,fit:"stretch",box:{x:120,y:0,width:100,height:100}}]});
  const plan=planImageReplacement(slide.elements[0],next,{fit:"cover"});slide.replaceElement("one",plan.replacement);
  const [first,second]=resolveLayout(deck).slides[0].elements;
  assert.ok(Math.abs(first.crop.left-1/6)<1e-12 && Math.abs(first.crop.right-1/6)<1e-12);
  assert.equal(first.crop.top,0);assert.equal(first.crop.bottom,0);assert.equal(second.assetId,"old");
  const zip=unzipSync((await generatePptx(deck)).bytes);const xml=strFromU8(zip["ppt/slides/slide1.xml"]);
  assert.match(xml,/<a:srcRect l="16667" t="0" r="16667" b="0"\/>/);
  assert.ok(Object.keys(zip).some(name=>name==="ppt/media/old.png"));assert.ok(Object.keys(zip).some(name=>name==="ppt/media/next.jpg"));
  slide.replaceElement("one",plan.previous);assert.deepEqual(slide.elements[0].crop,{left:.1,right:.2});
});


test("visible-frame replacement exports the previously resolved contain frame",async()=>{
  const oldBytes=await readFile(new URL("../../../tests/fixtures/exif-orientation6-normalized.png",import.meta.url));
  const newBytes=await readFile(new URL("../../../tests/fixtures/transparent-replacement.png",import.meta.url));
  const deck=createPresentation();
  deck.registerAsset({id:"old",kind:"image",source:{type:"bytes",value:oldBytes},mimeType:"image/png",width:2,height:3});
  deck.registerAsset({id:"next",kind:"image",source:{type:"bytes",value:newBytes},mimeType:"image/png",width:4,height:2});
  const slide=deck.addSlide({id:"slide",elements:[{type:"image",id:"pic",assetId:"old",fit:"contain",box:{x:0,y:0,width:120,height:120}}]});
  const oldFrame=resolveLayout(deck).slides[0].elements[0].box;
  assert.deepEqual(oldFrame,{x:20,y:0,width:80,height:120});
  const plan=planVisibleImageReplacement(deck,"slide","pic","next","stretch");slide.replaceElement("pic",plan.replacement);
  assert.deepEqual(resolveLayout(deck).slides[0].elements[0].box,oldFrame);
  const zip=unzipSync((await generatePptx(deck)).bytes),xml=strFromU8(zip["ppt/slides/slide1.xml"]);
  assert.match(xml,/<a:off x="254000" y="0"\/><a:ext cx="1016000" cy="1524000"\/>/);
  assert.deepEqual(zip["ppt/media/next.png"],new Uint8Array(newBytes));
  slide.replaceElement("pic",plan.previous);assert.deepEqual(resolveLayout(deck).slides[0].elements[0].box,oldFrame);
});


test("composed visible-frame replacement exports the resolved snapshot without implicit recomposition",async()=>{
  const oldBytes=await readFile(new URL("../../../tests/fixtures/exif-orientation6-normalized.png",import.meta.url));
  const newBytes=await readFile(new URL("../../../tests/fixtures/transparent-replacement.png",import.meta.url));
  const deck=createPresentation();
  deck.registerAsset({id:"old",kind:"image",source:{type:"bytes",value:oldBytes},mimeType:"image/png",width:2,height:3});
  deck.registerAsset({id:"next",kind:"image",source:{type:"bytes",value:newBytes},mimeType:"image/png",width:4,height:2});
  const slide=deck.addSlide({id:"slide",elements:[{type:"image",id:"pic",assetId:"old",fit:"contain",box:{x:0,y:0,width:120,height:120}}]});
  const composition=[{type:"align",container:{kind:"slide",id:"slide"},elementIds:["pic"],area:{x:100,y:100,width:200,height:200},horizontal:"end",vertical:"end"}];
  const plan=planVisibleImageReplacement(deck,"slide","pic","next","stretch",{composition});
  assert.deepEqual(plan.replacement.box,{x:200,y:180,width:80,height:120});slide.replaceElement("pic",plan.replacement);
  const zip=unzipSync((await generatePptx(deck)).bytes),xml=strFromU8(zip["ppt/slides/slide1.xml"]);
  assert.match(xml,/<a:off x="2540000" y="2286000"\/><a:ext cx="1016000" cy="1524000"\/>/);
  slide.replaceElement("pic",plan.previous);assert.deepEqual(slide.elements[0].box,{x:0,y:0,width:120,height:120});
});


test("shared master replacement writes one shared image without duplicating slide artwork",async()=>{
 const old=await readFile(new URL("../../../tests/fixtures/exif-orientation6-normalized.png",import.meta.url));
 const next=await readFile(new URL("../../../tests/fixtures/transparent-replacement.png",import.meta.url));
 const deck=createPresentation();deck.registerAsset({id:"old",kind:"image",source:{type:"bytes",value:old},mimeType:"image/png",width:2,height:3});
 deck.registerAsset({id:"new",kind:"image",source:{type:"bytes",value:next},mimeType:"image/png",width:4,height:2});
 deck.defineSlideMaster({id:"master",name:"Shared",elements:[{id:"art",type:"image",assetId:"old",fit:"contain",box:{x:0,y:0,width:120,height:120}}]});
 deck.defineSlideLayout({id:"layout",name:"Layout",masterId:"master"});deck.addSlide({id:"first",layoutId:"layout"});deck.addSlide({id:"second",layoutId:"layout"});
 const plan=planSharedVisibleImageReplacement(deck,{kind:"master",id:"master"},"art","new","stretch");assert.equal(plan.kind,"master");deck.replaceSlideMaster(plan.ownerId,plan.replacement);
 const archive=unzipSync((await generatePptx(deck)).bytes);
 const masterFiles=Object.keys(archive).filter(n=>/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n));
 const artwork=masterFiles.map(name=>({name,xml:strFromU8(archive[name])})).filter(item=>item.xml.includes('name="art"'));assert.equal(artwork.length,1);
 assert.match(artwork[0].xml,/<a:off x="254000" y="0"\/><a:ext cx="1016000" cy="1524000"\/>/);
 const embed=artwork[0].xml.match(/<a:blip r:embed="([^"]+)"/);assert.ok(embed);
 const rels=strFromU8(archive[`ppt/slideMasters/_rels/${artwork[0].name.split("/").at(-1)}.rels`]);
 assert.ok(rels.includes(`Id="${embed[1]}"`) && rels.includes('Target="../media/new.png"'));
 assert.match(rels,new RegExp(`<Relationship Id="${embed[1]}"[^>]*Target="[.][.]/media/new[.]png"`));
 assert.deepEqual(archive["ppt/media/new.png"],new Uint8Array(next));
 const layouts=[];
 for(const name of ["ppt/slides/slide1.xml","ppt/slides/slide2.xml"]) {
   assert.doesNotMatch(strFromU8(archive[name]),/<p:pic>/);
   const slideRels=strFromU8(archive[`ppt/slides/_rels/${name.split("/").at(-1)}.rels`]);
   const layoutTarget=slideRels.match(/Type="[^"]*\/slideLayout" Target="([^"]+)"/);assert.ok(layoutTarget);
   layouts.push(path.posix.normalize(path.posix.join("ppt/slides",layoutTarget[1])));
 }
 assert.equal(layouts[0],layouts[1]);
 const layoutRels=strFromU8(archive[`ppt/slideLayouts/_rels/${layouts[0].split("/").at(-1)}.rels`]);
 const masterTarget=layoutRels.match(/Type="[^"]*\/slideMaster" Target="([^"]+)"/);assert.ok(masterTarget);
 assert.equal(path.posix.normalize(path.posix.join("ppt/slideLayouts",masterTarget[1])),artwork[0].name);
 deck.replaceSlideMaster(plan.ownerId,plan.previous);assert.equal(deck.masters.find(m=>m.id==="master").elements[0].assetId,"old");
});
