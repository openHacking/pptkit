import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, planImageReplacement } from "../packages/core/dist/index.js";
import { resolveLayout, planNestedVisibleImageReplacement } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";

async function pixelSamples(page,svg) {
  return page.evaluate(async svg => {
    const image = new Image(), url=URL.createObjectURL(new Blob([svg],{type:"image/svg+xml"}));
    try {image.src=url;await image.decode();const canvas=document.createElement("canvas");canvas.width=500;canvas.height=250;
      const ctx=canvas.getContext("2d");ctx.fillStyle="white";ctx.fillRect(0,0,500,250);ctx.drawImage(image,0,0);
      return [[50,10],[50,80],[150,80],[300,60],[300,160]].map(([x,y])=>Array.from(ctx.getImageData(x,y,1,1).data));
    } finally {URL.revokeObjectURL(url)}
  },svg);
}

test("transparent per-image replacement spans layout, preview, export, import and safe undo",async t=>{
  const old=await readFile("tests/fixtures/exif-orientation6-normalized.png");
  const transparent=await readFile("tests/fixtures/transparent-replacement.png");
  const doc=createPresentation({size:{width:500,height:250}});
  const shared=doc.registerAsset({id:"shared",kind:"image",source:{type:"bytes",value:old},mimeType:"image/png",width:2,height:3});
  const replacement=doc.registerAsset({id:"transparent",kind:"image",source:{type:"bytes",value:transparent},mimeType:"image/png",width:4,height:2});
  const slide=doc.addSlide({elements:[{type:"image",id:"target",assetId:shared.id,fit:"stretch",box:{x:0,y:0,width:200,height:200}},
    {type:"image",id:"other",assetId:shared.id,fit:"stretch",box:{x:250,y:0,width:200,height:200}}]});
  const prior=await renderPresentationToSvg(doc); const plan=planImageReplacement(slide.elements[0],replacement,{fit:"contain"});
  slide.replaceElement("target",plan.replacement);
  const layout=resolveLayout(doc);assert.deepEqual(layout.slides[0].elements[0].box,{x:0,y:50,width:200,height:100});
  assert.equal(slide.elements[1].assetId,"shared");const preview=await renderPresentationToSvg(doc);assert.deepEqual(preview.warnings,[]);
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();const requests=[];
  await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
  const alpha=await page.evaluate(async bytes=>{
    const image=await createImageBitmap(new Blob([Uint8Array.from(bytes)],{type:"image/png"}));
    const canvas=document.createElement("canvas");canvas.width=4;canvas.height=2;const ctx=canvas.getContext("2d");ctx.drawImage(image,0,0);image.close();
    return [Array.from(ctx.getImageData(0,0,1,1).data),Array.from(ctx.getImageData(3,0,1,1).data)];
  },[...transparent]);
  assert.deepEqual(alpha,[[255,0,0,255],[0,0,0,0]]);
  const beforePixels=await pixelSamples(page,prior.slides[0].svg),afterPixels=await pixelSamples(page,preview.slides[0].svg);
  assert.deepEqual(afterPixels[0],[255,255,255,255]);assert.deepEqual(afterPixels[1],[255,0,0,255]);assert.deepEqual(afterPixels[2],[255,255,255,255]);
  assert.deepEqual(afterPixels.slice(3),beforePixels.slice(3));assert.deepEqual(requests,[]);
  const exported=await generatePptx(doc);const session=await importPptx(exported.bytes,{projection:"content",editSaveProfile:"text-runs-v1"});
  try {
    const images=session.document.slides[0].elements.filter(e=>e.type==="image");assert.equal(images.length,2);
    const imported=images.find(e=>e.name==="target"), other=images.find(e=>e.name==="other");assert.ok(imported&&other);
    const importedAsset=session.document.assets.find(a=>a.id===imported.assetId),otherAsset=session.document.assets.find(a=>a.id===other.assetId);
    assert.deepEqual(Buffer.from(importedAsset.source.value,"base64"),transparent);assert.deepEqual(Buffer.from(otherAsset.source.value,"base64"),old);
    assert.deepEqual(imported.box,{x:0,y:50,width:200,height:100});assert.equal(imported.fit,"stretch");
    const targetAsset=session.document.registerAsset({id:"new-target",kind:"image",source:{type:"bytes",value:old},mimeType:"image/png",width:2,height:3});
    const importedSlide=session.document.slides[0], edit=planImageReplacement(imported,targetAsset,{fit:"cover"});
    importedSlide.replaceElement(imported.id,edit.replacement);
    assert.equal((await session.prepareSave({mode:"edit"})).status,"blocked");assert.equal((await session.prepareSave({mode:"no-op"})).status,"blocked");
    importedSlide.replaceElement(imported.id,edit.previous);session.document.removeAsset(targetAsset.id);
    const restored=await session.prepareSave({mode:"no-op"});assert.equal(restored.status,"ready");assert.deepEqual((await session.save(restored.plan)).bytes,exported.bytes);
  } finally {session.dispose()}
  slide.replaceElement("target",plan.previous);const undone=await renderPresentationToSvg(doc);
  assert.deepEqual(await pixelSamples(page,undone.slides[0].svg),beforePixels);
});


test("nested replacement survives fresh PPTX export and import while retaining group ownership",async()=>{
  const old=await readFile("tests/fixtures/exif-orientation6-normalized.png"), replacement=await readFile("tests/fixtures/transparent-replacement.png");
  const doc=createPresentation();doc.registerAsset({id:"old",kind:"image",source:{type:"bytes",value:old},mimeType:"image/png",width:2,height:3});
  doc.registerAsset({id:"new",kind:"image",source:{type:"bytes",value:replacement},mimeType:"image/png",width:4,height:2});
  const slide=doc.addSlide({id:"slide",elements:[{type:"group",id:"group",box:{x:100,y:50,width:240,height:240},coordinateSize:{width:120,height:120},children:[
    {type:"image",id:"nested",assetId:"old",fit:"contain",box:{x:0,y:0,width:120,height:120}},
    {type:"image",id:"other",assetId:"old",fit:"stretch",box:{x:0,y:0,width:20,height:20}}]}]});
  const plan=planNestedVisibleImageReplacement(doc,"slide","nested","new","stretch");
  slide.replaceElement(plan.rootElementId,plan.replacement);assert.equal(slide.elements[0].children[1].assetId,"old");
  assert.deepEqual(resolveLayout(doc).slides[0].elements[0].children[0].box,{x:20,y:0,width:80,height:120});
  const output=await generatePptx(doc);const session=await importPptx(output.bytes,{projection:"content"});
  try {
    const group=session.document.slides[0].elements.find(e=>e.type==="group");assert.ok(group);
    const image=group.children.find(e=>e.name==="nested"),other=group.children.find(e=>e.name==="other");assert.ok(image&&other);
    assert.deepEqual(image.box,{x:20,y:0,width:80,height:120});
    const a=session.document.assets.find(x=>x.id===image.assetId),b=session.document.assets.find(x=>x.id===other.assetId);
    assert.deepEqual(Buffer.from(a.source.value,"base64"),replacement);assert.deepEqual(Buffer.from(b.source.value,"base64"),old);
  } finally {session.dispose()}
  slide.replaceElement(plan.rootElementId,plan.previous);assert.equal(slide.elements[0].children[0].assetId,"old");
});

test("composed ancestor geometry and replacement survive preview and fresh PPTX import",async t=>{
  const old=await readFile('tests/fixtures/exif-orientation6-normalized.png'),replacement=await readFile('tests/fixtures/transparent-replacement.png');
  const doc=createPresentation({size:{width:500,height:250}});doc.registerAsset({id:'old',kind:'image',source:{type:'bytes',value:old},mimeType:'image/png',width:2,height:3});doc.registerAsset({id:'new',kind:'image',source:{type:'bytes',value:replacement},mimeType:'image/png',width:4,height:2});
  const slide=doc.addSlide({id:'slide',elements:[{type:'group',id:'group',box:{x:20,y:20,width:160,height:160},coordinateSize:{width:100,height:100},transform:{rotation:30,flipH:true},children:[{type:'image',id:'nested',assetId:'old',fit:'contain',box:{x:0,y:0,width:80,height:80}},{type:'image',id:'sibling',assetId:'old',box:{x:80,y:0,width:20,height:20}}]}]});
  const rules=[{type:'constraints',container:{kind:'slide',id:'slide'},elements:[{elementId:'group',x:{value:180},width:{value:200}}]}];
  const expected=resolveLayout(doc,{composition:rules}).slides[0].elements[0],before=slide.elements[0];
  const plan=planNestedVisibleImageReplacement(doc,'slide','nested','new','stretch',{composition:rules,commitAncestorGeometry:true});slide.replaceElement(plan.rootElementId,plan.replacement);
  const actual=resolveLayout(doc).slides[0].elements[0];assert.deepEqual(actual.box,expected.box);assert.deepEqual(actual.children[0].box,expected.children[0].box);
  const svg=await renderPresentationToSvg(doc);assert.ok(svg.slides[0].svg.includes('rotate(30'));
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
  const painted=await page.evaluate(async svg=>{const image=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));try{image.src=url;await image.decode();const c=document.createElement('canvas');c.width=500;c.height=250;const ctx=c.getContext('2d');ctx.fillStyle='white';ctx.fillRect(0,0,500,250);ctx.drawImage(image,0,0);const data=ctx.getImageData(0,0,500,250).data;let left=0,moved=0;for(let y=0;y<250;y++)for(let x=0;x<500;x++){const i=(y*500+x)*4;if(data[i]<245||data[i+1]<245||data[i+2]<245){if(x<50)left++;if(x>=160)moved++;}}return {left,moved}}finally{URL.revokeObjectURL(url)}},svg.slides[0].svg);
  assert.equal(painted.left,0);assert.ok(painted.moved>100);
  const output=await generatePptx(doc),session=await importPptx(output.bytes,{projection:'content'});
  try{const group=session.document.slides[0].elements.find(e=>e.type==='group');assert.deepEqual(group.box,actual.box);assert.deepEqual(group.transform,actual.transform);const image=group.children.find(e=>e.name==='nested');for(const key of ['x','y','width','height'])assert.ok(Math.abs(image.box[key]-actual.children[0].box[key])<=1/12700);assert.deepEqual(Buffer.from(session.document.assets.find(a=>a.id===image.assetId).source.value,'base64'),replacement);}finally{session.dispose()}
  slide.replaceElement(plan.rootElementId,plan.previous);assert.deepEqual(slide.elements[0],before);
});

test("ancestor composition commit plans agree between Chromium and Node",async t=>{
 const create=core=>{const d=core.createPresentation();d.registerAsset({id:'old',kind:'image',source:{type:'url',value:'old'},width:4,height:2});d.registerAsset({id:'new',kind:'image',source:{type:'url',value:'new'},width:2,height:4});d.addSlide({id:'s',elements:[{type:'group',id:'g',box:{x:10,y:20,width:100,height:100},coordinateSize:{width:100,height:100},children:[{type:'image',id:'i',assetId:'old',fit:'contain',box:{x:0,y:0,width:100,height:100}}]}]});return d};
 const rules=[{type:'constraints',container:{kind:'slide',id:'s'},elements:[{elementId:'g',x:{value:200}}]}],expected=planNestedVisibleImageReplacement(create({createPresentation}),'s','i','new','stretch',{composition:rules,commitAncestorGeometry:true});
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of ['core','layout'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(rules=>{const d=PPTKit.createPresentation();d.registerAsset({id:'old',kind:'image',source:{type:'url',value:'old'},width:4,height:2});d.registerAsset({id:'new',kind:'image',source:{type:'url',value:'new'},width:2,height:4});d.addSlide({id:'s',elements:[{type:'group',id:'g',box:{x:10,y:20,width:100,height:100},coordinateSize:{width:100,height:100},children:[{type:'image',id:'i',assetId:'old',fit:'contain',box:{x:0,y:0,width:100,height:100}}]}]});return PPTKit.planNestedVisibleImageReplacement(d,'s','i','new','stretch',{composition:rules,commitAncestorGeometry:true})},rules);
 assert.deepEqual(actual,expected);
});
