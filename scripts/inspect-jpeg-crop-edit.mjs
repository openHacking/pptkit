import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { chromium } from "@playwright/test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
const dir=process.argv[2]; if (!dir || process.argv.length!==3) throw Error("Expected evidence directory.");
const input=await readFile(join(dir,"crop-native-edit.pptx"));const session=await importPptx(input,{projection:"content"});
try {
  const images=session.document.slides[0].elements.filter(element=>element.type==="image");
  const mapped=images.find(image=>image.name==="mapped"),control=images.find(image=>image.name==="control");
  assert.equal(images.length,2);assert.ok(mapped&&control);assert.equal(mapped.fit,"crop");
  assert.deepEqual(mapped.crop,{left:.25326,top:0,right:0,bottom:.33333});
  assert.ok(Math.abs(mapped.box.x-180.78251968503938)<1e-8 && Math.abs(mapped.box.width-179.21740157480315)<1e-8);
  assert.equal(mapped.box.y,130);assert.equal(mapped.box.height,240);assert.deepEqual(control.box,{x:600,y:130,width:240,height:240});
  for(const [image,name,width,height] of [[mapped,"browser-normalized.png",200,300],[control,"crop-control.png",200,200]]){
    const asset=session.document.assets.find(value=>value.id===image.assetId);assert.ok(asset);assert.equal(asset.width,width);assert.equal(asset.height,height);
    assert.deepEqual(Buffer.from(asset.source.value,"base64"),await readFile(join(dir,name)));
  }
  const prepared=await session.prepareSave({mode:"no-op"});assert.equal(prepared.status,"ready");const saved=await session.save(prepared.plan);assert.deepEqual(Buffer.from(saved.bytes),input);
  const preview=await renderPresentationToSvg(session.document);assert.equal(preview.slides.length,1);
  const svg=preview.slides[0].svg; await writeFile(join(dir,"crop-native-edit.svg"),svg);
  const browser=await chromium.launch({headless:true});let screenshot,requests=[];
  try {const page=await browser.newPage({viewport:{width:1920,height:1080}});await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
    await page.setContent(`<style>html,body{margin:0} svg{width:1920px;height:1080px;display:block}</style>${svg}`);
    await page.evaluate(async()=>{await Promise.all([...document.images].map(image=>image.decode().catch(()=>{})))});
    screenshot=await page.screenshot();await writeFile(join(dir,"crop-native-edit-preview.png"),screenshot);
  } finally {await browser.close()}
  assert.deepEqual(requests,[]);const hash=data=>createHash("sha256").update(data).digest("hex");
  const report={inputSha256:hash(input),svgSha256:hash(svg),previewSha256:hash(screenshot),mappedCrop:mapped.crop,mappedBox:mapped.box,controlBox:control.box,noOpByteIdentical:true,requests,warnings:preview.warnings};
  await writeFile(join(dir,"crop-native-edit-report.json"),JSON.stringify(report,null,2)+"\n");console.log(JSON.stringify(report));
} finally {session.dispose()}
