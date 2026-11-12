import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { chromium } from "@playwright/test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";

const path=process.argv[2];if(!path || process.argv.length!==3)throw Error("Expected the named PowerPoint-edited PPTX path.");
const bytes=await readFile(path),session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v1"});
const master=session.document.masters.find(m=>m.elements.some(e=>e.name==="logo"));assert.ok(master);
assert.deepEqual(master.elements.map(e=>e.name),["blue","logo"]);
assert.equal(master.elements[0].style.fill.color,"0000FF");assert.equal(master.elements[1].crop.left,.2087);
const result=await renderPresentationToSvg(session.document),browser=await chromium.launch();
try {
 const page=await browser.newPage();await page.route("**/*",route=>route.abort());
 const samples=await page.evaluate(async({slides,width,height})=>{
  const sample=[];for(const svg of slides){const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:"image/svg+xml"}));await new Promise((resolve,reject)=>{img.onload=resolve;img.onerror=reject;img.src=url});const canvas=document.createElement("canvas");canvas.width=width;canvas.height=height;const c=canvas.getContext("2d");c.drawImage(img,0,0);sample.push([125,180,260].map(x=>[...c.getImageData(x,220,1,1).data]));URL.revokeObjectURL(url)}return sample;
 },{slides:result.slides.map(s=>s.svg),width:result.width,height:result.height});
 const colors=samples.map(row=>row.map(([r,g,b,a])=>a===255 && r>240 && g<15 && b<15?"red":a===255 && b>240 && r<15 && g<15?"blue":a===255 && g>240 && r<15 && b<15?"green":a===255 && r>240 && g>240 && b>240?"white":"other"));
 assert.deepEqual(colors,[["blue","red","blue"],["blue","red","blue"],["white","white","white"],["green","green","green"]]);
 console.log(JSON.stringify({sourceSha256:createHash("sha256").update(bytes).digest("hex"),masterId:master.id,slideSize:[result.width,result.height],coordinates:[[125,220],[180,220],[260,220]],colors,warnings:result.warnings.length},null,2));
} finally {await browser.close();session.dispose()}
