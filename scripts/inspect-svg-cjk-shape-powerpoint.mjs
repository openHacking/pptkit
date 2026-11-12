import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
const sha = b => createHash("sha256").update(b).digest("hex");
execFileSync("python3", ["scripts/research/verify-cjk-shape-tabs-powerpoint.py"], { stdio: "pipe" });
const sourcePath="docs/testing/tab-cjk-shape-powerpoint-evidence.json", sourceBytes=await readFile(sourcePath), source=JSON.parse(sourceBytes);
const inputs=source.pages.flatMap(p=>p.rows.map(r=>({align:p.kind,word:r.word,bold:r.bold,characters:r.characters})));
const paths=["core","layout","svg-renderer"].map(n=>[`@pptkit/${n}`,`packages/${n}/dist/browser/index.js`]);
const bundles=await Promise.all(paths.map(async ([name,file])=>{const bytes=await readFile(file);return {name,file,sha256:sha(bytes),url:`data:text/javascript;base64,${bytes.toString("base64")}`};}));
const fontPath="/System/Library/Fonts/Hiragino Sans GB.ttc", fontSha256=sha(await readFile(fontPath));
const browser=await chromium.launch();let observations,fonts,version;
try {
 version=browser.version();const page=await browser.newPage();await page.route("**/*",r=>r.abort());await page.setContent(`<script type="importmap">${JSON.stringify({imports:Object.fromEntries(bundles.map(b=>[b.name,b.url]))})}</script>`);
 observations=await page.evaluate(async inputs=>{
  const {createPresentation}=await import("@pptkit/core"),{renderPresentationToSvg}=await import("@pptkit/svg-renderer");
  const all=[];
  for(const [index,item] of inputs.entries()) {
   const deck=createPresentation({size:{width:960,height:540}});deck.addSlide().addElement({id:`shape-${index}`,type:"shape",shape:"rect",box:{x:60,y:120,width:700,height:65},style:{fill:{type:"solid",color:"F2F5FA"},stroke:{paint:{type:"none"}}},text:{frame:{margin:0,wrap:false},content:[{style:{tabStops:[{position:280,align:item.align}]},runs:[{text:"合计\t",style:{fontFamily:"Hiragino Sans GB",fontSize:28}},{text:item.word,style:{fontFamily:"Hiragino Sans GB",fontSize:28,bold:item.bold}}]}]}});
   const rendered=await renderPresentationToSvg(deck),holder=document.createElement('div');holder.innerHTML=rendered.slides[0].svg;document.body.append(holder);
   const parts=[...holder.querySelectorAll('text[data-pptkit-tab-part]')],suffix=parts.filter(n=>n.dataset.pptkitTabPart!=='prefix'),chars=[];
   for(const el of suffix)for(let i=0;i<el.getNumberOfChars();i++){const a=el.getStartPositionOfChar(i),b=el.getEndPositionOfChar(i);chars.push({char:el.textContent[i],left:a.x-60,right:b.x-60,baseline:a.y-120});}
   all.push({chars,text:parts.map(n=>n.textContent).join(''),fill:holder.querySelector(`[data-pptkit-element-id="shape-${index}"] rect`)?.getAttribute('fill'),warnings:rendered.warnings.map(w=>w.code)});
  }
  return all;
 },inputs);
 const cdp=await page.context().newCDPSession(page);await cdp.send('DOM.enable');await cdp.send('CSS.enable');const {root}=await cdp.send('DOM.getDocument');fonts=[];
 for(const index of [0,2]) {const {nodeId}=await cdp.send('DOM.querySelector',{nodeId:root.nodeId,selector:`[data-pptkit-element-id="shape-${index}"] text[data-pptkit-tab-part="suffix"]`});fonts.push((await cdp.send('CSS.getPlatformFontsForNode',{nodeId})).fonts);}
} finally {await browser.close();}
assert.deepEqual(fonts.map(row=>row.map(f=>f.postScriptName)),[["HiraginoSansGB-W3"],["HiraginoSansGB-W6"]]);
const items=inputs.map((input,i)=>{const observed=observations[i];assert.equal(observed.text,`合计\t${input.word}`);assert.equal(observed.fill,'#F2F5FA');assert.ok(observed.warnings.includes('text-tab-layout-unsupported'));assert.deepEqual(observed.chars.map(c=>c.char),input.characters.map(c=>c.char));return {align:input.align,word:input.word,bold:input.bold,chars:input.characters.map((native,j)=>({char:native.char,native,svg:observed.chars[j],leftDifference:observed.chars[j].left-native.left,rightDifference:observed.chars[j].right-native.right,baselineDifference:observed.chars[j].baseline-native.baseline}))};});
const chars=items.flatMap(i=>i.chars),maxHorizontal=Math.max(...chars.flatMap(c=>[Math.abs(c.leftDifference),Math.abs(c.rightDifference)])),maxBaseline=Math.max(...chars.map(c=>Math.abs(c.baselineDifference)));
const result={scope:'First-open PowerPoint PDF versus Chromium SVG suffix characters in 12 Chinese-prefix shape samples with Hiragino Sans GB; no edit/reopen, Japanese/Korean or other-font claim.',source:sourcePath,sourceSha256:sha(sourceBytes),runnerSha256:sha(await readFile('scripts/inspect-svg-cjk-shape-powerpoint.mjs')),browser:version,bundles:bundles.map(({url,...rest})=>rest),font:{file:fontPath,sha256:fontSha256},selectedFonts:fonts,items,maxHorizontal,maxBaseline};
await writeFile('docs/testing/tab-svg-cjk-shape-powerpoint-evidence.json',JSON.stringify(result,null,2)+'\n');console.log(JSON.stringify({samples:items.length,characters:chars.length,maxHorizontal,maxBaseline}));
