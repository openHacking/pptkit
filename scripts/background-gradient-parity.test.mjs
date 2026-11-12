import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, prepareLayoutRebind } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const gradient=(angle,scaled=false)=>({type:"linearGradient",angle,scaled,stops:[{position:0,color:"FF0000",opacity:.5},{position:1,color:"0000FF"}]});
function build(core){const d=core.createPresentation({size:{width:200,height:100}});d.defineSlideMaster({id:"m",name:"Gradient",background:gradient(0)});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m"});d.addSlide({id:"one",layoutId:"l"});d.addSlide({id:"two",layoutId:"l",background:gradient(90)});d.addSlide({id:"three",layoutId:"l",background:gradient(45,true)});return d}

test("linear background gradient geometry, inherited ownership, OOXML and browser parity",async t=>{
 const d=build({createPresentation}),normalized=normalizePresentation(d);assert.deepEqual(normalized.slides.map(s=>s.backgroundSource),["master","slide","slide"]);
 const svg=await renderPresentationToSvg(d);assert.deepEqual(svg.warnings.filter(w=>w.code==="background-gradient-approximation").map(w=>w.path),["slides.0.background","slides.1.background","slides.2.background"]);assert.equal(svg.status,"rendered-with-warnings");const bytes=(await generatePptx(d)).bytes,entries=readGeneratedZip(bytes),xml=name=>entries.get(name).toString("utf8");
 const imported=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v7"});
 assert.deepEqual(normalizePresentation(imported.document).slides.map(s=>s.background),normalized.slides.map(s=>s.background));
 assert.deepEqual(normalizePresentation(imported.document).slides.map(s=>s.backgroundSource),["master","slide","slide"]);
 imported.dispose();
 const master=[...entries].find(([name,data])=>name.startsWith("ppt/slideMasters/slideMaster") && data.toString("utf8").includes('name="Gradient"'))[1].toString("utf8");
 assert.match(master,/<a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"\/>/);assert.match(master,/<a:gs pos="100000"><a:srgbClr val="0000FF">/);assert.match(master,/<a:lin ang="0" scaled="0"\/>/);
 assert.doesNotMatch(xml("ppt/slides/slide1.xml"),/<p:bg>/);assert.match(xml("ppt/slides/slide2.xml"),/<a:lin ang="5400000" scaled="0"\/>/);assert.match(xml("ppt/slides/slide3.xml"),/<a:lin ang="2700000" scaled="1"\/>/);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const core=await readFile("packages/core/dist/browser/index.js"),layout=await readFile("packages/layout/dist/browser/index.js"),renderer=await readFile("packages/svg-renderer/dist/browser/index.js");
 const coreUrl=`data:text/javascript;base64,${core.toString("base64")}`,layoutUrl=`data:text/javascript;base64,${layout.toString("base64")}`;
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{"@pptkit/core":coreUrl,"@pptkit/layout":layoutUrl}})}</script>`);
 const result=await page.evaluate(async({coreUrl,rendererUrl})=>{const core=await import(coreUrl),renderer=await import(rendererUrl);const gradient=(angle,scaled=false)=>({type:"linearGradient",angle,scaled,stops:[{position:0,color:"FF0000",opacity:.5},{position:1,color:"0000FF"}]});const d=core.createPresentation({size:{width:200,height:100}});d.defineSlideMaster({id:"m",name:"Gradient",background:gradient(0)});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m"});d.addSlide({id:"one",layoutId:"l"});d.addSlide({id:"two",layoutId:"l",background:gradient(90)});d.addSlide({id:"three",layoutId:"l",background:gradient(45,true)});return (await renderer.renderPresentationToSvg(d)).slides.map(s=>s.svg)}, {coreUrl,rendererUrl:`data:text/javascript;base64,${renderer.toString("base64")}`});
 assert.deepEqual(result,svg.slides.map(s=>s.svg));assert.deepEqual(requests,[]);
 const pixels=await page.evaluate(async svgs=>{const out=[];for(const svg of svgs){const image=new Image(),url=URL.createObjectURL(new Blob([svg],{type:"image/svg+xml"}));await new Promise((ok,fail)=>{image.onload=ok;image.onerror=fail;image.src=url});const c=document.createElement("canvas");c.width=200;c.height=100;const ctx=c.getContext("2d");ctx.drawImage(image,0,0);out.push([[10,50],[190,50],[100,10],[100,90],[160,10],[40,70]].map(([x,y])=>[...ctx.getImageData(x,y,1,1).data]));URL.revokeObjectURL(url)}return out},result);
 assert.ok(pixels[0][0][3] < pixels[0][1][3]);
 assert.ok(pixels[0][0][0]>pixels[0][1][0] && pixels[0][0][2]<pixels[0][1][2]);assert.ok(pixels[1][2][0]>pixels[1][3][0] && pixels[1][2][2]<pixels[1][3][2]);
 assert.ok(pixels[2][0][0]>pixels[2][1][0] && pixels[2][2][0]>pixels[2][3][0]);
 assert.ok(pixels[2][4].every((c,i)=>Math.abs(c-pixels[2][5][i])<=2),"scaled 45° has constant color along normalized-coordinate diagonals");
});

test("theme gradient stops survive export and resolve in the selected owner's theme",async()=>{
 const d=createPresentation({size:{width:200,height:100}});
 const gradient={type:"linearGradient",angle:0,stops:[{position:0,color:{theme:"accent1"},opacity:.4},{position:1,color:{theme:"accent2"}}]};
 d.defineSlideMaster({id:"red-master",name:"Red",theme:{colors:{accent1:"AA0000",accent2:"0000AA"}},background:gradient});
 d.defineSlideMaster({id:"green-master",name:"Green",theme:{colors:{accent1:"00AA00",accent2:"AAAA00"}},background:gradient});
 d.defineSlideLayout({id:"red-layout",name:"Red layout",masterId:"red-master"});d.defineSlideLayout({id:"green-layout",name:"Green layout",masterId:"green-master"});
 d.addSlide({id:"red",layoutId:"red-layout"});d.addSlide({id:"green",layoutId:"green-layout"});
 const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/stop-color="#AA0000" stop-opacity="0.4"/);assert.match(svg.slides[1].svg,/stop-color="#00AA00" stop-opacity="0.4"/);
 const pptx=(await generatePptx(d)).bytes;const entries=readGeneratedZip(pptx),masters=[...entries].filter(([k])=>/^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(k)).map(([,v])=>v.toString('utf8'));
 assert.equal(masters.filter(xml=>xml.includes('<a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr>')).length,2);
 const session=await importPptx(pptx,{projection:"content",masterArtwork:"artwork-v7"});const imported=normalizePresentation(session.document).slides.map(s=>s.background.stops);
 assert.deepEqual(imported.map(stops=>stops.map(s=>[s.color,s.opacity])),[[["AA0000",.4],["0000AA",1]],[["00AA00",.4],["AAAA00",1]]]);session.dispose();
});

test("PPTX reports distinct gradient stops collapsed by position precision at each authored owner",async t=>{
 const d=createPresentation();const narrow={type:"linearGradient",angle:0,stops:[{position:0,color:"000000"},{position:.5,color:"FF0000"},{position:.500004,color:"0000FF"},{position:1,color:"FFFFFF"}]};
 const hard={...narrow,stops:[{position:0,color:"000000"},{position:.5,color:"FF0000"},{position:.5,color:"0000FF"},{position:1,color:"FFFFFF"}]};
 d.defineSlideMaster({id:"m",name:"Master",background:narrow});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m",background:narrow});
 d.addSlide({id:"s",layoutId:"l",background:narrow});d.addSlide({id:"duplicate",layoutId:"l",background:hard});d.addSlide({id:"inherited",layoutId:"l"});
 const result=await generatePptx(d),warnings=result.warnings.filter(w=>w.code==="background-gradient-stops-collapsed");
 assert.equal(result.status,"generated-with-warnings");assert.equal(warnings.length,3);assert.deepEqual(warnings.map(w=>[w.masterId??w.layoutId??w.slideId,w.path]),[["m","masters.0.background.stops.2.position"],["l",`layouts.${d.layouts.findIndex(l=>l.id==="l")}.background.stops.2.position`],["s","slides.0.background.stops.2.position"]]);
 const xml=readGeneratedZip(result.bytes).get("ppt/slides/slide1.xml").toString("utf8");assert.equal((xml.match(/<a:gs pos="50000">/g)??[]).length,2);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const core=await readFile("packages/core/dist/browser/index.js"),layout=await readFile("packages/layout/dist/browser/index.js"),exporter=await readFile("packages/pptx-exporter/dist/browser/index.js");
 const url=bytes=>`data:text/javascript;base64,${bytes.toString("base64")}`,coreUrl=url(core),exporterUrl=url(exporter);
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{"@pptkit/core":coreUrl,"@pptkit/layout":url(layout)}})}</script>`);
 const browserWarnings=await page.evaluate(async({coreUrl,exporterUrl,narrow,hard})=>{const core=await import(coreUrl),exporter=await import(exporterUrl),d=core.createPresentation();d.defineSlideMaster({id:"m",name:"Master",background:narrow});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m",background:narrow});d.addSlide({id:"s",layoutId:"l",background:narrow});d.addSlide({id:"duplicate",layoutId:"l",background:hard});d.addSlide({id:"inherited",layoutId:"l"});return (await exporter.generatePptx(d)).warnings.filter(w=>w.code==="background-gradient-stops-collapsed")},{coreUrl,exporterUrl,narrow,hard});assert.deepEqual(browserWarnings,warnings);
});


test("rebound gradient appearance stays attached to the intended slide through PPTX export/import",async()=>{
 const d=createPresentation();const gradient={type:"linearGradient",angle:45,stops:[{position:0,color:{theme:"accent1"}},{position:1,color:"0000FF",opacity:.5}]};
 d.defineSlideMaster({id:"red",name:"Red",theme:{colors:{accent1:"AA0000"}},background:gradient});d.defineSlideMaster({id:"green",name:"Green",theme:{colors:{accent1:"00AA00"}},background:gradient});
 d.defineSlideLayout({id:"a",name:"A",masterId:"red"});d.defineSlideLayout({id:"b",name:"B",masterId:"green"});d.addSlide({id:"preserve",layoutId:"a"});d.addSlide({id:"adopt",layoutId:"a"});
 const preserved=prepareLayoutRebind(d,[{slideId:"preserve",layoutId:"b"}],{localAppearance:"preserve"});assert.equal(preserved.status,"ready");d.applyLayoutRebind(preserved);
 const adopted=prepareLayoutRebind(d,[{slideId:"adopt",layoutId:"b"}],{localAppearance:"adopt"});assert.equal(adopted.status,"ready");d.applyLayoutRebind(adopted);
 const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/stop-color="#AA0000"/);assert.match(svg.slides[1].svg,/stop-color="#00AA00"/);
 const bytes=(await generatePptx(d)).bytes,session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v7"});const slides=normalizePresentation(session.document).slides;
 assert.deepEqual(slides.map(s=>s.backgroundSource),["slide","master"]);assert.deepEqual(slides.map(s=>s.background.stops[0].color),["AA0000","00AA00"]);assert.deepEqual(slides.map(s=>s.background.stops[1].opacity),[.5,.5]);session.dispose();
});

test("gradient per-stop origin records match between Node and browser",async t=>{
 const make=core=>{const d=core.createPresentation({id:"origin-parity"});d.defineSlideMaster({id:'m',name:'M',background:{type:'linearGradient',angle:0,stops:[{position:0,color:{theme:'accent1'}},{position:1,color:'0000FF',opacity:0}]}});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({id:'s',layoutId:'l'});return d};
 const node=(await import('../packages/core/dist/index.js')).resolvePresentationInheritance(make({createPresentation})).origins;
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),bundle=await readFile('packages/core/dist/browser/index.js'),coreUrl=`data:text/javascript;base64,${bundle.toString('base64')}`;
 const actual=await page.evaluate(async url=>{const core=await import(url),d=core.createPresentation({id:"origin-parity"});d.defineSlideMaster({id:'m',name:'M',background:{type:'linearGradient',angle:0,stops:[{position:0,color:{theme:'accent1'}},{position:1,color:'0000FF',opacity:0}]}});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({id:'s',layoutId:'l'});return core.resolvePresentationInheritance(d).origins},coreUrl);
 assert.deepEqual(actual,node);assert.equal(actual.filter(x=>x.target.kind==='slide'&&x.property.startsWith('background.stops.')).length,6);
});
