import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { inspectPptxMasterPictures, importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { structureFixture } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { pack, contentTypes } from "../packages/pptx-parser/test/opc-fixture.mjs";

test("master picture inventory matches browser and Node on shared and nested artwork",async t=>{
 const png=await readFile("tests/fixtures/transparent-replacement.png"),d=createPresentation();
 d.registerAsset({id:"img",kind:"image",source:{type:"bytes",value:png},mimeType:"image/png",width:4,height:2});
 d.defineSlideMaster({id:"m",name:"Master",elements:[{id:"group",type:"group",box:{x:10,y:20,width:100,height:80},coordinateSize:{width:100,height:100},transform:{rotation:15,flipH:true},children:[{id:"inner",type:"group",box:{x:3,y:4,width:60,height:50},coordinateSize:{width:80,height:70},transform:{rotation:-5},children:[{id:"pic",type:"image",assetId:"img",fit:"crop",crop:{left:.1},box:{x:5,y:10,width:40,height:30}}]}]}]});
 d.defineSlideLayout({id:"l",name:"Layout",masterId:"m"});d.addSlide({id:"one",layoutId:"l"});d.addSlide({id:"two",layoutId:"l",showInheritedShapes:false});
 const bytes=(await generatePptx(d)).bytes,node=await inspectPptxMasterPictures(bytes),entry=await readFile("packages/pptx-parser/dist/browser/index.js");
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];
 await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const result=await page.evaluate(async({bundle,data})=>{const api=await import(bundle);return await api.inspectPptxMasterPictures(Uint8Array.from(data))},
  {bundle:`data:text/javascript;base64,${entry.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(result,node);const part=node.masters.find(m=>m.pictures.some(p=>p.name==="pic"));assert.ok(part);
 assert.equal(part.slideIds.length,2);assert.equal(part.pictures[0].groupElementIndices.length,2);assert.equal(part.pictures[0].crop.left,"10000");assert.deepEqual(requests,[]);
 const imported=await importPptx(bytes,{projection:"content",masterArtwork:"pictures-v1"}),group=imported.document.masters.find(m=>m.elements.length).elements[0];
 assert.equal(group.type,"group");assert.equal(group.transform.rotation,15);assert.equal(group.transform.flipH,true);assert.equal(group.children[0].transform.rotation,-5);assert.equal(group.children[0].children[0].type,"image");
 const browserGroup=await page.evaluate(async({bundle,data})=>{const api=await import(bundle),s=await api.importPptx(Uint8Array.from(data),{projection:"content",masterArtwork:"pictures-v1"});return s.document.masters.find(m=>m.elements.length).elements[0]}, {bundle:`data:text/javascript;base64,${entry.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(browserGroup,group);
 const sourceSvg=await renderPresentationToSvg(d),importSvg=await renderPresentationToSvg(imported.document);
 assert.deepEqual(importSvg.slides.map(s=>(s.svg.match(/<image\b/g)||[]).length),[1,0]);
 const picture=async svg=>{await page.setContent(`<body style="margin:0">${svg}</body>`);return await page.screenshot()};
 assert.deepEqual(await picture(importSvg.slides[0].svg),await picture(sourceSvg.slides[0].svg));
 const round=(await generatePptx(imported.document)).bytes,reimported=await importPptx(round,{projection:"content",masterArtwork:"pictures-v1"});
 assert.deepEqual(await picture((await renderPresentationToSvg(reimported.document)).slides[0].svg),await picture(sourceSvg.slides[0].svg));
 imported.dispose();reimported.dispose();
});

test("shared direct artwork survives explicit import and fresh export with browser parity",async t=>{
 const png=await readFile("tests/fixtures/transparent-replacement.png"),d=createPresentation();
 d.registerAsset({id:"img",kind:"image",source:{type:"bytes",value:png},mimeType:"image/png",width:4,height:2});
 d.defineSlideMaster({id:"m",name:"Master",elements:[{id:"back",type:"shape",shape:"rect",box:{x:0,y:0,width:100,height:60},style:{fill:{type:"solid",color:"0000FF"},stroke:{paint:{type:"none"}}}},{id:"pic",type:"image",assetId:"img",fit:"crop",crop:{left:.1},box:{x:5,y:10,width:40,height:30}}]});
 d.defineSlideLayout({id:"l",name:"Layout",masterId:"m",elements:[{id:"lp",type:"image",assetId:"img",fit:"stretch",box:{x:60,y:10,width:20,height:20}}]});
 d.defineSlideLayout({id:"l2",name:"No master artwork",masterId:"m",showMasterShapes:false,elements:[{id:"lp2",type:"image",assetId:"img",fit:"stretch",box:{x:80,y:10,width:20,height:20}}]});
 d.addSlide({id:"one",layoutId:"l"});d.addSlide({id:"two",layoutId:"l"});d.addSlide({id:"hidden-art",layoutId:"l",showInheritedShapes:false});d.addSlide({id:"hidden-master",layoutId:"l2"});
 const bytes=(await generatePptx(d)).bytes,session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v1"});
 const summarize=s=>({masters:s.document.masters.map(m=>({id:m.id,elements:m.elements})),layouts:s.document.layouts.filter(l=>l.masterId).map(l=>({id:l.id,masterId:l.masterId,elements:l.elements})),slides:s.document.slides.map(x=>({id:x.id,layoutId:x.layoutId,show:x.showInheritedShapes})),assets:s.document.assets.length});
 const summary=summarize(session);assert.deepEqual(summary.masters.flatMap(m=>m.elements).map(e=>e.type),["shape","image"]);assert.equal(summary.layouts.flatMap(l=>l.elements).length,2);assert.equal(summary.assets,1);
 const svg=await renderPresentationToSvg(session.document);assert.deepEqual(svg.slides.map(s=>(s.svg.match(/<image\b/g)||[]).length),[2,2,0,1]);
 const rebuilt=(await generatePptx(session.document)).bytes,again=await importPptx(rebuilt,{projection:"content",masterArtwork:"artwork-v1"});
 assert.deepEqual(again.document.masters.flatMap(m=>m.elements).map(e=>e.type),["shape","image"]);assert.equal(again.document.layouts.flatMap(l=>l.elements).length,2);
 assert.deepEqual((await renderPresentationToSvg(again.document)).slides.map(s=>(s.svg.match(/<image\b/g)||[]).length),[2,2,0,1]);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const bundle=await readFile("packages/pptx-parser/dist/browser/index.js");
 const browserSummary=await page.evaluate(async({url,data})=>{const api=await import(url),s=await api.importPptx(Uint8Array.from(data),{projection:"content",masterArtwork:"artwork-v1"});return {masters:s.document.masters.map(m=>({id:m.id,elements:m.elements})),layouts:s.document.layouts.filter(l=>l.masterId).map(l=>({id:l.id,masterId:l.masterId,elements:l.elements})),slides:s.document.slides.map(x=>({id:x.id,layoutId:x.layoutId,show:x.showInheritedShapes})),assets:s.document.assets.length}},{url:`data:text/javascript;base64,${bundle.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(browserSummary,summary);assert.deepEqual(requests,[]);
 const screenshot=async svg=>{await page.setContent(`<body style="margin:0">${svg}</body>`);return page.screenshot()};const original=await renderPresentationToSvg(d);
 for(let i=0;i<4;i++) assert.deepEqual(await screenshot(svg.slides[i].svg),await screenshot(original.slides[i].svg));
 session.dispose();again.dispose();
});

test("shared shape text survives opt-in import and fresh export with browser parity",async t=>{
 const d=createPresentation();d.defineSlideMaster({id:"m",name:"Master",elements:[{id:"label",type:"shape",shape:"rect",box:{x:20,y:30,width:300,height:80},style:{fill:{type:"solid",color:"0000FF"},stroke:{paint:{type:"none"}}},text:{content:[{runs:[{text:"Shared text",style:{fontFamily:"Arial",fontSize:24,bold:true}}]}]}}]});
 d.defineSlideLayout({id:"l",name:"Layout",masterId:"m"});d.addSlide({id:"one",layoutId:"l"});d.addSlide({id:"hidden",layoutId:"l",showInheritedShapes:false});
 const bytes=(await generatePptx(d)).bytes,session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v3"});
 const shape=session.document.masters.find(m=>m.elements.length).elements[0];assert.equal(shape.text.content[0].runs[0].text,"Shared text");
 const svg=await renderPresentationToSvg(session.document);assert.match(svg.slides[0].svg,/Shared text/);assert.doesNotMatch(svg.slides[1].svg,/Shared text/);
 const round=(await generatePptx(session.document)).bytes,again=await importPptx(round,{projection:"content",masterArtwork:"artwork-v3"});assert.equal(again.document.masters.find(m=>m.elements.length).elements[0].text.content[0].runs[0].text,"Shared text");
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const bundle=await readFile("packages/pptx-parser/dist/browser/index.js"),browserShape=await page.evaluate(async({url,data})=>{const api=await import(url),s=await api.importPptx(Uint8Array.from(data),{projection:"content",masterArtwork:"artwork-v3"});return s.document.masters.find(m=>m.elements.length).elements[0]}, {url:`data:text/javascript;base64,${bundle.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(browserShape,shape);assert.deepEqual(requests,[]);session.dispose();again.dispose();
});

test("master scoped theme text resolves identically in browser and Node",async t=>{
 const f=structureFixture();const text='<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>Theme</a:t></a:r></a:p></p:txBody>';
 const shape=`<p:sp><p:nvSpPr><p:cNvPr id="2" name="text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></p:spPr>${text}</p:sp>`;
 f.parts["deck/master.xml"]=`<p:sldMaster xmlns:p="${f.P}" xmlns:a="${f.A}" xmlns:r="${f.R}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg><p:spTree>${shape}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst></p:sldMaster>`;
 f.parts["deck/theme.xml"]=`<a:theme xmlns:a="${f.A}"><a:themeElements><a:clrScheme name="theme"><a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme><a:fontScheme name="theme"><a:majorFont><a:latin typeface="Theme Font"/></a:majorFont></a:fontScheme><a:fmtScheme name="matrix"><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"><a:alphaMod val="50000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
 const bytes=pack(f.parts,contentTypes(undefined,Object.entries(f.types).map(([PartName,type])=>({PartName:`/${PartName}`,ContentType:`application/vnd.openxmlformats-officedocument.${type}+xml`}))));
 const session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v6"}),element=session.document.masters[0].elements[0],node={text:element.text.content[0].runs[0].style,shape:element.style,background:session.document.masters[0].background};
 assert.equal(node.text.color,"123456");assert.equal(node.text.fontFamily,"Theme Font");assert.deepEqual(node.shape.fill,{type:"solid",color:"123456",opacity:.5});assert.equal(node.shape.stroke.paint.color,"123456");
 assert.deepEqual(node.background,{type:"solid",color:"123456",opacity:.5});
 const rebuilt=(await generatePptx(session.document)).bytes,round=await importPptx(rebuilt,{projection:"content",masterArtwork:"artwork-v6"});
 const owner=round.document.masters.find(m=>m.elements.length);assert.deepEqual(owner.elements[0].style.fill,node.shape.fill);assert.deepEqual(owner.background,node.background);round.dispose();
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const bundle=await readFile("packages/pptx-parser/dist/browser/index.js"),result=await page.evaluate(async({url,data})=>{const api=await import(url),s=await api.importPptx(Uint8Array.from(data),{projection:"content",masterArtwork:"artwork-v6"}),element=s.document.masters[0].elements[0];return {text:element.text.content[0].runs[0].style,shape:element.style,background:s.document.masters[0].background}},{url:`data:text/javascript;base64,${bundle.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(result,node);assert.deepEqual(requests,[]);session.dispose();
});

test("explicit master, layout and slide backgrounds survive projection and browser roundtrip",async t=>{
 const d=createPresentation();d.defineSlideMaster({id:"m",name:"Master",background:{type:"solid",color:"123456"}});
 d.defineSlideLayout({id:"lm",name:"Master background",masterId:"m"});d.defineSlideLayout({id:"ll",name:"Layout background",masterId:"m",background:{type:"solid",color:"AA0000"}});
 d.addSlide({id:"m-slide",layoutId:"lm"});d.addSlide({id:"l-slide",layoutId:"ll",showInheritedShapes:false});d.addSlide({id:"s-slide",layoutId:"ll",background:{type:"solid",color:"00BB00"}});
 const bytes=(await generatePptx(d)).bytes,session=await importPptx(bytes,{projection:"content",masterArtwork:"artwork-v5"});
 const snapshot=s=>({masters:s.document.masters.map(m=>m.background),layouts:s.document.layouts.filter(l=>l.masterId).map(l=>l.background),slides:s.document.slides.map(x=>x.background)});
 const node=snapshot(session),svg=await renderPresentationToSvg(session.document),source=await renderPresentationToSvg(d);
 const round=(await generatePptx(session.document)).bytes,again=await importPptx(round,{projection:"content",masterArtwork:"artwork-v5"});
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route("**/*",route=>{requests.push(route.request().url());return route.abort()});
 const bundle=await readFile("packages/pptx-parser/dist/browser/index.js"),result=await page.evaluate(async({url,data})=>{const api=await import(url),s=await api.importPptx(Uint8Array.from(data),{projection:"content",masterArtwork:"artwork-v5"});return {masters:s.document.masters.map(m=>m.background),layouts:s.document.layouts.filter(l=>l.masterId).map(l=>l.background),slides:s.document.slides.map(x=>x.background)}},{url:`data:text/javascript;base64,${bundle.toString("base64")}`,data:[...bytes]});
 assert.deepEqual(result,node);assert.deepEqual(requests,[]);
 const screenshot=async value=>{await page.setContent(`<body style="margin:0">${value}</body>`);return page.screenshot()};const reopened=await renderPresentationToSvg(again.document);
 for(let i=0;i<3;i++){const expected=await screenshot(source.slides[i].svg);assert.deepEqual(await screenshot(svg.slides[i].svg),expected);assert.deepEqual(await screenshot(reopened.slides[i].svg),expected)}
 session.dispose();again.dispose();
});
