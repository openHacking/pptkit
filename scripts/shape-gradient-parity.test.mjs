import test from 'node:test';
import assert from 'node:assert/strict';
import { chromium } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { createPresentation, normalizePresentation, validatePresentation, resolvePresentationInheritance } from '../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

test('shape gradients compose stop opacity, object opacity and nested group opacity in SVG and PPTX',async()=>{
 const d=createPresentation({size:{width:200,height:100},theme:{colors:{accent1:'AA0000'}}});const slide=d.addSlide({id:'s'});
 const fill={type:'linearGradient',angle:45,scaled:true,stops:[{position:0,color:{theme:'accent1'},opacity:.5},{position:1,color:'0000FF'}]};
 slide.addElement({id:'g',type:'group',box:{x:10,y:10,width:100,height:60},coordinateSize:{width:100,height:60},opacity:.5,children:[{id:'sh',type:'shape',shape:'rect',box:{x:5,y:5,width:80,height:40},opacity:.5,transform:{rotation:30},style:{fill}}]});
 assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);const normal=normalizePresentation(d).slides[0].elements[0].children[0];assert.deepEqual(normal.style.fill.stops[0],{position:0,color:{theme:'accent1'},opacity:.5});
 const svg=await renderPresentationToSvg(d);assert.ok(svg.warnings.some(w=>w.code==='shape-gradient-approximation'&&w.elementId==='sh'));assert.match(svg.slides[0].svg,/<linearGradient[^>]+gradientUnits="userSpaceOnUse"/);assert.match(svg.slides[0].svg,/stop-color="#AA0000" stop-opacity="0.5"/);assert.match(svg.slides[0].svg,/rotate\(30/);
 const result=await generatePptx(d),xml=readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString('utf8');assert.match(xml,/<a:lin ang="2700000" scaled="1"\/>/);assert.match(xml,/<a:schemeClr val="accent1"><a:alpha val="12500"\/>/);assert.match(xml,/<a:srgbClr val="0000FF"><a:alpha val="25000"\/>/);assert.match(xml,/rot="1800000"/);
});

test('shape gradients reject invalid stops and do not expand line or table fill support',()=>{
 const d=createPresentation(),s=d.addSlide();const bad={type:'linearGradient',angle:0,stops:[{position:.8,color:'000000'},{position:.1,color:'FFFFFF'}]};s.addElement({id:'s',type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10},style:{fill:bad}});assert.ok(validatePresentation(d).some(x=>x.path.endsWith('style.fill')&&x.code==='invalid-paint'));
 const other=createPresentation(),page=other.addSlide();const gradient={type:'linearGradient',angle:0,stops:[{position:0,color:'000000'},{position:1,color:'FFFFFF'}]};page.addElement({id:'line',type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10},style:{stroke:{paint:gradient}}});assert.ok(validatePresentation(other).some(x=>x.path.endsWith('style.stroke.paint')&&x.code==='invalid-paint'));
});


test('shape gradient stop collapse warning identifies the nested authored shape',async()=>{
 const d=createPresentation(),s=d.addSlide({id:'s'});s.addElement({id:'g',type:'group',box:{x:0,y:0,width:30,height:30},coordinateSize:{width:30,height:30},children:[{id:'shape',type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20},style:{fill:{type:'linearGradient',angle:0,stops:[{position:.5,color:'FF0000'},{position:.500004,color:'0000FF'}]}}}]});
 const result=await generatePptx(d);assert.deepEqual(result.warnings.filter(w=>w.code==='shape-gradient-stops-collapsed').map(w=>[w.slideId,w.elementId,w.path]),[['s','shape','slides.0.elements.0.children.0.style.fill.stops.1.position']]);
});


test('rendered shape gradient uses its own offset frame and clips to its geometry',async t=>{
 const d=createPresentation({size:{width:200,height:100}}),s=d.addSlide();s.addElement({id:'shape',type:'shape',shape:'rect',box:{x:50,y:20,width:100,height:60},style:{fill:{type:'linearGradient',angle:0,stops:[{position:0,color:'FF0000'},{position:1,color:'0000FF'}]},stroke:{paint:{type:'none'}}}});
 const svg=(await renderPresentationToSvg(d)).slides[0].svg,browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const pixels=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const c=document.createElement('canvas');c.width=200;c.height=100;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const result=[[55,50],[145,50],[20,50]].map(([x,y])=>[...ctx.getImageData(x,y,1,1).data]);URL.revokeObjectURL(url);return result},svg);
 assert.ok(pixels[0][0]>pixels[0][2]);assert.ok(pixels[1][2]>pixels[1][0]);assert.deepEqual(pixels[2],[255,255,255,255]);
});


test('shared authored shape gradient exports and imports under artwork-v8',async()=>{
 const d=createPresentation(),fill={type:'linearGradient',angle:30,scaled:false,stops:[{position:0,color:'FF0000',opacity:.4},{position:1,color:'0000FF'}]};
 d.defineSlideMaster({id:'m',name:'Master',elements:[{id:'shared',type:'shape',shape:'ellipse',box:{x:30,y:30,width:100,height:60},style:{fill,stroke:{paint:{type:'none'}}}}]});d.defineSlideLayout({id:'l',name:'Layout',masterId:'m'});d.addSlide({id:'s',layoutId:'l'});
 const bytes=(await generatePptx(d)).bytes,old=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v7'}),current=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v8'});
 assert.ok(old.document.masters.every(m=>m.elements.length===0));const owner=current.document.layouts.find(l=>l.id===current.document.slides[0].layoutId).masterId;const shape=current.document.masters.find(m=>m.id===owner).elements[0];assert.equal(shape.shape,'ellipse');assert.deepEqual(shape.style.fill,{type:'linearGradient',angle:30,scaled:false,stops:[{position:0,color:'FF0000',opacity:.4},{position:1,color:'0000FF',opacity:1}]});
 const svg=await renderPresentationToSvg(current.document);assert.ok(svg.warnings.some(w=>w.code==='shape-gradient-approximation'));assert.match(svg.slides[0].svg,/stop-color="#FF0000" stop-opacity="0.4"/);old.dispose();current.dispose();
});

test('artwork-v9 maps slide and grouped gradients, preserves older behavior and rejects unsupported gradient shapes',async()=>{
 const {createRequire}=await import('node:module');const require=createRequire(new URL('../packages/pptx-parser/package.json',import.meta.url));const {unzipSync,zipSync,strFromU8,strToU8}=require('fflate');
 const d=createPresentation({theme:{colors:{accent1:'123456'}}}),s=d.addSlide({id:'s'}),fill={type:'linearGradient',angle:90,stops:[{position:0,color:{theme:'accent1'},opacity:.4},{position:1,color:'0000FF'}]};
 const shape=id=>({id,type:'shape',shape:'rect',box:{x:0,y:0,width:50,height:30},style:{fill,stroke:{paint:{type:'none'}}}});
 s.addElement(shape('direct'));s.addElement({id:'group',type:'group',box:{x:50,y:0,width:50,height:30},coordinateSize:{width:50,height:30},children:[shape('nested')]});
 const bytes=(await generatePptx(d)).bytes,old=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v8'}),now=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9'});
 assert.equal(old.document.slides[0].elements[0].style.fill.type,'none');const actual=[now.document.slides[0].elements[0],now.document.slides[0].elements[1].children[0]];
 const traces=now.report.findings.filter(f=>f.feature==='drawing.gradient-dependencies');assert.equal(traces.length,2);for(const trace of traces){assert.ok(trace.resolvedFrom.some(x=>x.path.includes('gs[1]')));assert.ok(trace.resolvedFrom.some(x=>x.part.includes('/theme/')));assert.ok(Object.isFrozen(trace.resolvedFrom));}
 for(const e of actual)assert.deepEqual(e.style.fill,{type:'linearGradient',angle:90,scaled:false,stops:[{position:0,color:'123456',opacity:.4},{position:1,color:'0000FF',opacity:1}]});
 const ready=await now.prepareSave({mode:'no-op'});assert.deepEqual((await now.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await now.prepareSave({mode:'edit'})).status,'blocked');old.dispose();now.dispose();
 const files=unzipSync(bytes),xml=strFromU8(files['ppt/slides/slide1.xml']);files['ppt/slides/slide1.xml']=strToU8(xml.replace('<a:gradFill rotWithShape="1">','<a:gradFill flip="x" rotWithShape="1">'));const bad=await importPptx(zipSync(files),{projection:'content',masterArtwork:'artwork-v9'});
 assert.equal(bad.document.slides[0].elements.length,1);assert.ok(bad.report.findings.some(f=>f.tier==='unsupported'&&f.reason.includes('Shape gradient is outside')));assert.equal(bad.report.findings.filter(f=>f.feature==='drawing.gradient-dependencies').length,1);bad.dispose();
 await assert.rejects(importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9',sessionLimits:{maxTextStyleDependencies:1}}),{code:'import-limit-exceeded'});
});


test('shape gradient origin report agrees between Node and browser',async t=>{
 const build=core=>{const d=core.createPresentation({id:'shape-origin'}),s=d.addSlide({id:'s'});s.addElement({id:'shape',type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20},style:{fill:{type:'linearGradient',angle:0,stops:[{position:0,color:'FF0000'},{position:1,color:'0000FF',opacity:0}]}}});return d};
 const expected=resolvePresentationInheritance(build({createPresentation})).origins.filter(x=>x.target.id==='shape'&&x.property.startsWith('style.fill.'));
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),bytes=await readFile('packages/core/dist/browser/index.js'),url=`data:text/javascript;base64,${bytes.toString('base64')}`;
 const actual=await page.evaluate(async url=>{const core=await import(url),d=core.createPresentation({id:'shape-origin'}),s=d.addSlide({id:'s'});s.addElement({id:'shape',type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20},style:{fill:{type:'linearGradient',angle:0,stops:[{position:0,color:'FF0000'},{position:1,color:'0000FF',opacity:0}]}}});return core.resolvePresentationInheritance(d).origins.filter(x=>x.target.id==='shape'&&x.property.startsWith('style.fill.'))},url);
 assert.equal(actual.length,10);assert.deepEqual(actual,expected);
});
