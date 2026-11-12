import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

function deck(shadow, extra={}) {
 const d=createPresentation({size:{width:100,height:80},theme:{colors:{accent1:'334455'}}});
 d.addSlide({id:'s'}).addElement({id:'box',type:'shape',shape:'rect',box:{x:20,y:20,width:20,height:20},style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow},...extra});return d;
}

test('shape shadow normalizes, traces and emits editable theme-aware DrawingML',async()=>{
 const d=deck({color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90},{opacity:.5});
 assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);
 const shadow=normalizePresentation(d).slides[0].elements[0].style.shadow;assert.deepEqual(shadow,{color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90});
 assert.ok(resolvePresentationInheritance(d).origins.some(x=>x.target.id==='box'&&x.property==='style.shadow.color'));
 const xml=readGeneratedZip((await generatePptx(d)).bytes).get('ppt/slides/slide1.xml').toString('utf8');
 assert.match(xml,/<a:outerShdw blurRad="76200" dist="101600" dir="5400000" algn="ctr" rotWithShape="1"><a:schemeClr val="accent1"><a:alpha val="40000"\/>/);
 const svg=await renderPresentationToSvg(d);assert.ok(svg.warnings.some(x=>x.code==='shape-shadow-approximation'));assert.match(svg.slides[0].svg,/flood-color="#334455" flood-opacity="0.4"/);
});

test('invalid shape shadow values are rejected while boundary values are accepted',()=>{
 const base={color:'000000',blur:0,distance:0,direction:0};
 for(const shadow of [{...base,color:'bad'},{...base,opacity:-.1},{...base,opacity:NaN},{...base,blur:-1},{...base,blur:1001},{...base,distance:Infinity},{...base,distance:1001},{...base,direction:360},{...base,direction:-1}]) assert.ok(validatePresentation(deck(shadow)).some(x=>x.code==='invalid-shape-shadow'));
 assert.ok(!validatePresentation(deck({...base,blur:1000,distance:1000,direction:359.99,opacity:0})).some(x=>x.code==='invalid-shape-shadow'));
 assert.equal(normalizePresentation(deck(base)).slides[0].elements[0].style.shadow.opacity,1);
});

test('export warns when translucent source paint or object opacity can alter shadow appearance',async()=>{
 const shadow={color:'000000',blur:0,distance:2,direction:0};
 for(const extra of [{opacity:.5},{style:{fill:{type:'solid',color:'FF0000',opacity:.5},stroke:{paint:{type:'none'}},shadow}},{style:{fill:{type:'linearGradient',angle:0,stops:[{position:0,color:'FF0000',opacity:.5},{position:1,color:'000000'}]},stroke:{paint:{type:'none'}},shadow}},{style:{fill:{type:'pattern',pattern:'horizontal',foreground:'FF0000',background:'000000',foregroundOpacity:.5},stroke:{paint:{type:'none'}},shadow}},{style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'solid',color:'000000',opacity:.5}},shadow}}]){
  const result=await generatePptx(deck(shadow,extra));assert.equal(result.status,'generated-with-warnings');assert.ok(result.warnings.some(w=>w.code==='shape-shadow-transparency-application-difference'&&w.slideId==='s'&&w.elementId==='box'));
 }
 assert.ok(!(await generatePptx(deck(shadow))).warnings.some(w=>w.code==='shape-shadow-transparency-application-difference'));
});

test('opt-in estimated shadow overflow is forwarded by preview and export without changing content',async()=>{
 const d=createPresentation({size:{width:100,height:80}});d.addSlide({id:'s'}).addElement({id:'box',type:'shape',shape:'rect',box:{x:80,y:20,width:10,height:10},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',blur:0,distance:15,direction:0}}});
 const options={layout:{diagnostics:{shadowBounds:true}}},plain=await generatePptx(d),checked=await generatePptx(d,options);assert.deepEqual(checked.bytes,plain.bytes);assert.ok(checked.warnings.some(w=>w.code==='shadow-slide-overflow-approximate'&&w.elementId==='box'));
 const before=await renderPresentationToSvg(d),after=await renderPresentationToSvg(d,options);assert.deepEqual(after.slides,before.slides);assert.ok(after.warnings.some(w=>w.code==='shadow-slide-overflow-approximate'&&w.elementId==='box'));
});

test('import reports shadow as unsupported and no-op save preserves its source bytes',async()=>{
 const bytes=(await generatePptx(deck({color:'000000',blur:2,distance:2,direction:90}))).bytes;
 const session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});
 assert.equal(session.document.slides[0].elements[0].style.shadow,undefined);
 assert.ok(session.report.findings.some(f=>f.feature==='drawing.unmapped-property'&&f.operation==='import'&&f.tier==='unsupported'&&f.source.path.includes('effectLst')));
 const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));session.dispose();
});

test('artwork-v11 projects scoped shadows while older and unsupported profiles retain reports',async()=>{
 const {createRequire}=await import('node:module');const {unzipSync,zipSync,strFromU8,strToU8}=createRequire(new URL('../packages/pptx-parser/package.json',import.meta.url))('fflate');
 const d=createPresentation({theme:{colors:{accent1:'334455'}}}),shape=id=>({id,type:'shape',shape:'rect',box:{x:0,y:0,width:30,height:30},style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow:{color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90}}});
 d.defineSlideMaster({id:'m',name:'M',elements:[shape('shared')]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});const slide=d.addSlide({id:'s',layoutId:'l'});slide.addElement(shape('local'));slide.addElement({id:'g',type:'group',box:{x:40,y:0,width:30,height:30},coordinateSize:{width:30,height:30},children:[shape('nested')]});
 const bytes=(await generatePptx(d)).bytes,old=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'}),current=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v11'});
 assert.equal(old.document.slides[0].elements[0].style.shadow,undefined);const masterId=current.document.layouts.find(l=>l.id===current.document.slides[0].layoutId).masterId;
 const shapes=[current.document.masters.find(m=>m.id===masterId).elements[0],current.document.slides[0].elements[0],current.document.slides[0].elements[1].children[0]];for(const value of shapes)assert.deepEqual(value.style.shadow,{color:'334455',opacity:.4,blur:6,distance:8,direction:90});
 assert.equal(current.report.findings.filter(f=>f.feature.endsWith('shadow-dependencies')).length,3);const ready=await current.prepareSave({mode:'no-op'});assert.deepEqual((await current.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await current.prepareSave({mode:'edit'})).status,'blocked');old.dispose();current.dispose();
 const files=unzipSync(bytes),path='ppt/slides/slide1.xml';files[path]=strToU8(strFromU8(files[path]).replace('<a:outerShdw ','<a:outerShdw sx="90000" '));const unsupported=await importPptx(zipSync(files),{projection:'content',masterArtwork:'artwork-v11'});assert.equal(unsupported.document.slides[0].elements[0].style.shadow,undefined);assert.ok(unsupported.report.findings.some(f=>f.tier==='unsupported'&&f.source.path?.includes('effectLst')));unsupported.dispose();
 await assert.rejects(importPptx(bytes,{projection:'content',masterArtwork:'artwork-v11',sessionLimits:{maxTextStyleDependencies:1}}),{code:'import-limit-exceeded'});
});

test('LibreOfficeDev resave imports zero-blur fallback and explicit blur with source evidence',async()=>{
 const bytes=await readFile('tests/fixtures/shape-shadow-libreoffice.pptx'),s=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v11'});
 const shadows=s.document.slides.map(slide=>slide.elements.find(element=>element.type==='shape').style.shadow);
 assert.deepEqual(shadows.map(x=>x.blur),[0,0,18]);assert.deepEqual(shadows.map(x=>x.direction),[0,0,90]);assert.ok(shadows.every(x=>x.color==='000000'&&x.opacity===.5&&Math.abs(x.distance-60)<.02));
 const fallbacks=s.report.findings.filter(f=>f.operation==='import'&&f.reason.includes('Missing shadow blur radius'));assert.equal(fallbacks.length,2);assert.ok(fallbacks.every(f=>f.source.part.startsWith('/ppt/slides/')));
 const ready=await s.prepareSave({mode:'no-op'});assert.deepEqual((await s.save(ready.plan)).bytes,new Uint8Array(bytes));s.dispose();
});

test('LibreOfficeDev shadow subset remains stable through three projection and rebuild rounds',async()=>{
 let bytes=await readFile('tests/fixtures/shape-shadow-libreoffice.pptx'),expected;
 for(let round=0;round<4;round++){
  const session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v11'});
  try {
   const values=session.document.slides.map(slide=>slide.elements.find(element=>element.type==='shape').style.shadow);
   if(round===0)expected=structuredClone(values);else assert.deepEqual(values,expected);
   assert.equal(session.report.findings.filter(f=>f.operation==='import'&&f.reason.includes('Missing shadow blur radius')).length,round===0?2:0);
   const ready=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(ready.plan)).bytes,new Uint8Array(bytes));
   if(round<3){const output=await generatePptx(session.document);assert.ok(output.warnings.some(w=>w.code==='shape-shadow-transparency-application-difference'));const parts=readGeneratedZip(output.bytes);for(let i=1;i<=3;i++){const xml=parts.get(`ppt/slides/slide${i}.xml`).toString('utf8');assert.match(xml,new RegExp(`<a:outerShdw blurRad="${i===3?228600:0}" dist="762120" dir="${i===3?5400000:0}" algn="ctr" rotWithShape="1"><a:srgbClr val="000000"><a:alpha val="50000"/>`));}bytes=output.bytes;}
  }finally{session.dispose();}
 }
});

test('browser preview places an unblurred shadow in the authored direction',async t=>{
 const d=deck({color:'000000',opacity:.5,blur:0,distance:10,direction:0});const svg=(await renderPresentationToSvg(d)).slides[0].svg;
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const pixels=await page.evaluate(async svg=>{const image=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{image.onload=ok;image.onerror=fail;image.src=url});const canvas=document.createElement('canvas');canvas.width=100;canvas.height=80;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);const pixel=(x,y)=>[...ctx.getImageData(x,y,1,1).data];const result=[pixel(25,25),pixel(45,25),pixel(15,25)];URL.revokeObjectURL(url);return result},svg);
 assert.deepEqual(pixels[0],[255,0,0,255]);assert.ok(pixels[1][0]>=125&&pixels[1][0]<=130&&pixels[1][1]===pixels[1][0]);assert.deepEqual(pixels[2],[255,255,255,255]);
});

test('browser bundles retain grouped shadow theme and opacity behavior',async t=>{
 const create=core=>{const d=core.createPresentation({size:{width:100,height:80},theme:{colors:{accent1:'334455'}}});d.addSlide({id:'s'}).addElement({id:'group',type:'group',box:{x:10,y:10,width:60,height:60},coordinateSize:{width:60,height:60},opacity:.5,children:[{id:'box',type:'shape',shape:'rect',box:{x:10,y:10,width:20,height:20},opacity:.5,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow:{color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90}}}]});return d};
 const d=create({createPresentation}),expected=await renderPresentationToSvg(d),exported=await generatePptx(d),xml=readGeneratedZip(exported.bytes).get('ppt/slides/slide1.xml').toString('utf8');assert.match(xml,/<a:outerShdw[^>]*><a:schemeClr val="accent1"><a:alpha val="40000"\/>/);assert.ok(exported.warnings.some(w=>w.code==='shape-shadow-transparency-application-difference'&&w.elementId==='box'));
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),core=await readFile('packages/core/dist/browser/index.js'),layout=await readFile('packages/layout/dist/browser/index.js'),renderer=await readFile('packages/svg-renderer/dist/browser/index.js');const url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(core);
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl,'@pptkit/layout':url(layout)}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,rendererUrl})=>{const core=await import(coreUrl),renderer=await import(rendererUrl),d=core.createPresentation({size:{width:100,height:80},theme:{colors:{accent1:'334455'}}});d.addSlide({id:'s'}).addElement({id:'group',type:'group',box:{x:10,y:10,width:60,height:60},coordinateSize:{width:60,height:60},opacity:.5,children:[{id:'box',type:'shape',shape:'rect',box:{x:10,y:10,width:20,height:20},opacity:.5,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'}},shadow:{color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90}}}]});return renderer.renderPresentationToSvg(d)},{coreUrl,rendererUrl:url(renderer)});assert.deepEqual(actual,expected);
});

test('artwork-v11 import agrees across Node, browser ESM and global bundles',async t=>{
 const bytes=(await generatePptx(deck({color:{theme:'accent1'},opacity:.4,blur:6,distance:8,direction:90}))).bytes,options={projection:'content',masterArtwork:'artwork-v11'},node=await importPptx(bytes,options);
 const expected={shadow:node.document.slides[0].elements[0].style.shadow,findings:node.report.findings.filter(f=>f.feature==='drawing.shadow-dependencies')};node.dispose();
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort()});
 await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});const bundle=await readFile('packages/pptx-parser/dist/browser/index.js');
 const outputs=await page.evaluate(async({url,bytes,options})=>{const esm=await import(url),results=[];for(const api of [esm,globalThis.PPTKit]){const s=await api.importPptx(Uint8Array.from(bytes),options);try{const plan=await s.prepareSave({mode:'no-op'}),saved=await s.save(plan.plan);if(saved.bytes.length!==bytes.length||saved.bytes.some((x,i)=>x!==bytes[i]))throw Error('source changed');results.push({shadow:s.document.slides[0].elements[0].style.shadow,findings:s.report.findings.filter(f=>f.feature==='drawing.shadow-dependencies')});}finally{s.dispose();}}return results},{url:`data:text/javascript;base64,${bundle.toString('base64')}`,bytes:[...bytes],options});
 for(const result of outputs)assert.deepEqual(result,expected);assert.deepEqual(requests,[]);
});

test('shadow-boundary analysis agrees in Node and browser for transformed groups',async t=>{
 const make=core=>{const d=core.createPresentation({size:{width:100,height:100}});d.addSlide({id:'s'}).addElement({id:'g',type:'group',box:{x:50,y:0,width:100,height:100},coordinateSize:{width:50,height:50},children:[{id:'child',type:'shape',shape:'rect',box:{x:15,y:10,width:10,height:10},transform:{rotation:90},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',blur:2,distance:10,direction:0}}}]});return d};
 const options={diagnostics:{shadowBounds:true}},expected=resolveLayout(make({createPresentation}),options).analysis,browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),core=await readFile('packages/core/dist/browser/index.js'),layout=await readFile('packages/layout/dist/browser/index.js');const url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(core);
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,layoutUrl,options})=>{const core=await import(coreUrl),layout=await import(layoutUrl),d=core.createPresentation({size:{width:100,height:100}});d.addSlide({id:'s'}).addElement({id:'g',type:'group',box:{x:50,y:0,width:100,height:100},coordinateSize:{width:50,height:50},children:[{id:'child',type:'shape',shape:'rect',box:{x:15,y:10,width:10,height:10},transform:{rotation:90},style:{stroke:{paint:{type:'none'},width:0},shadow:{color:'000000',blur:2,distance:10,direction:0}}}]});return layout.resolveLayout(d,options).analysis},{coreUrl,layoutUrl:url(layout),options});
 assert.deepEqual(actual,expected);assert.equal(actual.diagnostics[0].code,'shadow-slide-overflow-approximate');
});
