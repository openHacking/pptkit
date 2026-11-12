import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation } from '../packages/core/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const build=core=>{const d=core.createPresentation({id:'patterns',size:{width:200,height:100},theme:{colors:{accent1:'AA0000'}}}),s=d.addSlide({id:'s'});for(const [i,pattern] of ['horizontal','vertical','cross'].entries())s.addElement({id:`p${i}`,type:'shape',shape:'rect',box:{x:10+i*60,y:20,width:50,height:50},opacity:.5,style:{fill:{type:'pattern',pattern,foreground:{theme:'accent1'},background:'0000FF',foregroundOpacity:.4,backgroundOpacity:.8},stroke:{paint:{type:'none'}}}});return d};
test('shape pattern fills preserve theme, opacity and editable OOXML with explicit SVG approximation',async t=>{
 const d=build({createPresentation}),normalized=normalizePresentation(d);assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);assert.equal(normalized.slides[0].elements[0].style.fill.pattern,'horizontal');
 const origins=resolvePresentationInheritance(d).origins;assert.equal(origins.find(x=>x.target.id==='p0'&&x.property==='style.fill.foregroundOpacity').source.path,'slides.0.elements.0.style.fill.foregroundOpacity');
 const svg=await renderPresentationToSvg(d);assert.equal(svg.warnings.filter(w=>w.code==='shape-pattern-approximation').length,3);assert.equal((svg.slides[0].svg.match(/<pattern /g)??[]).length,3);
 const pptx=(await generatePptx(d)).bytes,xml=readGeneratedZip(pptx).get('ppt/slides/slide1.xml').toString('utf8');for(const preset of ['horz','vert','cross'])assert.match(xml,new RegExp(`<a:pattFill prst="${preset}">`));assert.match(xml,/<a:fgClr><a:schemeClr val="accent1"><a:alpha val="20000"\/>/);assert.match(xml,/<a:bgClr><a:srgbClr val="0000FF"><a:alpha val="40000"\/>/);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),core=await readFile('packages/core/dist/browser/index.js'),layout=await readFile('packages/layout/dist/browser/index.js'),renderer=await readFile('packages/svg-renderer/dist/browser/index.js');const url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(core);
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl,'@pptkit/layout':url(layout)}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,rendererUrl})=>{const core=await import(coreUrl),renderer=await import(rendererUrl),d=core.createPresentation({id:'patterns',size:{width:200,height:100},theme:{colors:{accent1:'AA0000'}}}),s=d.addSlide({id:'s'});for(const [i,pattern] of ['horizontal','vertical','cross'].entries())s.addElement({id:`p${i}`,type:'shape',shape:'rect',box:{x:10+i*60,y:20,width:50,height:50},opacity:.5,style:{fill:{type:'pattern',pattern,foreground:{theme:'accent1'},background:'0000FF',foregroundOpacity:.4,backgroundOpacity:.8},stroke:{paint:{type:'none'}}}});return renderer.renderPresentationToSvg(d)},{coreUrl,rendererUrl:url(renderer)});assert.deepEqual(actual,svg);
});
test('invalid patterns and unsupported paint domains are diagnosed',()=>{
 const d=createPresentation(),s=d.addSlide(),good={type:'pattern',pattern:'cross',foreground:'FF0000',background:'000000'};for(const [i,fill] of [{...good,pattern:'unknown'},{...good,foreground:'bad'},{...good,foregroundOpacity:2},{...good,backgroundOpacity:NaN}].entries())s.addElement({id:`s${i}`,type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10},style:{fill}});assert.equal(validatePresentation(d).filter(x=>x.code==='invalid-paint').length,4);
 const e=createPresentation(),slide=e.addSlide();slide.setBackground(good);assert.ok(validatePresentation(e).some(x=>x.code==='invalid-paint'&&x.path.endsWith('background')));
});


test('artwork-v10 imports shared and local pattern fills with scoped colors, opacity and source dependencies',async()=>{
 const {createRequire}=await import('node:module');const {unzipSync,zipSync,strFromU8,strToU8}=createRequire(new URL('../packages/pptx-parser/package.json',import.meta.url))('fflate');
 const d=createPresentation({theme:{colors:{accent1:'AA0000'}}}),fill={type:'pattern',pattern:'cross',foreground:{theme:'accent1'},background:'0000FF',foregroundOpacity:.4,backgroundOpacity:.8},shape=id=>({id,type:'shape',shape:'rect',box:{x:0,y:0,width:30,height:30},style:{fill,stroke:{paint:{type:'none'}}}});
 d.defineSlideMaster({id:'m',name:'M',elements:[shape('shared')]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});const slide=d.addSlide({id:'s',layoutId:'l'});slide.addElement(shape('local'));slide.addElement({id:'g',type:'group',box:{x:40,y:0,width:30,height:30},coordinateSize:{width:30,height:30},children:[shape('nested')]});
 const bytes=(await generatePptx(d)).bytes,old=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v9'}),current=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});
 assert.equal(old.document.slides[0].elements[0].style.fill.type,'none');const owner=current.document.layouts.find(l=>l.id===current.document.slides[0].layoutId).masterId;
 const shapes=[current.document.masters.find(m=>m.id===owner).elements[0],current.document.slides[0].elements[0],current.document.slides[0].elements[1].children[0]];for(const shape of shapes)assert.deepEqual(shape.style.fill,{type:'pattern',pattern:'cross',foreground:'AA0000',background:'0000FF',foregroundOpacity:.4,backgroundOpacity:.8});
 const traces=current.report.findings.filter(f=>f.feature.endsWith('pattern-dependencies'));assert.equal(traces.length,3);assert.ok(traces.every(f=>f.resolvedFrom.some(r=>r.path.includes('fgClr'))&&f.resolvedFrom.some(r=>r.path.includes('bgClr'))));const ready=await current.prepareSave({mode:'no-op'});assert.deepEqual((await current.save(ready.plan)).bytes,new Uint8Array(bytes));assert.equal((await current.prepareSave({mode:'edit'})).status,'blocked');old.dispose();current.dispose();
 const files=unzipSync(bytes);files['ppt/slides/slide1.xml']=strToU8(strFromU8(files['ppt/slides/slide1.xml']).replace('<a:pattFill prst="cross">','<a:pattFill prst="smCheck">'));const unsupported=await importPptx(zipSync(files),{projection:'content',masterArtwork:'artwork-v10'});assert.equal(unsupported.document.slides[0].elements.length,1);assert.ok(unsupported.report.findings.some(f=>f.tier==='unsupported'&&f.reason.includes('Shape pattern is outside')));unsupported.dispose();
 await assert.rejects(importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10',sessionLimits:{maxTextStyleDependencies:1}}),{code:'import-limit-exceeded'});
});

test('browser raster distinguishes horizontal, vertical and cross tiles inside shape bounds',async t=>{
 const d=createPresentation({size:{width:200,height:100}}),slide=d.addSlide();for(const [i,pattern] of ['horizontal','vertical','cross'].entries())slide.addElement({id:`p${i}`,type:'shape',shape:'rect',box:{x:10+i*60,y:20,width:48,height:48},style:{fill:{type:'pattern',pattern,foreground:'FF0000',background:'0000FF',foregroundOpacity:.5},stroke:{paint:{type:'none'}}}});
 const svg=(await renderPresentationToSvg(d)).slides[0].svg,browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();const result=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const c=document.createElement('canvas');c.width=200;c.height=100;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const color=(x,y)=>[...ctx.getImageData(x,y,1,1).data];const out=[0,1,2].map(i=>{const x=10+i*60;return {row:Array.from({length:24},(_,n)=>color(x+n,32)),column:Array.from({length:24},(_,n)=>color(x+12,20+n))}});URL.revokeObjectURL(url);return out},svg);
 const red=colors=>colors.map(c=>c[0]);const unique=values=>new Set(values).size;assert.equal(unique(red(result[0].row)),1);assert.ok(unique(red(result[0].column))>1);assert.ok(unique(red(result[1].row))>1);assert.equal(unique(red(result[1].column)),1);assert.ok(unique(red(result[2].row))>1&&unique(red(result[2].column))>1);
});

test('cross pattern applies foreground opacity once at intersections',async t=>{
 const d=createPresentation({size:{width:960,height:540}}),slide=d.addSlide();slide.addElement({id:'cross',type:'shape',shape:'rect',box:{x:180,y:120,width:600,height:300},style:{fill:{type:'pattern',pattern:'cross',foreground:'FF0000',background:'0000FF',foregroundOpacity:.5},stroke:{paint:{type:'none'}}}});
 const svg=(await renderPresentationToSvg(d)).slides[0].svg,browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const [intersection,line]=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const c=document.createElement('canvas');c.width=960;c.height=540;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const red=(x,y)=>ctx.getImageData(x,y,1,1).data[0];const result=[red(404,200),red(400,200)];URL.revokeObjectURL(url);return result},svg);
 assert.ok(line>=60&&line<=68,`line red channel ${line}`);assert.ok(intersection>=90&&intersection<=102,`intersection red channel ${intersection}`);
});

test('diagonal patterns export, import, and render in opposite directions',async t=>{
 const d=createPresentation({size:{width:160,height:80}}),slide=d.addSlide();for(const [i,pattern] of ['downwardDiagonal','upwardDiagonal'].entries())slide.addElement({id:`diagonal-${i}`,type:'shape',shape:'rect',box:{x:16+i*64,y:16,width:48,height:48},style:{fill:{type:'pattern',pattern,foreground:'FF0000',background:'0000FF',foregroundOpacity:.5},stroke:{paint:{type:'none'}}}});
 assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);const bytes=(await generatePptx(d)).bytes,xml=readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString('utf8');assert.match(xml,/prst="dnDiag"/);assert.match(xml,/prst="upDiag"/);
 const session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});assert.deepEqual(session.document.slides[0].elements.map(e=>e.style.fill.pattern),['downwardDiagonal','upwardDiagonal']);session.dispose();
 const rendered=await renderPresentationToSvg(d);assert.equal(rendered.warnings.filter(w=>w.code==='shape-pattern-approximation').length,2);const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const pixels=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const c=document.createElement('canvas');c.width=160;c.height=80;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const red=(x,y)=>ctx.getImageData(x,y,1,1).data[0];const result=[[red(26,26),red(26,29),red(34,26),red(26,34)],[red(90,26),red(90,29),red(98,26),red(90,34)]];URL.revokeObjectURL(url);return result},rendered.slides[0].svg);
 assert.ok(pixels[0][0]>pixels[0][1],`downward ${pixels[0]}`);assert.ok(pixels[1][1]>pixels[1][0],`upward ${pixels[1]}`);for(const values of pixels){assert.equal(values[0],values[2]);assert.equal(values[0],values[3]);}
});

test('diagonal cross is editable, round-trips and paints its crossing only once',async t=>{
 const d=createPresentation({size:{width:80,height:80}}),slide=d.addSlide();slide.addElement({id:'x',type:'shape',shape:'rect',box:{x:16,y:16,width:48,height:48},style:{fill:{type:'pattern',pattern:'diagonalCross',foreground:'FF0000',background:'0000FF',foregroundOpacity:.5},stroke:{paint:{type:'none'}}}});
 const bytes=(await generatePptx(d)).bytes;assert.match(readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString('utf8'),/prst="diagCross"/);const s=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});assert.equal(s.document.slides[0].elements[0].style.fill.pattern,'diagonalCross');s.dispose();
 const result=await renderPresentationToSvg(d);assert.ok(result.warnings.some(w=>w.code==='shape-pattern-approximation'));const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const reds=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const c=document.createElement('canvas');c.width=80;c.height=80;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const result=[];for(let y=24;y<32;y++)for(let x=24;x<32;x++)result.push(ctx.getImageData(x,y,1,1).data[0]);URL.revokeObjectURL(url);return result},result.slides[0].svg);
 assert.ok(Math.max(...reds)<=130,`foreground opacity compounded: ${Math.max(...reds)}`);assert.ok(Math.max(...reds)>70&&Math.min(...reds)===0);assert.ok(reds[2*8+2]>0&&reds[5*8+2]>0,'both diagonal directions must be visible');
});


test('percentage patterns round-trip all DrawingML presets and approximate their nominal coverage',async t=>{
 const values=[5,10,20,25,30,40,50,60,70,75,80,90],d=createPresentation({size:{width:960,height:80}}),slide=d.addSlide();
 for(const [i,value] of values.entries())slide.addElement({id:`pct${value}`,type:'shape',shape:'rect',box:{x:i*80,y:0,width:80,height:80},style:{fill:{type:'pattern',pattern:`percent${value}`,foreground:'FF0000',background:'0000FF'},stroke:{paint:{type:'none'}}}});
 assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);
 const bytes=(await generatePptx(d)).bytes,xml=readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString('utf8');
 for(const value of values)assert.match(xml,new RegExp(`prst="pct${value}"`));
 const session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});assert.deepEqual(session.document.slides[0].elements.map(e=>e.style.fill.pattern),values.map(value=>`percent${value}`));session.dispose();
 const result=await renderPresentationToSvg(d);assert.equal(result.warnings.filter(w=>w.code==='shape-pattern-approximation').length,values.length);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 const coverage=await page.evaluate(async svg=>{const img=new Image(),url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));await new Promise((ok,fail)=>{img.onload=ok;img.onerror=fail;img.src=url});const canvas=document.createElement('canvas');canvas.width=9600;canvas.height=800;const ctx=canvas.getContext('2d');ctx.drawImage(img,0,0,9600,800);const result=Array.from({length:12},(_,i)=>{const pixels=ctx.getImageData(i*800,0,800,800).data;let total=0;for(let j=0;j<pixels.length;j+=4)total+=pixels[j]/255;return total/(pixels.length/4)});URL.revokeObjectURL(url);return result},result.slides[0].svg);
 for(const [i,value] of values.entries())assert.ok(Math.abs(coverage[i]-value/100)<.025,`${value}% rendered ${coverage[i]}`);
});

test('percentage patterns preserve scoped colors and transparency through grouped Node and browser output',async t=>{
 const build=core=>{const d=core.createPresentation({id:'percent-group',size:{width:160,height:80},theme:{colors:{accent1:'AA0000',accent2:'0000AA'}}}),s=d.addSlide({id:'s'});s.addElement({id:'group',type:'group',box:{x:20,y:10,width:120,height:60},coordinateSize:{width:120,height:60},opacity:.5,transform:{rotation:15},children:[5,90].map((n,i)=>({id:`p${n}`,type:'shape',shape:'rect',box:{x:i*60,y:0,width:50,height:50},opacity:.5,style:{fill:{type:'pattern',pattern:`percent${n}`,foreground:{theme:'accent1'},background:{theme:'accent2'},foregroundOpacity:.4,backgroundOpacity:.8},stroke:{paint:{type:'none'}}}}))});return d};
 const deck=build({createPresentation});assert.equal(validatePresentation(deck).filter(x=>x.severity==='error').length,0);
 const node=await renderPresentationToSvg(deck),bytes=(await generatePptx(deck)).bytes,xml=readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString('utf8');
 for(const n of [5,90])assert.match(xml,new RegExp(`prst="pct${n}"`));
 assert.equal((xml.match(/<a:schemeClr val="accent1"><a:alpha val="10000"\/>/g)??[]).length,2);assert.equal((xml.match(/<a:schemeClr val="accent2"><a:alpha val="20000"\/>/g)??[]).length,2);
 const imported=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});assert.deepEqual(imported.document.slides[0].elements[0].children.map(e=>e.style.fill),[5,90].map(n=>({type:'pattern',pattern:`percent${n}`,foreground:'AA0000',background:'0000AA',foregroundOpacity:.1,backgroundOpacity:.2})));imported.dispose();
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),core=await readFile('packages/core/dist/browser/index.js'),layout=await readFile('packages/layout/dist/browser/index.js'),renderer=await readFile('packages/svg-renderer/dist/browser/index.js'),url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(core);
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl,'@pptkit/layout':url(layout)}})}</script>`);
 const browserOutput=await page.evaluate(async({coreUrl,rendererUrl,make})=>{const core=await import(coreUrl),renderer=await import(rendererUrl);return renderer.renderPresentationToSvg((0,eval)(`(${make})`)(core))},{coreUrl,rendererUrl:url(renderer),make:build.toString()});assert.deepEqual(browserOutput,node);
});

test('PowerPoint-saved percentage presets survive bounded import and three rebuilds',async()=>{
 const values=[5,10,20,25,30,40,50,60,70,75,80,90];let bytes=await readFile('tests/fixtures/percent-pattern-powerpoint-resave.pptx');
 for(let round=0;round<=3;round++){
  const session=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v10'});
  try{
   assert.deepEqual(session.document.slides[0].elements.map(e=>({name:e.name,fill:e.style.fill})),values.map(n=>({name:`percent${n}`,fill:{type:'pattern',pattern:`percent${n}`,foreground:'000000',background:'FFFFFF',foregroundOpacity:1,backgroundOpacity:1}})));
   assert.equal(session.report.findings.filter(f=>f.feature.endsWith('pattern-dependencies')).length,12);
   const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));
   if(round<3)bytes=(await generatePptx(session.document)).bytes;
  }finally{session.dispose();}
 }
});
