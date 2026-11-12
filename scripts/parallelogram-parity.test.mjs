import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { createPresentation, validatePresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
const { unzipSync, zipSync, strFromU8, strToU8 } = createRequire(new URL('../packages/pptx-parser/package.json', import.meta.url))('fflate');
const shape=(id,box)=>({id,type:'shape',shape:'parallelogram',box,style:{fill:{type:'solid',color:'FF0000'},stroke:{paint:{type:'none'},width:0}}});

test('content-sized parallelogram uses the same resolved frame in SVG and exported PPTX',async()=>{
 const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({...shape('target',{x:10,y:20,width:20,height:20}),slant:.4,text:{content:'Label',frame:{margin:5,wrap:false}}});
 const layout={intrinsicSizing:[{elementId:'target',width:{type:'content'},height:{type:'content'}}],textMeasurement:{measure:()=>({bounds:{x:0,y:0,width:30,height:20}})}};
 const resolved=resolveLayout(d,layout),box=resolved.slides[0].elements[0].box;
 assert.ok(Math.abs(box.width-480/7)<1e-8);assert.ok(Math.abs(box.height-360/7)<1e-8);
 const svg=await renderPresentationToSvg(d,{layout});const match=svg.slides[0].svg.match(/<foreignObject x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/);assert.ok(match);
 assert.ok(Math.abs(Number(match[3])-40)<1e-8);assert.ok(Math.abs(Number(match[4])-30)<1e-8);
 const exported=await generatePptx(d,{layout}),session=await importPptx(exported.bytes,{projection:'content'});
 try{const imported=session.document.slides[0].elements[0].box;assert.ok(Math.abs(imported.width-box.width)<1/12700);assert.ok(Math.abs(imported.height-box.height)<1/12700);}finally{session.dispose();}
});

test('wrapped parallelogram height is shared by layout, SVG and PPTX',async()=>{
 const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({...shape('target',{x:10,y:20,width:100,height:20}),slant:.4,text:{content:'Wrapped',frame:{margin:0,wrap:true}}});
 const layout={intrinsicSizing:[{elementId:'target',height:{type:'content'}}],textMeasurement:{measure:r=>({bounds:{x:0,y:0,width:10,height:r.availableWidth>=80?20:40}})}};
 const resolved=resolveLayout(d,layout),height=resolved.slides[0].elements[0].box.height;assert.ok(height>50);
 const svg=await renderPresentationToSvg(d,{layout}),match=svg.slides[0].svg.match(/<foreignObject x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/);assert.ok(match);assert.ok(Number(match[4])>=40-1e-8);
 const exported=await generatePptx(d,{layout}),session=await importPptx(exported.bytes,{projection:'content'});
 try{assert.ok(Math.abs(session.document.slides[0].elements[0].box.height-height)<1/12700);}finally{session.dispose();}
});

test('wrapped parallelogram with two content axes keeps geometry across SVG and PPTX',async()=>{
 const d=createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({...shape('target',{x:10,y:20,width:20,height:20}),slant:.4,text:{content:'Wrapped',frame:{margin:0,wrap:true}}});
 const layout={intrinsicSizing:[{elementId:'target',width:{type:'content',max:80},height:{type:'content'}}],textMeasurement:{measure:r=>({bounds:r.wrap?{x:0,y:0,width:Math.min(30,r.availableWidth),height:r.availableWidth>=63?20:40}:{x:0,y:0,width:100,height:20}})}};
 const box=resolveLayout(d,layout).slides[0].elements[0].box;assert.equal(box.width,80);assert.ok(box.height>=80-1e-8);
 const svg=await renderPresentationToSvg(d,{layout}),match=svg.slides[0].svg.match(/<foreignObject x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/);assert.ok(match);assert.ok(Number(match[3])>=30);assert.ok(Number(match[4])>=40-1e-8);
 const exported=await generatePptx(d,{layout}),session=await importPptx(exported.bytes,{projection:'content'});
 try{const imported=session.document.slides[0].elements[0].box;assert.ok(Math.abs(imported.width-box.width)<1/12700);assert.ok(Math.abs(imported.height-box.height)<1/12700);}finally{session.dispose();}
});

test('wrapped two-axis sizing agrees between Node and browser layout bundles',async t=>{
 const {readFile}=await import('node:fs/promises'),{chromium}=await import('@playwright/test');
 const make=core=>{const d=core.createPresentation({size:{width:400,height:300}});d.addSlide({id:'s'}).addElement({id:'p',type:'shape',shape:'parallelogram',slant:.4,box:{x:10,y:20,width:20,height:20},text:{content:'Wrapped',frame:{margin:0,wrap:true}}});return d};
 const options=()=>({intrinsicSizing:[{elementId:'p',width:{type:'content',max:80},height:{type:'content'}}],textMeasurement:{measure:r=>({bounds:r.wrap?{x:0,y:0,width:Math.min(30,r.availableWidth),height:r.availableWidth>=63?20:40}:{x:0,y:0,width:100,height:20}})}});
 const expected=resolveLayout(make({createPresentation}),options()),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(await readFile('packages/core/dist/browser/index.js')),layoutUrl=url(await readFile('packages/layout/dist/browser/index.js'));
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,layoutUrl,makeSource,optionsSource})=>{const core=await import(coreUrl),layout=await import(layoutUrl);return layout.resolveLayout((0,eval)(`(${makeSource})`)(core),(0,eval)(`(${optionsSource})`)())},{coreUrl,layoutUrl,makeSource:make.toString(),optionsSource:options.toString()});
 assert.deepEqual(actual,expected);
});

test('default parallelogram uses shorter-side inset in SVG and native OOXML through three fresh rounds',async()=>{
 const d=createPresentation({size:{width:300,height:200}}),slide=d.addSlide({id:'s'});
 slide.addElement(shape('wide',{x:10,y:10,width:100,height:40}));slide.addElement({...shape('tall',{x:150,y:10,width:40,height:100}),transform:{rotation:30,flipH:true}});
 slide.addElement({...shape('text',{x:10,y:120,width:100,height:40}),text:{content:'Label'}});
 const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/points="10,50 20,10 110,10 100,50"/);assert.match(svg.slides[0].svg,/points="150,110 160,10 190,10 180,110"/);assert.ok(svg.warnings.some(w=>w.code==='shape-text-frame-approximation'&&w.elementId==='text'));
 let bytes=(await generatePptx(d)).bytes;
 for(let i=0;i<=3;i++){
  const xml=readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();assert.equal((xml.match(/prst="parallelogram"/g)??[]).length,3);assert.match(xml,/rot="1800000" flipH="1"/);
  const s=await importPptx(bytes,{projection:'content'});try{assert.deepEqual(s.document.slides[0].elements.map(e=>e.shape),['parallelogram','parallelogram','parallelogram']);const plan=await s.prepareSave({mode:'no-op'});assert.deepEqual((await s.save(plan.plan)).bytes,new Uint8Array(bytes));if(i<3)bytes=(await generatePptx(s.document)).bytes;}finally{s.dispose();}
 }
});

test('out-of-profile parallelogram stays source-only, and unsupported native connector attachment falls back',async()=>{
 const d=createPresentation(),slide=d.addSlide({id:'s'});slide.addElement(shape('target',{x:20,y:20,width:100,height:40}));slide.addElement({id:'edge',type:'connector',start:{elementId:'target',anchor:'right'},end:{x:200,y:40}});
 assert.equal(resolveLayout(d).slides[0].elements.find(e=>e.id==='edge').start.x,120);
 const generated=await generatePptx(d,{connectorMode:'native'});assert.ok(generated.warnings.some(w=>w.code==='native-connector-fallback'&&w.elementId==='edge'));
 const files=unzipSync(generated.bytes),path='ppt/slides/slide1.xml',xml=strFromU8(files[path]);assert.ok(xml.includes('<a:prstGeom prst="parallelogram"><a:avLst/></a:prstGeom>'));
 files[path]=strToU8(xml.replace('<a:prstGeom prst="parallelogram"><a:avLst/></a:prstGeom>','<a:prstGeom prst="parallelogram"><a:avLst><a:gd name="adj" fmla="val 120000"/></a:avLst></a:prstGeom>'));
 const source=zipSync(files),s=await importPptx(source,{projection:'content'});try{assert.ok(!s.document.slides[0].elements.some(e=>e.shape==='parallelogram'));assert.ok(s.report.findings.some(f=>f.reason.includes('Parallelogram adjustments')));const plan=await s.prepareSave({mode:'no-op'});assert.deepEqual((await s.save(plan.plan)).bytes,new Uint8Array(source));}finally{s.dispose();}
 const bad=createPresentation(),page=bad.addSlide();page.addElement({...shape('bad',{x:0,y:0,width:10,height:10}),shape:'unknown'});assert.ok(validatePresentation(bad).some(f=>f.code==='invalid-shape-kind'));
});

test('shared unadjusted parallelograms project, while out-of-profile shared geometry remains source-only',async()=>{
 const d=createPresentation();d.defineSlideMaster({id:'m',name:'M',elements:[shape('master',{x:0,y:0,width:20,height:80})]});d.defineSlideLayout({id:'l',name:'L',masterId:'m',elements:[shape('layout',{x:0,y:0,width:80,height:20})]});d.addSlide({id:'s',layoutId:'l'});
 const bytes=(await generatePptx(d)).bytes,s=await importPptx(bytes,{projection:'content',masterArtwork:'artwork-v1'});
 try{const l=s.document.layouts.find(x=>x.id===s.document.slides[0].layoutId),m=s.document.masters.find(x=>x.id===l.masterId);assert.equal(l.elements[0].shape,'parallelogram');assert.equal(m.elements[0].shape,'parallelogram');}finally{s.dispose();}
 const parts=unzipSync(bytes);for(const path of ['ppt/slideMasters/slideMaster2.xml','ppt/slideLayouts/slideLayout2.xml']){const xml=strFromU8(parts[path]);assert.ok(xml.includes('prst="parallelogram"'));parts[path]=strToU8(xml.replace('<a:prstGeom prst="parallelogram"><a:avLst/></a:prstGeom>','<a:prstGeom prst="parallelogram"><a:avLst><a:gd name="adj" fmla="val 120000"/></a:avLst></a:prstGeom>'));}
 const source=zipSync(parts),adjusted=await importPptx(source,{projection:'content',masterArtwork:'artwork-v1'});try{const l=adjusted.document.layouts.find(x=>x.id===adjusted.document.slides[0].layoutId),m=adjusted.document.masters.find(x=>x.id===l.masterId);assert.equal(l.elements.length,0);assert.equal(m.elements.length,0);assert.ok(adjusted.report.findings.some(f=>f.reason.includes('adjustments are outside')));}finally{adjusted.dispose();}
});

test('parallelogram renderer output is consistent across Node and browser bundles',async t=>{
 const { readFile }=await import('node:fs/promises'),{ chromium }=await import('@playwright/test');
 const make=core=>{const d=core.createPresentation({id:'parallel',size:{width:200,height:100}}),s=d.addSlide({id:'s'});s.addElement({id:'shape',type:'shape',shape:'parallelogram',box:{x:10,y:10,width:150,height:50},transform:{rotation:15,flipV:true},slant:.4,text:{content:'Browser parity'}});return d};
 const expected=await renderPresentationToSvg(make({createPresentation})),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(await readFile('packages/core/dist/browser/index.js')),layoutUrl=url(await readFile('packages/layout/dist/browser/index.js')),rendererUrl=url(await readFile('packages/svg-renderer/dist/browser/index.js'));
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl,'@pptkit/layout':layoutUrl}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,rendererUrl,source})=>{const core=await import(coreUrl),renderer=await import(rendererUrl);return renderer.renderPresentationToSvg((0,eval)(`(${source})`)(core))},{coreUrl,rendererUrl,source:make.toString()});assert.deepEqual(actual,expected);
});

test('PowerPoint-saved defaults import with their transforms and preserve original bytes',async()=>{
 const {readFile}=await import('node:fs/promises');const bytes=await readFile('tests/fixtures/parallelogram-powerpoint-resave.pptx'),s=await importPptx(bytes,{projection:'content'});
 try{const shapes=s.document.slides[0].elements;assert.deepEqual(shapes.map(e=>[e.name,e.shape,e.transform.rotation,e.transform.flipH]),[['wide','parallelogram',0,false],['tall','parallelogram',0,false],['rotated','parallelogram',30,true]]);const plan=await s.prepareSave({mode:'no-op'});assert.deepEqual((await s.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{s.dispose();}
});

test('explicit slant supports bounded geometry, quantization and local/shared literal import',async()=>{
 const d=createPresentation({size:{width:300,height:200}}),base=shape('x',{x:0,y:0,width:100,height:40});
 d.defineSlideMaster({id:'m',name:'M',elements:[{...base,id:'master',slant:.4}]});d.defineSlideLayout({id:'l',name:'L',masterId:'m',elements:[{...base,id:'layout',slant:1}]});const slide=d.addSlide({id:'s',layoutId:'l'});
 for(const [i,slant] of [0,.4,1,.333333].entries())slide.addElement({...shape(`s${i}`,{x:i*70,y:80,width:60,height:40}),slant});
 const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/points="0,120 0,80 60,80 60,120"/);assert.match(svg.slides[0].svg,/points="70,120 86,80 130,80 114,120"/);assert.match(svg.slides[0].svg,/points="140,120 180,80 200,80 160,120"/);
 const result=await generatePptx(d);assert.deepEqual(result.warnings.filter(w=>w.code==='shape-slant-quantized').map(w=>w.elementId),['s3']);const parts=readGeneratedZip(result.bytes);assert.match(parts.get('ppt/slides/slide1.xml').toString(),/name="adj" fmla="val 33333"/);
 const session=await importPptx(result.bytes,{projection:'content',masterArtwork:'artwork-v1'});try{assert.deepEqual(session.document.slides[0].elements.map(e=>e.slant),[0,.4,1,.33333]);const l=session.document.layouts.find(x=>x.id===session.document.slides[0].layoutId),m=session.document.masters.find(x=>x.id===l.masterId);assert.equal(l.elements[0].slant,1);assert.equal(m.elements[0].slant,.4);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(result.bytes));}finally{session.dispose();}
 const invalid=createPresentation(),p=invalid.addSlide();for(const [i,slant] of [-.01,1.01,NaN,Infinity].entries())p.addElement({...base,id:`bad${i}`,slant});p.addElement({...base,id:'wrong',shape:'rect',slant:.4});assert.equal(validatePresentation(invalid).filter(x=>x.code==='invalid-shape-slant').length,5);
});

test('unsupported slant formulas and extra adjustment content never project as defaults',async()=>{
 const d=createPresentation(),s=d.addSlide();s.addElement({...shape('target',{x:0,y:0,width:100,height:40}),slant:.4});const source=(await generatePptx(d)).bytes;
 const original='<a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst>';
 for(const replacement of [
  '<a:avLst><a:gd name="adj" fmla="*/ 1 2 3"/></a:avLst>',
  '<a:avLst><a:gd name="adj" fmla="val 100001"/></a:avLst>',
  '<a:avLst><a:gd name="other" fmla="val 40000"/></a:avLst>',
  '<a:avLst><a:gd name="adj" fmla="val 40000"/><a:gd name="other" fmla="val 1"/></a:avLst>',
  '<a:avLst><a:gd name="adj" fmla="val 40000" extra="x"/></a:avLst>',
  '<a:avLst><a:gd name="adj" fmla="val 40000"><a:gd name="x" fmla="val 1"/></a:gd></a:avLst>',
 ]){
  const files=unzipSync(source),path='ppt/slides/slide1.xml',xml=strFromU8(files[path]);assert.ok(xml.includes(original));files[path]=strToU8(xml.replace(original,replacement));const bytes=zipSync(files),session=await importPptx(bytes,{projection:'content'});
  try{assert.equal(session.document.slides[0].elements.length,0);assert.ok(session.report.findings.some(f=>f.reason.includes('Parallelogram adjustments')));const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{session.dispose();}
 }
});

test('source-bound slant edits are blocked by preservation save',async()=>{
 const d=createPresentation(),slide=d.addSlide({id:'s'});slide.addElement({...shape('target',{x:0,y:0,width:100,height:40}),slant:.4});const bytes=(await generatePptx(d)).bytes;
 const session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v1'});try{const target=session.document.slides[0].elements[0];session.document.slides[0].replaceElement(target.id,{...target,slant:.6});assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');assert.equal((await session.prepareSave({mode:'no-op'})).status,'blocked');}finally{session.dispose();}
});

test('PowerPoint-saved explicit slants import and preserve their source bytes',async()=>{
 const {readFile}=await import('node:fs/promises');const bytes=await readFile('tests/fixtures/parallelogram-slant-powerpoint-resave.pptx'),session=await importPptx(bytes,{projection:'content'});
 try{assert.deepEqual(session.document.slides[0].elements.map(e=>[e.name,e.shape,e.slant,e.transform.rotation,e.transform.flipH]),[['zero','parallelogram',0,0,false],['mid','parallelogram',.4,0,false],['max-rotated','parallelogram',1,30,true]]);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{session.dispose();}
});

test('literal slant tolerates XML whitespace, signed values and leading zeroes while retaining strict bounds',async()=>{
 const d=createPresentation(),slide=d.addSlide();slide.addElement({...shape('target',{x:0,y:0,width:100,height:40}),slant:.4});const source=(await generatePptx(d)).bytes;
 for(const [formula,expected] of [[' val  +040000 ',.4],['val&#9;000000',0],['val -0',0],['val +100000',1],['val -1',undefined],['val +100001',undefined],['val 1.0',undefined],['val 99999999999999999999999999',undefined],['&#160;val 40000',undefined]]){
  const files=unzipSync(source),path='ppt/slides/slide1.xml',xml=strFromU8(files[path]);files[path]=strToU8(xml.replace('fmla="val 40000"',`fmla="${formula}"`));const bytes=zipSync(files),session=await importPptx(bytes,{projection:'content'});
  try{if(expected===undefined){assert.equal(session.document.slides[0].elements.length,0);assert.ok(session.report.findings.some(f=>f.reason.includes('Parallelogram adjustments')));}else{assert.equal(session.document.slides[0].elements[0].slant,expected);if(expected===0)assert.ok(!Object.is(session.document.slides[0].elements[0].slant,-0));}const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{session.dispose();}
 }
});

test('literal slant lexical projection agrees in Node and both browser entries',async t=>{
 const {readFile}=await import('node:fs/promises'),{chromium}=await import('@playwright/test');const d=createPresentation(),slide=d.addSlide();slide.addElement({...shape('target',{x:0,y:0,width:100,height:40}),slant:.4});const source=(await generatePptx(d)).bytes;
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort()});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});await page.evaluate(async code=>{globalThis.parser=await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(code)))}`)},await readFile('packages/pptx-parser/dist/browser/index.js','utf8'));
 for(const formula of [' val  +040000 ','val&#9;000000','val -0','val -1','&#160;val 40000']){
  const files=unzipSync(source),path='ppt/slides/slide1.xml';files[path]=strToU8(strFromU8(files[path]).replace('fmla="val 40000"',`fmla="${formula}"`));const bytes=zipSync(files),node=await importPptx(bytes,{projection:'content'});let expected;try{expected=node.document.slides[0].elements.map(e=>({shape:e.shape,slant:e.slant}));}finally{node.dispose();}
  const actual=await page.evaluate(async bytes=>{const out=[];for(const api of [globalThis.parser,globalThis.PPTKit]){const s=await api.importPptx(Uint8Array.from(bytes),{projection:'content'});try{out.push(s.document.slides[0].elements.map(e=>({shape:e.shape,slant:e.slant})));}finally{s.dispose();}}return out},[...bytes]);for(const result of actual)assert.deepEqual(result,expected);
 }
 assert.deepEqual(requests,[]);
});

test('native inherited slant source and PowerPoint resave project both owner levels and preserve bytes',async()=>{
 const {readFile}=await import('node:fs/promises');for(const name of ['source','resave']){const bytes=await readFile(`tests/fixtures/inherited-slant-powerpoint-${name}.pptx`),session=await importPptx(bytes,{projection:'content'});
  try{assert.deepEqual(session.document.slides[0].elements.map(e=>[e.name,e.shape,e.slant]),[['zero','parallelogram',.4],['mid','parallelogram',.8],['max-rotated','parallelogram',1]]);assert.deepEqual(session.report.findings.filter(f=>f.feature==='drawing.placeholder-slant').map(f=>f.source.part),['/ppt/slideLayouts/slideLayout1.xml','/ppt/slideMasters/slideMaster1.xml']);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{session.dispose();}}
});

test('inherited native slants and provenance agree across Node and browser import entries',async t=>{
 const {readFile}=await import('node:fs/promises'),{chromium}=await import('@playwright/test');const bytes=await readFile('tests/fixtures/inherited-slant-powerpoint-resave.pptx'),node=await importPptx(bytes,{projection:'content'});let expected;
 try{expected={shapes:node.document.slides[0].elements.map(e=>[e.name,e.slant]),sources:node.report.findings.filter(f=>f.feature==='drawing.placeholder-slant').map(f=>f.source)};}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort()});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});await page.evaluate(async code=>{globalThis.parser=await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(code)))}`)},await readFile('packages/pptx-parser/dist/browser/index.js','utf8'));
 const actual=await page.evaluate(async bytes=>{const result=[];for(const api of [globalThis.parser,globalThis.PPTKit]){const s=await api.importPptx(Uint8Array.from(bytes),{projection:'content'});try{result.push({shapes:s.document.slides[0].elements.map(e=>[e.name,e.slant]),sources:s.report.findings.filter(f=>f.feature==='drawing.placeholder-slant').map(f=>f.source)});}finally{s.dispose();}}return result},[...bytes]);for(const entry of actual)assert.deepEqual(entry,expected);assert.deepEqual(requests,[]);
});

test('inherited slant changes do not gain source preservation write access',async()=>{
 const {readFile}=await import('node:fs/promises'),bytes=await readFile('tests/fixtures/inherited-slant-powerpoint-resave.pptx'),session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v1'});
 try{const slide=session.document.slides[0],target=slide.elements[0];slide.replaceElement(target.id,{...target,slant:.5});assert.equal((await session.prepareSave({mode:'edit'})).status,'blocked');assert.equal((await session.prepareSave({mode:'no-op'})).status,'blocked');}finally{session.dispose();}
});

test('a mixed text and slant edit is atomic for local and inherited geometries',async()=>{
 const {readFile}=await import('node:fs/promises'),{body}=await import('../packages/pptx-parser/test/content-fixture.mjs');
 const local=createPresentation(), slide=local.addSlide();slide.addElement({...shape('local',{x:0,y:0,width:100,height:40}),slant:.4,text:{content:'Original'}});
 const inherited=unzipSync(await readFile('tests/fixtures/inherited-slant-powerpoint-source.pptx')),path='ppt/slides/slide1.xml',xml=strFromU8(inherited[path]),position=xml.indexOf('</p:spPr></p:sp>');assert.ok(position>=0);
 inherited[path]=strToU8(xml.slice(0,position)+xml.slice(position).replace('</p:spPr></p:sp>',`</p:spPr>${body()}</p:sp>`));
 for(const [label,bytes] of [['local',(await generatePptx(local)).bytes],['inherited',zipSync(inherited)]]){
  const original=unzipSync(bytes),s=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v2'});
  try{
   const target=s.document.slides[0].elements.find(e=>label==='local'||e.name==='zero');assert.equal(target.slant,.4);
   const changed=structuredClone(target);changed.text.content[0].runs[0].text='Permitted text';changed.slant=.6;s.document.slides[0].replaceElement(target.id,changed);
   assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked',label);
   assert.equal((await s.prepareSave({mode:'no-op'})).status,'blocked',label);
   changed.slant=.4;s.document.slides[0].replaceElement(target.id,changed);
   const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));
   const saved=(await s.save(plan.plan)).bytes,parts=unzipSync(saved);
   assert.deepEqual(Object.keys(parts).sort(),Object.keys(original).sort());
   for(const part of Object.keys(original))if(part!==path)assert.deepEqual(parts[part],original[part],`${label}: ${part}`);
   const reopened=await importPptx(saved,{projection:'content'});try{const shape=reopened.document.slides[0].elements.find(e=>label==='local'||e.name==='zero');assert.equal(shape.slant,.4);assert.equal(shape.text.content[0].runs[0].text,'Permitted text');}finally{reopened.dispose();}
  }finally{s.dispose();}
 }
});

test('shape-text preservation edits retain inherited slant and shared source parts through three rounds',async()=>{
 const {readFile}=await import('node:fs/promises'),{body}=await import('../packages/pptx-parser/test/content-fixture.mjs');const source=await readFile('tests/fixtures/inherited-slant-powerpoint-source.pptx'),parts=unzipSync(source),path='ppt/slides/slide1.xml';
 const original=strFromU8(parts[path]),position=original.indexOf('</p:spPr></p:sp>');assert.ok(position>=0);parts[path]=strToU8(original.slice(0,position)+original.slice(position).replace('</p:spPr></p:sp>',`</p:spPr>${body()}</p:sp>`));let bytes=zipSync(parts);
 const initial=unzipSync(bytes),layout=initial['ppt/slideLayouts/slideLayout1.xml'],master=initial['ppt/slideMasters/slideMaster1.xml'];
 for(let round=1;round<=3;round++){
  const session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v2'});
  try{
   const slide=session.document.slides[0],target=slide.elements.find(e=>e.name==='zero');assert.equal(target.slant,.4);assert.ok(session.report.textRuns.some(r=>r.elementId===target.id&&r.owner==='shape-text'&&r.status==='bound'));
   const changed=structuredClone(target);changed.text.content[0].runs[0].text=`Edited ${round}`;slide.replaceElement(target.id,changed);
   const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await session.save(plan.plan)).bytes;
   const saved=unzipSync(bytes);assert.deepEqual(saved['ppt/slideLayouts/slideLayout1.xml'],layout);assert.deepEqual(saved['ppt/slideMasters/slideMaster1.xml'],master);assert.deepEqual(Object.keys(saved).sort(),Object.keys(initial).sort());for(const name of Object.keys(initial))if(name!==path)assert.deepEqual(saved[name],initial[name],`unexpected change to ${name}`);
  }finally{session.dispose();}
 }
 const final=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v2'});try{const shape=final.document.slides[0].elements.find(e=>e.name==='zero');assert.equal(shape.slant,.4);assert.equal(shape.text.content[0].runs[0].text,'Edited 3');assert.equal(final.report.findings.find(f=>f.feature==='drawing.placeholder-slant'&&f.elementId===shape.id).source.part,'/ppt/slideLayouts/slideLayout1.xml');}finally{final.dispose();}
});

test('parallelogram SVG text frames use the preset geometry for varied proportions and slants',async()=>{
 for(const [width,height,slant,expected] of [[200,100,undefined,[37.083333333333336,33.541666666666664,145.83333333333334,72.91666666666667]],[100,200,.4,[35,70,50,100]],[100,100,0,[18.333333333333332,28.333333333333332,83.33333333333334,83.33333333333334]],[100,100,1,[60,70,0,0]],[.001,100,.75,[10.000395833333333,59.58333333333333,.00020833333333333335,20.833333333333336]]]){
  const deck=createPresentation({size:{width:400,height:300}}),slide=deck.addSlide({id:'s'}),input={...shape('label',{x:10,y:20,width,height}),...(slant===undefined?{}:{slant}),text:{content:'Frame'}};slide.addElement(input);
  const result=await renderPresentationToSvg(deck),match=result.slides[0].svg.match(/<foreignObject x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/);assert.ok(match,result.slides[0].svg);
  for(let i=0;i<4;i++)assert.ok(Math.abs(Number(match[i+1])-expected[i])<1e-9,`${width}x${height} slant ${slant} index ${i}`);
  assert.ok(result.warnings.some(w=>w.code==='shape-text-frame-approximation'&&w.elementId==='label'));
 }
});
