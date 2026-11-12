import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation,normalizePresentation,resolvePresentationInheritance,validatePresentation} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderPresentationToSvg} from '../packages/svg-renderer/dist/index.js';
import {importPptx} from '../packages/pptx-parser/dist/index.js';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
import {chromium} from '@playwright/test';
import {createRequire} from 'node:module';
import {readFile} from 'node:fs/promises';
const {unzipSync,zipSync,strFromU8,strToU8}=createRequire(new URL('../packages/pptx-parser/package.json',import.meta.url))('fflate');
const directions=[undefined,'rotate90','rotate270','horizontal'];
function deck(){const d=createPresentation({size:{width:400,height:200}});d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:10,y:20,width:320,height:100},columns:[80,80,80,80],rows:[{height:100,cells:directions.map((direction,i)=>({content:`Cell ${i}`,...(direction?{style:{textDirection:direction}}:{})}))}]});return d;}
test('table directions export native attributes, preview with warnings, and survive three import/export rounds',async()=>{
 let d=deck();
 for(let round=0;round<3;round++){
  const generated=await generatePptx(d),xml=readGeneratedZip(generated.bytes).get('ppt/slides/slide1.xml').toString();
  assert.deepEqual([...xml.matchAll(/<a:tcPr\b[^>]*>/g)].map(m=>m[0].match(/\bvert="([^"]+)"/)?.[1]),[undefined,'vert','vert270','horz']);
  const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/writing-mode:sideways-rl/);assert.match(svg.slides[0].svg,/writing-mode:sideways-lr/);assert.equal(svg.warnings.filter(w=>w.code==='table-text-direction-approximation').length,2);
  const session=await importPptx(generated.bytes,{projection:'content'});try{assert.deepEqual(session.document.slides[0].elements[0].rows[0].cells.map(c=>c.style?.textDirection),directions);const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual((await session.save(plan.plan)).bytes,generated.bytes);d=session.document;}finally{session.dispose();}
 }
});
test('rotated cells use approximate transformed measurement and verified row sizing',()=>{
 const d=deck(),provider={measure:()=>({bounds:{x:0,y:0,width:5,height:5}})},result=resolveLayout(d,{textMeasurement:provider,diagnostics:{textOverflow:true},tableSizing:[{elementId:'t',autoRows:[{row:0}]}]});
 assert.equal(result.analysis.diagnostics.filter(x=>x.code==='table-text-direction-sizing-approximate').length,2);
 assert.equal(result.analysis.diagnostics.filter(x=>x.code==='table-text-direction-measurement-approximate').length,2);
 assert.equal(result.analysis.measurements.length,4);assert.ok(result.slides[0].elements[0].rows[0].height>0);
});
test('regional direction cascade, explicit horizontal override and invalid inputs',()=>{
 const d=deck(),slide=d.slides[0],t=structuredClone(slide.elements[0]);t.style={wholeTable:{textDirection:'rotate90'},lastColumn:{textDirection:'rotate270'}};t.rows[0].cells[1].style={textDirection:'horizontal'};delete t.rows[0].cells[2].style;delete t.rows[0].cells[3].style;slide.replaceElement(t.id,t);
 assert.deepEqual(normalizePresentation(d).slides[0].elements[0].rows[0].cells.map(c=>c.style.textDirection),['rotate90','horizontal','rotate90','rotate270']);
 const bad=structuredClone(slide.elements[0]);bad.rows[0].cells[1].style.textDirection='upside';slide.replaceElement(bad.id,bad);assert.ok(validatePresentation(d).some(x=>x.code==='invalid-table-text-direction'));
});

test('Chromium rotates preview glyph runs in both directions',async t=>{
 const svg=(await renderPresentationToSvg(deck())).slides[0].svg,browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage({viewport:{width:400,height:200},deviceScaleFactor:1});await page.setContent(`<style>html,body{margin:0}</style>${svg}`);
 const samples=await page.evaluate(()=>[...document.querySelectorAll('[role="cell"]')].map(td=>{const node=td.querySelector('span');const range=document.createRange();range.selectNodeContents(node);const box=range.getBoundingClientRect(),wrapper=td.querySelector('[style*="writing-mode"]');return {width:box.width,height:box.height,mode:wrapper?getComputedStyle(wrapper).writingMode:'horizontal-tb'};}));
 assert.equal(samples[0].mode,'horizontal-tb');assert.equal(samples[1].mode,'sideways-rl');assert.equal(samples[2].mode,'sideways-lr');assert.equal(samples[3].mode,'horizontal-tb');
 assert.ok(samples[0].width>samples[0].height);assert.ok(samples[1].height>samples[1].width);assert.ok(samples[2].height>samples[2].width);
 const screenshot=await page.screenshot(),pixels=await page.evaluate(async base64=>{const img=new Image();img.src=`data:image/png;base64,${base64}`;await img.decode();const c=document.createElement('canvas');c.width=img.width;c.height=img.height;const x=c.getContext('2d');x.drawImage(img,0,0);const data=x.getImageData(0,0,img.width,img.height).data;return [0,1,2,3].map(i=>{let n=0;for(let y=25;y<115;y++)for(let x=15+80*i;x<85+80*i;x++){const p=(y*img.width+x)*4;if(data[p]<80&&data[p+1]<80&&data[p+2]<80)n++;}return n;});},screenshot.toString('base64'));
 assert.ok(pixels.every(n=>n>10),`text pixels missing: ${pixels}`);
});

test('resolved border-grid cells retain rotation in preview',async()=>{
 const d=deck(),slide=d.slides[0],table=structuredClone(slide.elements[0]);table.rows[0].cells[1].style.borders={left:{width:1}};slide.replaceElement(table.id,table);
 const svg=await renderPresentationToSvg(d);assert.match(svg.slides[0].svg,/role="cell"/);assert.match(svg.slides[0].svg,/writing-mode:sideways-rl/);assert.match(svg.slides[0].svg,/writing-mode:sideways-lr/);assert.equal(svg.warnings.filter(w=>w.code==='table-text-direction-approximation').length,2);
});

test('rotated cell frames use grid geometry and diagnose mismatched table frames',async()=>{
 const d=deck(),slide=d.slides[0],table=structuredClone(slide.elements[0]);table.box.height=80;slide.replaceElement(table.id,table);
 const layout=resolveLayout(d);assert.ok(layout.slides[0].elements[0].cellLayout);assert.ok(layout.analysis.diagnostics.some(x=>x.code==='table-text-direction-frame-grid-mismatch'));
 const svg=await renderPresentationToSvg(d);assert.ok(svg.warnings.some(x=>x.code==='table-text-direction-frame-grid-mismatch'));
 const ordinary=deck(),normalSlide=ordinary.slides[0],plain=structuredClone(normalSlide.elements[0]);for(const cell of plain.rows[0].cells)delete cell.style;normalSlide.replaceElement(plain.id,plain);assert.equal(resolveLayout(ordinary).slides[0].elements[0].cellLayout,undefined);
});

test('direction origin reports agree in Node, browser ESM and global bundles',async t=>{
 const make=api=>{const d=api.createPresentation();d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:0,y:0,width:60,height:30},columns:[30,30],rows:[{height:30,cells:[{content:'A'},{content:'B',style:{textDirection:'horizontal'}}]}],style:{wholeTable:{textDirection:'rotate90'}}});return api.resolvePresentationInheritance(d).origins.filter(o=>o.target.id==='t'&&o.property.endsWith('.textDirection'));};
 const expected=make({createPresentation,resolvePresentationInheritance});assert.equal(expected.length,2);const browser=await chromium.launch();t.after(()=>browser.close());
 const esm=`data:text/javascript;base64,${(await readFile('packages/core/dist/browser/index.js')).toString('base64')}`,global=await readFile('packages/core/dist/browser/global.js','utf8');
 for(const format of ['esm','global']){const page=await browser.newPage();if(format==='global')await page.addScriptTag({content:global});const actual=await page.evaluate(async({esm,source})=>{const api=esm?await import(esm):PPTKit;return (0,eval)(`(${source})`)(api)}, {esm:format==='esm'?esm:null,source:make.toString()});assert.deepEqual(actual,expected);await page.close();}
});

test('compact rotated row sizing agrees between Node and browser layout bundles',async t=>{
 const make=api=>{const d=api.createPresentation({size:{width:100,height:100}});d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:0,y:0,width:50,height:20},columns:[50],rows:[{height:20,cells:[{content:'R',style:{textDirection:'rotate90',margin:5}}]}]});return d;};
 const options=()=>({textMeasurement:{measure:r=>({bounds:!r.wrap?{x:0,y:0,width:100,height:10}:{x:0,y:0,width:30,height:r.availableWidth>=30?20:50}})},tableSizing:[{elementId:'t',autoRows:[{row:0}],rotatedTextSizing:'compact'}]});
 const expected=resolveLayout(make({createPresentation}),options()),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),url=b=>`data:text/javascript;base64,${b.toString('base64')}`,coreUrl=url(await readFile('packages/core/dist/browser/index.js')),layoutUrl=url(await readFile('packages/layout/dist/browser/index.js'));
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':coreUrl}})}</script>`);
 const actual=await page.evaluate(async({coreUrl,layoutUrl,makeSource,optionsSource})=>{const core=await import(coreUrl),layout=await import(layoutUrl);return layout.resolveLayout((0,eval)(`(${makeSource})`)(core),(0,eval)(`(${optionsSource})`)())},{coreUrl,layoutUrl,makeSource:make.toString(),optionsSource:options.toString()});assert.deepEqual(actual,expected);
});

test('unsupported native direction remains source-owned and survives no-op bytes',async()=>{
 const generated=await generatePptx(deck()),parts=unzipSync(generated.bytes),path='ppt/slides/slide1.xml',xml=strFromU8(parts[path]);assert.ok(xml.includes('vert="vert270"'));parts[path]=strToU8(xml.replace('vert="vert270"','vert="futureVert"'));const bytes=zipSync(parts);
 const session=await importPptx(bytes,{projection:'content'});try{assert.notEqual(session.document.slides[0].elements[0].rows[0].cells[2].style?.textDirection,'rotate270');assert.ok(session.report.findings.some(f=>f.source?.part==='/ppt/slides/slide1.xml'&&JSON.stringify(f.source).includes('vert')),JSON.stringify(session.report.findings.filter(f=>f.source?.part==='/ppt/slides/slide1.xml'&&f.tier==='unsupported').slice(-10)));const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual((await session.save(plan.plan)).bytes,bytes);}finally{session.dispose();}
});

test('text edit does not erase an unsupported native cell direction',async()=>{
 const generated=await generatePptx(deck()),parts=unzipSync(generated.bytes),path='ppt/slides/slide1.xml';
 parts[path]=strToU8(strFromU8(parts[path]).replace('vert="vert270"','vert="futureVert"'));
 const bytes=zipSync(parts),session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v3'});
 try{
  const slide=session.document.slides[0],changed=structuredClone(slide.elements[0]);changed.rows[0].cells[2].content[0].runs[0].text='Preserved';slide.replaceElement(changed.id,changed);
  const plan=await session.prepareSave({mode:'edit'});
  assert.equal(plan.status,'ready',JSON.stringify(plan.report));const saved=await session.save(plan.plan),savedParts=unzipSync(saved.bytes);
  const savedXml=strFromU8(savedParts[path]);assert.match(savedXml,/<a:tc>(?:(?!<\/a:tc>)[\s\S])*Preserved(?:(?!<\/a:tc>)[\s\S])*<a:tcPr\b[^>]*\bvert="futureVert"/);
  for(const name of Object.keys(parts))if(name!==path)assert.deepEqual(savedParts[name],parts[name],`unexpected change to ${name}`);
 }finally{session.dispose();}
});

test('table text preservation keeps cell direction through three edits and blocks direction changes',async()=>{
 let bytes=(await generatePptx(deck())).bytes;
 for(let round=1;round<=3;round++){
  const s=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v3'});
  try{const slide=s.document.slides[0],changed=structuredClone(slide.elements[0]);changed.rows[0].cells[1].content[0].runs[0].text=`Edited ${round}`;changed.rows[0].cells[1].style.textDirection='rotate270';slide.replaceElement(changed.id,changed);assert.equal((await s.prepareSave({mode:'edit'})).status,'blocked');changed.rows[0].cells[1].style.textDirection='rotate90';slide.replaceElement(changed.id,changed);const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await s.save(plan.plan)).bytes;const xml=readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();assert.match(xml,/<a:tcPr\b[^>]*\bvert="vert"/);}finally{s.dispose();}
 }
 const final=await importPptx(bytes,{projection:'content'});try{const cell=final.document.slides[0].elements[0].rows[0].cells[1];assert.equal(cell.style.textDirection,'rotate90');assert.equal(cell.content[0].runs[0].text,'Edited 3');}finally{final.dispose();}
});

test('PowerPoint-resaved table retains quarter-turn directions and normalizes explicit horizontal',async()=>{
 const bytes=await readFile('tests/fixtures/table-text-direction-powerpoint-resave.pptx'),s=await importPptx(bytes,{projection:'content'});
 try{const cells=s.document.slides[0].elements[0].rows[0].cells;assert.deepEqual(cells.map(c=>c.content[0].runs[0].text),['DEFAULT','CLOCKWISE','COUNTER','HORIZONTAL']);assert.deepEqual(cells.map(c=>c.style?.textDirection),[undefined,'rotate90','rotate270',undefined]);const plan=await s.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual((await s.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{s.dispose();}
});

test('PowerPoint interactive edits retain text and direction across preservation saves',async()=>{
 let bytes=new Uint8Array(await readFile('tests/fixtures/table-text-direction-powerpoint-interactive.pptx'));
 for(let round=0;round<3;round++){
  const s=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v3'});
  try{
   const slide=s.document.slides[0],table=slide.elements[0],cells=table.rows[0].cells;
   assert.deepEqual(cells.map(c=>c.content.flatMap(p=>p.runs.map(r=>r.text)).join('')),['DEFAULT',round?`CWEDIT${round}`:'CWEDIT','CCWEDIT','HORIZONTAL']);
   assert.deepEqual(cells.map(c=>c.style?.textDirection),[undefined,'rotate90','rotate270',undefined]);
   const changed=structuredClone(table);changed.rows[0].cells[1].content[0].runs[0].text=`CWEDIT${round+1}`;slide.replaceElement(changed.id,changed);
   const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));
   const before=unzipSync(bytes),saved=(await s.save(plan.plan)).bytes,after=unzipSync(saved);
   assert.deepEqual(Object.keys(after).sort(),Object.keys(before).sort());
   for(const name of Object.keys(before))if(name!=='ppt/slides/slide1.xml')assert.deepEqual(after[name],before[name],`unexpected change to ${name}`);
   bytes=saved;
  }finally{s.dispose();}
 }
 const final=await importPptx(bytes,{projection:'content'});try{const cells=final.document.slides[0].elements[0].rows[0].cells;assert.deepEqual(cells.map(c=>c.content.flatMap(p=>p.runs.map(r=>r.text)).join('')),['DEFAULT','CWEDIT3','CCWEDIT','HORIZONTAL']);assert.deepEqual(cells.map(c=>c.style?.textDirection),[undefined,'rotate90','rotate270',undefined]);}finally{final.dispose();}
});
