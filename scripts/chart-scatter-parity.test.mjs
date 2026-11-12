import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, validatePresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { replaceArchiveEntries } from '../packages/pptx-parser/dist/preservation/replace-archive-entries.js';

function deck(mode='gap'){
 const p=createPresentation({id:'scatter'});p.addSlide({id:'slide'}).addElement({id:'xy',type:'chart',chartType:'scatter',box:{x:0,y:0,width:500,height:300},blankDisplay:mode,legend:{visible:false},axes:{x:{scale:{min:-10,max:10,majorUnit:5}},y:{scale:{min:-10,max:10,majorUnit:5}}},series:[{name:'red',xValues:[-8,null,2,8],yValues:[4,3,-4,6],color:'#FF0000',line:{paint:{type:'solid',color:'#FF0000'},width:2},marker:{shape:'circle',size:10}},{name:'blue',xValues:[7,-1,3,-7],yValues:[-6,4,2,-2],color:'#0000FF',marker:{shape:'circle',size:10}}]});return p;
}

test('scatter public model retains independent pairs and resolves numeric axes and gaps',()=>{
 for(const mode of ['gap','zero','span']){
  const p=deck(mode);assert.deepEqual(validatePresentation(p),[]);const chart=normalizePresentation(p).slides[0].elements[0];assert.equal(chart.chartType,'scatter');assert.deepEqual(chart.series[0].xValues,[-8,null,2,8]);assert.ok(!('categories' in chart));
  const l=resolveLayout(p).slides[0].elements[0].chartLayout;assert.equal(l.xScale.min,-10);assert.equal(l.valueScale.max,10);assert.deepEqual(l.scatter[0].points.map(p=>p.sourceIndex),mode==='zero'?[0,1,2,3]:[0,2,3]);assert.deepEqual(l.scatter[1].points.map(p=>p.sourceIndex),[0,1,2,3]);
 }
});

test('scatter produces two value axes, independent numeric caches and workbook cells',async()=>{
 for(const mode of ['gap','zero','span']){
  const p=deck(mode),svg=(await renderPresentationToSvg(p)).slides[0].svg;assert.doesNotMatch(svg,/NaN|Infinity/);assert.match(svg,/data-chart-role="scatter-marker"/);assert.match(svg,/data-chart-role="x-label"/);assert.match(svg,/data-chart-role="y-label"/);
  const bytes=(await generatePptx(p)).bytes,parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),book=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')),sheet=book.get('xl/worksheets/sheet1.xml').toString();
  assert.match(xml,/<c:scatterChart>/);assert.equal((xml.match(/<c:valAx>/g)??[]).length,2);assert.doesNotMatch(xml,/<c:catAx>/);assert.ok(xml.includes(`<c:dispBlanksAs val="${mode}"/>`));assert.ok(xml.includes('<c:f>Sheet1!$A$2:$A$5</c:f>'));assert.ok(xml.includes('<c:f>Sheet1!$C$2:$C$5</c:f>'));assert.ok(sheet.includes('<c r="A3"/>'));assert.ok(sheet.includes('<c r="D5"><v>-2</v></c>'));
  const session=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});try{const chart=session.document.slides[0].elements.find(e=>e.type==='chart');assert.equal(chart.chartType,'scatter');assert.equal(chart.blankDisplay,mode);assert.deepEqual(chart.series.map(s=>[s.xValues,s.yValues]),[[[-8,null,2,8],[4,3,-4,6]],[[7,-1,3,-7],[-6,4,2,-2]]]);assert.equal(chart.axes.x.scale.min,-10);assert.equal(chart.axes.y.scale.max,10);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),Buffer.from(bytes));}finally{session.dispose();}
 }
});

test('scatter authoring and SVG agree in Node and offline browser',async t=>{
 const input={id:'xy',type:'chart',chartType:'scatter',box:{x:0,y:0,width:500,height:300},axes:{x:{scale:{min:-10,max:10,majorUnit:5}},y:{scale:{min:-10,max:10,majorUnit:5}}},series:[{name:'s',xValues:[2,null,-5],yValues:[4,2,-1]}]};
 const p=createPresentation({id:'scatter'});p.addSlide({id:'slide'}).addElement(input);const expected={model:normalizePresentation(p),svg:(await renderPresentationToSvg(p)).slides[0].svg,bytes:[...(await generatePptx(p)).bytes]};
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});for(const n of ['core','layout','pptx-exporter','svg-renderer'])await page.addScriptTag({content:await readFile(`packages/${n}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async input=>{const p=PPTKit.createPresentation({id:'scatter'});p.addSlide({id:'slide'}).addElement(input);return {model:PPTKit.normalizePresentation(p),svg:(await PPTKit.renderPresentationToSvg(p)).slides[0].svg,bytes:[...(await PPTKit.generatePptx(p)).bytes]};},input);assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('scatter survives three fresh imports and coordinate preservation edits are refused',async()=>{
 let bytes=(await generatePptx(deck())).bytes;
 for(let round=0;round<3;round++){
  const session=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
  try{
   const slide=session.document.slides[0],chart=slide.elements.find(e=>e.type==='chart');assert.deepEqual(chart.series[0].xValues,[-8,null,2,8]);assert.deepEqual(chart.series[0].yValues,[4,3,-4,6]);assert.ok(session.report.chartValues.length>0);assert.ok(session.report.chartValues.every(b=>b.status==='unmapped'));
   for(const coordinate of ['xValues','yValues']){const next=structuredClone(chart);next.series[0][coordinate][0]=42;slide.replaceElement(chart.id,next);const plan=await session.prepareSave({mode:'edit'});assert.notEqual(plan.status,'ready');slide.replaceElement(chart.id,chart);}
   bytes=(await generatePptx(session.document)).bytes;
  }finally{session.dispose();}
 }
});

test('opt-in sparse scatter preserves blank positions while editing existing X/Y values over three rounds',async()=>{
 for(const mode of ['gap','zero','span'])for(const blanks of ['x','y','both']){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:mode,series:[{name:'S',xValues:blanks==='y'?[1,2,3]:[1,null,3],yValues:blanks==='x'?[4,5,6]:[4,null,6]}]});let bytes=(await generatePptx(p)).bytes;
  for(let round=0;round<3;round++){
   const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});
   try{const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);assert.ok(s.report.chartValues.every(b=>b.status==='bound'));assert.equal(s.report.chartValues.some(b=>b.coordinate==='x'&&b.pointIndex===1),blanks==='y');assert.equal(s.report.chartValues.some(b=>b.coordinate==='y'&&b.pointIndex===1),blanks==='x');
    const invalid=structuredClone(chart);invalid.series[0][blanks==='y'?'yValues':'xValues'][1]=99;slide.replaceElement(chart.id,invalid);assert.notEqual((await s.prepareSave({mode:'edit'})).status,'ready');slide.replaceElement(chart.id,chart);
    const doc=s.document,changes=[{type:'chart-point',slideId:slide.id,elementId:chart.id,seriesIndex:0,pointIndex:2,coordinate:'xValues',expected:chart.series[0].xValues[2],value:10+round},{type:'chart-point',slideId:slide.id,elementId:chart.id,seriesIndex:0,pointIndex:0,coordinate:'yValues',expected:chart.series[0].yValues[0],value:-10-round}];const undo=doc.applyEditOperation({schemaVersion:'pptkit-edit-v1',documentId:doc.id,changes}).inverse;const redo=doc.applyEditOperation(undo).inverse;doc.applyEditOperation(redo);const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await s.save(plan.plan)).bytes;
    const parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),sheet=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).get('xl/worksheets/sheet1.xml').toString();
    if(blanks!=='y'){assert.ok(sheet.includes('<c r="A3"/>'));assert.match(xml,/<c:xVal>.*?<c:pt idx="0">.*?<c:pt idx="2">/s);}if(blanks!=='x')assert.ok(sheet.includes('<c r="B3"/>'));
    const reopened=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});try{const result=reopened.document.slides[0].elements[0];assert.equal(result.series[0].xValues[2],10+round);assert.equal(result.series[0].yValues[0],-10-round);if(blanks!=='y')assert.equal(result.series[0].xValues[1],null);if(blanks!=='x')assert.equal(result.series[0].yValues[1],null);}finally{reopened.dispose();}
   }finally{s.dispose();}
  }
 }
});

test('sparse scatter v2 rejects conflicting blanks, shared cells and missing caches atomically',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'gap',series:[{name:'S',xValues:[1,null,3],yValues:[1,null,3]}]});const original=(await generatePptx(p)).bytes,parts=readGeneratedZip(original),chartPart='ppt/charts/chart1.xml',bookPart='ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx',xml=parts.get(chartPart).toString(),book=parts.get(bookPart),sheet=readGeneratedZip(book).get('xl/worksheets/sheet1.xml').toString();
 const alteredSheet=sheet.replace('<c r="A3"/>','<c r="A3"><v>8</v></c>');assert.notEqual(alteredSheet,sheet);const changedBook=await replaceArchiveEntries(book,[{name:'xl/worksheets/sheet1.xml',bytes:new TextEncoder().encode(alteredSheet)}],{maxOutputBytes:10_000_000});
 const conflict=await replaceArchiveEntries(original,[{name:bookPart,bytes:changedBook}],{maxOutputBytes:10_000_000});
 const aliasXml=xml.replace('Sheet1!$B$2:$B$4','Sheet1!$A$2:$A$4');assert.notEqual(aliasXml,xml);const alias=await replaceArchiveEntries(original,[{name:chartPart,bytes:new TextEncoder().encode(aliasXml)}],{maxOutputBytes:10_000_000});
 const missingXml=xml.replace(/(<c:xVal><c:numRef><c:f>[^<]+<\/c:f>)<c:numCache>.*?<\/c:numCache>/s,'$1');assert.notEqual(missingXml,xml);const missing=await replaceArchiveEntries(original,[{name:chartPart,bytes:new TextEncoder().encode(missingXml)}],{maxOutputBytes:10_000_000});
 for(const bytes of [conflict,alias,missing]){const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});try{assert.ok(s.report.chartValues.some(b=>b.status==='unmapped'));const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[0]=9;chart.series[0].yValues[0]=8;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});assert.notEqual(plan.status,'ready');}finally{s.dispose();}}
});

test('sparse scatter v2 source saves agree in Node and offline Chromium',async t=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'span',series:[{name:'S',xValues:[1,null,3],yValues:[4,5,6]}]});const bytes=(await generatePptx(p)).bytes,options={projection:'content',chartEditSaveProfile:'existing-values-v2'};
 const s=await importPptx(bytes,options);let expected;try{const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[2]=7;chart.series[0].yValues[0]=8;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready');expected=[...(await s.save(plan.plan)).bytes];}finally{s.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});
 const actual=await page.evaluate(async({bytes,options})=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),options);try{const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[2]=7;chart.series[0].yValues[0]=8;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});if(plan.status!=='ready')throw Error(JSON.stringify(plan.report));return [...(await s.save(plan.plan)).bytes];}finally{s.dispose();}},{bytes:[...bytes],options});assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('sparse scatter v2 retains signed zero in both cache and workbook',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'gap',series:[{name:'S',xValues:[-0,null,3],yValues:[4,5,6]}]});const s=await importPptx((await generatePptx(p)).bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});
 try{const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);assert.ok(Object.is(chart.series[0].xValues[0],-0));chart.series[0].xValues[2]=-0;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready');const bytes=(await s.save(plan.plan)).bytes,parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),sheet=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).get('xl/worksheets/sheet1.xml').toString();assert.match(xml,/<c:pt idx="2"><c:v>-0<\/c:v><\/c:pt>/);assert.ok(sheet.includes('<c r="A4"><v>-0</v></c>'));const reopened=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});try{assert.ok(Object.is(reopened.document.slides[0].elements[0].series[0].xValues[2],-0));assert.ok(reopened.report.chartValues.every(b=>b.status==='bound'));}finally{reopened.dispose();}}finally{s.dispose();}
});

test('sparse value profile does not widen category-chart edit authority',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'line',box:{x:0,y:0,width:400,height:300},categories:['A','B','C'],blankDisplay:'gap',series:[{name:'S',values:[1,null,3]}]});
 const s=await importPptx((await generatePptx(p)).bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});try{assert.ok(s.report.chartValues.every(b=>b.status==='unmapped'));const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].values[0]=8;slide.replaceElement(chart.id,chart);assert.notEqual((await s.prepareSave({mode:'edit'})).status,'ready');}finally{s.dispose();}
});

test('sparse scatter v2 combines atomically with bound text and first-slide number',async()=>{
 const p=createPresentation(),slide=p.addSlide();slide.addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'gap',series:[{name:'S',xValues:[1,null,3],yValues:[4,5,6]}]});slide.addElement({type:'text',box:{x:0,y:310,width:400,height:50},content:'before'});
 const options={projection:'content',chartEditSaveProfile:'existing-values-v2',editSaveProfile:'text-runs-v3',firstSlideNumberSaveProfile:'attribute-v2'},s=await importPptx((await generatePptx(p)).bytes,options);
 try{const owner=s.document.slides[0],chart=structuredClone(owner.elements[0]),text=structuredClone(owner.elements[1]);chart.series[0].xValues[2]=8;text.content[0].runs[0].text='after';owner.replaceElement(chart.id,chart);owner.replaceElement(text.id,text);s.document.setFirstSlideNumber(17);const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));const result=await importPptx((await s.save(plan.plan)).bytes,options);try{assert.equal(result.document.firstSlideNumber,17);assert.deepEqual(result.document.slides[0].elements[0].series[0].xValues,[1,null,8]);assert.equal(result.document.slides[0].elements[1].content[0].runs[0].text,'after');}finally{result.dispose();}}finally{s.dispose();}
});

test('dense scatter coordinate saves update distinct caches and workbook cells over three reopens',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:500,height:300},series:[{name:'A',xValues:[1,2],yValues:[3,4]},{name:'B',xValues:[5,6],yValues:[7,8]}]});
 let bytes=(await generatePptx(p)).bytes;
 for(let round=0;round<3;round++){
  const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
  try{
   const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);assert.equal(s.report.chartValues.length,8);assert.ok(s.report.chartValues.every(b=>b.status==='bound'));assert.equal(new Set(s.report.chartValues.map(b=>`${b.coordinate}:${b.seriesIndex}:${b.pointIndex}`)).size,8);
   chart.series[0].xValues[0]=10+round;chart.series[0].yValues[0]=20+round;chart.series[1].yValues[1]=-30-round;slide.replaceElement(chart.id,chart);
   const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));
   if(round===0){const changed=structuredClone(chart);changed.series[0].xValues[0]++;slide.replaceElement(chart.id,changed);await assert.rejects(s.save(plan.plan),{code:'save-snapshot-stale'});slide.replaceElement(chart.id,chart);}
   const saved=await s.save(plan.plan);assert.deepEqual(saved.report.changes.map(c=>c.chartValue?.coordinate).sort(),['x','y','y']);bytes=saved.bytes;
   const parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),sheet=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).get('xl/worksheets/sheet1.xml').toString();
   assert.match(xml,new RegExp(`<c:pt idx="0"><c:v>${10+round}</c:v></c:pt>`));assert.ok(sheet.includes(`<c r="A2"><v>${10+round}</v></c>`));assert.ok(sheet.includes(`<c r="B2"><v>${20+round}</v></c>`));assert.ok(sheet.includes(`<c r="D3"><v>${-30-round}</v></c>`));
   const reopened=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});try{const actual=reopened.document.slides[0].elements[0];assert.equal(actual.series[0].xValues[0],10+round);assert.equal(actual.series[0].yValues[0],20+round);assert.equal(actual.series[1].yValues[1],-30-round);}finally{reopened.dispose();}
  }finally{s.dispose();}
 }
});

test('scatter rejects conflicting caches and shared coordinate workbook cells',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:500,height:300},series:[{name:'A',xValues:[1,2],yValues:[1,2]}]});
 const original=(await generatePptx(p)).bytes,part='ppt/charts/chart1.xml',xml=readGeneratedZip(original).get(part).toString();
 for(const altered of [xml.replace('<c:xVal><c:numRef>','<c:xVal><c:numRef>').replace('<c:v>1</c:v>','<c:v>999</c:v>'),xml.replace('Sheet1!$B$2:$B$3','Sheet1!$A$2:$A$3')]){
  assert.notEqual(altered,xml);const bytes=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(altered)}],{maxOutputBytes:10_000_000});const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
  try{assert.ok(s.report.chartValues.length>0);assert.ok(s.report.chartValues.every(b=>b.status==='unmapped'));const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[0]=10;slide.replaceElement(chart.id,chart);assert.notEqual((await s.prepareSave({mode:'edit'})).status,'ready');}finally{s.dispose();}
 }
});

test('missing scatter cache uses a blank-aware workbook range only with a bounded counterpart',async()=>{
 for(const coordinate of ['x','y'])for(const mode of ['gap','zero','span']){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:mode,series:[{name:'S',xValues:coordinate==='x'?[1,null,3]:[1,2,3],yValues:coordinate==='y'?[4,null,6]:[4,5,6]}]});
  const original=(await generatePptx(p)).bytes,part='ppt/charts/chart1.xml',xml=readGeneratedZip(original).get(part).toString(),pattern=new RegExp(`(<c:${coordinate}Val><c:numRef><c:f>[^<]+</c:f>)<c:numCache>.*?</c:numCache>`,'s');
  const missing=xml.replace(pattern,'$1');assert.notEqual(missing,xml);
  const bytes=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(missing)}],{maxOutputBytes:10_000_000});
  const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
  try{const slide=s.document.slides[0],chart=slide.elements[0];assert.deepEqual(chart.series[0].xValues,coordinate==='x'?[1,null,3]:[1,2,3]);assert.deepEqual(chart.series[0].yValues,coordinate==='y'?[4,null,6]:[4,5,6]);assert.ok(s.report.findings.some(f=>f.reason.includes('workbook range with explicit blanks')));const plan=await s.prepareSave({mode:'no-op'});assert.deepEqual(Buffer.from((await s.save(plan.plan)).bytes),Buffer.from(bytes));const changed=structuredClone(chart);changed.series[0].xValues[0]=8;slide.replaceElement(chart.id,changed);assert.notEqual((await s.prepareSave({mode:'edit'})).status,'ready');}finally{s.dispose();}
  const unequal=missing.replace(coordinate==='x'?'$A$2:$A$4':'$B$2:$B$4',coordinate==='x'?'$A$2:$A$3':'$B$2:$B$3');assert.notEqual(unequal,missing);
  const bad=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(unequal)}],{maxOutputBytes:10_000_000});const s2=await importPptx(bad,{projection:'content'});try{assert.ok(s2.document.slides[0].elements.every(e=>e.type!=='chart'));}finally{s2.dispose();}
  const unspecified=missing.replace(`<c:dispBlanksAs val="${mode}"/>`,'');assert.notEqual(unspecified,missing);const noMode=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(unspecified)}],{maxOutputBytes:10_000_000});const s3=await importPptx(noMode,{projection:'content'});try{assert.ok(s3.document.slides[0].elements.every(e=>e.type!=='chart'));}finally{s3.dispose();}
  const noCount=xml.replace(new RegExp(`(<c:${coordinate}Val><c:numRef><c:f>[^<]+</c:f><c:numCache>.*?)(<c:ptCount val="3"/>)`,'s'),'$1');assert.notEqual(noCount,xml);const malformed=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(noCount)}],{maxOutputBytes:10_000_000});await assert.rejects(importPptx(malformed,{projection:'content'}),e=>e.code==='import-invalid-package');
  const other=coordinate==='x'?'y':'x',both=missing.replace(new RegExp(`(<c:${other}Val><c:numRef><c:f>[^<]+</c:f>)<c:numCache>.*?</c:numCache>`,'s'),'$1');assert.notEqual(both,missing);const noCaches=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(both)}],{maxOutputBytes:10_000_000});const s4=await importPptx(noCaches,{projection:'content'});try{const chart=s4.document.slides[0].elements[0];assert.equal(chart.chartType,'scatter');assert.equal(chart.series[0][coordinate==='x'?'xValues':'yValues'][1],null);}finally{s4.dispose();}
 }
});

test('counterpart-bounded scatter workbook fallback agrees in offline Chromium',async t=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'span',series:[{name:'S',xValues:[1,null,3],yValues:[4,5,6]}]});
 const original=(await generatePptx(p)).bytes,part='ppt/charts/chart1.xml',xml=readGeneratedZip(original).get(part).toString().replace(/(<c:xVal><c:numRef><c:f>[^<]+<\/c:f>)<c:numCache>.*?<\/c:numCache>/s,'$1'),bytes=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(xml)}],{maxOutputBytes:10_000_000});
 const node=await importPptx(bytes,{projection:'content'});let expected;try{expected=node.document.slides[0].elements[0].series;}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});
 const actual=await page.evaluate(async bytes=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),{projection:'content'});try{return s.document.slides[0].elements[0].series;}finally{s.dispose();}},[...bytes]);assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('signed scatter axis identities survive three dense source-bound X/Y saves',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:500,height:300},series:[{name:'S',xValues:[1,2],yValues:[3,4]}]});
 const original=(await generatePptx(p)).bytes,part='ppt/charts/chart1.xml',xml=readGeneratedZip(original).get(part).toString().replaceAll('<c:axId val="1"/>','<c:axId val="-2128940872"/>').replaceAll('<c:axId val="2"/>','<c:axId val="-2129643912"/>').replaceAll('<c:crossAx val="1"/>','<c:crossAx val="-2128940872"/>').replaceAll('<c:crossAx val="2"/>','<c:crossAx val="-2129643912"/>');
 assert.match(xml,/axId val="-2128940872"/);let bytes=await replaceArchiveEntries(original,[{name:part,bytes:new TextEncoder().encode(xml)}],{maxOutputBytes:10_000_000});
 for(let round=0;round<3;round++){
  const session=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
  try{assert.ok(session.report.chartValues.every(b=>b.status==='bound'));const slide=session.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[0]=5+round;chart.series[0].yValues[0]=8+round;slide.replaceElement(chart.id,chart);const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await session.save(plan.plan)).bytes;const saved=readGeneratedZip(bytes).get(part).toString();assert.equal((saved.match(/axId val="-2128940872"/g)??[]).length,2);assert.equal((saved.match(/axId val="-2129643912"/g)??[]).length,2);assert.match(saved,/crossAx val="-2128940872"/);assert.match(saved,/crossAx val="-2129643912"/);}finally{session.dispose();}
 }
});

test('dense scatter X/Y preservation matches offline browser output',async t=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:250},series:[{name:'S',xValues:[1,2],yValues:[3,4]}]});const bytes=(await generatePptx(p)).bytes,options={projection:'content',chartEditSaveProfile:'existing-values-v1'};
 const node=await importPptx(bytes,options);let expected;try{const slide=node.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[0]=9;chart.series[0].yValues[0]=10;slide.replaceElement(chart.id,chart);const plan=await node.prepareSave({mode:'edit'});assert.equal(plan.status,'ready');expected=[...(await node.save(plan.plan)).bytes];}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});
 const actual=await page.evaluate(async({bytes,options})=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),options);try{const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[0]=9;chart.series[0].yValues[0]=10;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});if(plan.status!=='ready')throw Error(JSON.stringify(plan.report));return [...(await s.save(plan.plan)).bytes];}finally{s.dispose();}},{bytes:[...bytes],options});assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('scatter content import and byte-exact no-op agree in Node and offline browser',async t=>{
 const bytes=(await generatePptx(deck('span'))).bytes,options={projection:'content',chartEditSaveProfile:'existing-values-v1'};
 const node=await importPptx(bytes,options);let expected;
 try{const c=node.document.slides[0].elements.find(e=>e.type==='chart'),plan=await node.prepareSave({mode:'no-op'});expected={series:c.series,axes:c.axes,blankDisplay:c.blankDisplay,bytes:[...(await node.save(plan.plan)).bytes]};}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});
 const actual=await page.evaluate(async({bytes,options})=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),options);try{const c=s.document.slides[0].elements.find(e=>e.type==='chart'),plan=await s.prepareSave({mode:'no-op'});return {series:c.series,axes:c.axes,blankDisplay:c.blankDisplay,bytes:[...(await s.save(plan.plan)).bytes]};}finally{s.dispose();}},{bytes:[...bytes],options});assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('scatter validates independent arrays, both axis scales and styles',()=>{
 const cases=[
  [{name:'s',xValues:[1],yValues:[]}],{code:'scatter-values-length'},
  [{name:'s',xValues:[,],yValues:[1]}],{code:'scatter-coordinate'},
 ];
 for(let i=0;i<cases.length;i+=2){const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},series:cases[i]});assert.ok(validatePresentation(p).some(d=>d.code===cases[i+1].code));}
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},axes:{x:{scale:{min:1,max:1,majorUnit:1}},y:{scale:{min:0,max:1,majorUnit:0}}},series:[{name:'',xValues:[1],yValues:[2],marker:{shape:'bad'}}]});const d=validatePresentation(p);assert.ok(d.some(x=>x.code==='chart-series-name'));assert.ok(d.some(x=>x.code==='invalid-chart-marker'));assert.ok(d.some(x=>x.path.endsWith('axes.x.scale')));assert.ok(d.some(x=>x.path.endsWith('axes.y.scale')));
});

test('scatter auto axes, empty series and null endpoints stay finite and ordered',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'span',series:[{name:'empty',xValues:[],yValues:[]},{name:'s',xValues:[null,7,-2,null],yValues:[null,4,-1,1]}]});
 const chart=normalizePresentation(p).slides[0].elements[0],l=resolveLayout(p).slides[0].elements[0].chartLayout;assert.deepEqual(l.scatter[0].points,[]);assert.deepEqual(l.scatter[1].points.map(p=>p.sourceIndex),[1,2]);assert.ok([l.xScale.min,l.xScale.max,l.valueScale.min,l.valueScale.max].every(Number.isFinite));assert.equal(chart.blankDisplay,'span');assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/NaN|Infinity/);
});

test('scatter auto axes handle constant extrema and subnormal values without invalid coordinates',async()=>{
 for(const value of [Number.MAX_VALUE,-Number.MAX_VALUE,Number.MIN_VALUE,-Number.MIN_VALUE]){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},series:[{name:'s',xValues:[value],yValues:[value]}]});
  const l=resolveLayout(p).slides[0].elements[0].chartLayout;for(const axis of [l.xScale,l.valueScale])assert.ok([axis.min,axis.max,axis.majorUnit,...axis.ticks].every(Number.isFinite)&&axis.min<axis.max);assert.ok(l.scatter[0].points.every(p=>Number.isFinite(p.x)&&Number.isFinite(p.y)));assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/NaN|Infinity/);
 }
});

test('scatter export keeps different series lengths and empty series distinct',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},series:[{name:'long',xValues:[1,2,3],yValues:[4,null,6]},{name:'short',xValues:[7],yValues:[8]},{name:'empty',xValues:[],yValues:[]}]});
 const bytes=(await generatePptx(p)).bytes,parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),sheet=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).get('xl/worksheets/sheet1.xml').toString();
 assert.ok(xml.includes('<c:f>Sheet1!$A$2:$A$4</c:f>'));assert.ok(xml.includes('<c:f>Sheet1!$C$2:$C$2</c:f>'));assert.equal((xml.match(/<c:numLit>/g)??[]).length,2);assert.ok(sheet.includes('<c r="B3"/>'));assert.ok(!sheet.includes('<c r="C3"'));assert.ok(sheet.includes('<c r="F1"'));
});
