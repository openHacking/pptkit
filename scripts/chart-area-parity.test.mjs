import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout, resolveNormalizedLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { chartFixture, chartXml, series, encodeChart } from "../packages/pptx-parser/test/chart-fixture.mjs";

const close=(a,b)=>assert.ok(Math.abs(a-b)<1e-8,`${a} != ${b}`);
function deck(grouping="standard", blankDisplay, values=[[4,2,null,-2,3],[2,-3,3,4,1]], scale) {
  const p=createPresentation({id:'area'}), s=p.addSlide({id:'slide'});
  s.addElement({id:'chart',type:'chart',chartType:'area',grouping,...(blankDisplay?{blankDisplay}:{}),box:{x:0,y:0,width:500,height:300},categories:['A','B','C','D','E'],series:values.map((value,i)=>({name:`S${i}`,values:value,color:i?'#0000FF':'#FF0000'})),axes:{value:{scale:scale ?? (grouping==='percentStacked'?{min:-1,max:1,majorUnit:.5}:{min:-10,max:10,majorUnit:5})}},legend:{visible:false}});return p;
}
function yvalue(y,plot,scale){return scale.min+(plot.y+plot.height-y)/plot.height*(scale.max-scale.min);}

test('stacked area warnings match layout endpoints and survive fresh export and import', async () => {
 for (const grouping of ['stacked', 'percentStacked']) {
  const scale=grouping==='stacked'?{min:-1,max:5,majorUnit:1}:{min:-.1,max:.8,majorUnit:.1};
  const p=deck(grouping,'zero',[[4,2,0,-2,3],[2,-3,3,4,1]],scale);
  const chart=p.slides[0].elements[0];
  const layout=resolveLayout(p).slides[0].elements[0].chartLayout;
  const expected=[];
  for (const [seriesIndex,area] of layout.areas.entries()) for (const [categoryIndex,point] of area.polygons[0].slice(0,5).entries()) {
   const value=yvalue(point.y,layout.plotBox,layout.valueScale);
   if(value < chart.axes.value.scale.min-1e-8 || value > chart.axes.value.scale.max+1e-8) expected.push(`series.${seriesIndex}.values.${categoryIndex}`);
  }
  const paths=doc=>validatePresentation(doc).filter(d=>d.code==='chart-value-outside-fixed-axis').map(d=>d.path.match(/series\.\d+\.values\.\d+$/)[0]).sort();
  expected.sort(); assert.deepEqual(paths(p),expected); assert.ok(expected.length>0);
  const session=await importPptx((await generatePptx(p)).bytes,{projection:'content'});
  try { assert.deepEqual(paths(session.document),expected); } finally { session.dispose(); }
 }
});

test('area gap, zero and signed stacked bands have the expected geometry',()=>{
 for(const [group,mode] of [['standard','gap'],['standard','zero'],['stacked','zero'],['percentStacked','zero']]){
  const p=deck(group,mode), layout=resolveLayout(p).slides[0].elements[0].chartLayout, areas=layout.areas;
  assert.equal(layout.categoryAxisPlacement,'onCategory');assert.equal(areas.length,2); assert.equal(areas[0].polygons.length,mode==='gap'?2:1);
  if(group==='stacked'||group==='percentStacked'){
   const polygon=areas[1].polygons[0], top=polygon.slice(0,5), bottom=polygon.slice(5).reverse();
   const expected=group==='stacked'?[[4,6],[2,-1],[0,3],[-2,2],[3,4]]:[[4/6,1],[2/5,-1/5],[0,1],[-2/6,2/6],[3/4,1]];
   for(let i=0;i<5;i++){close(yvalue(bottom[i].y,layout.plotBox,layout.valueScale),expected[i][0]);close(yvalue(top[i].y,layout.plotBox,layout.valueScale),expected[i][1]);}
  }
 }
});

test('area chart cache, workbook, preview and import survive three fresh rounds',async()=>{
 for(const [group,mode] of [['standard','gap'],['standard','zero'],['stacked','zero'],['percentStacked','zero']]){
  const p=deck(group,mode), svg=(await renderPresentationToSvg(p)).slides[0].svg;assert.match(svg,/data-chart-role="area"/);assert.doesNotMatch(svg,/NaN|Infinity/);
  if(group==='percentStacked')assert.match(svg,/>50%<\/text>/);
  let bytes=(await generatePptx(p)).bytes;
  for(let round=0;round<3;round++){
   const parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),book=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')),sheet=book.get('xl/worksheets/sheet1.xml').toString();
   assert.match(xml,new RegExp(`<c:areaChart><c:grouping val="${group}"`));assert.ok(xml.includes(`<c:dispBlanksAs val="${mode}"/>`));assert.ok(sheet.includes('<c r="B4"/>'));assert.match(xml,/<c:crossBetween val="midCat"\/>/);
   if(group==='percentStacked')assert.match(xml,/<c:numFmt formatCode="0%" sourceLinked="0"\/>/);
   const session=await importPptx(bytes,{projection:'content'});try{const chart=session.document.slides[0].elements[0];assert.equal(chart.chartType,'area');assert.equal(chart.grouping,group);assert.equal(chart.blankDisplay,mode);assert.deepEqual(chart.series.map(s=>s.values),[[4,2,null,-2,3],[2,-3,3,4,1]]);const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),Buffer.from(bytes));bytes=(await generatePptx(session.document)).bytes;}finally{session.dispose();}
  }
 }
});

test('area grouping, unsupported blank strategies and stack overflow are rejected',()=>{
 assert.ok(validatePresentation(deck('stacked','gap')).some(d=>d.code==='invalid-chart-blank-display'));
 assert.ok(validatePresentation(deck('percentStacked','span')).some(d=>d.code==='invalid-chart-blank-display'));
 assert.ok(validatePresentation(deck('bad')).some(d=>d.code==='invalid-chart-grouping'));
 assert.ok(validatePresentation(deck('stacked','zero',[[Number.MAX_VALUE,0,0,0,0],[Number.MAX_VALUE,0,0,0,0]])).some(d=>d.code==='area-stack-overflow'));
 const normalized=normalizePresentation(deck('stacked',undefined)).slides[0].elements[0];assert.equal(normalized.blankDisplay,'zero');
});

test('area import leaves unsupported stacked gap source-only and no-op intact',async()=>{
 const bytes=(await generatePptx(deck('stacked','zero'))).bytes,parts=readGeneratedZip(bytes),key='ppt/charts/chart1.xml';parts.set(key,Buffer.from(parts.get(key).toString().replace('<c:dispBlanksAs val="zero"/>','<c:dispBlanksAs val="gap"/>')));const source=zip([...parts].map(([name,body])=>({name,body}))).bytes;
 const session=await importPptx(source,{projection:'content'});try{assert.ok(!session.document.slides[0].elements.some(e=>e.type==='chart'));const p=await session.prepareSave({mode:'no-op'});assert.deepEqual(Buffer.from((await session.save(p.plan)).bytes),Buffer.from(source));}finally{session.dispose();}
});

test('area Node and browser global import/export and SVG agree',async t=>{
 const source=(await generatePptx(deck('stacked','zero'))).bytes,node=await importPptx(source,{projection:'content'});let expected;try{expected={model:normalizePresentation(node.document),bytes:[...(await generatePptx(node.document)).bytes],svg:(await renderPresentationToSvg(node.document)).slides[0].svg};}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});for(const name of ['core','layout','pptx-exporter','pptx-parser','svg-renderer'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async input=>{const session=await PPTKit.importPptx(Uint8Array.from(input),{projection:'content'});try{return {model:PPTKit.normalizePresentation(session.document),bytes:[...(await PPTKit.generatePptx(session.document)).bytes],svg:(await PPTKit.renderPresentationToSvg(session.document)).slides[0].svg};}finally{session.dispose();}},[...source]);assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('percent area handles all-zero and very large signed inputs without overflow',()=>{
 for(const values of [[[0,0,0,0,0],[0,0,0,0,0]],[[Number.MAX_VALUE,0,0,0,0],[-Number.MAX_VALUE,0,0,0,0]]]){
  const p=deck('percentStacked','zero',values), layout=resolveLayout(p).slides[0].elements[0].chartLayout;
  assert.ok(layout.areas.flatMap(a=>a.polygons.flat()).every(p=>Number.isFinite(p.x)&&Number.isFinite(p.y)));
  if(values[0][0]){const polygon=layout.areas[1].polygons[0];close(yvalue(polygon[0].y,layout.plotBox,layout.valueScale),0);close(yvalue(polygon.at(-1).y,layout.plotBox,layout.valueScale),.5);}
 }
});

test('area blanks retain category positions and single points do not become filled spans',async()=>{
 const p=deck('standard','gap',[[null,2,null,null,null],[null,null,null,null,null]]),l=resolveLayout(p).slides[0].elements[0].chartLayout;
 assert.equal(l.areas[0].polygons.length,1);assert.equal(l.areas[0].polygons[0].length,2);assert.deepEqual(l.areas[1].polygons,[]);
 assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/data-chart-role="area"/);
 const s=await importPptx((await generatePptx(p)).bytes,{projection:'content'});try{assert.deepEqual(s.document.slides[0].elements[0].series.map(x=>x.values),[[null,2,null,null,null],[null,null,null,null,null]]);}finally{s.dispose();}
});

test('unsupported area span policy remains source-only even with dense values',async()=>{
 const bytes=(await generatePptx(deck('standard','zero',[[4,2,1,-2,3],[2,-3,3,4,1]]))).bytes,parts=readGeneratedZip(bytes),key='ppt/charts/chart1.xml';
 parts.set(key,Buffer.from(parts.get(key).toString().replace('<c:dispBlanksAs val="zero"/>','<c:dispBlanksAs val="span"/>')));
 const source=zip([...parts].map(([name,body])=>({name,body}))).bytes,session=await importPptx(source,{projection:'content'});
 try{assert.ok(!session.document.slides[0].elements.some(e=>e.type==='chart'));const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),Buffer.from(source));}finally{session.dispose();}
});

test('automatic area scales keep finite ticks and geometry for extreme finite values',async()=>{
 for(const magnitude of [1e300,Number.MAX_VALUE]){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'area',box:{x:0,y:0,width:500,height:300},categories:['a','b'],series:[{name:'s',values:[magnitude,-magnitude]}]});
  const l=resolveLayout(p).slides[0].elements[0].chartLayout;assert.ok([l.valueScale.min,l.valueScale.max,l.valueScale.majorUnit,...l.valueScale.ticks].every(Number.isFinite));assert.ok(l.valueScale.ticks.some(v=>v>0));assert.ok(l.valueScale.ticks.some(v=>v<0));
  assert.ok(l.areas[0].polygons.flat().every(p=>Number.isFinite(p.x)&&Number.isFinite(p.y)));assert.ok(l.areas[0].polygons[0][0].y < l.areas[0].polygons[0][1].y);assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/NaN|Infinity/);
 }
});


test('transported stacked area without blankDisplay retains zero geometry',()=>{
 const normalized=normalizePresentation(deck('stacked','zero'));delete normalized.slides[0].elements[0].blankDisplay;
 const area=resolveNormalizedLayout(normalized).slides[0].elements[0].chartLayout.areas[0];assert.equal(area.polygons.length,1);assert.equal(area.polygons[0].length,10);
});


test('extreme manual percent scales produce finite SVG coordinates and readable labels',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'area',grouping:'percentStacked',box:{x:0,y:0,width:500,height:300},categories:['a','b'],series:[{name:'s',values:[1,2]}],axes:{value:{scale:{min:-Number.MAX_VALUE,max:Number.MAX_VALUE,majorUnit:Number.MAX_VALUE/2}}}});
 const svg=(await renderPresentationToSvg(p)).slides[0].svg;assert.doesNotMatch(svg,/NaN|Infinity/);assert.match(svg,/e\+310%/);
});


test('independent area fixtures import both OOXML dialects and supported blank modes',async()=>{
 for(const strict of [false,true])for(const [group,mode] of [['standard','gap'],['standard','zero'],['stacked','zero'],['percentStacked','zero']]){
  const data=series({values:[0,2]}).replace('<c:pt idx="1"><c:v>2</c:v></c:pt>','');
  const xml=chartXml({family:'area',data,after:`<c:dispBlanksAs val="${mode}"/>`}).replace('<c:areaChart>',`<c:areaChart><c:grouping val="${group}"/>`);
  const source=encodeChart(chartFixture({strict,xml})),s=await importPptx(source,{projection:'content'});
  try{const chart=s.document.slides[0].elements[0];assert.equal(chart.chartType,'area');assert.equal(chart.grouping,group);assert.equal(chart.blankDisplay,mode);assert.deepEqual(chart.series[0].values,[0,null]);const plan=await s.prepareSave({mode:'no-op'});assert.deepEqual(Buffer.from((await s.save(plan.plan)).bytes),Buffer.from(source));}finally{s.dispose();}
 }
});

test('dense area numeric source edits synchronize cache and workbook over three reopens',async()=>{
 for(const grouping of ['standard','stacked','percentStacked']){
  let bytes=(await generatePptx(deck(grouping,'zero',[[4,2,1,-2,3],[2,-3,3,4,1]]))).bytes;
  for(let round=0;round<3;round++){
   const s=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v1'});
   try{assert.equal(s.report.chartValues.length,10);assert.ok(s.report.chartValues.every(v=>v.status==='bound'));const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].values[1]=10+round;slide.replaceElement(chart.id,chart);
    const plan=await s.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await s.save(plan.plan)).bytes;
    const parts=readGeneratedZip(bytes),xml=parts.get('ppt/charts/chart1.xml').toString(),book=readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'));
    assert.ok(xml.includes(`<c:pt idx="1"><c:v>${10+round}</c:v></c:pt>`));assert.ok(book.get('xl/worksheets/sheet1.xml').toString().includes(`<c r="B3"><v>${10+round}</v></c>`));
   }finally{s.dispose();}
  }
 }
});

test('small nonzero values retain distinct auto-scale ticks, geometry and labels',async()=>{
 for(const magnitude of [1e-10,1e-15,1e-300,1e-323])for(const kind of ['area','line','bar']){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:kind,box:{x:0,y:0,width:400,height:300},categories:['a','b'],series:[{name:'s',values:[magnitude,magnitude*2]}]});
  const l=resolveLayout(p).slides[0].elements[0].chartLayout,scale=l.valueScale;
  assert.ok(Number.isFinite(scale.min)&&Number.isFinite(scale.max)&&Number.isFinite(scale.majorUnit)&&scale.max>scale.min,`${kind} ${magnitude}`);
  assert.ok(scale.max<1e-6,`${kind} ${magnitude}: ${scale.max}`);assert.ok(scale.ticks.some(v=>v!==0),`${kind} ${magnitude} ticks`);
  const points=kind==='area'?l.areas[0].polygons[0].slice(0,2):kind==='line'?l.lines[0].points:l.bars.map(b=>({y:b.box.y}));
  assert.ok(points[1].y<points[0].y,`${kind} ${magnitude} geometry`);
  const svg=(await renderPresentationToSvg(p)).slides[0].svg;assert.doesNotMatch(svg,/NaN|Infinity/);assert.match(svg,/data-chart-role="value-label">[^<]*[1-9]/);
 }
});
