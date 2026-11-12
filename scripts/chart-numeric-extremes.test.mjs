import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, validatePresentation } from '../packages/core/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';

const close=(actual,expected)=>assert.ok(Math.abs(actual-expected)<1e-8,`${actual} != ${expected}`);

test('pie ratios survive tiny and huge finite values without collapsing or overflowing',async()=>{
 for(const magnitude of [1e-300,Number.MIN_VALUE,Number.MAX_VALUE]){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'pie',box:{x:0,y:0,width:400,height:300},categories:['A','B'],series:[{name:'S',values:[magnitude,magnitude]}]});
  const l=resolveLayout(p).slides[0].elements[0].chartLayout;
  for(const slice of l.pie.slices)close(slice.endAngle-slice.startAngle,Math.PI);
  assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/NaN|Infinity/);
 }
});

test('percent-stacked bars normalize tiny and huge values separately by sign',async()=>{
 for(const magnitude of [1e-300,Number.MIN_VALUE,Number.MAX_VALUE]){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'bar',grouping:'percentStacked',box:{x:0,y:0,width:400,height:300},categories:['A'],series:[{name:'positive',values:[magnitude]},{name:'negative',values:[-magnitude]},{name:'positive 2',values:[magnitude]}]});
  const l=resolveLayout(p).slides[0].elements[0].chartLayout;
  close(l.bars[0].box.height/l.plotBox.height,.25);close(l.bars[1].box.height/l.plotBox.height,.5);close(l.bars[2].box.height/l.plotBox.height,.25);
  assert.doesNotMatch((await renderPresentationToSvg(p)).slides[0].svg,/NaN|Infinity/);
 }
});

test('tiny and extreme chart previews agree in Node and browser',async t=>{
 const { chromium }=await import('@playwright/test'),{readFile}=await import('node:fs/promises');
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});
 for(const name of ['core','layout','svg-renderer'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 for(const [kind,magnitude] of [['area',1e-300],['pie',Number.MAX_VALUE],['bar',Number.MIN_VALUE]]){
  const chart=kind==='bar'?{type:'chart',chartType:kind,grouping:'percentStacked',box:{x:0,y:0,width:400,height:300},categories:['A'],series:[{name:'P',values:[magnitude]},{name:'N',values:[-magnitude]}]}:{type:'chart',chartType:kind,box:{x:0,y:0,width:400,height:300},categories:['A','B'],series:[{name:'S',values:[magnitude,magnitude]}]};
  const p=createPresentation({id:'numeric'});p.addSlide({id:'slide'}).addElement({...chart,id:'chart'});const expected=(await renderPresentationToSvg(p)).slides[0].svg;
  const actual=await page.evaluate(async c=>{const p=PPTKit.createPresentation({id:'numeric'});p.addSlide({id:'slide'}).addElement({...c,id:'chart'});return (await PPTKit.renderPresentationToSvg(p)).slides[0].svg;},chart);
  assert.equal(actual,expected);assert.doesNotMatch(actual,/NaN|Infinity/);
 }
 assert.deepEqual(requests,[]);
});

test('closely spaced tiny values keep distinct ticks and labels',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'line',box:{x:0,y:0,width:400,height:300},categories:['A','B'],series:[{name:'S',values:[1e-10,1.000000000001e-10]}]});
 const l=resolveLayout(p).slides[0].elements[0].chartLayout;assert.ok(new Set(l.valueScale.ticks).size>2);assert.ok(l.lines[0].points[1].y<l.lines[0].points[0].y);
 const svg=(await renderPresentationToSvg(p)).slides[0].svg,labels=[...svg.matchAll(/data-chart-role="value-label">([^<]+)<\/text>/g)].map(m=>m[1]);assert.ok(new Set(labels).size>2);
});

test('small tick intervals retain distinct labels at ordinary magnitudes',async()=>{
 const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'line',box:{x:0,y:0,width:400,height:300},categories:['A','B'],series:[{name:'S',values:[1,1.00001]}]});
 const svg=(await renderPresentationToSvg(p)).slides[0].svg,labels=[...svg.matchAll(/data-chart-role="value-label">([^<]+)<\/text>/g)].map(m=>m[1]);assert.ok(new Set(labels).size>2,JSON.stringify(labels));
});


test('stacked bars reject unrepresentable same-sign sums but accept opposite signs and percentages',()=>{
 for(const [grouping,values,invalid] of [['stacked',[Number.MAX_VALUE,Number.MAX_VALUE],true],['stacked',[-Number.MAX_VALUE,-Number.MAX_VALUE],true],['stacked',[Number.MAX_VALUE,-Number.MAX_VALUE],false],['percentStacked',[Number.MAX_VALUE,Number.MAX_VALUE],false]]){
  const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'bar',grouping,box:{x:0,y:0,width:400,height:300},categories:['A'],series:values.map((value,i)=>({name:`S${i}`,values:[value]}))});
  assert.equal(validatePresentation(p).some(d=>d.code==='bar-stack-overflow'),invalid);
 }
});


test('overflowing native stacked charts remain source-only across both dialects',async()=>{
 const {chartFixture,chartXml,series,encodeChart}=await import('../packages/pptx-parser/test/chart-fixture.mjs');const {importPptx}=await import('../packages/pptx-parser/dist/index.js');
 for(const strict of [false,true])for(const family of ['bar','area']){
  const data=series({idx:0,order:0,values:[Number.MAX_VALUE,1]})+series({idx:1,order:1,values:[Number.MAX_VALUE,1]});
  const xml=chartXml({family,data}).replace('<c:grouping val="clustered"/>','<c:grouping val="stacked"/>').replace('<c:areaChart>','<c:areaChart><c:grouping val="stacked"/>');
  const source=encodeChart(chartFixture({strict,xml})),session=await importPptx(source,{projection:'content'});
  try{assert.ok(!session.document.slides[0].elements.some(e=>e.type==='chart'));const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),Buffer.from(source));}finally{session.dispose();}
 }
});
