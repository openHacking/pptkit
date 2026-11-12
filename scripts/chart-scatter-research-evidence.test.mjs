import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { scatterGeometry } from '../packages/layout/dist/chart/scatter-geometry.js';
import { snapshotScatterData } from '../packages/core/dist/chart/scatter-data.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { chromium } from '@playwright/test';

const base='.tmp/native-scatter-research/powerpoint-grywvvg1/', record='docs/testing/chart-scatter-research-powerpoint-evidence.json';
const verify=file=>execFileSync('python3',['scripts/research/verify-scatter-powerpoint.py',...(file?[file]:[])],{encoding:'utf8',stdio:['ignore','pipe','pipe']});
const values=[{xValues:[-8,-3,2,8],yValues:[4,null,-4,6]},{xValues:[-7,-1,3,7],yValues:[-2,4,2,-6]}];
const axis={min:-10,max:10,majorUnit:5,ticks:[-10,-5,0,5,10]};

test('native scatter research source, PDF and one reopen verify',()=>{
 const result=JSON.parse(verify());assert.equal(result.charts,4);assert.equal(result.reopenRounds,1);assert.ok(result.maxMarkerResidual<.2);assert.ok(result.maxEndpointDifference<.3);
});

test('internal scatter geometry matches native centers and blank segment counts',async()=>{
 const data=JSON.parse(await readFile(record,'utf8'));
 for(const sample of data.samples){
  const {xIntercept,xPerUnit,yIntercept,yPerUnit}=sample.fit;
  const plot={x:xIntercept-10*xPerUnit,y:yIntercept+10*yPerUnit,width:20*xPerUnit,height:-20*yPerUnit};
  const snapshot=snapshotScatterData(values,sample.mode);const result=scatterGeometry(snapshot.series,plot,axis,axis,snapshot.blankDisplay);
  for(let s=0;s<2;s++){
   assert.equal(result[s].points.length,sample.markers[s].length);
   for(let i=0;i<result[s].points.length;i++){
    const p=result[s].points[i],q=sample.markers[s][i];assert.ok(Math.hypot(p.x-q[0],p.y-q[1])<.2);
   }
   const segments=sample.slide===1?0:result[s].segments.reduce((sum,segment)=>sum+Math.max(0,segment.length-1),0);
   assert.equal(segments,sample.lineCounts[s]);
  }
 }
});

test('signed-axis native scatter projects coordinates and preserves packages byte-for-byte',async()=>{
 for(const file of ['scatter-research.pptx','scatter-research-round-1.pptx']){
  const bytes=await readFile(base+file),session=await importPptx(bytes,{projection:'content'});
  try{assert.equal(session.document.slides.length,4);for(const [index,slide] of session.document.slides.entries()){const chart=slide.elements.find(e=>e.type==='chart');assert.equal(chart.chartType,'scatter');assert.equal(chart.blankDisplay,['gap','gap','zero','span'][index]);assert.deepEqual(chart.series.map(s=>({xValues:s.xValues,yValues:s.yValues})),values);assert.equal(chart.series[0].line===false,index===0);}const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),bytes);}finally{session.dispose();}
 }
});

test('signed-axis native scatter import agrees in Node and offline Chromium',async t=>{
 const bytes=await readFile(base+'scatter-research-round-1.pptx'),node=await importPptx(bytes,{projection:'content'});let expected;
 try{expected=node.document.slides.map(s=>{const c=s.elements.find(e=>e.type==='chart');return {blankDisplay:c.blankDisplay,series:c.series,axes:c.axes};});}finally{node.dispose();}
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});
 const actual=await page.evaluate(async bytes=>{const s=await PPTKit.importPptx(Uint8Array.from(bytes),{projection:'content'});try{return s.document.slides.map(slide=>{const c=slide.elements.find(e=>e.type==='chart');return {blankDisplay:c.blankDisplay,series:c.series,axes:c.axes};});}finally{s.dispose();}},[...bytes]);assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});

test('PowerPoint-saved sparse scatter sources retain existing X/Y edits across three programmatic rounds',async()=>{
 let bytes=await readFile(base+'scatter-research-round-1.pptx');
 for(let round=0;round<3;round++){
  const session=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});
  try{assert.equal(session.report.chartValues.length,60);assert.ok(session.report.chartValues.every(b=>b.status==='bound'));for(const slide of session.document.slides){const chart=structuredClone(slide.elements.find(e=>e.type==='chart'));assert.equal(chart.series[0].yValues[1],null);chart.series[0].xValues[0]=-6+round;chart.series[0].yValues[2]=-3+round;slide.replaceElement(chart.id,chart);}const plan=await session.prepareSave({mode:'edit'});assert.equal(plan.status,'ready',JSON.stringify(plan.report));bytes=(await session.save(plan.plan)).bytes;}finally{session.dispose();}
 }
 const reopened=await importPptx(bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});try{for(const slide of reopened.document.slides){const chart=slide.elements.find(e=>e.type==='chart');assert.equal(chart.series[0].xValues[0],-4);assert.equal(chart.series[0].yValues[2],-1);assert.equal(chart.series[0].yValues[1],null);}}finally{reopened.dispose();}
});

test('scatter evidence rejects altered geometry, package and capture records',async()=>{
 const original=JSON.parse(await readFile(record,'utf8')),dir=await mkdtemp(path.join(tmpdir(),'pptkit-scatter-evidence-'));
 try{for(const mutate of [r=>{r.samples[1].markers[0][0][0]+=1;},r=>{r.packages[1].sha256='0'.repeat(64);},r=>{r.captures[0].screenshotSha256='0'.repeat(64);}]){
  const bad=structuredClone(original);mutate(bad);const file=path.join(dir,'bad.json');await writeFile(file,JSON.stringify(bad));assert.throws(()=>verify(file));
 }}finally{await rm(dir,{recursive:true,force:true});}
});
