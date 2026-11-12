import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { scatterGeometry } from '../packages/layout/dist/chart/scatter-geometry.js';
import { snapshotScatterData } from '../packages/core/dist/chart/scatter-data.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const base='.tmp/native-scatter-missing-x/powerpoint-tp3m_1ti/',record='docs/testing/chart-scatter-missing-x-powerpoint-evidence.json';
const verify=file=>execFileSync('python3',['scripts/research/verify-scatter-missing-x-powerpoint.py',...(file?[file]:[])],{encoding:'utf8',stdio:['ignore','pipe','pipe']});
const values=[{xValues:[-8,null,2,8],yValues:[4,3,-4,6]},{xValues:[7,-1,3,-7],yValues:[-6,4,2,-2]}],axis={min:-10,max:10,majorUnit:5,ticks:[-10,-5,0,5,10]};

test('missing X, unsorted series, cache, workbook and one native reopen verify',()=>{
 const result=JSON.parse(verify());assert.equal(result.charts,4);assert.equal(result.reopenRounds,1);assert.ok(result.maxMarkerResidual<.2);assert.ok(result.maxEndpointDifference<.3);
});

test('scatter geometry matches native missing-X positions and unsorted path order',async()=>{
 const data=JSON.parse(await readFile(record,'utf8'));
 for(const sample of data.samples){
  const {xIntercept,xPerUnit,yIntercept,yPerUnit}=sample.fit,plot={x:xIntercept-10*xPerUnit,y:yIntercept+10*yPerUnit,width:20*xPerUnit,height:-20*yPerUnit};
  const snapshot=snapshotScatterData(values,sample.mode);const out=scatterGeometry(snapshot.series,plot,axis,axis,snapshot.blankDisplay);
  for(let j=0;j<2;j++){
   assert.equal(out[j].points.length,sample.markers[j].length);
   for(let i=0;i<out[j].points.length;i++){const p=out[j].points[i],q=sample.markers[j][i];assert.ok(Math.hypot(p.x-q[0],p.y-q[1])<.2);}
   const count=sample.slide===1?0:out[j].segments.reduce((sum,segment)=>sum+Math.max(0,segment.length-1),0);assert.equal(count,sample.lineCounts[j]);
  }
 }
});

test('signed-axis native missing-X samples project coordinates and preserve bytes',async()=>{
 for(const name of ['scatter-missing-x.pptx','scatter-missing-x-round-1.pptx']){
  const bytes=await readFile(base+name),session=await importPptx(bytes,{projection:'content'});
  try{assert.equal(session.document.slides.length,4);for(const [index,slide] of session.document.slides.entries()){const chart=slide.elements.find(e=>e.type==='chart');assert.equal(chart.chartType,'scatter');assert.equal(chart.blankDisplay,['gap','gap','zero','span'][index]);assert.deepEqual(chart.series.map(s=>({xValues:s.xValues,yValues:s.yValues})),values);assert.equal(chart.series[0].line===false,index===0);}const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),bytes);}finally{session.dispose();}
 }
});

test('missing-X evidence rejects altered position, data source and screenshot records',async()=>{
 const original=JSON.parse(await readFile(record,'utf8')),dir=await mkdtemp(path.join(tmpdir(),'pptkit-scatter-missing-x-'));
 try{for(const mutate of [r=>{r.samples[2].markers[0][1][0]+=1;},r=>{r.packages[1].sha256='0'.repeat(64);},r=>{r.captures[0].screenshotSha256='0'.repeat(64);}]){
  const bad=structuredClone(original);mutate(bad);const file=path.join(dir,'bad.json');await writeFile(file,JSON.stringify(bad));assert.throws(()=>verify(file));
 }}finally{await rm(dir,{recursive:true,force:true});}
});
