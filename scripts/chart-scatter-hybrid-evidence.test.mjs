import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { scatterWorkbook } from '../packages/pptx-exporter/dist/ooxml/scatter-workbook.js';
import { scatterSeriesXml } from '../packages/pptx-exporter/dist/ooxml/scatter-series.js';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

const base='.tmp/native-scatter-hybrid/powerpoint-V2fSwr/',record='docs/testing/chart-scatter-hybrid-powerpoint-evidence.json';
const verify=file=>execFileSync('python3',['scripts/research/verify-scatter-hybrid-powerpoint.py',...(file?[file]:[])],{encoding:'utf8',stdio:['ignore','pipe','pipe']});

test('hybrid native shell, PPTKit caches/workbooks and control geometry verify',()=>{
 const result=JSON.parse(verify());assert.equal(result.charts,4);assert.equal(result.reopenRounds,1);assert.equal(result.maxMarkerBoxDifference,0);assert.equal(result.maxLineEndpointDifference,0);
});
test('signed-axis hybrid scatter projects coordinates and survives original-byte no-op',async()=>{
 for(const file of ['scatter-hybrid.pptx','scatter-hybrid-round-1.pptx']){
  const bytes=await readFile(base+file),session=await importPptx(bytes,{projection:'content'});
  try{assert.equal(session.document.slides.length,4);for(const [index,slide] of session.document.slides.entries()){const chart=slide.elements.find(e=>e.type==='chart');assert.equal(chart.chartType,'scatter');assert.equal(chart.blankDisplay,['gap','gap','zero','span'][index]);assert.deepEqual(chart.series.map(s=>[s.xValues,s.yValues]),[[[-8,null,2,8],[4,3,-4,6]],[[7,-1,3,-7],[-6,4,2,-2]]]);assert.equal(chart.series[0].line===false,index===0);}const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),bytes);}finally{session.dispose();}
 }
});
test('hybrid evidence rejects altered hashes, source package and native measurements',async()=>{
 const original=JSON.parse(await readFile(record,'utf8')),dir=await mkdtemp(path.join(tmpdir(),'pptkit-scatter-hybrid-'));
 try{for(const mutate of [r=>{r.comparisons[2].series[0].maxMarkerBoxDifference=1;},r=>{r.packages[1].sha256='0'.repeat(64);},r=>{r.captures[0].screenshotSha256='0'.repeat(64);},r=>{r.referenceBuilderSha256='0'.repeat(64);}]){
  const bad=structuredClone(original);mutate(bad);const file=path.join(dir,'bad.json');await writeFile(file,JSON.stringify(bad));assert.throws(()=>verify(file));
 }}finally{await rm(dir,{recursive:true,force:true});}
});


test('current builders reproduce the embedded hybrid series and workbook bytes',async()=>{
 const parts=readGeneratedZip(await readFile(base+'scatter-hybrid.pptx'));
 const paint=color=>({type:'solid',color,opacity:1}),stroke=color=>({paint:paint(color),width:2,dash:'solid',beginArrow:'none',endArrow:'none'}),noStroke={paint:{type:'none'},width:1,dash:'solid',beginArrow:'none',endArrow:'none'};
 const data=[{name:'RED',xValues:[-8,null,2,8],yValues:[4,3,-4,6],color:'#FF0000'},{name:'BLUE',xValues:[7,-1,3,-7],yValues:[-6,4,2,-2],color:'#0000FF'}],book=Buffer.from(scatterWorkbook(data));
 for(let i=1;i<=4;i++){
  const xml=parts.get(`ppt/charts/chart${i}.xml`).toString(),actual=xml.match(/<c:ser>.*?<\/c:ser><c:ser>.*?<\/c:ser>/s)?.[0];
  const series=data.map(d=>({...d,line:i===1?false:stroke(d.color),marker:{shape:'circle',size:10,fill:paint(d.color),stroke:noStroke}}));assert.equal(actual,scatterSeriesXml(series));
  const rel=parts.get(`ppt/charts/_rels/chart${i}.xml.rels`).toString(),target=rel.match(/<Relationship\b[^>]*Type="[^"]*\/package"[^>]*Target="([^"]+)"/)?.[1];assert.ok(target);
  const workbook=path.posix.normalize(path.posix.join('ppt/charts',target));assert.deepEqual(parts.get(workbook),book);
 }
});
