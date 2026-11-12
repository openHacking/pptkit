import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { resolveLayout } from '../packages/layout/dist/index.js';

const base='.tmp/native-area-authored/powerpoint-ohuZBk/';
const recordPath='docs/testing/chart-area-authored-powerpoint-evidence.json';
const verify=(record)=>execFileSync('python3',['scripts/research/verify-area-authored-powerpoint.py',...(record?[record]:[])],{encoding:'utf8',stdio:['ignore','pipe','pipe']});

test('native evidence checks first and third prints and all saved packages',()=>{
 const result=JSON.parse(verify());assert.equal(result.reopenRounds,3);assert.equal(result.maxReopenVertexDifference,0);assert.ok(result.maxSlidePointDifference<.8);
});

test('PowerPoint saved charts retain source data and no-op bytes in each round',async()=>{
 for(const name of ['area-authored.pptx','area-authored-round-1.pptx','area-authored-round-2.pptx','area-authored-round-3.pptx']){
  const bytes=await readFile(base+name),session=await importPptx(bytes,{projection:'content'});
  try{
   const charts=session.document.slides.map(slide=>slide.elements.find(e=>e.type==='chart'));
   assert.deepEqual(charts.map(c=>[c.chartType,c.grouping,c.blankDisplay]),[['area','standard','gap'],['area','standard','zero'],['area','stacked','zero'],['area','percentStacked','zero']]);
   for(const chart of charts)assert.deepEqual(chart.series.map(s=>s.values),[[4,2,null,-2,3],[2,-3,3,4,1]]);
   const plan=await session.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes),bytes);
  }finally{session.dispose();}
 }
});

test('recorded layout is recomputed from current public importer and layout engine',async()=>{
 const expected=JSON.parse(await readFile(base+'layout.json','utf8')),session=await importPptx(await readFile(base+'area-authored.pptx'),{projection:'content'});
 try{
  const slides=resolveLayout(session.document).slides.map(s=>{const e=s.elements.find(e=>e.type==='chart');return {box:e.box,grouping:e.grouping,blankDisplay:e.blankDisplay,plotBox:e.chartLayout.plotBox,valueScale:e.chartLayout.valueScale,areas:e.chartLayout.areas};});
  assert.deepEqual(slides,expected.slides);
 }finally{session.dispose();}
});

test('evidence rejects altered print geometry, saved workbook records and capture hashes',async()=>{
 const evidence=JSON.parse(await readFile(recordPath,'utf8')),dir=await mkdtemp(path.join(tmpdir(),'pptkit-area-evidence-'));
 try{for(const mutate of [r=>{r.comparisons[4].polygons[0].maxSlidePointDifference+=1;},r=>{r.packages[3].charts[0].workbookPart='missing.xlsx';},r=>{r.captures.at(-1).screenshotSha256='0'.repeat(64);},r=>{r.reopen3PdfSha256='0'.repeat(64);}]){
  const bad=structuredClone(evidence);mutate(bad);const file=path.join(dir,'bad.json');await writeFile(file,JSON.stringify(bad));assert.throws(()=>verify(file));
 }}finally{await rm(dir,{recursive:true,force:true});}
});
