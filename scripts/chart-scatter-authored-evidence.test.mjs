import assert from 'node:assert/strict';
import test from 'node:test';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
const base='.tmp/native-scatter-authored/powerpoint-PmQkqF/',record='docs/testing/chart-scatter-authored-powerpoint-evidence.json';
const verify=file=>execFileSync('python3',['scripts/research/verify-scatter-authored-powerpoint.py',...(file?[file]:[])],{encoding:'utf8',stdio:['ignore','pipe','pipe']});
test('authored native scatter matches layout within the measured approximation',()=>{
 execFileSync('node',['scripts/inspect-native-scatter-authored-layout.mjs'],{stdio:'pipe'});const result=JSON.parse(verify());assert.equal(result.charts,4);assert.equal(result.reopenRounds,1);assert.ok(result.maxMarkerDifference<3.2);assert.ok(result.maxLineDifference<3.2);
});
test('authored and PowerPoint-saved scatter project independent coordinates and preserve exact no-op',async()=>{
 for(const file of ['scatter-authored.pptx','scatter-authored-round-1.pptx']){const bytes=await readFile(base+file),s=await importPptx(bytes,{projection:'content'});try{assert.equal(s.document.slides.length,4);for(const slide of s.document.slides){const chart=slide.elements.find(e=>e.type==='chart');assert.equal(chart.chartType,'scatter');assert.deepEqual(chart.series.map(x=>[x.xValues,x.yValues]),[[[-8,null,2,8],[4,3,-4,6]],[[7,-1,3,-7],[-6,4,2,-2]]]);}const plan=await s.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual(Buffer.from((await s.save(plan.plan)).bytes),bytes);}finally{s.dispose();}}
});
test('authored evidence rejects altered measurements, package hashes and captures',async()=>{
 const original=JSON.parse(await readFile(record,'utf8')),dir=await mkdtemp(path.join(tmpdir(),'pptkit-scatter-authored-'));try{for(const mutate of [r=>{r.comparisons[0].series[0].maxMarkerDifference=0;},r=>{r.packages[1].sha256='0'.repeat(64);},r=>{r.captures[0].screenshotSha256='0'.repeat(64);}]){const bad=structuredClone(original);mutate(bad);const file=path.join(dir,'bad.json');await writeFile(file,JSON.stringify(bad));assert.throws(()=>verify(file));}}finally{await rm(dir,{recursive:true,force:true});}
});
