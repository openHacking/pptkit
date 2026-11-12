import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { inspectScatterData, snapshotScatterData } from '../packages/core/dist/chart/scatter-data.js';
import { scatterGeometry } from '../packages/layout/dist/chart/scatter-geometry.js';

test('internal scatter data and geometry match Node and offline Chromium',async t=>{
 const sources=await Promise.all(['packages/core/dist/chart/scatter-data.js','packages/layout/dist/chart/scatter-geometry.js'].map(f=>readFile(f,'utf8')));
 const axis={min:-10,max:10,majorUnit:5,ticks:[]},plot={x:10,y:20,width:200,height:100};
 const data=[{xValues:[-8,null,2,8],yValues:[4,3,-4,6]},{xValues:[7,-1,3,-7],yValues:[-6,4,2,-2]},{xValues:[],yValues:[]}];
 const expected=['gap','zero','span'].map(mode=>{const snap=snapshotScatterData(data,mode);return {snap,layout:scatterGeometry(snap.series,plot,axis,axis,snap.blankDisplay)};});
 const malformed=[{xValues:[1,,NaN],yValues:[Infinity,undefined]}],issues=inspectScatterData(malformed);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort();});
 const actual=await page.evaluate(async ({sources,data,malformed,axis,plot})=>{
  const [core,layout]=await Promise.all(sources.map(source=>import(`data:text/javascript;charset=utf-8,${encodeURIComponent(source)}`)));
  return {results:['gap','zero','span'].map(mode=>{const snap=core.snapshotScatterData(data,mode);return {snap,layout:layout.scatterGeometry(snap.series,plot,axis,axis,snap.blankDisplay)};}),issues:core.inspectScatterData(malformed)};
 },{sources,data,malformed,axis,plot});
 assert.deepEqual(actual.results,expected);assert.deepEqual(actual.issues,issues);assert.deepEqual(requests,[]);
});

test('internal scatter workbook byte output matches Node and offline Chromium',async t=>{
 const { rollup }=await import('rollup'),{ nodeResolve }=await import('@rollup/plugin-node-resolve'),{ scatterWorkbook,scatterWorkbookRanges }=await import('../packages/pptx-exporter/dist/ooxml/scatter-workbook.js');
 const bundle=await rollup({input:'packages/pptx-exporter/dist/ooxml/scatter-workbook.js',plugins:[nodeResolve({browser:true})]});
 const code=(await bundle.generate({format:'es'})).output[0].code;await bundle.close();
 const data=[{name:'First',xValues:[-8,null,-0],yValues:[4,3,null]},{name:'Second',xValues:[7,-1],yValues:[-6,4]}];
 const expected={bytes:[...scatterWorkbook(data)],ranges:scatterWorkbookRanges(data)};
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort();});
 const actual=await page.evaluate(async ({code,data})=>{const m=await import(`data:text/javascript;charset=utf-8,${encodeURIComponent(code)}`);return {bytes:[...m.scatterWorkbook(data)],ranges:m.scatterWorkbookRanges(data)};},{code,data});
 assert.deepEqual(actual,expected);assert.deepEqual(requests,[]);
});
