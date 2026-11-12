import assert from 'node:assert/strict';
import test from 'node:test';
import { inspectScatterData, snapshotScatterData } from '../dist/chart/scatter-data.js';

test('scatter data preserves independent series lengths, order, signed zero and explicit blanks',()=>{
 const input=[{xValues:[1,null,-0],yValues:[3,4,0]},{xValues:[7],yValues:[null]}],before=structuredClone(input);
 assert.deepEqual(inspectScatterData(input),[]);const out=snapshotScatterData(input);assert.equal(out.blankDisplay,'gap');assert.deepEqual(out.series,before);assert.ok(Object.is(out.series[0].xValues[2],-0));
 input[0].xValues[0]=99;assert.equal(out.series[0].xValues[0],1);out.series[0].yValues[0]=100;assert.equal(input[0].yValues[0],3);
 assert.equal(snapshotScatterData([{xValues:[1],yValues:[0]}]).blankDisplay,'zero');for(const mode of ['gap','zero','span'])assert.equal(snapshotScatterData(before,mode).blankDisplay,mode);
});

test('scatter validation locates missing, sparse, malformed and mismatched data',()=>{
 assert.deepEqual(inspectScatterData([]),[{code:'scatter-empty-series'}]);assert.deepEqual(inspectScatterData(new Array(1)),[{code:'scatter-missing-series',seriesIndex:0}]);
 assert.deepEqual(inspectScatterData([{xValues:[],yValues:[]}]),[]);assert.equal(snapshotScatterData([{xValues:[],yValues:[]}]).blankDisplay,'zero');
 assert.deepEqual(inspectScatterData([null,{xValues:null,yValues:[]}]),[{code:'scatter-missing-series',seriesIndex:0},{code:'scatter-invalid-values',seriesIndex:1,coordinate:'xValues'}]);
 const input=[{xValues:[1,,NaN],yValues:[Infinity,undefined]}];assert.deepEqual(inspectScatterData(input),[
  {code:'scatter-values-length',seriesIndex:0},
  {code:'scatter-coordinate',seriesIndex:0,coordinate:'xValues',pointIndex:1},
  {code:'scatter-coordinate',seriesIndex:0,coordinate:'xValues',pointIndex:2},
  {code:'scatter-coordinate',seriesIndex:0,coordinate:'yValues',pointIndex:0},
  {code:'scatter-coordinate',seriesIndex:0,coordinate:'yValues',pointIndex:1},
 ]);assert.throws(()=>snapshotScatterData(input),RangeError);assert.throws(()=>snapshotScatterData([{xValues:[0],yValues:[0]}],'invalid'),RangeError);
});

test('scatter validation checks every index without category-sized allocation',()=>{
 const x=Array.from({length:10000},(_,i)=>i),y=x.map(i=>i%11===0?null:-i);const input=[{xValues:x,yValues:y},{xValues:[1,2],yValues:[3,4]}];
 assert.deepEqual(inspectScatterData(input),[]);const out=snapshotScatterData(input);assert.equal(out.series[0].xValues.length,10000);assert.equal(out.series[1].xValues.length,2);assert.equal(out.blankDisplay,'gap');
});
