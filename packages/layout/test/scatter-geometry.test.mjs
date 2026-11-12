import assert from 'node:assert/strict';
import test from 'node:test';
import { scatterGeometry } from '../dist/chart/scatter-geometry.js';

const axis={min:-10,max:10,majorUnit:5,ticks:[-10,-5,0,5,10]},plot={x:10,y:20,width:200,height:100};
const series=[{xValues:[-8,-3,2,8],yValues:[4,null,-4,6]},{xValues:[-7,-1,3,7],yValues:[-2,4,2,-6]}];
test('scatter maps independent numeric X and Y values with gap, zero and span',()=>{
 for(const [mode,indices,segments] of [['gap',[0,2,3],[[0],[2,3]]],['zero',[0,1,2,3],[[0,1,2,3]]],['span',[0,2,3],[[0,2,3]]]]){
  const result=scatterGeometry(series,plot,axis,axis,mode);assert.deepEqual(result[0].points.map(p=>p.sourceIndex),indices);assert.deepEqual(result[0].segments.map(s=>s.map(p=>p.sourceIndex)),segments);
  for(const p of result[0].points){assert.equal(p.x,10+(series[0].xValues[p.sourceIndex]+10)*10);assert.equal(p.y,120-((series[0].yValues[p.sourceIndex]??0)+10)*5);assert.equal(p.blank,mode==='zero'&&p.sourceIndex===1?true:undefined);}
  assert.deepEqual(result[1].points.map(p=>p.sourceIndex),[0,1,2,3]);assert.notEqual(result[0].points[0].x,result[1].points[0].x);
 }
});
test('scatter retains point order and explicit blank identity',()=>{
 const data=[{xValues:[2,-1,2],yValues:[3,null,-5]}],original=structuredClone(data),out=scatterGeometry(data,plot,axis,axis,'zero');
 assert.deepEqual(out[0].points.map(p=>p.sourceIndex),[0,1,2]);assert.equal(out[0].points[0].x,out[0].points[2].x);assert.deepEqual(data,original);
 const before=out[0].segments[0][0].x;out[0].points[0].x=999;assert.equal(out[0].segments[0][0].x,before);assert.deepEqual(data,original);
});
test('scatter rejects malformed arrays, implicit holes and unrepresentable geometry',()=>{
 for(const input of [[{xValues:[1],yValues:[]}],[{xValues:[undefined],yValues:[1]}],[{xValues:[1],yValues:[undefined]}],[{xValues:[1],yValues:[NaN]}],[{xValues:new Array(1),yValues:[2]}]])assert.throws(()=>scatterGeometry(input,plot,axis,axis,'gap'),RangeError);
 assert.throws(()=>scatterGeometry(series,plot,{...axis,min:0,max:0},axis,'gap'),RangeError);
 assert.throws(()=>scatterGeometry([{xValues:[1],yValues:[null]}],plot,axis,{...axis,min:0,max:0},'gap'),RangeError);
 assert.throws(()=>scatterGeometry(new Array(1),plot,axis,axis,'gap'),RangeError);
 assert.throws(()=>scatterGeometry([{xValues:[Number.MAX_VALUE],yValues:[0]}],plot,{...axis,min:0,max:Number.MIN_VALUE},axis,'gap'),RangeError);
 assert.throws(()=>scatterGeometry(series,plot,axis,axis,'invalid'),RangeError);
});
test('scatter handles opposite floating point extrema without overflow',()=>{
 const extreme={min:-Number.MAX_VALUE,max:Number.MAX_VALUE,majorUnit:1,ticks:[]},result=scatterGeometry([{xValues:[-Number.MAX_VALUE,0,Number.MAX_VALUE],yValues:[-Number.MAX_VALUE,0,Number.MAX_VALUE]}],plot,extreme,extreme,'gap');
 assert.deepEqual(result[0].points.map(({x,y})=>[x,y]),[[10,120],[110,70],[210,20]]);
});

test('scatter leaves representable off-axis points for clipping without distorting lines',()=>{
 const out=scatterGeometry([{xValues:[-20,20],yValues:[-20,20]}],plot,axis,axis,'gap')[0];
 assert.deepEqual(out.points.map(({x,y})=>[x,y]),[[-90,170],[310,-30]]);assert.equal(out.segments[0].length,2);
});

test('scatter missing X honors gap, zero and span and retains unsorted series order',()=>{
 const data=[{xValues:[-8,null,2,8],yValues:[4,3,-4,6]},{xValues:[7,-1,3,-7],yValues:[-6,4,2,-2]}];
 for(const mode of ['gap','zero','span']){
  const out=scatterGeometry(data,plot,axis,axis,mode);assert.deepEqual(out[0].points.map(p=>p.sourceIndex),mode==='zero'?[0,1,2,3]:[0,2,3]);
  assert.deepEqual(out[0].segments.map(s=>s.map(p=>p.sourceIndex)),mode==='gap'?[[0],[2,3]]:mode==='zero'?[[0,1,2,3]]:[[0,2,3]]);
  if(mode==='zero'){assert.equal(out[0].points[1].x,110);assert.equal(out[0].points[1].y,55);assert.equal(out[0].points[1].blank,true);}
  assert.deepEqual(out[1].points.map(p=>p.sourceIndex),[0,1,2,3]);assert.deepEqual(out[1].points.map(p=>p.x),[180,100,140,40]);
 }
});

test('leading, trailing and both-coordinate blanks preserve source positions without extrapolation',()=>{
 const data=[{xValues:[null,1,2,null],yValues:[null,3,4,null]}];
 for(const mode of ['gap','span']){
  const out=scatterGeometry(data,plot,axis,axis,mode)[0];assert.deepEqual(out.points.map(p=>p.sourceIndex),[1,2]);assert.deepEqual(out.segments.map(s=>s.map(p=>p.sourceIndex)),[[1,2]]);
 }
 const zero=scatterGeometry(data,plot,axis,axis,'zero')[0];assert.deepEqual(zero.points.map(p=>p.sourceIndex),[0,1,2,3]);assert.deepEqual(zero.points.filter(p=>p.blank).map(p=>[p.x,p.y]),[[110,70],[110,70]]);
});
