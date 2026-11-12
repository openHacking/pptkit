import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation } from "@pptkit/core";
import { planNestedVisibleImageReplacement, resolveLayout } from "../dist/index.js";

function setup(){
  const doc=createPresentation();doc.registerAsset({id:"old",kind:"image",source:{type:"url",value:"old"},width:400,height:200});
  doc.registerAsset({id:"new",kind:"image",source:{type:"url",value:"new"},width:100,height:300});
  const slide=doc.addSlide({id:"slide",elements:[{id:"outer",type:"group",box:{x:100,y:40,width:400,height:400},coordinateSize:{width:200,height:200},transform:{rotation:30,flipH:true},children:[
    {id:"inner",type:"group",box:{x:0,y:0,width:200,height:200},coordinateSize:{width:200,height:200},children:[
      {id:"pic",type:"image",assetId:"old",fit:"contain",box:{x:10,y:20,width:100,height:100}},
      {id:"sibling",type:"image",assetId:"old",fit:"stretch",box:{x:120,y:20,width:50,height:50}}]},
    {id:"decoration",type:"shape",shape:"rect",box:{x:0,y:150,width:20,height:20}}]}]});
  return {doc,slide};
}
test("nested plan retains parent-local resolved frame and untouched group siblings with atomic undo",()=>{
  const {doc,slide}=setup(),before=slide.elements[0];
  const plan=planNestedVisibleImageReplacement(doc,"slide","pic","new","stretch");
  assert.equal(plan.rootElementId,"outer");assert.equal(plan.imageId,"pic");assert.notEqual(plan.previous,before);
  assert.ok(Object.isFrozen(plan.replacement.children[0].children[0].box));
  assert.deepEqual(plan.replacement.children[0].children[0].box,{x:10,y:45,width:100,height:50});
  assert.equal(plan.replacement.children[0].children[0].assetId,"new");
  assert.deepEqual(plan.replacement.children[0].children[1],before.children[0].children[1]);
  assert.deepEqual(plan.replacement.children[1],before.children[1]);assert.deepEqual(plan.replacement.transform,before.transform);
  slide.replaceElement(plan.rootElementId,plan.replacement);
  const image=resolveLayout(doc).slides[0].elements[0].children[0].children[0];assert.deepEqual(image.box,{x:10,y:45,width:100,height:50});
  slide.replaceElement(plan.rootElementId,plan.previous);assert.deepEqual(slide.elements[0],before);
});
test("nested composition applies in parent coordinates but refuses ancestor transforms",()=>{
  const {doc,slide}=setup();const before=slide.elements[0];
  const childRule=[{type:"align",container:{kind:"group",id:"inner"},elementIds:["pic"],area:{x:0,y:0,width:200,height:200},horizontal:"end",vertical:"center"}];
  const plan=planNestedVisibleImageReplacement(doc,"slide","pic","new","stretch",{composition:childRule});
  assert.deepEqual(plan.replacement.children[0].children[0].box,{x:100,y:75,width:100,height:50});
  const parentRule=[{type:"align",container:{kind:"slide",id:"slide"},elementIds:["outer"],horizontal:"end"}];
  assert.throws(()=>planNestedVisibleImageReplacement(doc,"slide","pic","new","stretch",{composition:parentRule}),/ancestor group frame/);
  assert.deepEqual(slide.elements[0],before);
});
test("nested planning rejects unsupported identity, missing assets and unresolved fit without mutation",()=>{
  const {doc,slide}=setup(),before=slide.elements;
  for(const args of [["slide","outer","new","stretch"],["slide","missing","new","stretch"],["slide","pic","missing","stretch"],["missing","pic","new","stretch"],["slide","pic","new","bad"]]) assert.throws(()=>planNestedVisibleImageReplacement(doc,...args));
  doc.replaceAsset("old",{kind:"image",source:{type:"url",value:"old"}});
  assert.throws(()=>planNestedVisibleImageReplacement(doc,"slide","pic","new","stretch"),/unresolved/);assert.deepEqual(slide.elements,before);
});

test("explicit ancestor geometry commit preserves composed nested frame and moves siblings atomically",()=>{
 const {doc,slide}=setup(),before=slide.elements[0];
 const rules=[
  {type:'constraints',container:{kind:'slide',id:'slide'},elements:[{elementId:'outer',x:{value:250},width:{value:300}}]},
  {type:'constraints',container:{kind:'group',id:'outer'},elements:[{elementId:'inner',x:{value:20},width:{value:150}}]},
 ];
 const composed=resolveLayout(doc,{composition:rules}).slides[0].elements[0],plan=planNestedVisibleImageReplacement(doc,'slide','pic','new','stretch',{composition:rules,commitAncestorGeometry:true});
 assert.deepEqual(plan.committedAncestorIds,['outer','inner']);assert.ok(Object.isFrozen(plan.committedAncestorIds));
 assert.deepEqual(plan.replacement.children[0].children[1],before.children[0].children[1]);assert.deepEqual(plan.replacement.children[1],before.children[1]);
 slide.replaceElement(plan.rootElementId,plan.replacement);const result=resolveLayout(doc).slides[0].elements[0];
 assert.deepEqual(result.box,composed.box);assert.deepEqual(result.transform,composed.transform);
 assert.deepEqual(result.children[0].box,composed.children[0].box);assert.deepEqual(result.children[0].children[0].box,composed.children[0].children[0].box);
 assert.equal(result.children[0].children[0].assetId,'new');
 slide.replaceElement(plan.rootElementId,plan.previous);assert.deepEqual(slide.elements[0],before);
 assert.throws(()=>planNestedVisibleImageReplacement(doc,'slide','pic','new','stretch',{composition:rules,commitAncestorGeometry:'yes'}),TypeError);
});
