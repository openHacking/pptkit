import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation } from "@pptkit/core";
import { planVisibleImageReplacement, resolveLayout } from "../dist/index.js";

function setup(fit="contain",width=400,height=200){
  const doc=createPresentation();const original=doc.registerAsset({id:"old",kind:"image",source:{type:"url",value:"old"},...(width===undefined?{}:{width}),...(height===undefined?{}:{height})});
  const target=doc.registerAsset({id:"new",kind:"image",source:{type:"url",value:"new"},width:100,height:300});
  const slide=doc.addSlide({id:"slide",elements:[{type:"image",id:"pic",assetId:original.id,fit,box:{x:10,y:20,width:100,height:100}},
    {type:"image",id:"other",assetId:original.id,fit:"stretch",box:{x:150,y:20,width:50,height:50}}]});
  return {doc,slide,target};
}
test("visible-frame plan resolves current contain geometry and restores exact authoring state",()=>{
  const {doc,slide}=setup(),original=slide.elements[0];
  const plan=planVisibleImageReplacement(doc,"slide","pic","new","preserve");
  assert.deepEqual(plan.replacement.box,{x:10,y:45,width:100,height:50});assert.equal(plan.replacement.fit,"stretch");
  assert.equal(plan.replacement.assetId,"new");assert.deepEqual(plan.previous,original);assert.ok(Object.isFrozen(plan.previous.box));
  slide.replaceElement("pic",plan.replacement);assert.deepEqual(resolveLayout(doc).slides[0].elements[0].box,plan.replacement.box);
  assert.equal(slide.elements[1].assetId,"old");slide.replaceElement("pic",plan.previous);assert.deepEqual(slide.elements[0],original);
});
test("visible-frame policies retain source resolved frame while reset or preserve crop explicitly",()=>{
  const {doc,slide}=setup("cover");const visible=resolveLayout(doc).slides[0].elements[0];
  const preserve=planVisibleImageReplacement(doc,"slide","pic","new","preserve");
  assert.deepEqual(preserve.replacement.box,visible.box);assert.equal(preserve.replacement.fit,"crop");assert.deepEqual(preserve.replacement.crop,visible.crop);
  for(const fit of ["stretch","contain","cover"]){
    const plan=planVisibleImageReplacement(doc,"slide","pic","new",fit);assert.equal(plan.replacement.fit,fit);assert.equal("crop" in plan.replacement,false);
    assert.deepEqual(plan.replacement.box,visible.box);
  }
  slide.replaceElement("pic",{...slide.elements[0],box:{x:30,y:20,width:200,height:100}});
  const fresh=planVisibleImageReplacement(doc,"slide","pic","new","stretch");assert.deepEqual(fresh.replacement.box,{x:30,y:20,width:200,height:100});
});
test("visible-frame planning rejects missing identity, target and unresolved original fit",()=>{
  const {doc,slide}=setup();const before=slide.elements;
  for(const args of [["missing","pic","new","stretch"],["slide","missing","new","stretch"],["slide","pic","missing","stretch"],["slide","pic","new","bogus"]]) assert.throws(()=>planVisibleImageReplacement(doc,...args));
  assert.throws(()=>planVisibleImageReplacement(doc,"","pic","new","stretch"),TypeError);
  assert.deepEqual(slide.elements,before);
  const unresolved=setup("contain");unresolved.doc.replaceAsset("old",{kind:"image",source:{type:"url",value:"old"}});
  assert.throws(()=>planVisibleImageReplacement(unresolved.doc,"slide","pic","new","stretch"),/unresolved/);
  const zero=setup("cover");zero.slide.replaceElement("pic",{...zero.slide.elements[0],box:{x:0,y:0,width:0,height:1}});
  assert.throws(()=>planVisibleImageReplacement(zero.doc,"slide","pic","new","stretch"),/unresolved/);
  const zeroStretch=setup("stretch");zeroStretch.slide.replaceElement("pic",{...zeroStretch.slide.elements[0],box:{x:0,y:0,width:0,height:1}});
  assert.throws(()=>planVisibleImageReplacement(zeroStretch.doc,"slide","pic","new","stretch"),/positive visible frame/);
});


test("visible-frame plan evaluates detached composition before fitting and retains authored undo",()=>{
  const {doc,slide}=setup();const before=slide.elements[0];
  const composition=[{type:"align",container:{kind:"slide",id:"slide"},elementIds:["pic"],area:{x:200,y:100,width:200,height:200},horizontal:"end",vertical:"center"}];
  const plan=planVisibleImageReplacement(doc,"slide","pic","new","stretch",{composition});
  assert.deepEqual(plan.replacement.box,{x:300,y:175,width:100,height:50});assert.deepEqual(plan.previous,before);
  composition[0].area.x=0;assert.equal(plan.replacement.box.x,300);
  slide.replaceElement("pic",plan.replacement);assert.deepEqual(resolveLayout(doc).slides[0].elements[0].box,plan.replacement.box);
  slide.replaceElement("pic",plan.previous);assert.deepEqual(slide.elements[0],before);
});

test("visible-frame composition refuses failed rules and unsupported options",()=>{
  const {doc,slide}=setup();const before=slide.elements;
  const missing=[{type:"align",container:{kind:"slide",id:"slide"},elementIds:["missing"],horizontal:"center"}];
  assert.throws(()=>planVisibleImageReplacement(doc,"slide","pic","new","stretch",{composition:missing}),/diagnostics/);
  assert.throws(()=>planVisibleImageReplacement(doc,"slide","pic","new","stretch",{tablePagination:[]}),TypeError);
  let reads=0;const options={get composition(){reads++;return []}};
  planVisibleImageReplacement(doc,"slide","pic","new","stretch",options);assert.equal(reads,1);
  assert.deepEqual(slide.elements,before);
});
