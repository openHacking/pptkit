import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, planImageReplacement } from "../dist/index.js";

function setup() {
  const doc=createPresentation();
  const old=doc.registerAsset({id:"old",kind:"image",source:{type:"url",value:"old"},width:400,height:200});
  const target=doc.registerAsset({id:"new",kind:"image",source:{type:"url",value:"new"},width:100,height:300});
  const slide=doc.addSlide({elements:[{id:"one",type:"image",assetId:old.id,fit:"crop",crop:{left:.1,right:.2,top:.3,bottom:.1},box:{x:2,y:3,width:100,height:80}}, {id:"two",type:"image",assetId:old.id,fit:"contain",box:{x:200,y:0,width:100,height:100}}]});
  return {doc,slide,target};
}
test("per-image replacement retains crop, shared asset isolation and undo snapshot",()=>{
  const {doc,slide,target}=setup(), before=slide.elements[0];
  const plan=planImageReplacement(before,target,{fit:"preserve"});
  assert.notEqual(plan.previous,before);assert.notEqual(plan.replacement.box,before.box);
  assert.ok(Object.isFrozen(plan.replacement.crop));assert.equal(plan.replacement.id,"one");
  assert.deepEqual(plan.replacement.crop,before.crop);assert.equal(plan.replacement.assetId,"new");
  assert.equal(slide.elements[0],before);slide.replaceElement("one",plan.replacement);
  assert.equal(slide.elements[1].assetId,"old");assert.equal(doc.getAsset("old").id,"old");
  slide.replaceElement("one",plan.previous);assert.deepEqual(slide.elements[0],before);
});
test("explicit fit and size policies reset crop but retain authored position and identity",()=>{
  const {slide,target}=setup(),image=slide.elements[0];
  for(const fit of ["stretch","contain","cover"]){
    const plan=planImageReplacement(image,target,{fit});assert.equal(plan.replacement.fit,fit);
    assert.equal("crop" in plan.replacement,false);assert.deepEqual(plan.replacement.box,image.box);
  }
  const result=planImageReplacement(image,target,{fit:"size",width:42,height:17});
  assert.equal(result.replacement.fit,"stretch");assert.deepEqual(result.replacement.box,{x:2,y:3,width:42,height:17});
  assert.equal("crop" in result.replacement,false);
});
test("invalid policies, dimensions, crops and uncloneable state reject without mutation",()=>{
  const {slide,target}=setup(),image=slide.elements[0],baseline=slide.elements;
  const dimensionless={...target,width:undefined,height:undefined};
  for(const fit of ["contain","cover"]) assert.throws(()=>planImageReplacement(image,dimensionless,{fit}),RangeError);
  assert.throws(()=>planImageReplacement({...image,fit:"contain"},dimensionless,{fit:"preserve"}),RangeError);
  assert.throws(()=>planImageReplacement(image,{...target,width:Number.MAX_VALUE,height:Number.MIN_VALUE},{fit:"cover"}),RangeError);
  assert.throws(()=>planImageReplacement({...image,box:{x:0,y:0,width:1e-200,height:1}}, {...target,width:1e200,height:1}, {fit:"cover"}),RangeError);
  assert.throws(()=>planImageReplacement({...image,box:{x:0,y:0,width:1,height:1}}, {...target,width:1e20,height:1}, {fit:"cover"}),RangeError);
  assert.throws(()=>planImageReplacement({...image,box:{...image.box,width:0}},target,{fit:"contain"}),RangeError);
  for(const size of [0,-1,NaN,Infinity]) assert.throws(()=>planImageReplacement(image,target,{fit:"size",width:size,height:1}),RangeError);
  for(const crop of [{left:1},{left:.5,right:.5},{top:NaN}]) assert.throws(()=>planImageReplacement({...image,crop},target,{fit:"preserve"}),RangeError);
  assert.throws(()=>planImageReplacement(image,target,{fit:"unknown"}),TypeError);
  assert.throws(()=>planImageReplacement({...image,box:{...image.box,width:-1}},target,{fit:"stretch"}),RangeError);
  assert.throws(()=>planImageReplacement(image,{...target,id:""},{fit:"stretch"}),TypeError);
  assert.throws(()=>planImageReplacement({...image,customData:{bad:()=>{}}},target,{fit:"stretch"}));
  assert.deepEqual(slide.elements,baseline);
});
