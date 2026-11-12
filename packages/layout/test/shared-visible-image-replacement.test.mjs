import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation } from "@pptkit/core";
import { planSharedVisibleImageReplacement, resolveLayout } from "../dist/index.js";
function setup(){
 const d=createPresentation();d.registerAsset({id:"old",kind:"image",source:{type:"url",value:"old"},width:400,height:200});
 d.registerAsset({id:"new",kind:"image",source:{type:"url",value:"new"},width:100,height:300});
 d.defineSlideMaster({id:"m",name:"Master",elements:[{type:"image",id:"master-img",assetId:"old",fit:"contain",box:{x:10,y:20,width:100,height:100}}]});
 d.defineSlideLayout({id:"l",name:"Layout",masterId:"m",showMasterShapes:false,elements:[{type:"group",id:"group",box:{x:0,y:0,width:200,height:200},coordinateSize:{width:100,height:100},transform:{rotation:20},children:[
  {type:"image",id:"layout-img",assetId:"old",fit:"contain",box:{x:0,y:0,width:100,height:100}},
  {type:"image",id:"sibling",assetId:"old",fit:"stretch",box:{x:10,y:10,width:20,height:20}}]}]});
 d.addSlide({id:"s1",layoutId:"l"});d.addSlide({id:"s2",layoutId:"l"});return d;
}
test("master image plan updates shared owner with authored undo and leaves layout untouched",()=>{
 const d=setup(),before=d.masters.find(m=>m.id==="m"),other=d.layouts.find(l=>l.id==="l");
 const p=planSharedVisibleImageReplacement(d,{kind:"master",id:"m"},"master-img","new","stretch");
 assert.equal(p.kind,"master");assert.equal(p.ownerId,"m");assert.notEqual(p.previous,before);assert.ok(Object.isFrozen(p.replacement.elements[0].box));
 assert.deepEqual(p.replacement.elements[0].box,{x:10,y:45,width:100,height:50});assert.deepEqual(p.previous,before);
 d.replaceSlideMaster(p.ownerId,p.replacement);assert.equal(d.masters.find(m=>m.id==="m").elements[0].assetId,"new");
 assert.deepEqual(d.layouts.find(l=>l.id==="l"),other);assert.equal(d.slides[0].elements.length,0);assert.equal(d.slides[1].elements.length,0);
 assert.equal(d.getAssetReferences("old").length,2);d.replaceSlideMaster(p.ownerId,p.previous);assert.deepEqual(d.masters.find(m=>m.id==="m"),before);
});
test("nested layout image plan keeps definition metadata, siblings and group coordinates",()=>{
 const d=setup(),before=d.layouts.find(l=>l.id==="l");
 const p=planSharedVisibleImageReplacement(d,{kind:"layout",id:"l"},"layout-img","new","stretch");
 assert.equal(p.kind,"layout");assert.equal(p.replacement.masterId,"m");assert.equal(p.replacement.showMasterShapes,false);
 assert.deepEqual(p.replacement.elements[0].children[0].box,{x:0,y:25,width:100,height:50});
 assert.deepEqual(p.replacement.elements[0].children[1],before.elements[0].children[1]);
 d.replaceSlideLayout(p.ownerId,p.replacement);const image=resolveLayout(d).layouts.find(l=>l.id==="l").elements[0].children[0];
 assert.equal(image.assetId,"new");assert.deepEqual(image.box,{x:0,y:25,width:100,height:50});
 d.replaceSlideLayout(p.ownerId,p.previous);assert.deepEqual(d.layouts.find(l=>l.id==="l"),before);
});
test("shared plans reject wrong owners, absent images and changing ancestor groups",()=>{
 const d=setup(),before=d.layouts.find(l=>l.id==="l");
 for(const [owner,image,asset] of [[{kind:"master",id:"m"},"layout-img","new"],[{kind:"layout",id:"l"},"master-img","new"],[{kind:"master",id:"bad"},"master-img","new"],[{kind:"layout",id:"l"},"layout-img","bad"]]) assert.throws(()=>planSharedVisibleImageReplacement(d,owner,image,asset,"stretch"));
 const rules=[{type:"align",container:{kind:"layout",id:"l"},elementIds:["group"],horizontal:"end"}];
 assert.throws(()=>planSharedVisibleImageReplacement(d,{kind:"layout",id:"l"},"layout-img","new","stretch",{composition:rules}),/ancestor group frame/);
 assert.deepEqual(d.layouts.find(l=>l.id==="l"),before);
});


test("shared group composition is captured once and only the selected image is planned",()=>{
 const d=setup(),before=d.layouts.find(l=>l.id==="l");let reads=0;
 const composition=[{type:"align",container:{kind:"group",id:"group"},elementIds:["layout-img"],area:{x:20,y:10,width:200,height:200},horizontal:"end",vertical:"center"}];
 const options={get composition(){reads++;return composition}};
 const p=planSharedVisibleImageReplacement(d,{kind:"layout",id:"l"},"layout-img","new","stretch",options);
 assert.equal(reads,1);assert.deepEqual(p.replacement.elements[0].children[0].box,{x:120,y:85,width:100,height:50});
 composition[0].area.x=0;assert.equal(p.replacement.elements[0].children[0].box.x,120);
 assert.deepEqual(p.replacement.elements[0].children[1],before.elements[0].children[1]);assert.deepEqual(d.layouts.find(l=>l.id==="l"),before);
});

test("explicit shared ancestor commit changes the shared group and preserves undo",()=>{
 const d=setup(),before=d.layouts.find(l=>l.id==='l');
 const rules=[{type:'constraints',container:{kind:'layout',id:'l'},elements:[{elementId:'group',x:{value:150},width:{value:300}}]}];
 const composed=resolveLayout(d,{composition:rules}).layouts.find(l=>l.id==='l').elements[0];
 const p=planSharedVisibleImageReplacement(d,{kind:'layout',id:'l'},'layout-img','new','stretch',{composition:rules,commitAncestorGeometry:true});
 assert.deepEqual(p.committedAncestorIds,['group']);assert.ok(Object.isFrozen(p.committedAncestorIds));
 assert.deepEqual(p.replacement.elements[0].children[1],before.elements[0].children[1]);
 d.replaceSlideLayout(p.ownerId,p.replacement);const after=resolveLayout(d).layouts.find(l=>l.id==='l').elements[0];
 assert.deepEqual(after.box,composed.box);assert.deepEqual(after.children[0].box,composed.children[0].box);assert.equal(after.children[0].assetId,'new');
 assert.equal(d.slides.length,2);d.replaceSlideLayout(p.ownerId,p.previous);assert.deepEqual(d.layouts.find(l=>l.id==='l'),before);
 assert.throws(()=>planSharedVisibleImageReplacement(d,{kind:'layout',id:'l'},'layout-img','new','stretch',{composition:rules,commitAncestorGeometry:1}),TypeError);
});
