import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation } from '../dist/index.js';
import { masterArtworkDocument } from '../../../tests/fixtures/master-artwork.mjs';
const box = {x:0,y:0,width:100,height:40};
const shape = id => ({ ...(id===undefined?{}:{id}),type:'shape',shape:'rect',box });
const invalid = code => e => e.name==='PresentationValidationError' && e.diagnostics.some(d=>d.code===code);

test('masters own immutable registered artwork without copying objects into layouts or slides', () => {
  const d=masterArtworkDocument(createPresentation), r=resolvePresentationInheritance(d), n=r.presentation;
  assert.deepEqual(n,normalizePresentation(d)); assert.equal(d.masters[0].elements.length,4); assert.ok(Object.isFrozen(d.masters[0].elements[3].children[0].box));
  assert.deepEqual(n.slides.map(s=>s.elements.length),[3,3,3,3,3,3]); assert.deepEqual(n.layouts.filter(l=>l.masterId).map(l=>l.elements.length),[3,3,3]);
  assert.equal(n.masters[0].elements[0].id,'red-master-shape'); assert.equal(n.masters[1].elements[0].id,'blue-master-shape');
  assert.equal(r.origins.find(o=>o.target.id==='red-master-text'&&o.property==='content.0.runs.0.style.color').source.path,'masters.0.elements.1.content.0.runs.0.style.color');
  n.masters[0].elements[0].box.x=99; assert.equal(d.masters[0].elements[0].box.x,20);
});

test('initial master object IDs participate in the global store, including nested/automatic identities', () => {
  const input=[{id:'m',name:'M',elements:[shape(undefined),{id:'g',type:'group',box,coordinateSize:{width:100,height:40},children:[shape('child')]}]}];
  const d=createPresentation({masters:input});input[0].elements[0].box.x=70;
  assert.equal(d.masters[0].elements[0].id,'element-1');assert.equal(d.masters[0].elements[0].box.x,0);
  assert.throws(()=>d.addSlide().addElement(shape('child')),/Duplicate/);
  d.defineSlideLayout({id:'l',name:'L'});assert.throws(()=>d.replaceSlideLayout('l',{name:'Collision',elements:[shape('g')]}),/Duplicate/);
  assert.throws(()=>createPresentation({masters:[{id:'a',name:'A',elements:[shape('same')]},{id:'b',name:'B',elements:[shape('same')]}]}),/Duplicate/);
});

test('master replacements stage full collections and allocator rollback, then release only the replaced ownership', () => {
  const d=createPresentation({masters:[{id:'m',name:'M',elements:[shape('old')]}]});const original=d.masters[0];
  assert.throws(()=>d.replaceSlideMaster('m',{name:'Bad',elements:[{type:'text',content:'Missing frame'}]}),invalid('missing-element-box'));
  assert.equal(d.masters[0],original);
  d.replaceSlideMaster('m',{name:'Good',elements:[shape(undefined)]});assert.equal(d.masters[0].elements[0].id,'element-1');
  assert.equal(d.addSlide().addElement(shape('old')).id,'old');assert.throws(()=>d.addSlide().addElement(shape('element-1')),/Duplicate/);
  d.replaceSlideMaster('m',{name:'No artwork'});assert.equal(d.masters[0].elements,undefined);assert.equal(d.addSlide().addElement(shape('element-1')).id,'element-1');
});

test('master edits reject missing resources, invalid prompt bindings and broken dependent connectors atomically', () => {
  const d=createPresentation({masters:[{id:'m',name:'M',elements:[shape('anchor')]}]});d.defineSlideLayout({id:'layout-blank',name:'Default',masterId:'m'});
  d.addSlide({elements:[{id:'link',type:'connector',start:{elementId:'anchor'},end:{x:200,y:100}}]});
  const old=d.masters[0];
  assert.throws(()=>d.replaceSlideMaster('m',{name:'Break anchor'}),invalid('missing-connector-target'));assert.equal(d.masters[0],old);
  assert.throws(()=>d.replaceSlideMaster('m',{name:'Missing image',elements:[{id:'logo',type:'image',assetId:'absent',box}]}),invalid('missing-image-asset'));assert.equal(d.masters[0],old);
  assert.throws(()=>d.replaceSlideMaster('m',{name:'Prompt',elements:[{id:'prompt',type:'text',content:'X',placeholderKey:'body'}]}),invalid('missing-placeholder'));
  assert.equal(d.masters[0],old);d.replaceSlideMaster('m',{name:'Moved',elements:[{...shape('anchor'),box:{...box,x:30}}]});assert.equal(normalizePresentation(d).masters[0].elements[0].box.x,30);
});

test('master asset references protect hidden and unused shared resources and update after replacement', () => {
  const d=masterArtworkDocument(createPresentation), refs=d.getAssetReferences('logo');
  assert.deepEqual(refs.filter(r=>r.masterId).map(r=>r.path),['masters.0.elements.2.assetId','masters.1.elements.2.assetId']);
  assert.throws(()=>d.removeAsset('logo'),/referenced/);
  d.registerAsset({id:'unused-logo',kind:'image',source:d.getAsset('logo').source});d.defineSlideMaster({id:'unused',name:'Unused',elements:[{id:'hidden-unused',type:'image',assetId:'unused-logo',box,hidden:true}]});
  assert.equal(d.getAssetReferences('unused-logo')[0].masterId,'unused');assert.throws(()=>d.removeAsset('unused-logo'),/referenced/);assert.equal(d.removeUnusedAssets().length,0);
  const {elements:_art,...without}=d.masters[0];d.replaceSlideMaster('red',without);
  assert.equal(d.getAssetReferences('logo').filter(r=>r.masterId).length,1);assert.equal(d.getAsset('logo').id,'logo');
  assert.ok(Object.isFrozen(refs[0]));
});

test('visibility declarations preserve false/true/absence, copies and strict managed slide setters', () => {
  const d=masterArtworkDocument(createPresentation), n=normalizePresentation(d);
  assert.equal(n.slides[0].showInheritedShapes,undefined);assert.equal(n.slides[1].showInheritedShapes,false);assert.equal(n.slides[4].showInheritedShapes,true);
  const old=d.slides[1];assert.equal(d.duplicateSlide(old.id).showInheritedShapes,false);
  old.setShowInheritedShapes(true);assert.equal(old.showInheritedShapes,true);old.setShowInheritedShapes();assert.equal(normalizePresentation(d).slides[1].showInheritedShapes,undefined);
  for(const bad of [null,0,'false',{}]) {assert.throws(()=>old.setShowInheritedShapes(bad),TypeError);assert.equal(old.showInheritedShapes,undefined);}
  const invalidDoc=createPresentation();invalidDoc.defineSlideLayout({id:'bad',name:'Bad',showMasterShapes:'no'});invalidDoc.addSlide({showInheritedShapes:0});
  assert.equal(validatePresentation(invalidDoc).filter(d=>d.code==='invalid-artwork-visibility').length,2);
});

test('master definition capture is plain-data only and rejects malformed artwork without invoking accessors', () => {
  const d=createPresentation({masters:[{id:'m',name:'M'}]});let calls=0;
  assert.throws(()=>d.replaceSlideMaster('m',{name:'Bad',get elements(){calls++;return[];}}),TypeError);assert.equal(calls,0);
  for(const elements of [null,{},new Array(1),[{type:'unknown'}],[{type:'group',children:null}]]) assert.throws(()=>d.replaceSlideMaster('m',{name:'Bad',elements}));
  assert.equal(d.masters[0].name,'M');assert.equal(d.masters[0].elements,undefined);
});
