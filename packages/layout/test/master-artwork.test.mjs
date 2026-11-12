import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation } from '@pptkit/core';
import { resolveLayout, resolveNormalizedLayout } from '../dist/index.js';
import { masterArtworkDocument } from '../../../tests/fixtures/master-artwork.mjs';

test('layout retains each master collection once and resolves images/groups without altering source geometry', () => {
  const d=masterArtworkDocument(createPresentation), n=normalizePresentation(d), before=structuredClone(n), r=resolveNormalizedLayout(n);
  assert.deepEqual(n,before);assert.equal(r.masters[0].elements.length,4);assert.equal(r.masters[0].elements[3].children.length,1);
  assert.deepEqual(r.slides.map(s=>s.elements.length),[3,3,3,3,3,3]);assert.equal(r.slides[1].showInheritedShapes,false);assert.equal(r.layouts.find(l=>l.id==='red-off').showMasterShapes,false);
  r.masters[0].elements[0].box.x=800;assert.equal(n.masters[0].elements[0].box.x,20);
});

test('master-owned tables/charts, shared sizing/analysis and composition use their own resolved theme', () => {
  const d=createPresentation({theme:{fonts:{body:'Arial'}}});
  d.defineSlideMaster({id:'m',name:'M',theme:{fonts:{body:'Georgia'}},elements:[
    {id:'mt',type:'text',box:{x:0,y:0,width:100,height:30},content:'Master'},
    {id:'table',type:'table',box:{x:0,y:50,width:100,height:30},columns:[100],rows:[{cells:[{content:'Cell',style:{borders:{left:{width:2}}}}]}]},
    {id:'chart',type:'chart',chartType:'bar',box:{x:150,y:0,width:200,height:150},categories:['A'],series:[{name:'S',values:[1]}]},
  ]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({layoutId:'l'});d.addSlide({layoutId:'l',showInheritedShapes:false});
  const families=[];const r=resolveLayout(d,{tableSizing:[{elementId:'table',autoRows:[{row:0}]}],intrinsicSizing:[{elementId:'mt',height:{type:'content'}}],composition:[{type:'align',container:{kind:'master',id:'m'},elementIds:['mt'],horizontal:'end'}],textMeasurement:{measure(request){families.push(request.paragraphs[0].runs[0].style.fontFamily);return{bounds:{x:0,y:0,width:30,height:20}};}}});
  const master=r.masters[0];assert.ok(families.length>=2);assert.ok(families.every(f=>f==='Georgia'));assert.equal(master.elements[0].box.x,860);
  assert.ok(master.elements[1].borderGrid);assert.ok(master.elements[2].chartLayout);assert.ok(r.analysis.measurements.some(m=>m.masterId==='m'&&m.elementId==='mt'));
  assert.equal(r.slides[0].elements.length,0);assert.equal(r.slides[1].elements.length,0);
});

test('slide and layout connectors resolve against master geometry even when the inherited layer is hidden', () => {
  const d=createPresentation({masters:[{id:'m',name:'M',elements:[{id:'target',type:'shape',shape:'rect',box:{x:20,y:30,width:40,height:50}}]}]});
  d.defineSlideLayout({id:'l',name:'L',masterId:'m',showMasterShapes:false,elements:[{id:'layout-link',type:'connector',start:{elementId:'target',anchor:'right'},end:{x:200,y:200}}]});
  d.addSlide({layoutId:'l',showInheritedShapes:false,elements:[{id:'slide-link',type:'connector',start:{elementId:'target',anchor:'bottom'},end:{x:200,y:200}}]});
  const r=resolveLayout(d);assert.deepEqual(r.layouts.find(l=>l.id==='l').elements[0].start,{x:60,y:55});assert.deepEqual(r.slides[0].elements[0].start,{x:40,y:80});
});
