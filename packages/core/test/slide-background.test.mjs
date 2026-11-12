import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation, validatePresentation, resolvePresentationInheritance } from "../dist/index.js";

test("managed slide background snapshots local overrides and can restore inheritance atomically",()=>{
 const d=createPresentation();d.defineSlideMaster({id:"m",name:"Master",background:{type:"solid",color:"112233"}});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m",background:{type:"solid",color:"445566"}});
 const s=d.addSlide({layoutId:"l"}),paint={type:"solid",color:"AABBCC",opacity:.5};s.setBackground(paint);paint.color="000000";
 assert.equal(s.background.color,"AABBCC");assert.ok(Object.isFrozen(s.background));assert.equal(normalizePresentation(d).slides[0].background.color,"AABBCC");
 const before=s.background;assert.throws(()=>s.setBackground(new Proxy({},{ownKeys(){throw Error("snapshot failed")}})),{name:"DataCloneError"});assert.equal(s.background,before);
 s.setBackground();assert.equal(s.background,undefined);const resolved=normalizePresentation(d).slides[0];assert.equal(resolved.background.color,"445566");assert.equal(resolved.backgroundSource,"layout");
});

test("linear background gradients preserve ordered stops, ownership and inheritance",()=>{
 const source={type:"linearGradient",angle:45,scaled:true,stops:[{position:0,color:"FF0000"},{position:.5,color:{theme:"accent1"},opacity:.25},{position:1,color:"0000FF"}]};
 const d=createPresentation();d.defineSlideMaster({id:"m",name:"Master",background:source});d.defineSlideLayout({id:"l",name:"Layout",masterId:"m"});const s=d.addSlide({layoutId:"l"});source.stops[0].color="000000";
 let normalized=normalizePresentation(d).slides[0];assert.equal(normalized.backgroundSource,"master");assert.equal(normalized.background.stops[0].color,"FF0000");assert.equal(normalized.background.stops[0].opacity,1);assert.equal(normalized.background.scaled,true);
 assert.ok(resolvePresentationInheritance(d).origins.some(o=>o.target.kind==="slide" && o.property==="background.type" && o.source.kind==="master"));
 s.setBackground({type:"solid",color:"123456"});assert.equal(normalizePresentation(d).slides[0].background.type,"solid");s.setBackground();assert.equal(normalizePresentation(d).slides[0].background.type,"linearGradient");
});

test("invalid linear backgrounds are diagnosed without changing other paint domains",()=>{
 const d=createPresentation(),slide=d.addSlide();const valid={type:"linearGradient",angle:0,stops:[{position:0,color:"000000"},{position:1,color:"FFFFFF"}]};
 for(const value of [{...valid,angle:360},{...valid,angle:NaN},{...valid,scaled:"yes"},{...valid,stops:[]},{...valid,stops:[{position:.8,color:"000000"},{position:.2,color:"FFFFFF"}]},{...valid,stops:[{position:0,color:"bad"},{position:1,color:"FFFFFF"}]},{...valid,stops:[{position:0,color:"000000",opacity:2},{position:1,color:"FFFFFF"}]}]) {slide.setBackground(value);assert.ok(validatePresentation(d).some(x=>x.code==="invalid-paint" && x.path.endsWith("background")));}
 slide.setBackground(valid);assert.ok(!validatePresentation(d).some(x=>x.code==="invalid-paint"));
});

test("gradient provenance identifies owner, defaults and replacement without stale solid fields",()=>{
 const d=createPresentation();const grad={type:'linearGradient',angle:45,stops:[{position:0,color:'FF0000'},{position:1,color:'0000FF'}]};
 d.defineSlideMaster({id:'m',name:'Master',background:grad});d.defineSlideLayout({id:'l',name:'Layout',masterId:'m'});const slide=d.addSlide({id:'s',layoutId:'l'});
 const origins=()=>resolvePresentationInheritance(d).origins.filter(x=>x.target.kind==='slide' && x.target.id==='s');
 let values=origins();for(const [property,path] of [['background.angle','masters.0.background.angle'],['background.stops','masters.0.background.stops']]) assert.deepEqual(values.find(x=>x.property===property).source,{kind:'master',id:'m',path});
 assert.deepEqual(values.find(x=>x.property==='background.scaled').source,{kind:'default',path:'defaults.background.linearGradient.scaled'});assert.ok(!values.some(x=>x.property==='background.color'||x.property==='background.opacity'));
 slide.setBackground({...grad,scaled:true});values=origins();assert.deepEqual(values.find(x=>x.property==='background.scaled').source,{kind:'slide',id:'s',path:'slides.0.background.scaled'});assert.deepEqual(values.find(x=>x.property==='background.stops').source,{kind:'slide',id:'s',path:'slides.0.background.stops'});
 slide.setBackground({type:'solid',color:'123456'});values=origins();assert.ok(!values.some(x=>['background.angle','background.scaled','background.stops'].includes(x.property)));assert.equal(values.find(x=>x.property==='background.color').source.kind,'slide');
 slide.setBackground();assert.equal(origins().find(x=>x.property==='background.angle').source.kind,'master');
});

test('gradient stop origins distinguish inherited declarations, explicit zero and omitted opacity',()=>{
 const d=createPresentation();const grad={type:'linearGradient',angle:0,stops:[{position:0,color:{theme:'accent1'}},{position:.5,color:'FF0000',opacity:0},{position:.5,color:'0000FF',opacity:.25}]};
 d.defineSlideMaster({id:'m',name:'Master',background:grad});d.defineSlideLayout({id:'l',name:'Layout',masterId:'m'});const slide=d.addSlide({id:'s',layoutId:'l'});
 const report=resolvePresentationInheritance(d);assert.deepEqual(report.presentation,normalizePresentation(d));assert.ok(Object.isFrozen(report.origins));const rows=report.origins.filter(x=>x.target.kind==='slide'&&x.target.id==='s'),source=p=>rows.find(x=>x.property===p).source;
 assert.deepEqual(source('background.stops.0.position'),{kind:'master',id:'m',path:'masters.0.background.stops.0.position'});assert.deepEqual(source('background.stops.0.color'),{kind:'master',id:'m',path:'masters.0.background.stops.0.color'});assert.deepEqual(source('background.stops.0.opacity'),{kind:'default',path:'defaults.background.linearGradient.stop.opacity'});
 assert.deepEqual(source('background.stops.1.opacity'),{kind:'master',id:'m',path:'masters.0.background.stops.1.opacity'});assert.equal(report.presentation.slides[0].background.stops[1].opacity,0);assert.deepEqual(source('background.stops.2.position'),{kind:'master',id:'m',path:'masters.0.background.stops.2.position'});
 slide.setBackground({type:'linearGradient',angle:90,stops:[{position:0,color:'FFFFFF',opacity:1},{position:1,color:'000000'}]});const next=resolvePresentationInheritance(d).origins.filter(x=>x.target.kind==='slide'&&x.target.id==='s');assert.equal(next.find(x=>x.property==='background.stops.0.opacity').source.kind,'slide');assert.ok(!next.some(x=>x.property.startsWith('background.stops.2.')));
 slide.setBackground({type:'solid',color:'FFFFFF'});assert.ok(!resolvePresentationInheritance(d).origins.some(x=>x.target.kind==='slide'&&x.target.id==='s'&&x.property.startsWith('background.stops.')));
});

test('shape gradient origins identify nested and shared authored fields without inventing inheritance',()=>{
 const fill={type:'linearGradient',angle:45,stops:[{position:0,color:{theme:'accent1'}},{position:1,color:'0000FF',opacity:0}]};const shape=(id)=>({id,type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20},style:{fill}});
 const d=createPresentation();d.defineSlideMaster({id:'m',name:'Master',elements:[shape('master-shape')]});d.defineSlideLayout({id:'l',name:'Layout',masterId:'m',elements:[shape('layout-shape')]});const s=d.addSlide({id:'s',layoutId:'l'});s.addElement({id:'group',type:'group',box:{x:0,y:0,width:20,height:20},coordinateSize:{width:20,height:20},children:[shape('child')]});
 const report=resolvePresentationInheritance(d);assert.deepEqual(report.presentation,normalizePresentation(d));const source=(id,property)=>report.origins.find(x=>x.target.kind==='element'&&x.target.id===id&&x.property===property)?.source;
 assert.deepEqual(source('master-shape','style.fill.stops.0.color'),{kind:'element',id:'master-shape',path:'masters.0.elements.0.style.fill.stops.0.color'});
 assert.deepEqual(source('layout-shape','style.fill.angle'),{kind:'element',id:'layout-shape',path:`layouts.${d.layouts.findIndex(l=>l.id==='l')}.elements.0.style.fill.angle`});
 assert.deepEqual(source('child','style.fill.stops.1.opacity'),{kind:'element',id:'child',path:'slides.0.elements.0.children.0.style.fill.stops.1.opacity'});assert.deepEqual(source('child','style.fill.stops.0.opacity'),{kind:'default',path:'defaults.shape.fill.linearGradient.stop.opacity'});assert.deepEqual(source('child','style.fill.scaled'),{kind:'default',path:'defaults.shape.fill.linearGradient.scaled'});
 const group=s.elements[0];s.replaceElement('group',{...group,children:[{...group.children[0],style:{fill:{type:'solid',color:'FFFFFF'}}}]});assert.ok(!resolvePresentationInheritance(d).origins.some(x=>x.target.id==='child'&&x.property.startsWith('style.fill.stops.')));
});

test('shape solid and stroke origins distinguish defaults, explicit zero and explicit none',()=>{
 const d=createPresentation(),s=d.addSlide({id:'s'});s.addElement({id:'solid',type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20},style:{fill:{type:'solid',color:'AA0000',opacity:0},stroke:{paint:{type:'solid',color:'0000FF'},width:0,dash:'dash',beginArrow:'none'}}});s.addElement({id:'none',type:'shape',shape:'rect',box:{x:30,y:0,width:20,height:20},style:{fill:{type:'none'},stroke:{paint:{type:'none'}}}});s.addElement({id:'defaults',type:'shape',shape:'ellipse',box:{x:60,y:0,width:20,height:20}});
 const report=resolvePresentationInheritance(d);assert.deepEqual(report.presentation,normalizePresentation(d));const row=(id,property)=>report.origins.find(o=>o.target.kind==='element'&&o.target.id===id&&o.property===property);
 for(const prop of ['fill.type','fill.color','fill.opacity','stroke.paint.type','stroke.paint.color','stroke.width','stroke.dash','stroke.beginArrow']) assert.deepEqual(row('solid',`style.${prop}`).source,{kind:'element',id:'solid',path:`slides.0.elements.0.style.${prop}`});
 assert.deepEqual(row('solid','style.stroke.paint.opacity').source,{kind:'default',path:'defaults.shape.style.stroke.paint.opacity'});assert.deepEqual(row('none','style.fill.type').source,{kind:'element',id:'none',path:'slides.0.elements.1.style.fill.type'});assert.equal(row('none','style.fill.color'),undefined);assert.equal(row('none','style.stroke.paint.color'),undefined);
 assert.deepEqual(row('defaults','style.fill.type').source,{kind:'default',path:'defaults.shape.style.fill.type'});assert.deepEqual(row('defaults','style.stroke.width').source,{kind:'default',path:'defaults.shape.style.stroke.width'});assert.equal(row('defaults','style.fill.color'),undefined);
});
