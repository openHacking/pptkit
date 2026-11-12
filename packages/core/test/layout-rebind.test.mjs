import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation,normalizePresentation,prepareLayoutRebind,LayoutRebindError,validatePresentation,resolvePresentationInheritance,TEXT_FONT_SCRIPT_CODES} from '../dist/index.js';
import {layoutRebindDocument} from '../../../tests/fixtures/layout-rebind.mjs';
const options={localAppearance:'adopt'};
const request=(slideId='page-0',layoutId='target',extra={})=>({slideId,layoutId,...extra});
const snapshot=d=>JSON.stringify({n:normalizePresentation(d),slides:d.slides.map(s=>({id:s.id,layoutId:s.layoutId,background:s.background,elements:s.elements,notes:s.notes,eventTimeline:s.eventTimeline}))});

test('logical key precedes unique role, never array order or ordinary object names',()=>{
 const d=layoutRebindDocument(createPresentation),before=snapshot(d),p=prepareLayoutRebind(d,[request()],options);assert.equal(p.status,'ready');assert.equal(snapshot(d),before);
 assert.deepEqual(p.matches.map(m=>[m.elementId,m.method,m.targetKey]),[['title-0','key','title'],['body-0','kind','copy'],['image-0','kind','photo'],['ordinary-0','ordinary',undefined],['group-0','ordinary',undefined],['connector-0','ordinary',undefined]]);
 const slide=d.slides[0],untouched=d.slides[2],old=untouched.elements;assert.deepEqual(d.applyLayoutRebind(p),{status:'applied',changedSlideIds:['page-0']});assert.equal(d.slides[0],slide);assert.equal(slide.layoutId,'target');assert.equal(untouched.elements[0],old[0]);assert.equal(slide.notes,'Notes 0');
 const n=normalizePresentation(d).slides[0];assert.deepEqual(n.elements[1].box,{x:40,y:120,width:320,height:90});assert.equal(n.elements[1].content[0].runs[0].style.fontSize,20);assert.equal(n.elements[1].content[0].runs[0].style.bold,false);assert.equal(n.elements[3].placeholderKey,undefined);assert.equal(d.slides[0].elements[1].content[0].runs[0].action.url,'https://example.com/body');
 assert.ok(p.findings.some(f=>f.code==='font-availability-unverified'));assert.equal(validatePresentation(d).filter(f=>f.severity==='error').length,0);
});

test('ambiguity, same-key role conflicts and missing roles are reviewable and never partially commit',()=>{
 const d=layoutRebindDocument(createPresentation),before=snapshot(d),p=prepareLayoutRebind(d,[request('page-2'),request('page-0','choices')],options);assert.equal(p.status,'needs-review');assert.deepEqual(p.requests.map(r=>r.slideId),['page-0','page-2']);
 assert.deepEqual(p.matches.find(m=>m.elementId==='body-0').candidates,['one','two']);assert.equal(d.applyLayoutRebind({...p,status:'ready'}).status,'needs-review');assert.equal(snapshot(d),before);
 const fixed=prepareLayoutRebind(d,[request('page-0','choices',{bindings:{'body-0':'one'}}),request('page-2')],options);assert.equal(fixed.status,'ready');assert.deepEqual(d.applyLayoutRebind(fixed).changedSlideIds,['page-0','page-2']);assert.equal(d.slides[1].layoutId,'source');
 const e=layoutRebindDocument(createPresentation),missing=prepareLayoutRebind(e,[request('page-0','empty')],options);assert.equal(missing.status,'needs-review');assert.equal(prepareLayoutRebind(e,[request('page-0','choices')],{...options,unmatched:'detach'}).status,'needs-review');
 e.replaceSlideLayout('target',{name:'Role changed',masterId:'blue',placeholders:[{key:'title',kind:'body',box:{x:0,y:0,width:100,height:50}}]});const conflict=prepareLayoutRebind(e,[request()],{...options,unmatched:'detach'});assert.ok(conflict.findings.some(f=>f.code==='placeholder-kind-conflict'));assert.equal(conflict.status,'needs-review');
});

test('explicit detach and explicit role changes retain content and resolved geometry without manufacturing prompts',()=>{
 const d=layoutRebindDocument(createPresentation),old=normalizePresentation(d).slides[0];const p=prepareLayoutRebind(d,[request('page-0','empty')],{...options,unmatched:'detach'});assert.equal(p.status,'ready');d.applyLayoutRebind(p);const n=normalizePresentation(d).slides[0];
 for(let i=0;i<3;i++){assert.equal(n.elements[i].placeholderKey,undefined);assert.deepEqual(n.elements[i].box,old.elements[i].box);}assert.equal(n.elements[0].plainText,old.elements[0].plainText);assert.equal(n.elements[2].assetId,'logo');
 const e=layoutRebindDocument(createPresentation),explicit=prepareLayoutRebind(e,[request('page-0','choices',{bindings:{'body-0':null,'image-0':'one'}})],options);assert.equal(explicit.status,'ready');assert.ok(explicit.findings.some(f=>f.code==='explicit-role-change'));e.applyLayoutRebind(explicit);assert.equal(e.slides[0].elements[1].placeholderKey,undefined);assert.equal(e.slides[0].elements[2].placeholderKey,'one');
});

test('preserve materializes foreground theme/text/table/chart values and background while target artwork stays shared',()=>{
 const d=layoutRebindDocument(createPresentation),old=normalizePresentation(d),p=prepareLayoutRebind(d,[request()],{localAppearance:'preserve'});assert.equal(p.status,'ready');assert.ok(p.findings.some(f=>f.code==='table-style-materialized'));d.applyLayoutRebind(p);const n=normalizePresentation(d),s=n.slides[0];
 for(let i=0;i<5;i++)assert.deepEqual(s.elements[i].box,old.slides[0].elements[i].box);
 assert.equal(s.elements[0].content[0].runs[0].style.fontFamily,'Arial');assert.equal(s.elements[0].content[0].runs[0].style.color,'CC2200');assert.equal(s.elements[1].content[0].runs[0].style.scriptFontFamilies.Hans,'Arial');assert.equal(s.elements[3].style.fill.color,'CC2200');assert.equal(s.background.color,'FFF8F0');
 const group=s.elements[4];assert.equal(group.children[0].rows[0].cells[0].style.fill.color,'CC2200');assert.equal(group.children[0].styleSource,undefined);assert.equal(group.children[1].series[0].color,old.slides[0].elements[4].children[1].series[0].color);assert.equal(group.children[1].style.textStyle.fontFamily,'Arial');assert.equal(group.children[1].series[0].marker.shape,old.slides[0].elements[4].children[1].series[0].marker.shape);
 assert.equal(s.elements.length,6);assert.equal(s.layoutId,'target');assert.equal(n.layouts.find(l=>l.id==='target').masterId,'blue');assert.equal(d.getAssetReferences('logo').length,3);assert.throws(()=>d.removeAsset('logo'),/referenced/);
});

test('adopt resets only explicit requested overrides, reports overflow/shrink and never changes text or font sizes to fit',()=>{
 const d=layoutRebindDocument(createPresentation),s=d.slides[0];s.replaceElement('body-0',{...s.elements[1],box:{x:8,y:8,width:400,height:150},frame:{margin:7},textStylePreset:undefined,content:[{style:{spaceAfter:9},runs:[{text:'Long '.repeat(80),style:{fontFamily:'Arial',fontSize:40,bold:true},action:{type:'url',url:'https://example.com/keep'}}]}]});
 const target=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...target,placeholders:target.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,frame:{...ph.textStyle.frame,autoFit:{mode:'shrink',fontScale:0.8}}}}:ph)});s.replaceElement('ordinary-0',{...s.elements[3],text:{content:[{runs:[{text:'Direct',style:{fontSize:11,color:'112233',bold:true}}]}]}});
 const p=prepareLayoutRebind(d,[request()],{...options,reset:['geometry','text','background']});assert.equal(p.status,'ready');assert.ok(p.findings.some(f=>f.code==='text-overflow-estimate'));d.applyLayoutRebind(p);const n=normalizePresentation(d).slides[0].elements[1];assert.equal(n.content[0].runs[0].text,'Long '.repeat(80));assert.equal(n.content[0].runs[0].style.fontSize,20);assert.equal(n.content[0].runs[0].style.bold,true);assert.equal(n.frame.autoFit.fontScale,0.8);assert.ok(p.findings.some(f=>f.code==='target-auto-fit-shrink'));assert.equal(normalizePresentation(d).slides[0].elements[3].text.content[0].runs[0].style.color,'112233');assert.equal(normalizePresentation(d).slides[0].elements[3].text.content[0].runs[0].style.fontSize,11);assert.equal(n.content[0].runs[0].action.url,'https://example.com/keep');assert.deepEqual(n.box,{x:40,y:120,width:320,height:90});assert.equal(d.slides[0].background,undefined);
 assert.throws(()=>prepareLayoutRebind(d,[request()],{localAppearance:'preserve',reset:['text']}),TypeError);assert.throws(()=>prepareLayoutRebind(d,[request()],{...options,reset:['text','text']}),TypeError);
});

test('inherited connector references need explicit target IDs and preserving dependent geometry does not guess',()=>{
 const d=layoutRebindDocument(createPresentation);d.slides[0].addElement({id:'inherited-link',type:'connector',start:{elementId:'red-anchor',anchor:'left'},end:{x:500,y:80}});const before=snapshot(d);
 const blocked=prepareLayoutRebind(d,[request()],options);assert.equal(blocked.status,'needs-review');assert.ok(blocked.findings.some(f=>f.code==='missing-inherited-reference'));d.applyLayoutRebind(blocked);assert.equal(snapshot(d),before);
 const blue=d.masters.find(m=>m.id==='blue');d.replaceSlideMaster('blue',{...blue,elements:[...blue.elements,{id:'blue-group',type:'group',box:{x:500,y:30,width:40,height:40},coordinateSize:{width:40,height:40},children:[{id:'blue-descendant',type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10}}]}]});const nested=prepareLayoutRebind(d,[request('page-0','target',{references:{'red-anchor':'blue-descendant'}})],options);assert.equal(nested.status,'needs-review');assert.ok(nested.findings.some(f=>f.code==='inherited-reference-scope'));
 const references={'red-anchor':'blue-anchor'},p=prepareLayoutRebind(d,[request('page-0','target',{references})],options);assert.equal(p.status,'ready');d.applyLayoutRebind(p);assert.equal(d.slides[0].elements.at(-1).start.elementId,'blue-anchor');
 const e=layoutRebindDocument(createPresentation);e.slides[0].addElement({id:'inherited-link',type:'connector',start:{elementId:'red-anchor'},end:{x:5,y:5}});assert.equal(prepareLayoutRebind(e,[request('page-0','target',{references})],{localAppearance:'preserve'}).status,'needs-review');
});

test('plans capture finite data, reject stale unrelated edits, ignore forged reports and retain ownership across atomic batches',()=>{
 const d=layoutRebindDocument(createPresentation),requests=[request(),request('page-1')],p=prepareLayoutRebind(d,requests,options);requests[0].layoutId='empty';assert.equal(p.requests[0].layoutId,'target');assert.ok(Object.isFrozen(p.matches[0].candidates));assert.throws(()=>{p.options.localAppearance='preserve';},TypeError);
 d.slides[2].setNotes('Changed outside selection');const before=snapshot(d);assert.throws(()=>d.applyLayoutRebind(p),e=>e instanceof LayoutRebindError&&e.code==='stale-plan');assert.equal(snapshot(d),before);
 const fresh=prepareLayoutRebind(d,[request(),request('page-1')],options),handles=d.slides.slice();assert.equal(d.applyLayoutRebind(JSON.parse(JSON.stringify(fresh))).status,'applied');assert.equal(d.slides[0],handles[0]);assert.equal(d.slides[1].hidden,true);assert.equal(d.slides[1].showInheritedShapes,false);
 assert.throws(()=>d.slides[2].addElement({id:'body-0',type:'text',box:{x:0,y:0,width:10,height:10},content:'Collision'}),/Duplicate/);d.removeSlide('page-0');assert.equal(d.slides[2-1].addElement({id:'body-0',type:'text',box:{x:0,y:0,width:10,height:10},content:'Reused'}).id,'body-0');
});

test('accessors/malformed requests never execute or mutate; repeat/default no-ops preserve pointers and permit cancellation',()=>{
 const d=layoutRebindDocument(createPresentation),before=snapshot(d);let calls=0;assert.throws(()=>prepareLayoutRebind(d,[{slideId:'page-0',get layoutId(){calls++;return'target';}}],options),TypeError);assert.equal(calls,0);
 for(const r of [[request(),request()], [request('absent')], [request('page-0','absent')], [{...request(),bindings:{unknown:'copy'}}], [{...request(),references:{'red-anchor':'absent'}}],new Array(1)])assert.throws(()=>prepareLayoutRebind(d,r,options));assert.equal(snapshot(d),before);
 const p=prepareLayoutRebind(d,[request()],options);assert.throws(()=>d.applyLayoutRebind({get requests(){calls++;return p.requests;},options,sourceSnapshot:p.sourceSnapshot}),e=>e.code==='invalid-plan');assert.equal(calls,0);assert.equal(snapshot(d),before);d.applyLayoutRebind(p);
 const element=d.slides[0].elements[0],repeat=prepareLayoutRebind(d,[request()],options);assert.deepEqual(repeat.changedSlideIds,[]);d.applyLayoutRebind(repeat);assert.equal(d.slides[0].elements[0],element);
 const empty=createPresentation(),s=empty.addSlide();const implicit=prepareLayoutRebind(empty,[{slideId:s.id,layoutId:'layout-blank'}],{localAppearance:'preserve'});assert.deepEqual(implicit.changedSlideIds,[]);empty.applyLayoutRebind(implicit);assert.equal(s.layoutId,undefined);
});


test('review cancellation consumes no IDs; invalidated source data is stale before normalization and authored default layout stays implicit',()=>{
 const d=layoutRebindDocument(createPresentation),control=layoutRebindDocument(createPresentation),shape={type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10}};
 const blocked=prepareLayoutRebind(d,[request('page-0','choices'),request('page-1')],options);assert.equal(blocked.status,'needs-review');assert.deepEqual(blocked.changes,[]);d.applyLayoutRebind(blocked);assert.equal(d.slides[0].elements[0].id,'title-0');assert.equal(d.slides[2].addElement(shape).id,control.slides[2].addElement(shape).id);
 const p=prepareLayoutRebind(d,[request()],options),old=d.slides[0].elements[3];d.slides[0].replaceElement(old.id,{...old,opacity:NaN});const broken=d.slides[0].elements[3];assert.throws(()=>d.applyLayoutRebind(p),e=>e.code==='stale-plan');assert.equal(d.slides[0].elements[3],broken);assert.equal(d.slides[0].layoutId,'source');
 const empty=createPresentation();empty.defineSlideLayout({id:'layout-blank',name:'Authored default'});const slide=empty.addSlide(),same=prepareLayoutRebind(empty,[{slideId:slide.id,layoutId:'layout-blank'}],options);assert.deepEqual(same.changedSlideIds,[]);empty.applyLayoutRebind(same);assert.equal(slide.layoutId,undefined);
});

test('read-only data plans need no mutation methods and rebinding retains local timing/actions/notes across nested identities',()=>{
 const d=layoutRebindDocument(createPresentation),slide=d.slides[0];slide.setEventTimeline({stages:[{id:'stage',target:{pageId:slide.id,elementId:'chart-0'},trigger:{type:'object-click',elementId:'body-0'},timing:{kind:'span',id:'time',durationMs:20,fill:'hold'}}]});slide.setTransition({type:'fade',durationMs:300});
 const timing=slide.eventTimeline,transition=slide.transition,timelinePlan=slide.getEventTimelinePlan(),p=prepareLayoutRebind(d,[request()],{localAppearance:'preserve'}),data=JSON.parse(p.sourceSnapshot);assert.equal(data.slides[0].setNotes,undefined);assert.deepEqual(normalizePresentation(data),normalizePresentation(d));assert.equal(prepareLayoutRebind(data,[request()],{localAppearance:'preserve'}).sourceSnapshot,p.sourceSnapshot);
 d.applyLayoutRebind(p);assert.equal(slide.eventTimeline,timing);assert.equal(slide.transition,transition);assert.equal(slide.notes,'Notes 0');assert.deepEqual(slide.getEventTimelinePlan(),timelinePlan);assert.equal(slide.elements[1].content[0].runs[0].action.url,'https://example.com/body');assert.equal(slide.elements[4].children[1].id,'chart-0');
});

test('selective font and color reset follows target while retaining live preset and unrelated overrides',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{fontFamily:'Georgia',color:'ABCDEF',fontSize:37,italic:true}}}}));
 for(const i of [0,1]) {const s=d.slides[i],e=s.elements[1];s.replaceElement(e.id,{...e,textStylePreset:'local',content:[{style:{spaceAfter:0},runs:[{text:'keep',style:{fontFamily:'Arial',color:'123456',bold:false,colorOpacity:0},action:{type:'url',url:'https://example.com'}}]}]});}
 const before=JSON.stringify(d.slides[2].elements),p=prepareLayoutRebind(d,[request(),request('page-1')],{localAppearance:'adopt',reset:['text.fontFamily','text.color']});assert.equal(p.status,'ready');assert.equal(d.applyLayoutRebind(p).status,'applied');assert.equal(JSON.stringify(d.slides[2].elements),before);
 const author=d.slides[0].elements[1];assert.equal(author.textStylePreset,'local');assert.deepEqual(author.textStylePresetReset,['fontFamily','color']);assert.deepEqual(author.content[0].runs[0].style,{bold:false,colorOpacity:0});assert.equal(author.content[0].runs[0].action.url,'https://example.com');
 const style=normalizePresentation(d).slides[0].elements[1].content[0].runs[0].style;assert.deepEqual(style.fontFamily,{theme:'body'});assert.deepEqual(style.color,{theme:'accent1'});assert.equal(style.fontSize,37);assert.equal(style.italic,true);assert.equal(style.bold,false);assert.equal(style.colorOpacity,0);
 const data=JSON.parse(p.sourceSnapshot);data.slides[0].elements[1]=JSON.parse(JSON.stringify({...author,placeholderKey:'body-old'}));data.textStylePresets.local.run.fontSize=51;data.textStylePresets.local.run.italic=false;const updated=normalizePresentation(data).slides[0].elements[1].content[0].runs[0].style;assert.equal(updated.fontSize,51);assert.equal(updated.italic,false);assert.notEqual(updated.fontFamily,'Georgia');
 const repeat=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.fontFamily','text.color']});assert.deepEqual(repeat.changedSlideIds,[]);assert.throws(()=>prepareLayoutRebind(d,[request()],{localAppearance:'preserve',reset:['text.color']}),TypeError);
 const invalid=JSON.parse(JSON.stringify(data));invalid.slides[0].elements[1].textStylePresetReset=['color','color'];assert.ok(validatePresentation(invalid).some(x=>x.code==='invalid-text-style-preset-reset'));invalid.slides[0].elements[1].textStylePresetReset=new Array(1);assert.ok(validatePresentation(invalid).some(x=>x.code==='invalid-text-style-preset-reset'));
});


test('selective reset handles shape text and reports inherited origins without changing ordinary shapes',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{fontFamily:'Georgia',color:'ABCDEF',fontSize:31}}}}));const s=d.slides[0],body=s.elements[1];s.replaceElement(body.id,{id:body.id,type:'shape',shape:'rect',placeholderKey:'body-old',text:{textStylePreset:'local',content:[{runs:[{text:'shape',style:{fontFamily:'Arial',color:'123456',italic:false}}]}]}});
 const p=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.color']});assert.equal(p.status,'ready');d.applyLayoutRebind(p);const t=s.elements[1].text;assert.deepEqual(t.textStylePresetReset,['color']);assert.equal(t.content[0].runs[0].style.fontFamily,'Arial');assert.equal(t.content[0].runs[0].style.italic,false);const resolved=resolvePresentationInheritance(d),style=resolved.presentation.slides[0].elements[1].text.content[0].runs[0].style;assert.deepEqual(style.color,{theme:'accent1'});assert.equal(style.fontSize,31);const origin=resolved.origins.find(o=>o.target.id==='body-0'&&o.property==='text.content.0.runs.0.style.color');assert.match(origin.source.path,/layouts.*textStyle\.run\.color/);assert.equal(s.elements[3].text.content,'Ordinary 0');
 const invalid=JSON.parse(p.sourceSnapshot);invalid.slides[0].elements[1].text.textStylePresetReset='color';assert.ok(validatePresentation(invalid).some(x=>x.code==='invalid-text-style-preset-reset'));
});

test('every run property independently resumes target inheritance, including null font barriers',()=>{
 const target={fontFamily:'Target',eastAsianFontFamily:'Target EA',complexScriptFontFamily:'Target CS',scriptFontFamilies:{Hans:'Target Hans',Arab:'Target Arab'},fontSize:25,bold:true,italic:true,underline:true,strike:true,color:'123456',colorOpacity:0.35,language:'fr-FR'};
 const preset={fontFamily:'Preset',eastAsianFontFamily:null,complexScriptFontFamily:null,scriptFontFamilies:null,fontSize:36,bold:false,italic:false,underline:false,strike:false,color:'654321',colorOpacity:0,language:'en-US'};
 const keys=Object.keys(target);
 for(const key of keys){
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:preset}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:target}}:ph)});
  s.replaceElement(e.id,{...e,textStylePreset:'local',content:[{runs:[{text:'test',style:preset,action:{type:'url',url:'https://example.com'}}]}]});
  const before=normalizePresentation(d).slides[0].elements[1].content[0].runs[0].style,plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:[`text.${key}`]});assert.equal(plan.status,'ready',key);d.applyLayoutRebind(plan);
  const author=s.elements[1],run=normalizePresentation(d).slides[0].elements[1].content[0].runs[0].style;assert.equal(author.textStylePreset,'local',key);assert.deepEqual(author.textStylePresetReset,[key],key);assert.equal(Object.hasOwn(author.content[0].runs[0].style,key),false,key);assert.deepEqual(run[key],target[key],key);for(const other of keys.filter(k=>k!==key))assert.deepEqual(run[other],before[other],`${key} changed ${other}`);assert.equal(author.content[0].runs[0].action.url,'https://example.com');
  const origins=resolvePresentationInheritance(d).origins,property=`content.0.runs.0.style.${key==='scriptFontFamilies'?'scriptFontFamilies.Hans':key}`,origin=origins.find(o=>o.target.id==='body-0'&&o.property===property);assert.ok(origin,key);assert.match(origin.source.path,/layouts.*textStyle\.run\./,key);
 }
 const d=layoutRebindDocument(createPresentation);for(const value of ['text.scriptFontFamilies.Bogus','text.toString','text.fontsize'])assert.throws(()=>prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:[value]}),TypeError);
});

test('paragraph and frame selectors independently inherit target while preserving preset and local values',()=>{
 const target={paragraph:{align:'right',indent:27,hanging:8,lineSpacing:1.4,spaceBefore:4,spaceAfter:5,bullet:{type:'number',style:'arabicParen',startAt:3}},frame:{margin:{left:3,right:4,top:5,bottom:6},verticalAlign:'bottom',wrap:false,autoFit:{mode:'shrink',fontScale:0.7}}};
 const local={paragraph:{align:'left',indent:0,hanging:0,lineSpacing:1,spaceBefore:0,spaceAfter:0,bullet:{type:'bullet',character:'+'}},frame:{margin:9,verticalAlign:'top',wrap:true,autoFit:{mode:'none'}}};
 for(const section of ['paragraph','frame'])for(const key of Object.keys(target[section])){
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local}})),slide=d.slides[0],element=slide.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,...target}}:ph)});
  slide.replaceElement(element.id,{...element,textStylePreset:'local',frame:local.frame,content:[{style:local.paragraph,runs:[{text:'test',style:{bold:false}}]}]});const before=normalizePresentation(d).slides[0].elements[1],plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:[`text.${section}.${key}`]});assert.equal(plan.status,'ready',`${section}.${key}`);d.applyLayoutRebind(plan);
  const author=slide.elements[1],elementAfter=normalizePresentation(d).slides[0].elements[1],style=section==='frame'?elementAfter.frame:elementAfter.content[0].style,prior=section==='frame'?before.frame:before.content[0].style;assert.deepEqual(style[key],target[section][key],`${section}.${key}`);for(const other of Object.keys(target[section]).filter(k=>k!==key))assert.deepEqual(style[other],prior[other],`${section}.${key} changed ${other}`);assert.deepEqual(author.textStylePresetReset,[`${section}.${key}`]);assert.equal(author.textStylePreset,'local');assert.equal(elementAfter.content[0].runs[0].style.bold,false);
  const prop=section==='frame'?`frame.${key==='margin'?'margin.left':key==='autoFit'?'autoFit.mode':key}`:`content.0.style.${key==='bullet'?'bullet.type':key}`,origin=resolvePresentationInheritance(d).origins.find(o=>o.target.id==='body-0'&&o.property===prop);assert.ok(origin,prop);assert.match(origin.source.path,/layouts.*textStyle\./,prop);
 }
});

test('shape paragraphs and frames reset atomically while plans remain stale-safe',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{paragraph:{align:'right'},frame:{wrap:false}}}})),s=d.slides[0],e=s.elements[1];s.replaceElement(e.id,{id:e.id,type:'shape',shape:'rect',placeholderKey:'body-old',text:{textStylePreset:'local',frame:{wrap:false,margin:0},content:[{style:{align:'right',spaceAfter:0},runs:[{text:'first'},{text:'second',style:{bold:false}}]}]}});
 const options={localAppearance:'adopt',reset:['text.paragraph.align','text.frame.wrap']},stale=prepareLayoutRebind(d,[request(),request('page-1')],options);d.slides[2].setNotes('changed');const before=snapshot(d);assert.throws(()=>d.applyLayoutRebind(stale),e=>e.code==='stale-plan');assert.equal(snapshot(d),before);
 const plan=prepareLayoutRebind(d,[request(),request('page-1')],options);assert.equal(plan.status,'ready');d.applyLayoutRebind(plan);const author=s.elements[1].text,normalized=normalizePresentation(d).slides[0].elements[1].text;assert.deepEqual(author.textStylePresetReset,['paragraph.align','frame.wrap']);assert.equal(author.frame.margin,0);assert.equal(author.content[0].style.spaceAfter,0);assert.equal(normalized.content[0].style.align,'left');assert.equal(normalized.frame.wrap,true);assert.equal(normalized.content[0].runs.map(r=>r.text).join(''),'firstsecond');assert.equal(normalized.content[0].runs[1].style.bold,false);
});

test('individual margin resets retain other scalar or partial declarations and their origins',()=>{
 const sides=['top','right','bottom','left'], target={top:1,right:2,bottom:3,left:4};
 for(const kind of ['text','shape'])for(const side of sides)for(const localMargin of [9,{top:0,right:11,bottom:12,left:13},undefined]){
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{frame:{margin:7}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');
  d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,frame:{margin:target}}}:ph)});
  const text={textStylePreset:'local',...(localMargin===undefined?{}:{frame:{margin:localMargin}}),content:'content'};
  s.replaceElement(e.id,kind==='text'?{id:e.id,type:'text',placeholderKey:'body-old',...text}:{id:e.id,type:'shape',shape:'rect',placeholderKey:'body-old',text});
  const before=(kind==='text'?normalizePresentation(d).slides[0].elements[1]:normalizePresentation(d).slides[0].elements[1].text).frame.margin;
  const plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:[`text.frame.margin.${side}`]});assert.equal(plan.status,'ready');d.applyLayoutRebind(plan);
  const author=kind==='text'?s.elements[1]:s.elements[1].text, result=resolvePresentationInheritance(d),out=kind==='text'?result.presentation.slides[0].elements[1]:result.presentation.slides[0].elements[1].text;
  assert.deepEqual(author.textStylePresetReset,[`frame.margin.${side}`]);assert.equal(out.frame.margin[side],target[side]);
  for(const other of sides.filter(x=>x!==side)) assert.equal(out.frame.margin[other],before[other],`${kind} ${side} ${other}`);
  const prefix=kind==='text'?'frame':'text.frame', selected=result.origins.find(o=>o.target.id==='body-0'&&o.property===`${prefix}.margin.${side}`);
  assert.match(selected.source.path,/layouts.*textStyle\.frame\.margin/);
  if(localMargin===undefined){const other=sides.find(x=>x!==side),origin=result.origins.find(o=>o.target.id==='body-0'&&o.property===`${prefix}.margin.${other}`);assert.equal(origin.source.path,'textStylePresets["local"].frame.margin');}
  s.replaceElement('body-0',kind==='text'?{...s.elements[1],frame:{...author.frame,margin:{...author.frame?.margin,[side]:0}}}:{...s.elements[1],text:{...author,frame:{...author.frame,margin:{...author.frame?.margin,[side]:0}}}});
  const after=normalizePresentation(d).slides[0].elements[1];assert.equal((kind==='text'?after:after.text).frame.margin[side],0);
 }
});

test('margin side and whole resets compose independent of order and reject unsupported sides',()=>{
 for(const reset of [['text.frame.margin','text.frame.margin.left'],['text.frame.margin.left','text.frame.margin']]){
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{frame:{margin:7}}}})),s=d.slides[0],e=s.elements[1];s.replaceElement(e.id,{...e,textStylePreset:'local',frame:{margin:9}});
  const plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset});d.applyLayoutRebind(plan);const target=normalizePresentation(d).layouts.find(l=>l.id==='target').placeholders.find(p=>p.key==='copy').textStyle.frame.margin;assert.deepEqual(normalizePresentation(d).slides[0].elements[1].frame.margin,target);
 }
 const d=layoutRebindDocument(createPresentation);assert.throws(()=>prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.frame.margin.center']}),TypeError);
});

test('margin side markers without a preset remain live and independent across repeated rebinds',()=>{
 const d=layoutRebindDocument(createPresentation),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,frame:{margin:{top:1,right:2,bottom:3,left:4}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:undefined,frame:{margin:9},textStylePresetReset:['frame.margin.top']});
 let plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.frame.margin.left','text.frame.margin.bottom']});d.applyLayoutRebind(plan);assert.deepEqual(s.elements[1].textStylePresetReset,['frame.margin.top','frame.margin.left','frame.margin.bottom']);assert.deepEqual(normalizePresentation(d).slides[0].elements[1].frame.margin,{top:9,right:9,bottom:3,left:4});
 plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.frame.margin.top']});d.applyLayoutRebind(plan);assert.deepEqual(normalizePresentation(d).slides[0].elements[1].frame.margin,{top:1,right:9,bottom:3,left:4});assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);
});

test('partial preset margin masks only the selected side and keeps inherited gaps',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{frame:{margin:{left:8,right:0}}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,frame:{margin:{top:1,right:2,bottom:3,left:4}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:'local',frame:undefined,content:'item'});const plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.frame.margin.left']});d.applyLayoutRebind(plan);const result=resolvePresentationInheritance(d);assert.deepEqual(result.presentation.slides[0].elements[1].frame.margin,{top:1,right:0,bottom:3,left:4});const origin=result.origins.find(o=>o.target.id==='body-0'&&o.property==='frame.margin.right');assert.equal(origin.source.path,'textStylePresets["local"].frame.margin.right');
});

test('each script key can resume target inheritance without changing other keys or clear barriers',()=>{
 const codes=TEXT_FONT_SCRIPT_CODES,target=Object.fromEntries(codes.map(code=>[code,`Target ${code}`]));
 for(const kind of ['text','shape'])for(const code of codes)for(const barrier of ['map','null']){
  const preset=barrier==='null'?null:Object.fromEntries(codes.map(k=>[k,`Preset ${k}`]));
  const local=barrier==='null'?null:Object.fromEntries(codes.map(k=>[k,k===code?'Local chosen':null]));
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{scriptFontFamilies:preset}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:{scriptFontFamilies:target}}}:ph)});
  const text={textStylePreset:'local',content:[{runs:[{text:'script',style:{scriptFontFamilies:local,bold:false}}]}]};s.replaceElement(e.id,kind==='text'?{id:e.id,type:'text',placeholderKey:'body-old',...text}:{id:e.id,type:'shape',shape:'rect',placeholderKey:'body-old',text});
  const before=normalizePresentation(d).slides[0].elements[1],prior=(kind==='text'?before:before.text).content[0].runs[0].style.scriptFontFamilies??{};const plan=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:[`text.scriptFontFamilies.${code}`]});assert.equal(plan.status,'ready');d.applyLayoutRebind(plan);
  const author=kind==='text'?s.elements[1]:s.elements[1].text,result=resolvePresentationInheritance(d),normalized=result.presentation.slides[0].elements[1],run=(kind==='text'?normalized:normalized.text).content[0].runs[0];assert.equal(run.style.scriptFontFamilies[code],target[code]);for(const other of codes.filter(k=>k!==code))assert.equal(run.style.scriptFontFamilies[other],prior[other],`${kind} ${barrier} ${code} ${other}`);assert.equal(run.style.bold,false);assert.deepEqual(author.textStylePresetReset,[`scriptFontFamilies.${code}`]);
  const prefix=kind==='text'?'content':'text.content',origin=result.origins.find(o=>o.target.id==='body-0'&&o.property===`${prefix}.0.runs.0.style.scriptFontFamilies.${code}`);assert.match(origin.source.path,/layouts.*textStyle\.run\.scriptFontFamilies/);
 }
});

test('script map null preset provenance, later local override and whole-map ordering remain correct',()=>{
 for(const reset of [['text.scriptFontFamilies','text.scriptFontFamilies.Hans'],['text.scriptFontFamilies.Hans','text.scriptFontFamilies']]){
  const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{scriptFontFamilies:null}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:{scriptFontFamilies:{Hans:'Target Han',Arab:'Target Arab'}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:'local',content:[{runs:[{text:'x',style:{scriptFontFamilies:null}}]}]});const p=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset});d.applyLayoutRebind(p);assert.deepEqual(normalizePresentation(d).slides[0].elements[1].content[0].runs[0].style.scriptFontFamilies,{Hans:'Target Han',Arab:'Target Arab'});
 }
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{scriptFontFamilies:null}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:{scriptFontFamilies:{Hans:'Target Han',Arab:'Target Arab'}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:'local',content:'x'});let p=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.scriptFontFamilies.Hans']});d.applyLayoutRebind(p);let result=resolvePresentationInheritance(d);assert.deepEqual(result.presentation.slides[0].elements[1].content[0].runs[0].style.scriptFontFamilies,{Hans:'Target Han'});let origin=result.origins.find(o=>o.target.id==='body-0'&&o.property==='content.0.runs.0.style.scriptFontFamilies.Arab');assert.equal(origin.state,'cleared');assert.equal(origin.source.path,'textStylePresets["local"].run.scriptFontFamilies');
 const author=s.elements[1];s.replaceElement(author.id,{...author,content:[{runs:[{text:'x',style:{scriptFontFamilies:{Hans:'Later',Arab:'Local Arab'}}}]}]});assert.deepEqual(normalizePresentation(d).slides[0].elements[1].content[0].runs[0].style.scriptFontFamilies,{Hans:'Later',Arab:'Local Arab'});assert.equal(validatePresentation(d).filter(x=>x.severity==='error').length,0);
});

test('script leaf reset preserves sparse map values and source identity',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{scriptFontFamilies:{Hans:'Preset Han',Arab:'Preset Arabic',Hebr:null}}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:{scriptFontFamilies:{Hans:{theme:'body'},Arab:'Target Arabic',Hebr:'Target Hebrew'}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:'local',content:[{runs:[{text:'x',style:{scriptFontFamilies:{Hans:null}}}]}]});const p=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.scriptFontFamilies.Hans']});d.applyLayoutRebind(p);const result=resolvePresentationInheritance(d),map=result.presentation.slides[0].elements[1].content[0].runs[0].style.scriptFontFamilies;assert.deepEqual(map,{Hans:{theme:'body'},Arab:'Preset Arabic'});const origin=result.origins.find(o=>o.target.id==='body-0'&&o.property==='content.0.runs.0.style.scriptFontFamilies.Hebr');assert.equal(origin.state,'cleared');assert.equal(origin.source.path,'textStylePresets["local"].run.scriptFontFamilies.Hebr');
});

test('a selected script inherits a target clear and retains its cleared provenance',()=>{
 const d=layoutRebindDocument(init=>createPresentation({...init,textStylePresets:{local:{run:{scriptFontFamilies:{Hans:'Preset Han',Arab:'Preset Arabic'}}}}})),s=d.slides[0],e=s.elements[1],layout=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...layout,placeholders:layout.placeholders.map(ph=>ph.key==='copy'?{...ph,textStyle:{...ph.textStyle,run:{scriptFontFamilies:{Hans:null,Arab:'Target Arabic'}}}}:ph)});s.replaceElement(e.id,{...e,textStylePreset:'local',content:[{runs:[{text:'x',style:{scriptFontFamilies:{Hans:'Local Han'}}}]}]});const p=prepareLayoutRebind(d,[request()],{localAppearance:'adopt',reset:['text.scriptFontFamilies.Hans']});d.applyLayoutRebind(p);const result=resolvePresentationInheritance(d),map=result.presentation.slides[0].elements[1].content[0].runs[0].style.scriptFontFamilies;assert.deepEqual(map,{Arab:'Preset Arabic'});const origin=result.origins.find(o=>o.target.id==='body-0'&&o.property==='content.0.runs.0.style.scriptFontFamilies.Hans');assert.equal(origin.state,'cleared');assert.match(origin.source.path,/layouts.*textStyle\.run\.scriptFontFamilies\.Hans/);
});

test('gradient backgrounds survive layout rebinding with explicit preserve and adopt semantics',()=>{
 const make=()=>{const d=createPresentation();const gradient=(color,angle)=>({type:'linearGradient',angle,scaled:true,stops:[{position:0,color:{theme:'accent1'},opacity:.5},{position:1,color}]});
  d.defineSlideMaster({id:'red',name:'Red',theme:{colors:{accent1:'AA0000'}},background:gradient('0000FF',45)});d.defineSlideMaster({id:'green',name:'Green',theme:{colors:{accent1:'00AA00'}},background:gradient('FFFF00',90)});
  d.defineSlideLayout({id:'source',name:'Source',masterId:'red'});d.defineSlideLayout({id:'target',name:'Target',masterId:'green'});d.addSlide({id:'s',layoutId:'source'});return d;};
 const d=make(),before=normalizePresentation(d).slides[0],plan=prepareLayoutRebind(d,[{slideId:'s',layoutId:'target'}],{localAppearance:'preserve'});
 assert.equal(plan.status,'ready');assert.equal(d.slides[0].background,undefined);d.applyLayoutRebind(plan);const preserved=normalizePresentation(d).slides[0];assert.equal(preserved.backgroundSource,'slide');assert.equal(preserved.background.angle,before.background.angle);assert.deepEqual(preserved.background.stops.map(s=>s.color),['AA0000','0000FF']);assert.equal(preserved.background.stops[0].opacity,.5);
 const e=make(),adopt=prepareLayoutRebind(e,[{slideId:'s',layoutId:'target'}],{localAppearance:'adopt'});assert.equal(adopt.status,'ready');e.applyLayoutRebind(adopt);const inherited=normalizePresentation(e).slides[0];assert.equal(inherited.backgroundSource,'master');assert.equal(inherited.background.angle,90);assert.deepEqual(inherited.background.stops[0].color,{theme:'accent1'});assert.equal(e.slides[0].background,undefined);
 const f=make();f.slides[0].setBackground({type:'linearGradient',angle:10,stops:[{position:0,color:'FFFFFF'},{position:1,color:'000000'}]});const reset=prepareLayoutRebind(f,[{slideId:'s',layoutId:'target'}],{localAppearance:'adopt',reset:['background']});assert.equal(reset.status,'ready');f.applyLayoutRebind(reset);assert.equal(f.slides[0].background,undefined);assert.equal(normalizePresentation(f).slides[0].background.angle,90);
});
