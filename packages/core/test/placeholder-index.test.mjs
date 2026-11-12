import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation,normalizePresentation,prepareLayoutRebind,validatePresentation,PresentationValidationError,LayoutRebindError} from '../dist/index.js';
const box={x:10,y:20,width:200,height:60};
const ph=(key,kind='body',index)=>({key,kind,box,...(index===undefined?{}:{index})});
function document(source=[ph('old','body',0)],target=[ph('other','body',7),ph('renamed','body',0)]){
 const d=createPresentation();d.defineSlideLayout({id:'source',name:'Source',placeholders:source});d.defineSlideLayout({id:'target',name:'Target',placeholders:target});
 d.addSlide({id:'page',layoutId:'source',elements:[{id:'text',type:'text',placeholderKey:source[0].key,content:'Retained'}]});return d;
}
const plan=(d,extra={})=>prepareLayoutRebind(d,[{slideId:'page',layoutId:'target',...extra}],{localAppearance:'adopt'});

test('explicit zero/max indices are captured, normalized and immutable; absent indices stay absent',()=>{
 const source=[ph('zero','body',0),ph('max','image',4294967295),ph('absent','title')],d=document(source);
 source[0].index=99;
 const owned=d.layouts.find(l=>l.id==='source');assert.equal(owned.placeholders[0].index,0);assert.ok(Object.isFrozen(owned.placeholders[0]));
 const n=normalizePresentation(d).layouts.find(l=>l.id==='source');assert.deepEqual(n.placeholders.map(p=>p.index),[0,4294967295,undefined]);assert.equal(Object.hasOwn(n.placeholders[2],'index'),false);
 n.placeholders[0].index=88;assert.equal(owned.placeholders[0].index,0);
});

test('invalid and duplicate explicit indices have located diagnostics, including duplicates across roles',()=>{
 for(const value of [-1,1.5,4294967296,Number.MAX_SAFE_INTEGER,NaN,Infinity,null,'0']){
  const d=document([ph('old','body',value)]),errors=validatePresentation(d).filter(f=>f.code==='invalid-placeholder-index');
  assert.equal(errors.length,1);assert.equal(errors[0].layoutId,'source');assert.equal(errors[0].path,'layouts.0.placeholders.0.index');assert.throws(()=>normalizePresentation(d),PresentationValidationError);
 }
 const d=document([ph('old','body',0),ph('different','title',0)]),errors=validatePresentation(d).filter(f=>f.code==='duplicate-placeholder-index');
 assert.equal(errors.length,1);assert.equal(errors[0].path,'layouts.0.placeholders.1.index');assert.throws(()=>plan(d),PresentationValidationError);
});

test('logical key precedes declared index; index disambiguates renamed same-role candidates regardless of order',()=>{
 const d=document(),p=plan(d);assert.equal(p.status,'ready');assert.deepEqual(p.matches[0],{slideId:'page',elementId:'text',sourceKey:'old',targetKey:'renamed',sourceIndex:0,targetIndex:0,method:'index',candidates:['renamed']});
 const target=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...target,placeholders:[...target.placeholders].reverse()});
 assert.throws(()=>d.applyLayoutRebind(p),e=>e instanceof LayoutRebindError&&e.code==='stale-plan');
 const current=plan(d);assert.equal(current.matches[0].targetKey,'renamed');assert.equal(current.matches[0].method,'index');d.applyLayoutRebind(current);assert.equal(d.slides[0].elements[0].placeholderKey,'renamed');
 const keyed=document([ph('old','body',0)],[ph('renamed','body',0),ph('old','body',17)]),keyPlan=plan(keyed);assert.equal(keyPlan.matches[0].method,'key');assert.equal(keyPlan.matches[0].targetIndex,17);assert.ok(keyPlan.findings.some(f=>f.code==='placeholder-index-changed'));
 const max=document([ph('old','body',4294967295)],[ph('renamed','body',4294967295),ph('other','body',0)]);assert.equal(plan(max).matches[0].method,'index');
});

test('same-index wrong roles require review and cannot fall through to unique-role guessing; explicit decisions remain available',()=>{
 const d=document([ph('old','body',3)],[ph('conflict','title',3),ph('body','body',9)]),before=normalizePresentation(d),p=plan(d);
 assert.equal(p.status,'needs-review');assert.equal(p.matches[0].method,'unmatched');assert.deepEqual(p.matches[0].candidates,['conflict']);assert.ok(p.findings.some(f=>f.code==='placeholder-index-kind-conflict'));
 assert.equal(d.applyLayoutRebind(p).status,'needs-review');assert.deepEqual(normalizePresentation(d),before);
 const chosen=plan(d,{bindings:{text:'body'}});assert.equal(chosen.status,'ready');assert.equal(chosen.matches[0].method,'explicit');assert.equal(chosen.matches[0].targetIndex,9);assert.ok(chosen.findings.some(f=>f.code==='placeholder-index-changed'));
 const detached=plan(d,{bindings:{text:null}});assert.equal(detached.status,'ready');assert.equal(detached.matches[0].sourceIndex,3);assert.equal(Object.hasOwn(detached.matches[0],'targetIndex'),false);
});

test('undeclared numeric identities never come from source/target traversal positions and unique-role fallback stays compatible',()=>{
 const d=document([ph('old')],[ph('first','body',1),ph('second','body',2)]);const p=plan(d);assert.equal(p.status,'needs-review');assert.equal(p.matches[0].sourceIndex,undefined);assert.deepEqual(p.matches[0].candidates,['first','second']);
 const legacy=document([ph('old','body',0)],[ph('only')]);assert.equal(plan(legacy).matches[0].method,'kind');assert.equal(Object.hasOwn(plan(legacy).matches[0],'targetIndex'),false);
 const unindexed=document([ph('old')],[ph('only')]);assert.equal(Object.hasOwn(plan(unindexed).matches[0],'sourceIndex'),false);
});

test('index validation participates in atomic definition replacement and full-snapshot rebind freshness',()=>{
 const d=document(),source=d.layouts.find(l=>l.id==='source'),before=normalizePresentation(d),handle=d.slides[0],element=handle.elements[0];
 assert.throws(()=>d.replaceSlideLayout('source',{...source,placeholders:[ph('old','body',1),ph('duplicate','image',1)],elements:[{id:'new-owned',type:'shape',shape:'rect',box}]}),PresentationValidationError);
 assert.equal(d.layouts.find(l=>l.id==='source'),source);assert.equal(d.slides[0],handle);assert.equal(handle.elements[0],element);assert.deepEqual(normalizePresentation(d),before);
 const reused=handle.addElement({id:'new-owned',type:'shape',shape:'rect',box});assert.equal(reused.id,'new-owned');
 const p=plan(d);d.replaceSlideLayout('source',{...source,placeholders:[ph('old','body',2)]});assert.throws(()=>d.applyLayoutRebind(p),e=>e instanceof LayoutRebindError&&e.code==='stale-plan');
});

test('runtime key/kind admission rejects unsupported identity roles before matching and leaves failed replacements atomic',()=>{
 for(const [field,values,code]of [['key',['','   ',0,null,undefined],'invalid-placeholder-key'],['kind',['unknown','body ',0,null,undefined],'invalid-placeholder-kind']])for(const value of values){
  const d=createPresentation();d.defineSlideLayout({id:'bad',name:'Bad',placeholders:[{...ph('body','body',0),[field]:value}]});
  const errors=validatePresentation(d).filter(f=>f.code===code);assert.equal(errors.length,1);assert.equal(errors[0].path,`layouts.0.placeholders.0.${field}`);assert.throws(()=>normalizePresentation(d),PresentationValidationError);
 }
 const invalid=document([ph('old','unknown',0)],[ph('renamed','unknown',0)]);assert.throws(()=>plan(invalid),e=>e instanceof PresentationValidationError&&e.diagnostics.filter(f=>f.code==='invalid-placeholder-kind').length===2);
 const d=document(),source=d.layouts.find(l=>l.id==='source'),before=normalizePresentation(d);
 assert.throws(()=>d.replaceSlideLayout('source',{...source,placeholders:[ph('old','unknown',0)]}),PresentationValidationError);assert.equal(d.layouts.find(l=>l.id==='source'),source);assert.deepEqual(normalizePresentation(d),before);
 const safe=document([ph('__proto__','body',0)],[ph('renamed','body',0)]);assert.equal(plan(safe).matches[0].method,'index');assert.equal(plan(safe).matches[0].sourceKey,'__proto__');
 const malformedKey={toString:1,valueOf:2},malformed=createPresentation();malformed.defineSlideLayout({id:'bad-key',name:'Bad key',placeholders:[{...ph('one','body',0),key:malformedKey},{...ph('two','title',1),key:malformedKey}]});
 const keyErrors=validatePresentation(malformed).filter(f=>f.code==='invalid-placeholder-key');assert.equal(keyErrors.length,2);assert.throws(()=>normalizePresentation(malformed),PresentationValidationError,'Duplicate malformed keys must not be stringified before admission');
});
