import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation,prepareLayoutRebind} from '@pptkit/core';
import {generatePptx} from '../dist/index.js';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
import {placeholderIndexDocument} from '../../../tests/fixtures/placeholder-index.mjs';
const xmlParts=async d=>readGeneratedZip((await generatePptx(d)).bytes);
const indices=xml=>[...xml.matchAll(/<p:ph\b[^>]*\bidx="(\d+)"/g)].map(m=>Number(m[1]));

test('explicit zero/max numeric identities and mixed implicit allocation are shared by layout prompts and local text/picture bindings',async()=>{
 const d=placeholderIndexDocument(createPresentation),before=await xmlParts(d);
 const source=[...before].find(([n,b])=>n.startsWith('ppt/slideLayouts/')&&n.endsWith('.xml')&&b.toString().includes('name="Source"'))[1].toString();assert.deepEqual(indices(source),[0,4294967295,73,4]);
 const target=[...before].find(([n,b])=>n.startsWith('ppt/slideLayouts/')&&n.endsWith('.xml')&&b.toString().includes('name="Target"'))[1].toString();assert.deepEqual(indices(target),[17,1,2,4294967295,0,73]);
 const p=prepareLayoutRebind(d,[{slideId:'page-0',layoutId:'target'}],{localAppearance:'adopt'});assert.equal(p.status,'ready');assert.deepEqual(p.matches.map(m=>m.method),['key','index','index','key','ordinary']);d.applyLayoutRebind(p);
 const out=await xmlParts(d);assert.deepEqual(indices(out.get('ppt/slides/slide1.xml').toString()),[73,0,4294967295,1]);assert.deepEqual(indices(out.get('ppt/slides/slide2.xml').toString()),[73,0,4294967295,4]);
 const current=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...current,placeholders:[...current.placeholders].reverse()});const reordered=await xmlParts(d);assert.deepEqual(indices(reordered.get('ppt/slides/slide1.xml').toString()),[73,0,4294967295,5]);
 assert.equal(d.layouts.find(l=>l.id==='target').placeholders.find(p=>p.key==='legacy').index,undefined,'Implicit serialization index never becomes authoring identity');
});

test('dense reserved indices never collide with implicit prompts and invalid duplicates reject before packaging',async()=>{
 const d=createPresentation(),box={x:0,y:0,width:10,height:10},placeholders=[];
 for(let i=0;i<100;i++)placeholders.push({key:'implicit-'+i,kind:'body',box});
 for(let i=1;i<=100;i++)placeholders.push({key:'explicit-'+i,kind:'body',box,index:i});
 d.defineSlideLayout({id:'many',name:'Many',placeholders});const parts=await xmlParts(d),layout=[...parts].find(([n,b])=>n.startsWith('ppt/slideLayouts/')&&n.endsWith('.xml')&&b.toString().includes('name="Many"'))[1].toString(),values=indices(layout);
 assert.equal(new Set(values).size,200);assert.deepEqual(values.slice(0,100),Array.from({length:100},(_,i)=>i+101));assert.deepEqual(values.slice(100),Array.from({length:100},(_,i)=>i+1));
 d.defineSlideLayout({id:'invalid',name:'Invalid',placeholders:[{key:'one',kind:'title',box,index:0},{key:'two',kind:'image',box,index:0}]});await assert.rejects(generatePptx(d),e=>e.diagnostics.some(d=>d.code==='duplicate-placeholder-index'));
 const invalidKind=createPresentation();invalidKind.defineSlideLayout({id:'unsupported',name:'Unsupported',placeholders:[{key:'bad',kind:'unknown',box,index:0}]});await assert.rejects(generatePptx(invalidKind),e=>e.diagnostics.some(d=>d.code==='invalid-placeholder-kind'),'Unsupported roles must not silently serialize as native body prompts');
});
