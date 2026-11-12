import assert from 'node:assert/strict';
import test from 'node:test';
import {readFile} from 'node:fs/promises';
import {chromium} from '@playwright/test';
import {createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {importPptx} from '../packages/pptx-parser/dist/index.js';
import {layoutRebindDocument} from '../tests/fixtures/layout-rebind.mjs';
import {placeholderIndexDocument} from '../tests/fixtures/placeholder-index.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
async function snapshot(api,factory,mode){const d=factory(api.createPresentation),before=api.normalizePresentation(d),p=api.prepareLayoutRebind(d,[{slideId:'page-1',layoutId:'target'},{slideId:'page-0',layoutId:mode==='review'?'choices':'target'}],{localAppearance:mode==='preserve'?'preserve':'adopt',...(mode==='reset'?{reset:['geometry','text','background']}: {})}),applied=d.applyLayoutRebind(p),layout=api.resolveLayout(d),out=await api.generatePptx(d);return JSON.parse(JSON.stringify({before,plan:p,applied,normalized:api.normalizePresentation(d),origins:api.resolvePresentationInheritance(d).origins,layout,svg:await api.renderLayoutToSvg(layout),bytes:[...out.bytes],warnings:out.warnings}));}

test('review/adopt/preserve layout rebinding is equal across Node/ESM/global and actual XML retains scoped target prompts',async t=>{
 const api={createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind,resolveLayout,generatePptx,renderLayoutToSvg},expected={};for(const mode of ['review','adopt','preserve','reset'])expected[mode]=await snapshot(api,layoutRebindDocument,mode);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});const imports={};for(const n of ['core','layout','svg-renderer','pptx-exporter'])imports['@pptkit/'+n]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${n}/dist/browser/index.js`,'utf8')).toString('base64');await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
 for(const format of ['esm','global']){
  if(format==='global')for(const n of ['core','layout','svg-renderer','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${n}/dist/browser/global.js`,'utf8')});
  const actual=await page.evaluate(async({format,factory,snapshot})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','layout','svg-renderer','pptx-exporter'].map(n=>import('@pptkit/'+n)))),f=Function(`return (${factory})`)(),s=Function(`return (${snapshot})`)(),r={};for(const mode of ['review','adopt','preserve','reset'])r[mode]=await s(api,f,mode);return r;},{format,factory:layoutRebindDocument.toString(),snapshot:snapshot.toString()});assert.deepEqual(actual,expected);
 }
 assert.deepEqual(expected.review.normalized,expected.review.before);assert.equal(expected.review.applied.status,'needs-review');assert.equal(expected.adopt.applied.status,'applied');assert.deepEqual(expected.adopt.applied.changedSlideIds,['page-0','page-1']);
 for(const mode of ['adopt','preserve','reset']){
  assert.deepEqual(expected[mode].normalized.slides[2],expected[mode].before.slides[2]);
  const parts=readGeneratedZip(Uint8Array.from(expected[mode].bytes));
  const actual=await page.evaluate(({parts,mode})=>{
   const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',REL='http://schemas.openxmlformats.org/package/2006/relationships',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
   const parse=n=>new DOMParser().parseFromString(parts[n],'application/xml'),direct=(n,tag)=>[...n.children].find(c=>c.namespaceURI===P&&c.localName===tag),rel=(owner,suffix)=>{const name=owner.replace(/([^/]+)$/,'_rels/$1.rels'),edge=[...parse(name).getElementsByTagNameNS(REL,'Relationship')].find(n=>n.getAttribute('Type').endsWith('/'+suffix));return new URL(edge.getAttribute('Target'),'https://test.invalid/'+owner).pathname.slice(1);};
   const rows=[];
   for(let i=1;i<=3;i++){
    const part=`ppt/slides/slide${i}.xml`,slide=parse(part),layout=rel(part,'slideLayout'),l=parse(layout),master=rel(layout,'slideMaster'),m=parse(master),name=l.getElementsByTagNameNS(P,'cSld')[0].getAttribute('name');
    const prompts=[...l.getElementsByTagNameNS(P,'sp')].map(s=>({name:s.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name'),ph:s.getElementsByTagNameNS(P,'ph')[0]})).filter(n=>n.ph);
    const local=[...slide.getElementsByTagNameNS(P,'sp')].find(s=>s.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name')===`body-${i-1}`),ph=local.getElementsByTagNameNS(P,'ph')[0],target=prompts.find(n=>n.name==='copy'||n.name==='body-old');
    if(!target||target.ph.getAttribute('idx')!==ph.getAttribute('idx'))throw Error('Crossed native placeholder index');
    const text=local.getElementsByTagNameNS(A,'rPr')[0],xfrm=direct(local,'spPr').getElementsByTagNameNS(A,'xfrm')[0],offset=xfrm.getElementsByTagNameNS(A,'off')[0],size=xfrm.getElementsByTagNameNS(A,'ext')[0],background=slide.getElementsByTagNameNS(P,'bg')[0];
    const picture=[...slide.getElementsByTagNameNS(P,'pic')].find(s=>s.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name')===`image-${i-1}`),imagePh=picture.getElementsByTagNameNS(P,'ph')[0],imageTarget=prompts.find(n=>n.name==='photo'||n.name==='picture-old');
    if(!imageTarget||imageTarget.ph.getAttribute('idx')!==imagePh.getAttribute('idx'))throw Error('Crossed picture placeholder index');
    const tc=slide.getElementsByTagNameNS(A,'tcPr')[0],fill=[...tc.children].find(n=>n.localName==='solidFill').firstElementChild,theme=parse(rel(master,'theme')),tableColor=fill.localName==='srgbClr'?fill.getAttribute('val'):theme.getElementsByTagNameNS(A,fill.getAttribute('val'))[0].firstElementChild.getAttribute('val');
    const chart=parse(rel(part,'chart')),ser=chart.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart','ser')[0],chartColor=ser.getElementsByTagNameNS(A,'srgbClr')[0].getAttribute('val');
    rows.push({name,bold:text.getAttribute('b'),tableColor,chartColor,picturePlaceholder:true,master:m.getElementsByTagNameNS(P,'cSld')[0].getAttribute('name'),hidden:slide.documentElement.getAttribute('show'),showInherited:slide.documentElement.getAttribute('showMasterSp'),x:Number(offset.getAttribute('x')),y:Number(offset.getAttribute('y')),w:Number(size.getAttribute('cx')),h:Number(size.getAttribute('cy')),font:(()=>{const raw=text.getElementsByTagNameNS(A,'latin')[0].getAttribute('typeface');return raw==='+mn-lt'||raw==='+mj-lt'?theme.getElementsByTagNameNS(A,raw==='+mn-lt'?'minorFont':'majorFont')[0].getElementsByTagNameNS(A,'latin')[0].getAttribute('typeface'):raw;})(),size:Number(text.getAttribute('sz')),background:background?.textContent??null});
   }
   return rows;
  },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith('.xml')||n.endsWith('.rels')).map(([n,b])=>[n,b.toString()])),mode});
  for(const[i,row]of actual.entries()){
   assert.equal(row.name,i===2?'Source':'Target');assert.equal(row.master,i===2?'red':'blue');assert.equal(row.hidden,i===1?'0':null);assert.equal(row.showInherited,i===1?'0':null);
   const old=i===2||mode==='preserve';assert.deepEqual([row.x,row.y,row.w,row.h],(old?[20,90,350,100]:[40,120,320,90]).map(n=>n*12700));assert.equal(row.font,old?'Arial':'Courier New');assert.equal(row.size,old?1800:2000);assert.equal(row.tableColor,old?'CC2200':'0044CC');assert.equal(row.chartColor,old?'CC2200':'0044CC');assert.equal(row.picturePlaceholder,true);assert.equal(row.bold,mode==='reset'&&i<2?'1':'0');
  }
 }
 assert.deepEqual(requests,[]);
});

test('rebind no-op preserves original source bytes and new template rebinding does not gain source-write permission',async()=>{
 const bytes=(await generatePptx(layoutRebindDocument(createPresentation))).bytes,session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v3'});
 try{
  const d=session.document,p=prepareLayoutRebind(d,[{slideId:d.slides[0].id,layoutId:'layout-blank'}],{localAppearance:'preserve'});assert.deepEqual(p.changedSlideIds,[]);d.applyLayoutRebind(p);const ready=await session.prepareSave({mode:'no-op'});assert.equal(ready.status,'ready');assert.deepEqual((await session.save(ready.plan)).bytes,bytes);
  d.defineSlideLayout({id:'new-target',name:'New target'});const change=prepareLayoutRebind(d,[{slideId:d.slides[0].id,layoutId:'new-target'}],{localAppearance:'adopt'});assert.equal(change.status,'ready');d.applyLayoutRebind(change);const blocked=await session.prepareSave({mode:'edit'});assert.equal(blocked.status,'blocked');assert.ok(blocked.report.diagnostics.some(f=>f.code==='save-edit-unsupported'));await assert.rejects(session.save(ready.plan),e=>e.code==='save-snapshot-stale');
 }finally{session.dispose();}
});

async function indexedSnapshot(api,factory,mode){
 const d=factory(api.createPresentation),p=api.prepareLayoutRebind(d,[{slideId:'page-0',layoutId:'target'},{slideId:'page-1',layoutId:'target'}],{localAppearance:mode==='preserve'?'preserve':'adopt',...(mode==='reset'?{reset:['geometry','text','background']}: {})});
 d.applyLayoutRebind(p);const layout=api.resolveLayout(d),out=await api.generatePptx(d),session=await api.importPptx(out.bytes,{projection:'content'});
 try{const ready=await session.prepareSave({mode:'no-op'});if(ready.status!=='ready')throw Error('Source-only no-op blocked');const saved=await session.save(ready.plan);if([...saved.bytes].some((b,i)=>b!==out.bytes[i])||saved.bytes.length!==out.bytes.length)throw Error('Index changed source bytes');return JSON.parse(JSON.stringify({plan:p,normalized:api.normalizePresentation(d),origins:api.resolvePresentationInheritance(d).origins,layout,svg:await api.renderLayoutToSvg(layout),warnings:out.warnings,bytes:[...out.bytes],report:session.report}));}finally{session.dispose();}
}

test('declared numeric identities survive Node/ESM/global rebind, XML and Parser provenance without turning ordinals into Core identities',async t=>{
 const api={createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind,resolveLayout,generatePptx,renderLayoutToSvg,importPptx},expected={};for(const mode of ['adopt','preserve','reset'])expected[mode]=await indexedSnapshot(api,placeholderIndexDocument,mode);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});const imports={};for(const n of ['core','layout','svg-renderer','pptx-exporter','pptx-parser'])imports['@pptkit/'+n]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${n}/dist/browser/index.js`,'utf8')).toString('base64');await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
 for(const format of ['esm','global']){
  if(format==='global')for(const n of ['core','layout','svg-renderer','pptx-exporter','pptx-parser'])await page.addScriptTag({content:await readFile(`packages/${n}/dist/browser/global.js`,'utf8')});
  const actual=await page.evaluate(async({format,factory,snapshot})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','layout','svg-renderer','pptx-exporter','pptx-parser'].map(n=>import('@pptkit/'+n)))),f=Function(`return (${factory})`)(),s=Function(`return (${snapshot})`)(),result={};for(const mode of ['adopt','preserve','reset'])result[mode]=await s(api,f,mode);return result;},{format,factory:placeholderIndexDocument.toString(),snapshot:indexedSnapshot.toString()});assert.deepEqual(actual,expected);
 }
 for(const mode of ['adopt','preserve','reset']){
  const value=expected[mode];assert.equal(value.plan.status,'ready');assert.equal(value.plan.matches.filter(m=>m.method==='index').length,4);assert.equal(value.plan.matches.filter(m=>m.targetIndex===4294967295).length,2);
  for(const layout of value.normalized.layouts)for(const ph of layout.placeholders)if(ph.key==='legacy')assert.equal(Object.hasOwn(ph,'index'),false);
  const entries=value.report.placeholderInheritance.placeholders.filter(p=>p.level==='slide');assert.equal(entries.length,12);for(const p of entries){assert.equal(p.parent.status,'matched');assert.equal(p.parent.rule,'index');assert.equal(p.parent.key,p.index);}assert.equal(entries.filter(p=>p.index===0).length,3);assert.equal(entries.filter(p=>p.index===4294967295).length,3);
  const parts=readGeneratedZip(Uint8Array.from(value.bytes));const rows=await page.evaluate(({parts})=>{
   const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',parse=n=>new DOMParser().parseFromString(parts[n],'application/xml');
   const layouts=Object.keys(parts).filter(n=>/^ppt\/slideLayouts\/[^/]+\.xml$/.test(n)).map(n=>parse(n));
   const targets=Object.fromEntries(layouts.filter(l=>l.getElementsByTagNameNS(P,'cSld')[0].getAttribute('name')==='Target').flatMap(l=>[...l.getElementsByTagNameNS(P,'sp')].map(s=>[s.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name'),Number(s.getElementsByTagNameNS(P,'ph')[0].getAttribute('idx'))])));
   const local=[];for(let i=0;i<3;i++){const doc=parse(`ppt/slides/slide${i+1}.xml`),objects=[...doc.getElementsByTagNameNS(P,'sp'),...doc.getElementsByTagNameNS(P,'pic')],values={};for(const node of objects){const name=node.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name'),ph=node.getElementsByTagNameNS(P,'ph')[0];if(ph)values[name]=Number(ph.getAttribute('idx'));}const body=objects.find(n=>n.getElementsByTagNameNS(P,'cNvPr')[0].getAttribute('name')===`body-${i}`),off=body.getElementsByTagNameNS(A,'off')[0];local.push({values,x:Number(off.getAttribute('x')),y:Number(off.getAttribute('y'))});}
   return {targets,local};
  },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith('.xml')).map(([n,b])=>[n,b.toString()]))});
  assert.deepEqual(rows.targets,{decoy:17,legacy:1,reserved:2,photo:4294967295,'new-body':0,title:73});assert.equal(new Set(Object.values(rows.targets)).size,6);
  for(let i=0;i<3;i++){assert.deepEqual(rows.local[i].values,{['title-'+i]:73,['body-'+i]:0,['legacy-'+i]:i===2?4:1,['image-'+i]:4294967295});assert.deepEqual([rows.local[i].x,rows.local[i].y],(mode==='preserve'||i===2?[30,110]:[40,130]).map(n=>n*12700));}
 }
 assert.deepEqual(requests,[]);
});

async function selectiveResetSnapshot(api,factory){
 const d=factory(init=>api.createPresentation({...init,textStylePresets:{local:{run:{fontFamily:'Georgia',color:'ABCDEF',fontSize:37,italic:true,bold:false,scriptFontFamilies:{Hans:'Preset Han',Arab:'Preset Arabic'}},paragraph:{spaceAfter:9},frame:{wrap:false,margin:7,autoFit:{mode:'resize'}}}}}));const target=d.layouts.find(l=>l.id==='target');d.replaceSlideLayout('target',{...target,placeholders:target.placeholders.map(p=>p.key==='copy'?{...p,textStyle:{...p.textStyle,run:{...p.textStyle.run,scriptFontFamilies:{Hans:'Target Han',Arab:'Target Arabic'}}}}:p)});const slide=d.slides[0],body=slide.elements[1];slide.replaceElement(body.id,{...body,textStylePreset:'local',frame:{wrap:false,margin:9},content:[{style:{spaceAfter:8},runs:[{text:'Parity',style:{fontFamily:'Arial',color:'112233',bold:false,colorOpacity:0,scriptFontFamilies:{Hans:'Local Han',Arab:'Local Arabic'}}}]}]});const plan=api.prepareLayoutRebind(d,[{slideId:'page-0',layoutId:'target'}],{localAppearance:'adopt',reset:['text.fontFamily','text.color','text.bold','text.italic','text.underline','text.strike','text.language','text.paragraph.spaceAfter','text.frame.wrap','text.frame.autoFit','text.frame.margin.left','text.scriptFontFamilies.Hans']});d.applyLayoutRebind(plan);const layout=api.resolveLayout(d),pptx=await api.generatePptx(d,{nativePlaceholderText:"inherit"});return JSON.parse(JSON.stringify({plan,author:slide.elements[1],normalized:api.normalizePresentation(d),origins:api.resolvePresentationInheritance(d).origins,layout,svg:await api.renderLayoutToSvg(layout),bytes:[...pptx.bytes]}));
}
test('selective preset reset agrees across Node and browser consumers through generated presentation',async t=>{
 const api={createPresentation,normalizePresentation,resolvePresentationInheritance,prepareLayoutRebind,resolveLayout,generatePptx,renderLayoutToSvg},expected=await selectiveResetSnapshot(api,layoutRebindDocument);assert.deepEqual(expected.author.textStylePresetReset,['fontFamily','color','bold','italic','underline','strike','language','paragraph.spaceAfter','frame.wrap','frame.autoFit','frame.margin.left','scriptFontFamilies.Hans']);assert.equal(expected.normalized.slides[0].elements[1].frame.margin.right,9);const run=expected.normalized.slides[0].elements[1].content[0].runs[0].style;assert.deepEqual(run.fontFamily,{theme:'body'});assert.deepEqual(run.color,{theme:'accent1'});assert.equal(run.fontSize,37);assert.equal(run.colorOpacity,0);assert.deepEqual(run.scriptFontFamilies,{Hans:'Target Han',Arab:'Local Arabic'});
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),imports={};for(const n of ['core','layout','svg-renderer','pptx-exporter'])imports['@pptkit/'+n]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${n}/dist/browser/index.js`,'utf8')).toString('base64');await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
 for(const format of ['esm','global']){if(format==='global')for(const n of ['core','layout','svg-renderer','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${n}/dist/browser/global.js`,'utf8')});const actual=await page.evaluate(async({format,factory,snapshot})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','layout','svg-renderer','pptx-exporter'].map(n=>import('@pptkit/'+n))));return Function(`return (${snapshot})`)()(api,Function(`return (${factory})`)());},{format,factory:layoutRebindDocument.toString(),snapshot:selectiveResetSnapshot.toString()});assert.deepEqual(actual,expected);}
});
