import assert from 'node:assert/strict';
import test from 'node:test';
import {readFile} from 'node:fs/promises';
import {chromium} from '@playwright/test';
import {createPresentation,normalizePresentation,resolvePresentationInheritance} from '../packages/core/dist/index.js';
import {resolveLayout} from '../packages/layout/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {renderLayoutToSvg} from '../packages/svg-renderer/dist/index.js';
import {importPptx} from '../packages/pptx-parser/dist/index.js';
import {masterArtworkDocument} from '../tests/fixtures/master-artwork.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
async function snapshot(api,fixture){const d=fixture(api.createPresentation),r=api.resolvePresentationInheritance(d),n=api.normalizePresentation(d),l=api.resolveLayout(d),out=await api.generatePptx(d);if(JSON.stringify(r.presentation)!==JSON.stringify(n))throw Error('Traced IR changed');return{normalized:n,origins:r.origins,layout:l,svg:await api.renderLayoutToSvg(l),bytes:[...out.bytes],warnings:out.warnings};}

test('master artwork stays owned and visibility/z-order/media agree through public Node/browser output and independent DOM',async t=>{
  const expected=await snapshot({createPresentation,normalizePresentation,resolvePresentationInheritance,resolveLayout,generatePptx,renderLayoutToSvg},masterArtworkDocument);assert.deepEqual(expected.warnings,[]);
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});
  const imports={},globals=[];for(const name of ['core','layout','pptx-exporter','svg-renderer']){imports['@pptkit/'+name]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`,'utf8')).toString('base64');globals.push(await readFile(`packages/${name}/dist/browser/global.js`,'utf8'));}
  await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
  for(const format of ['esm','global']){if(format==='global')for(const content of globals)await page.addScriptTag({content});const actual=await page.evaluate(async({format,fixture,snapshot})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','layout','pptx-exporter','svg-renderer'].map(n=>import('@pptkit/'+n))));return Function('api',`return (${snapshot})(api,(${fixture}))`)(api);},{format,fixture:masterArtworkDocument.toString(),snapshot:snapshot.toString()});assert.deepEqual(actual,expected);}
  const parts=readGeneratedZip(Uint8Array.from(expected.bytes));assert.equal([...parts.keys()].filter(n=>n.startsWith('ppt/media/')).length,1);
  const truth=await page.evaluate(({parts,svgs})=>{
    const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const parse=s=>new DOMParser().parseFromString(s,'application/xml'),doc=name=>parse(parts[name]);
    const names=name=>[...doc(name).getElementsByTagNameNS(P,'cNvPr')].map(n=>n.getAttribute('name')).filter(Boolean);
    const rels=name=>[...doc(name).getElementsByTagName('Relationship')].filter(n=>n.getAttribute('Type').endsWith('/image')).map(n=>n.getAttribute('Target'));
    const masters=[2,3].map(i=>({names:names(`ppt/slideMasters/slideMaster${i}.xml`),images:rels(`ppt/slideMasters/_rels/slideMaster${i}.xml.rels`)}));
    const layouts=[2,3,4].map(i=>({names:names(`ppt/slideLayouts/slideLayout${i}.xml`),show:doc(`ppt/slideLayouts/slideLayout${i}.xml`).documentElement.getAttribute('showMasterSp'),images:rels(`ppt/slideLayouts/_rels/slideLayout${i}.xml.rels`),prompts:doc(`ppt/slideLayouts/slideLayout${i}.xml`).getElementsByTagNameNS(P,'ph').length}));
    const slides=[1,2,3,4,5,6].map(i=>({names:names(`ppt/slides/slide${i}.xml`),show:doc(`ppt/slides/slide${i}.xml`).documentElement.getAttribute('showMasterSp'),images:rels(`ppt/slides/_rels/slide${i}.xml.rels`)}));
    const preview=svgs.map(svg=>{const d=parse(svg);return{text:d.documentElement.textContent,images:d.getElementsByTagName('image').length,labels:[...d.querySelectorAll('[aria-label]')].map(n=>n.getAttribute('aria-label')),fills:[...d.querySelectorAll('rect')].map(n=>n.getAttribute('fill'))};});
    return{masters,layouts,slides,preview};
  },{parts:Object.fromEntries([...parts].filter(([n])=>n.endsWith('.xml')||n.endsWith('.rels')).map(([n,b])=>[n,b.toString()])),svgs:expected.svg.slides.map(s=>s.svg)});
  for(const[i,id]of ['red','blue'].entries()){assert.deepEqual(truth.masters[i].names,[id+'-master-shape',id+'-master-text',id+'-master-logo',id+'-master-group',id+'-group-shape']);assert.deepEqual(truth.masters[i].images,['../media/logo.png']);}
  assert.deepEqual(truth.layouts.map(l=>l.show),[null,'0',null]);assert.ok(truth.layouts.every(l=>l.prompts===1&&l.images[0]==='../media/logo.png'));
  assert.deepEqual(truth.slides.map(s=>s.show),[null,'0',null,null,'1',null]);assert.ok(truth.slides.every(s=>s.names.length===3&&s.images.length===0));
  assert.deepEqual(truth.preview.map(p=>p.images),[2,0,1,2,1,2]);
  for(const[i,master,layout]of [[0,'red','red-on'],[1,null,null],[2,null,'red-off'],[3,'blue','blue-on'],[4,null,'red-off'],[5,'red','red-on']]){const p=truth.preview[i];assert.equal(p.text.includes('MASTER '),master!==null);assert.equal(p.text.includes('LAYOUT '),layout!==null);if(master)assert.ok(p.labels.indexOf(master+'-master-shape')<p.labels.indexOf(layout+'-layout-shape'));if(layout)assert.ok(p.labels.indexOf(layout+'-layout-shape')<p.labels.indexOf('local-shape-'+i));}
  const session=await importPptx(Uint8Array.from(expected.bytes),{projection:'structure'});try{const p=await session.prepareSave({mode:'no-op'});assert.equal(p.status,'ready');assert.deepEqual([...(await session.save(p.plan)).bytes],expected.bytes);}finally{session.dispose();}
  assert.deepEqual(requests,[]);
});

test('shared master artwork replacement changes only visible bound artwork and retains bystander resources',async()=>{
  const d=masterArtworkDocument(createPresentation),before=await renderLayoutToSvg(resolveLayout(d)),oldOutput=readGeneratedZip((await generatePptx(d)).bytes),oldMaster=d.masters[0];
  d.replaceSlideMaster('red',{...oldMaster,elements:oldMaster.elements.filter(e=>e.id!=='red-master-logo').map(e=>e.id==='red-master-text'?{...e,content:e.content.map(p=>({...p,runs:p.runs.map(r=>({...r,text:'MASTER red updated'}))}))}:e.id==='red-master-shape'?{...e,box:{...e.box,x:40}}:e)});
  const after=await renderLayoutToSvg(resolveLayout(d)),output=readGeneratedZip((await generatePptx(d)).bytes);
  for(const i of [1,2,3,4])assert.equal(after.slides[i].svg,before.slides[i].svg);
  for(const i of [0,5]){assert.notEqual(after.slides[i].svg,before.slides[i].svg);assert.ok(after.slides[i].svg.includes('MASTER red updated'));}
  assert.equal(oldMaster.elements.length,4);assert.equal(d.masters[0].elements.length,3);assert.equal(d.getAssetReferences('logo').length,4);assert.throws(()=>d.removeAsset('logo'),/referenced/);
  assert.deepEqual([...output.keys()].sort(),[...oldOutput.keys()].sort());
  const changed=[...oldOutput].filter(([name,bytes])=>!bytes.equals(output.get(name))).map(([name])=>name).sort();
  assert.deepEqual(changed,['ppt/slideMasters/_rels/slideMaster2.xml.rels','ppt/slideMasters/slideMaster2.xml']);
});

test('new artwork/visibility authoring does not acquire source-save permission and reset restores exact no-op',async()=>{
  const bytes=(await generatePptx(masterArtworkDocument(createPresentation))).bytes,session=await importPptx(bytes,{projection:'content',editSaveProfile:'text-runs-v3'});
  try{
    const slide=session.document.slides[0];slide.setShowInheritedShapes(false);
    const blocked=await session.prepareSave({mode:'edit'});assert.equal(blocked.status,'blocked');assert.ok(blocked.report.diagnostics.some(d=>d.code==='save-edit-unsupported'));
    slide.setShowInheritedShapes();const ready=await session.prepareSave({mode:'no-op'});assert.equal(ready.status,'ready');assert.deepEqual((await session.save(ready.plan)).bytes,bytes);
    session.document.defineSlideMaster({id:'new-source-art',name:'No source binding',elements:[{id:'new-source-shape',type:'shape',shape:'rect',box:{x:0,y:0,width:20,height:20}}]});
    const added=await session.prepareSave({mode:'edit'});assert.equal(added.status,'blocked');assert.ok(added.report.diagnostics.some(d=>d.code==='save-edit-unsupported'));
    await assert.rejects(session.save(ready.plan),e=>e.code==='save-snapshot-stale');
  }finally{session.dispose();}
});
