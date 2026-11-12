import assert from 'node:assert/strict';
import test from 'node:test';
import {readFile} from 'node:fs/promises';
import {chromium} from '@playwright/test';
import {createPresentation,normalizePresentation} from '../packages/core/dist/index.js';
import {generatePptx} from '../packages/pptx-exporter/dist/index.js';
import {inspectPptx,importPptx} from '../packages/pptx-parser/dist/index.js';
import {presentationFormatDocuments} from '../tests/fixtures/presentation-format.mjs';
import {readGeneratedZip} from '../tests/support/read-generated-zip.mjs';
async function snapshot(api,fixture){const result=[];for(const s of fixture(api.createPresentation)){const out=await api.generatePptx(s.document);result.push({label:s.label,cx:s.cx,cy:s.cy,normalized:api.normalizePresentation(s.document),bytes:[...out.bytes],warnings:out.warnings});}return result;}

test('format descriptions and exact frame dimensions agree across Node/ESM/global and independent XML inspection',async t=>{
  const expected=await snapshot({createPresentation,normalizePresentation,generatePptx},presentationFormatDocuments),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort();});
  const imports={},globals=[];for(const name of ['core','layout','pptx-exporter']){imports['@pptkit/'+name]='data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`,'utf8')).toString('base64');globals.push(await readFile(`packages/${name}/dist/browser/global.js`,'utf8'));}
  await page.setContent(`<script type="importmap">${JSON.stringify({imports})}</script>`);
  for(const format of ['esm','global']){if(format==='global')for(const content of globals)await page.addScriptTag({content});const actual=await page.evaluate(async({format,fixture,snapshot})=>{const api=format==='global'?PPTKit:Object.assign({},...await Promise.all(['core','pptx-exporter'].map(n=>import('@pptkit/'+n))));return Function('api',`return (${snapshot})(api,(${fixture}))`)(api);},{format,fixture:presentationFormatDocuments.toString(),snapshot:snapshot.toString()});assert.deepEqual(actual,expected);}
  for(const scene of expected){assert.deepEqual(scene.warnings,[]);const bytes=Uint8Array.from(scene.bytes),parts=readGeneratedZip(bytes);const actual=await page.evaluate(({app,presentation})=>{const p=new DOMParser(),a=p.parseFromString(app,'application/xml'),d=p.parseFromString(presentation,'application/xml'),size=d.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main','sldSz')[0],ns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';return{label:a.getElementsByTagNameNS(ns,'PresentationFormat')[0].textContent,cx:Number(size.getAttribute('cx')),cy:Number(size.getAttribute('cy')),preset:size.getAttribute('type'),company:a.getElementsByTagNameNS(ns,'Company')[0].textContent};},{app:parts.get('docProps/app.xml').toString(),presentation:parts.get('ppt/presentation.xml').toString()});assert.deepEqual(actual,{label:scene.label,cx:scene.cx,cy:scene.cy,preset:null,company:'A & <B>'});assert.equal((await inspectPptx(bytes)).slides.length,1);const s=await importPptx(bytes,{projection:'structure'});try{const plan=await s.prepareSave({mode:'no-op'});assert.equal(plan.status,'ready');assert.deepEqual([...(await s.save(plan.plan)).bytes],scene.bytes);}finally{s.dispose();}}
  assert.deepEqual(requests,[]);
});
