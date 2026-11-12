import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { inspectPng, inspectApng, inspectPngIcc } from '../packages/pptx-exporter/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { renderPresentationToSvg } from '../packages/svg-renderer/dist/index.js';

test('invalid PNG ICC envelope diagnostics match Node and browser without changing bytes', async t => {
 const original=await readFile('tests/fixtures/png-icc-srgb.png');let offset=8, start=-1, end=-1;
 while(offset<original.length){const length=original.readUInt32BE(offset),name=original.toString('ascii',offset+4,offset+8),next=offset+length+12;if(name==='iCCP'){start=offset;end=next;break}offset=next}
 assert.ok(start>0);const bytes=Buffer.concat([original.subarray(0,end),original.subarray(start,end),original.subarray(end)]);
 const d=createPresentation();d.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:bytes}});d.addSlide({elements:[{type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}}]});
 const expected=await generatePptx(d,{pngDiagnostics:'warn'});assert.deepEqual(expected.warnings.map(item=>item.code),['png-icc-color-unverified','png-icc-inspection-incomplete']);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];page.on('request',request=>requests.push(request.url()));
 for(const name of ['core','layout','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async data=>{const d=PPTKit.createPresentation();d.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:Uint8Array.from(data)}});d.addSlide({elements:[{type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}}]});const result=await PPTKit.generatePptx(d,{pngDiagnostics:'warn'});return{bytes:[...result.bytes],warnings:result.warnings}},[...bytes]);
 assert.deepEqual(actual,{bytes:[...expected.bytes],warnings:expected.warnings});assert.deepEqual(requests,[]);
});

test('PNG inventory and corruption rejection agree in Node and browser ESM/global',async t=>{
 const bytes=await readFile('tests/fixtures/transparent-replacement.png'),expected=inspectPng(bytes),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort()});
 const dataUrl=b=>`data:text/javascript;base64,${b.toString('base64')}`,core=dataUrl(await readFile('packages/core/dist/browser/index.js')),layout=dataUrl(await readFile('packages/layout/dist/browser/index.js'));
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':core,'@pptkit/layout':layout}})}</script>`);await page.addScriptTag({content:await readFile('packages/pptx-exporter/dist/browser/global.js','utf8')});const bundle=await readFile('packages/pptx-exporter/dist/browser/index.js');
 const results=await page.evaluate(async({url,data})=>{const module=await import(url),result=[];for(const api of [module,globalThis.PPTKit]){const bytes=Uint8Array.from(data),info=api.inspectPng(bytes);bytes[17]^=1;let error='';try{api.inspectPng(bytes)}catch(e){error=e.message}result.push({info,error});}return result},{url:`data:text/javascript;base64,${bundle.toString('base64')}`,data:[...bytes]});
 for(const result of results){assert.deepEqual(result.info,expected);assert.match(result.error,/CRC/)}assert.deepEqual(requests,[]);
});

test('PNG ICC inventory and warnings agree in Node and browser',async t=>{
 const bytes=await readFile('tests/fixtures/png-icc-srgb.png'),d=createPresentation();d.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:bytes}});d.addSlide({id:'s',elements:[{id:'i',type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}}]});
 const expected=await generatePptx(d,{pngDiagnostics:'warn'}),info=inspectPng(bytes),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();
 for(const name of ['core','layout','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async data=>{const bytes=Uint8Array.from(data),d=PPTKit.createPresentation();d.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:bytes}});d.addSlide({id:'s',elements:[{id:'i',type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}}]});const result=await PPTKit.generatePptx(d,{pngDiagnostics:'warn'});let limit;try{PPTKit.inspectPngIcc(bytes,{maxCompressedProfileBytes:1})}catch(e){limit=e.name}return {info:PPTKit.inspectPng(bytes),icc:PPTKit.inspectPngIcc(bytes),bytes:[...result.bytes],warnings:result.warnings,limit}},[...bytes]);
 assert.deepEqual(actual,{info,icc:inspectPngIcc(bytes),bytes:[...expected.bytes],warnings:expected.warnings,limit:'RangeError'});assert.deepEqual(actual.warnings.map(w=>w.code),['png-icc-color-unverified']);
});

test('APNG frame inventory agrees in Node and browser ESM/global',async t=>{
 const bytes=await readFile('tests/fixtures/two-frame-apng.png'),expected=inspectApng(bytes),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),url=b=>`data:text/javascript;base64,${b.toString('base64')}`,core=url(await readFile('packages/core/dist/browser/index.js')),layout=url(await readFile('packages/layout/dist/browser/index.js'));
 await page.setContent(`<script type="importmap">${JSON.stringify({imports:{'@pptkit/core':core,'@pptkit/layout':layout}})}</script>`);await page.addScriptTag({content:await readFile('packages/pptx-exporter/dist/browser/global.js','utf8')});const bundle=url(await readFile('packages/pptx-exporter/dist/browser/index.js'));
 const results=await page.evaluate(async({bundle,data})=>{const esm=await import(bundle);return [esm,globalThis.PPTKit].map(api=>api.inspectApng(Uint8Array.from(data)))},{bundle,data:[...bytes]});for(const result of results)assert.deepEqual(result,expected);
});

test('APNG export diagnostic agrees in Node and browser while preserving package bytes',async t=>{
 const bytes=await readFile('tests/fixtures/two-frame-apng.png'),d=createPresentation();d.registerAsset({id:'animation',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:bytes}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'animation',box:{x:0,y:0,width:20,height:20}}]});const expected=await generatePptx(d,{pngDiagnostics:'warn'});
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of ['core','layout','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async data=>{const d=PPTKit.createPresentation();d.registerAsset({id:'animation',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:Uint8Array.from(data)}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'animation',box:{x:0,y:0,width:20,height:20}}]});const result=await PPTKit.generatePptx(d,{pngDiagnostics:'warn'});return {bytes:[...result.bytes],warnings:result.warnings,status:result.status}},[...bytes]);
 assert.deepEqual(actual,{bytes:[...expected.bytes],warnings:expected.warnings,status:expected.status});assert.deepEqual(actual.warnings.map(w=>w.code),['apng-animation-playback-unverified']);
 const session=await importPptx(expected.bytes,{projection:'content'});try{const asset=session.document.assets[0];assert.equal(asset.mimeType,'image/png');assert.deepEqual(Buffer.from(asset.source.value,'base64'),bytes);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,expected.bytes);}finally{session.dispose();}
});

test('LibreOfficeDev third headless resave retains APNG media on import and no-op save',async()=>{
 const bytes=await readFile('tests/fixtures/apng-libreoffice-resave.pptx'),image=await readFile('tests/fixtures/two-frame-apng.png'),session=await importPptx(bytes,{projection:'content'});
 try{assert.equal(session.document.slides.length,1);assert.equal(session.document.slides[0].elements.filter(x=>x.type==='image').length,1);const asset=session.document.assets[0];assert.equal(asset.mimeType,'image/png');assert.deepEqual(Buffer.from(asset.source.value,'base64'),image);assert.equal(inspectApng(image).animation.frames.length,2);const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));}finally{session.dispose();}
});

test('declared APNG preview warning and SVG agree between Node and browser',async t=>{
 const bytes=await readFile('tests/fixtures/two-frame-apng.png'),d=createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:bytes}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});const expected=await renderPresentationToSvg(d);
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of ['core','layout','svg-renderer'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async data=>{const d=PPTKit.createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:Uint8Array.from(data)}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});return PPTKit.renderPresentationToSvg(d)},[...bytes]);assert.deepEqual(actual,expected);assert.deepEqual(actual.warnings.map(w=>w.code),['apng-preview-playback-unverified']);
});

test('byte-backed PNG animation scan agrees between Node and browser',async t=>{
 const bytes=await readFile('tests/fixtures/two-frame-apng.png'),d=createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/png',source:{type:'bytes',value:bytes}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});const expected=await renderPresentationToSvg(d,{pngAnimationDiagnostics:'warn'});
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of ['core','layout','svg-renderer'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const actual=await page.evaluate(async data=>{const d=PPTKit.createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/png',source:{type:'bytes',value:Uint8Array.from(data)}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});return PPTKit.renderPresentationToSvg(d,{pngAnimationDiagnostics:'warn'})},[...bytes]);assert.deepEqual(actual,expected);assert.deepEqual(actual.warnings.map(w=>w.code),['apng-preview-playback-unverified']);
});

test('browser PNG inspectors isolate source mutations after capture',async t=>{
 const bytes=await readFile('tests/fixtures/two-frame-apng.png'),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();await page.addScriptTag({content:await readFile('packages/pptx-exporter/dist/browser/global.js','utf8')});
 const results=await page.evaluate(data=>[PPTKit.inspectPng,PPTKit.inspectApng].map(inspect=>{const bytes=Uint8Array.from(data),baseline=inspect(bytes);let count=0;const signal={get aborted(){if(++count===2)bytes[17]^=1;return false}};return {baseline,actual:inspect(bytes,{signal}),changed:bytes[17]!==data[17]}}),[...bytes]);
 for(const result of results){assert.deepEqual(result.actual,result.baseline);assert.equal(result.changed,true);}
});

test('browser registration and resolver enforce real byte view limits',async t=>{
 const bytes=await readFile('tests/fixtures/transparent-replacement.png'),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();for(const name of ['core','layout','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const result=await page.evaluate(async data=>{class Hostile extends Uint8Array{get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}const source=new Hostile(data),d=PPTKit.createPresentation();d.registerAsset({id:'a',kind:'image',source:{type:'path',value:'virtual'}});d.addSlide({id:'s',elements:[{id:'i',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});const exported=await PPTKit.generatePptx(d,{maxAssetBytes:1,resolveAsset:()=>({bytes:source,mimeType:'image/png'})});let shared='unavailable';if(typeof SharedArrayBuffer!=='undefined'){const input=new Hostile(new SharedArrayBuffer(data.length));input.set(data);try{PPTKit.createPresentation().registerAsset({id:'b',kind:'image',source:{type:'bytes',value:input}});shared='accepted'}catch(e){shared=e.message}}return {warnings:exported.warnings.map(w=>w.code),shared}},[...bytes]);
 assert.ok(result.warnings.includes('asset-read-failed'));if(result.shared!=='unavailable')assert.match(result.shared,/SharedArrayBuffer/);
});

test('embedded image/apng declaration imports as a reported byte asset and survives no-op save',async t=>{
 const {createRequire}=await import('node:module'),{unzipSync,zipSync,strFromU8,strToU8}=createRequire(new URL('../packages/pptx-parser/package.json',import.meta.url))('fflate');
 const image=await readFile('tests/fixtures/two-frame-apng.png'),d=createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/apng',source:{type:'bytes',value:image}});d.addSlide({id:'s',elements:[{id:'image',type:'image',assetId:'a',box:{x:0,y:0,width:20,height:20}}]});
 const original=(await generatePptx(d)).bytes,parts=unzipSync(original);parts['[Content_Types].xml']=strToU8(strFromU8(parts['[Content_Types].xml']).replace('ContentType="image/png"','ContentType="image/apng"'));const bytes=zipSync(parts),session=await importPptx(bytes,{projection:'content'});
 try{assert.equal(session.document.assets[0].mimeType,'image/apng');assert.deepEqual(Buffer.from(session.document.assets[0].source.value,'base64'),image);assert.ok(session.report.findings.some(f=>f.operation==='import'&&f.reason.includes('animated-image playback')));const svg=await renderPresentationToSvg(session.document);assert.ok(svg.warnings.some(w=>w.code==='apng-preview-playback-unverified'));const plan=await session.prepareSave({mode:'no-op'});assert.deepEqual((await session.save(plan.plan)).bytes,new Uint8Array(bytes));const rebuilt=await generatePptx(session.document,{pngDiagnostics:'warn'});assert.ok(rebuilt.warnings.some(w=>w.code==='apng-animation-playback-unverified'));
  const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});const bundle=await readFile('packages/pptx-parser/dist/browser/index.js');const result=await page.evaluate(async({url,data})=>{const esm=await import(url),out=[];for(const api of [esm,PPTKit]){const s=await api.importPptx(Uint8Array.from(data),{projection:'content'});try{const plan=await s.prepareSave({mode:'no-op'}),saved=await s.save(plan.plan);out.push({asset:s.document.assets[0],same:saved.bytes.length===data.length&&saved.bytes.every((b,i)=>b===data[i])})}finally{s.dispose()}}return out},{url:`data:text/javascript;base64,${bundle.toString('base64')}`,data:[...bytes]});for(const x of result){assert.deepEqual(x.asset,session.document.assets[0]);assert.equal(x.same,true);}
 }finally{session.dispose();}
 parts['ppt/media/a.png']=new Uint8Array(await readFile('tests/fixtures/exif-orientation6.jpg'));const mismatch=await importPptx(zipSync(parts),{projection:'content'});try{assert.equal(mismatch.document.assets.length,0);assert.equal(mismatch.document.slides[0].elements.length,0);assert.ok(mismatch.report.findings.some(f=>f.tier==='unsupported'&&f.reason.includes('MIME/signature/dimensions')))}finally{mismatch.dispose();}
});
