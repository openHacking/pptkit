import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { inspectZipArchive } from '../packages/pptx-parser/dist/index.js';
import { structureFixture, encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';

test('browser parser enforces real byte-view limits and copies without hostile iterator',async t=>{
 const bytes=encodeStructure(structureFixture()),expected=await inspectZipArchive(bytes),browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage();await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});const bundle=await readFile('packages/pptx-parser/dist/browser/index.js');
 const results=await page.evaluate(async({url,data})=>{const esm=await import(url),out=[];class Hostile extends Uint8Array{get byteLength(){return 0}get length(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}for(const api of [esm,PPTKit]){const input=new Hostile(data),info=await api.inspectZipArchive(input);let code='';try{await api.inspectZipArchive(input,{limits:{maxInputBytes:1}})}catch(e){code=e.code}const session=await api.importPptx(input,{projection:'structure'});try{const plan=await session.prepareSave({mode:'no-op'}),saved=await session.save(plan.plan);out.push({info,code,same:saved.bytes.length===data.length&&saved.bytes.every((b,i)=>b===data[i])})}finally{session.dispose()}}return out},{url:`data:text/javascript;base64,${bundle.toString('base64')}`,data:[...bytes]});
 for(const result of results){assert.deepEqual(result.info,expected);assert.equal(result.code,'archive-limit-exceeded');assert.equal(result.same,true)}
});
