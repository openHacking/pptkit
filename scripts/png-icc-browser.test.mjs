import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { chromium } from '@playwright/test';
import { extractPngIccProfile, inspectIccProfile } from '../packages/pptx-exporter/dist/node.js';

function chunk(name,data=Buffer.alloc(0)){
 const type=Buffer.from(name),length=Buffer.alloc(4),sum=Buffer.alloc(4);length.writeUInt32BE(data.length);let crc=0xffffffff;
 for(const byte of Buffer.concat([type,data])){crc^=byte;for(let i=0;i<8;i++)crc=(crc&1)?0xedb88320^(crc>>>1):crc>>>1;}sum.writeUInt32BE((crc^0xffffffff)>>>0);return Buffer.concat([length,type,data,sum]);
}
function fixture(compressed){
 const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(1);ihdr.writeUInt32BE(1,4);ihdr[8]=8;ihdr[9]=2;
 return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),chunk('iCCP',Buffer.concat([Buffer.from('Test\0\0'),compressed])),chunk('IDAT',Buffer.from([1])),chunk('IEND')]);
}
function plainFixture(){const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(1);ihdr.writeUInt32BE(1,4);ihdr[8]=8;ihdr[9]=2;return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),chunk('IDAT',Buffer.from([1])),chunk('IEND')]);}

test('browser PNG ICC extraction agrees with Node and rejects invalid zlib streams',async t=>{
 const real=await readFile('tests/fixtures/png-icc-srgb.png'),profile=Buffer.alloc(100000,65),compressed=deflateSync(profile),checksum=Buffer.from(compressed);checksum[checksum.length-1]^=1;
 const cases=[fixture(checksum),fixture(compressed.subarray(0,-1)),fixture(Buffer.concat([compressed,Buffer.from([0])])),fixture(Buffer.concat([compressed,compressed]))];
 const expected=extractPngIccProfile(real),adapter='data:text/javascript;base64,'+(await readFile('packages/pptx-exporter/dist/browser/dom.js')).toString('base64');
 const browser=await chromium.launch();t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',r=>{requests.push(r.request().url());return r.abort()});
 for(const name of ['core','layout','pptx-exporter'])await page.addScriptTag({content:await readFile(`packages/${name}/dist/browser/global.js`,'utf8')});
 const result=await page.evaluate(async({adapter,real,bomb,cases,plain})=>{
  const api=await import(adapter),bytes=Uint8Array.from(real),parsed=await api.extractPngIccProfileInBrowser(bytes);const failures=[];
  for(const input of cases)try{await api.extractPngIccProfileInBrowser(Uint8Array.from(input));failures.push('accepted')}catch(e){failures.push(e.name)}
  for(const [input,options] of [[bomb,{maxProfileBytes:99999}],[real,{maxProfileBytes:0}],[real,{maxBytes:1}]])try{await api.extractPngIccProfileInBrowser(Uint8Array.from(input),options);failures.push('accepted')}catch(e){failures.push(e.name)}
  const mutable=Uint8Array.from(real),pending=api.extractPngIccProfileInBrowser(mutable);mutable.fill(0);const detached=await pending;
  const exact=await api.extractPngIccProfileInBrowser(Uint8Array.from(bomb),{maxProfileBytes:100000});
  const aborted=new AbortController();aborted.abort('stop');let reason;try{await api.extractPngIccProfileInBrowser(bytes,{signal:aborted.signal})}catch(e){reason=e}
  const delayed=new AbortController(),late=api.extractPngIccProfileInBrowser(Uint8Array.from(bomb),{signal:delayed.signal});delayed.abort('late');let lateReason;try{await late}catch(e){lateReason=e}
  class Hostile extends Uint8Array{get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}
  const hostile=await api.extractPngIccProfileInBrowser(new Hostile(real));let shared='unavailable';if(typeof SharedArrayBuffer!=='undefined'){const source=new Uint8Array(new SharedArrayBuffer(real.length));source.set(real);try{await api.extractPngIccProfileInBrowser(source);shared='accepted'}catch(e){shared=e.message}}
  let reads=0;try{await api.extractPngIccProfileInBrowser(bytes,{get maxCompressedProfileBytes(){reads++;return 1}})}catch{}
  const original=globalThis.DecompressionStream;globalThis.DecompressionStream=undefined;let unavailable;try{await api.extractPngIccProfileInBrowser(bytes)}catch(e){unavailable=e.message}const none=await api.extractPngIccProfileInBrowser(Uint8Array.from(plain));globalThis.DecompressionStream=original;
  let badDirectory;try{const bad=new Uint8Array(parsed.profileBytes);bad[36]=0;PPTKit.inspectIccProfile(bad);badDirectory='accepted'}catch(e){badDirectory=e.name}
  return {inspection:parsed.inspection,directory:PPTKit.inspectIccProfile(parsed.profileBytes),badDirectory,profile:[...parsed.profileBytes],detached:[...detached.profileBytes],hostile:[...hostile.profileBytes],exact:exact.profileBytes.length,failures,reason,lateReason,unavailable,shared,reads,none:!('profileBytes' in none)};
 },{adapter,real:[...real],bomb:[...fixture(compressed)],cases:cases.map(x=>[...x]),plain:[...plainFixture()]});
 assert.deepEqual(result.inspection,expected.inspection);assert.deepEqual(result.profile,[...expected.profileBytes]);assert.deepEqual(result.detached,result.profile);
 assert.deepEqual(result.directory,inspectIccProfile(expected.profileBytes));assert.notEqual(result.badDirectory,'accepted');
 assert.ok(result.failures.slice(0,4).every(x=>x!=='accepted'));assert.deepEqual(result.failures.slice(4),['RangeError','RangeError','RangeError']);
 assert.equal(result.reason,'stop');assert.equal(result.lateReason,'late');assert.match(result.unavailable,/DecompressionStream/);assert.deepEqual(requests,[]);
 assert.deepEqual(result.hostile,result.profile);assert.equal(result.exact,100000);if(result.shared!=='unavailable')assert.match(result.shared,/SharedArrayBuffer/);assert.equal(result.reads,1);assert.equal(result.none,true);
});
