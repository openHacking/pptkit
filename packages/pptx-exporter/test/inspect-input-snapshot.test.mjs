import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspectGif, inspectJpegExifOrientation, inspectPng, inspectApng } from '../dist/index.js';
import { normalizeJpegOrientationInBrowser } from '../dist/dom.js';

test('GIF and JPEG inspectors use bounded snapshots despite hostile typed-array subclasses',async()=>{
 const gif=Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7','base64'),jpeg=await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg',import.meta.url));
 class Hostile extends Uint8Array{*[Symbol.iterator](){throw Error('unexpected iterator')}get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}}
 for(const [inspect,source] of [[inspectGif,gif],[inspectJpegExifOrientation,jpeg]]){
  const expected=inspect(source),hostile=new Hostile(source);assert.deepEqual(inspect(hostile),expected);assert.throws(()=>inspect(hostile,{maxBytes:1}),/maxBytes/);
  if(typeof SharedArrayBuffer!=='undefined'){const shared=new Hostile(new SharedArrayBuffer(source.length));shared.set(source);assert.throws(()=>inspect(shared),/SharedArrayBuffer/);}
  const mutable=Uint8Array.from(source);let reads=0;const signal={get aborted(){if(++reads===2)mutable[0]^=1;return false}};assert.deepEqual(inspect(mutable,{signal}),expected);assert.notEqual(mutable[0],source[0]);
 }
});

test('image inspectors reject shared buffers from another JavaScript realm',async()=>{
 const vm=await import('node:vm'),foreign=vm.runInNewContext('new SharedArrayBuffer(32)'),input=new Uint8Array(foreign);assert.equal(foreign instanceof SharedArrayBuffer,false);
 for(const inspect of [inspectGif,inspectJpegExifOrientation,inspectPng,inspectApng])assert.throws(()=>inspect(input),/SharedArrayBuffer/);
});

test('browser JPEG normalization checks actual view size and shared buffer before decoding',async()=>{
 const source=await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg',import.meta.url));class Hostile extends Uint8Array{get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}
 const input=new Hostile(source);await assert.rejects(normalizeJpegOrientationInBrowser(input,{maxInputBytes:1}),/maxInputBytes/);await assert.rejects(normalizeJpegOrientationInBrowser(input),/browser document/);
 const vm=await import('node:vm'),foreign=vm.runInNewContext(`new SharedArrayBuffer(${source.length})`),shared=new Hostile(foreign);shared.set(source);await assert.rejects(normalizeJpegOrientationInBrowser(shared),/SharedArrayBuffer/);
});
