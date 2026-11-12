import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspectIccProfile } from '../dist/index.js';
import { extractPngIccProfile, inspectIccProfile as inspectNode } from '../dist/node.js';

const fixture = async () => extractPngIccProfile(await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png',import.meta.url))).profileBytes;
test('ICC directory inventory describes the real profile and permits exact shared tag data',async()=>{
 const bytes=await fixture(),info=inspectIccProfile(bytes);assert.equal(info.byteLength,588);assert.equal(info.colorSpace,'RGB ');assert.equal(info.connectionSpace,'XYZ ');assert.equal(info.tags.length,11);
 const channels=info.tags.filter(tag=>['rTRC','gTRC','bTRC'].includes(tag.signature));assert.equal(channels.length,3);assert.equal(new Set(channels.map(tag=>`${tag.offset}:${tag.size}`)).size,1);
 assert.deepEqual(inspectNode(bytes),info);const wrapped=new Uint8Array(bytes.length+2);wrapped.set(bytes,1);assert.deepEqual(inspectIccProfile(wrapped.subarray(1,-1)),info);
});

test('ICC directory rejects malformed size, signature, extents, duplicate tags and partial overlap',async()=>{
 const base=Buffer.from(await fixture());
 const cases=[
  b=>b.writeUInt32BE(b.length-1,0), b=>b.write('bad!',36), b=>b.writeUInt32BE(0xffffffff,128),
  b=>b.writeUInt32BE(263,136), b=>b.writeUInt32BE(128,136), b=>b.writeUInt32BE(0,140),
  b=>b.writeUInt32BE(0xffffffff,140), b=>b.copy(b,144,132,136), b=>{b.writeUInt32BE(268,148);b.writeUInt32BE(20,152)},
 ];
 for(const mutate of cases){const bytes=Buffer.from(base);mutate(bytes);assert.throws(()=>inspectIccProfile(bytes));}
 for(const n of [0,1,127,128,131])assert.throws(()=>inspectIccProfile(base.subarray(0,n)),/Invalid ICC/);
});

test('ICC bounds, snapshots, cancellation and cross-realm shared buffers are enforced',async()=>{
 const bytes=await fixture();assert.throws(()=>inspectIccProfile(bytes,{maxBytes:bytes.length-1}),/maxBytes/);assert.throws(()=>inspectIccProfile(bytes,{maxTags:10}),/maxTags/);
 for(const key of ['maxBytes','maxTags'])for(const value of [0,-1,NaN,1.5,Infinity])assert.throws(()=>inspectIccProfile(bytes,{[key]:value}),RangeError);
 const c=new AbortController(),reason=new Error('cancel ICC');c.abort(reason);assert.throws(()=>inspectIccProfile(bytes,{signal:c.signal}),e=>e===reason);
 class Hostile extends Uint8Array {get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}
 assert.deepEqual(inspectIccProfile(new Hostile(bytes)),inspectIccProfile(bytes));assert.throws(()=>inspectIccProfile(new Hostile(bytes),{maxBytes:1}),/maxBytes/);
 if(typeof SharedArrayBuffer!=='undefined'){const vm=await import('node:vm'),foreign=vm.runInNewContext(`new SharedArrayBuffer(${bytes.length})`);assert.throws(()=>inspectIccProfile(new Uint8Array(foreign)),/SharedArrayBuffer/)}
});
