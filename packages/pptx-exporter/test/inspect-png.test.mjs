import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { inspectPng, inspectApng, inspectPngIcc } from '../dist/index.js';
import { inspectPng as inspectNode } from '../dist/node.js';
import { extractPngIccProfile } from '../dist/node.js';
import { generatePptx } from '../dist/index.js';
import { createPresentation } from '@pptkit/core';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const signature=Buffer.from([137,80,78,71,13,10,26,10]);
function chunk(name,data=Buffer.alloc(0)){
 const type=Buffer.from(name),length=Buffer.alloc(4),sum=Buffer.alloc(4);length.writeUInt32BE(data.length);let crc=0xffffffff;
 for(const byte of Buffer.concat([type,data])){crc^=byte;for(let i=0;i<8;i++)crc=(crc&1)?0xedb88320^(crc>>>1):crc>>>1;}sum.writeUInt32BE((crc^0xffffffff)>>>0);return Buffer.concat([length,type,data,sum]);
}
function ihdr({width=4,height=2,depth=8,color=6,compression=0,filter=0,interlace=0}={}){const b=Buffer.alloc(13);b.writeUInt32BE(width);b.writeUInt32BE(height,4);b.set([depth,color,compression,filter,interlace],8);return chunk('IHDR',b)}
const png=(...chunks)=>Buffer.concat([signature,...chunks]);
const idat=()=>chunk('IDAT',Buffer.from([1,2,3])),end=()=>chunk('IEND');

test('PNG inspection reports source dimensions and verifies real image CRC without mutating input',async()=>{
 const bytes=await readFile(new URL('../../../tests/fixtures/transparent-replacement.png',import.meta.url)),before=Buffer.from(bytes),result=inspectPng(bytes);
 assert.deepEqual({width:result.width,height:result.height,bitDepth:result.bitDepth,colorType:result.colorType},{width:4,height:2,bitDepth:8,colorType:6});assert.ok(result.idatBytes>0);assert.equal(result.animationChunks,false);assert.deepEqual(result.unknownCriticalChunks,[]);assert.deepEqual(bytes,before);assert.deepEqual(inspectNode(bytes),result);
 const wrapped=Buffer.concat([Buffer.from([1,2]),bytes,Buffer.from([3])]);assert.deepEqual(inspectPng(wrapped.subarray(2,-1)),result);
});

test('PNG inspection checks critical chunk order, metadata, CRC and full prefixes',()=>{
 const good=png(ihdr(),idat(),end());assert.equal(inspectPng(good).chunkCount,3);
 for(let i=0;i<good.length;i++)assert.throws(()=>inspectPng(good.subarray(0,i)),undefined,`prefix ${i}`);
 const corrupted=Buffer.from(good);corrupted[17]^=1;assert.throws(()=>inspectPng(corrupted),/CRC/);
 for(const bytes of [png(idat(),ihdr(),end()),png(ihdr(),ihdr(),idat(),end()),png(ihdr({width:0}),idat(),end()),png(ihdr({depth:1,color:6}),idat(),end()),png(ihdr({interlace:2}),idat(),end()),png(ihdr(),end()),png(ihdr(),idat(),chunk('tEXt'),idat(),end()),png(ihdr(),idat(),chunk('IEND',Buffer.from([0]))),Buffer.concat([good,Buffer.from([0])]),png(ihdr({color:3}),idat(),end()),png(ihdr({color:0}),chunk('PLTE',Buffer.alloc(3)),idat(),end())])assert.throws(()=>inspectPng(bytes));
 const indexed=png(ihdr({color:3,depth:1}),chunk('PLTE',Buffer.alloc(6)),idat(),end());assert.equal(inspectPng(indexed).colorType,3);
 assert.throws(()=>inspectPng(png(ihdr({color:3,depth:1}),chunk('PLTE',Buffer.alloc(9)),idat(),end())),/PLTE/);
});

test('PNG inventory reports uninterpreted extensions and enforces independent limits',()=>{
 const bytes=png(ihdr(),chunk('acTL',Buffer.alloc(8)),chunk('ABCD'),idat(),end()),info=inspectPng(bytes);assert.equal(info.animationChunks,true);assert.deepEqual(info.unknownCriticalChunks,['ABCD']);
 assert.throws(()=>inspectPng(bytes,{maxBytes:bytes.length-1}),/maxBytes/);assert.throws(()=>inspectPng(bytes,{maxChunks:4}),/maxChunks/);assert.throws(()=>inspectPng(bytes,{maxChunks:0}),RangeError);
 const controller=new AbortController();controller.abort(new Error('cancelled'));assert.throws(()=>inspectPng(bytes,{signal:controller.signal}),/cancelled/);
});

test('PNG ICC chunk inventory is presence only and export warns without modifying bytes', async()=>{
 const bytes=await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png',import.meta.url));
 assert.equal(inspectPng(bytes).iccProfileChunks,1); assert.equal(inspectApng(bytes).png.iccProfileChunks,1);
 const artificial=png(ihdr(),chunk('iCCP',Buffer.from([0])),chunk('iCCP'),idat(),end());
 assert.equal(inspectPng(artificial).iccProfileChunks,2); // Payload and multiplicity are deliberately not validated.
 const bad=Buffer.from(artificial),name=bad.indexOf(Buffer.from('iCCP'));bad[name+4]^=1;assert.throws(()=>inspectPng(bad),/CRC/);
 const d=createPresentation();d.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:bytes}});const slide=d.addSlide();
 for(let i=0;i<2;i++)slide.addElement({type:'image',assetId:'icc',box:{x:i*30,y:0,width:20,height:20}});
 const baseline=await generatePptx(d),diagnosed=await generatePptx(d,{pngDiagnostics:'warn'});
 assert.deepEqual(baseline.warnings,[]);assert.deepEqual(diagnosed.warnings.map(w=>[w.code,w.assetId]),[['png-icc-color-unverified','icc']]);
 assert.deepEqual(diagnosed.bytes,baseline.bytes);assert.deepEqual(readGeneratedZip(diagnosed.bytes).get('ppt/media/icc.png'),bytes);
});

test('PNG ICC envelope checks name, placement and size without decompressing the profile',async()=>{
 const real=await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png',import.meta.url));
 const info=inspectPngIcc(real);assert.equal(info.png.iccProfileChunks,1);assert.ok(info.profile.name.length>0);assert.ok(info.profile.compressedBytes>0);
 const bare=inspectPngIcc(png(ihdr(),idat(),end()));assert.equal(bare.profile,undefined);assert.equal(bare.hasSrgb,false);
 const iccp=(name=Buffer.from('Profile'),method=0,content=Buffer.from([1,2]))=>chunk('iCCP',Buffer.concat([name,Buffer.from([0,method]),content]));
 const valid=png(ihdr(),chunk('sRGB',Buffer.from([0])),chunk('cICP',Buffer.alloc(4)),iccp(),idat(),end());
 assert.deepEqual(inspectPngIcc(valid).profile,{name:'Profile',compressedBytes:2});assert.equal(inspectPngIcc(valid).hasSrgb,true);assert.equal(inspectPngIcc(valid).hasCicp,true);
 assert.throws(()=>inspectPngIcc(valid,{maxCompressedProfileBytes:1}),/maxCompressedProfileBytes/);
 for(const name of [Buffer.alloc(0),Buffer.alloc(80,65),Buffer.from(' A'),Buffer.from('A '),Buffer.from('A  B'),Buffer.from([0x80])])assert.throws(()=>inspectPngIcc(png(ihdr(),iccp(name),idat(),end())),/Invalid PNG ICC/);
 for(const invalid of [png(ihdr(),iccp(Buffer.from('A'),1),idat(),end()),png(ihdr(),iccp(Buffer.from('A'),0,Buffer.alloc(0)),idat(),end()),png(ihdr(),chunk('iCCP',Buffer.from('A')),idat(),end()),png(ihdr(),iccp(),iccp(),idat(),end()),png(ihdr(),idat(),iccp(),end()),png(ihdr(),chunk('PLTE',Buffer.alloc(3)),iccp(),idat(),end())])assert.throws(()=>inspectPngIcc(invalid),/Invalid PNG ICC/);
 const latin=inspectPngIcc(png(ihdr(),iccp(Buffer.from([0xc4,32,0xff])),idat(),end()));assert.equal(latin.profile.name,'Ä ÿ');
 const corrupt=Buffer.from(valid);corrupt[corrupt.indexOf(Buffer.from('Profile'))]^=1;assert.throws(()=>inspectPngIcc(corrupt),/CRC/);
 assert.throws(()=>inspectPngIcc(valid,{maxBytes:1}),/maxBytes/);assert.throws(()=>inspectPngIcc(valid,{maxChunks:1}),/maxChunks/);
 for(const value of [0,-1,1.5,NaN,Infinity])assert.throws(()=>inspectPngIcc(valid,{maxCompressedProfileBytes:value}),RangeError);
 const controller=new AbortController();controller.abort(new Error('cancel icc'));assert.throws(()=>inspectPngIcc(valid,{signal:controller.signal}),/cancel icc/);
 if(typeof SharedArrayBuffer!=='undefined')assert.throws(()=>inspectPngIcc(new Uint8Array(new SharedArrayBuffer(valid.length))),/SharedArrayBuffer/);
 class Hostile extends Uint8Array { get byteLength(){return 0} get buffer(){return new ArrayBuffer(0)} *[Symbol.iterator](){throw Error('iterator')} }
 assert.deepEqual(inspectPngIcc(new Hostile(valid)),inspectPngIcc(valid));assert.throws(()=>inspectPngIcc(new Hostile(valid),{maxBytes:1}),/maxBytes/);
});

test('PNG export diagnoses invalid ICC envelopes without changing media or claiming decompression', async () => {
 const iccp=(name=Buffer.from('Profile'),method=0,payload=Buffer.from([1,2]))=>chunk('iCCP',Buffer.concat([name,Buffer.from([0,method]),payload]));
 const cases=[png(ihdr(),iccp(Buffer.from('Bad  Name')),idat(),end()),png(ihdr(),iccp(Buffer.from('Profile'),1),idat(),end()),png(ihdr(),iccp(),iccp(),idat(),end()),png(ihdr(),idat(),iccp(),end())];
 for(const input of cases){
  const doc=createPresentation();doc.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:input}});doc.addSlide({elements:[{type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}},{type:'image',assetId:'icc',box:{x:30,y:0,width:20,height:20}}]});
  const baseline=await generatePptx(doc),checked=await generatePptx(doc,{pngDiagnostics:'warn'});
  assert.deepEqual(baseline.warnings,[]);assert.deepEqual(checked.warnings.map(item=>[item.code,item.assetId]),[['png-icc-color-unverified','icc'],['png-icc-inspection-incomplete','icc']]);
  assert.match(checked.warnings[1].message,/Invalid PNG ICC/);assert.deepEqual(readGeneratedZip(checked.bytes).get('ppt/media/icc.png'),input);
 }
 const unverified=png(ihdr(),iccp(),idat(),end()),doc=createPresentation();doc.registerAsset({id:'icc',kind:'image',source:{type:'bytes',value:unverified}});doc.addSlide({elements:[{type:'image',assetId:'icc',box:{x:0,y:0,width:20,height:20}}]});
 assert.deepEqual((await generatePptx(doc,{pngDiagnostics:'warn'})).warnings.map(item=>item.code),['png-icc-color-unverified']);
});

test('Node PNG ICC extraction verifies zlib and bounds decoded bytes without asserting ICC validity',async()=>{
 const real=await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png',import.meta.url)),jpeg=await readFile(new URL('../../../tests/fixtures/jpeg-icc-srgb.jpg',import.meta.url));
 const start=jpeg.indexOf(Buffer.from('ICC_PROFILE\0'));assert.ok(start>0);const expected=jpeg.subarray(start+14,start+14+588);
 const extracted=extractPngIccProfile(real);assert.deepEqual(Buffer.from(extracted.profileBytes),expected);assert.equal(extracted.inspection.profile.name.length>0,true);
 extracted.profileBytes[0]^=1;assert.deepEqual(Buffer.from(extractPngIccProfile(real).profileBytes),expected);
 assert.equal(extractPngIccProfile(png(ihdr(),idat(),end())).profileBytes,undefined);
 const profile=Buffer.alloc(100000,65),compressed=deflateSync(profile),envelope=data=>png(ihdr(),chunk('iCCP',Buffer.concat([Buffer.from('Name\0\0'),data])),idat(),end());
 assert.deepEqual(Buffer.from(extractPngIccProfile(envelope(compressed),{maxProfileBytes:profile.length}).profileBytes),profile);
 assert.throws(()=>extractPngIccProfile(envelope(compressed),{maxProfileBytes:profile.length-1}),/maxProfileBytes/);
 const checksum=Buffer.from(compressed);checksum[checksum.length-1]^=1;
 for(const data of [checksum,compressed.subarray(0,-1),Buffer.concat([compressed,Buffer.from([0])]),Buffer.from([1,2,3])])assert.throws(()=>extractPngIccProfile(envelope(data)));
 for(const value of [0,-1,NaN,Infinity,1.5])assert.throws(()=>extractPngIccProfile(real,{maxProfileBytes:value}),RangeError);
 const controller=new AbortController(),reason=new Error('abort extraction');controller.abort(reason);assert.throws(()=>extractPngIccProfile(real,{signal:controller.signal}),e=>e===reason);
 if(typeof SharedArrayBuffer!=='undefined')assert.throws(()=>extractPngIccProfile(new Uint8Array(new SharedArrayBuffer(real.length))),/SharedArrayBuffer/);
 let reads=0;const options={get maxCompressedProfileBytes(){reads++;return 1}};
 assert.throws(()=>extractPngIccProfile(real,options),/maxCompressedProfileBytes/);assert.equal(reads,1);
});

test('PNG unknown critical chunk inventory scales and keeps first-seen order',()=>{
 const names=Array.from({length:12000},(_,i)=>`X${String.fromCharCode(65+Math.floor(i/676),65+Math.floor(i/26)%26,65+i%26)}`);
 const bytes=png(ihdr(),...names.map(name=>chunk(name)),chunk(names[0]),idat(),end()),info=inspectPng(bytes);
 assert.deepEqual(info.unknownCriticalChunks,names);assert.equal(info.chunkCount,12004);assert.throws(()=>inspectPng(bytes,{maxChunks:12003}),/maxChunks/);
});

function animationControl(frames=2,plays=0){const data=Buffer.alloc(8);data.writeUInt32BE(frames);data.writeUInt32BE(plays,4);return chunk('acTL',data)}
function frame(seq,{width=4,height=2,x=0,y=0,disposal=0,blend=0,denominator=0}={}){const data=Buffer.alloc(26);data.writeUInt32BE(seq);data.writeUInt32BE(width,4);data.writeUInt32BE(height,8);data.writeUInt32BE(x,12);data.writeUInt32BE(y,16);data.writeUInt16BE(10,20);data.writeUInt16BE(denominator,22);data[24]=disposal;data[25]=blend;return chunk('fcTL',data)}
function frameData(seq){const data=Buffer.alloc(5);data.writeUInt32BE(seq);return chunk('fdAT',data)}

test('APNG inspection inventories actual Pillow frames without decoding pixels',async()=>{
 const bytes=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url)),result=inspectApng(bytes);assert.equal(result.png.width,3);assert.equal(result.png.height,2);assert.equal(result.animation.plays,2);assert.equal(result.animation.defaultImageIsFrame,true);assert.equal(result.animation.frames.length,2);assert.deepEqual(result.animation.frames.map(f=>f.source),['IDAT','fdAT']);assert.deepEqual(result.animation.frames.map(f=>[f.delayNumerator,f.delayDenominator]),[[1,10],[1,4]]);
 const staticResult=inspectApng(png(ihdr(),idat(),end()));assert.equal(staticResult.animation,undefined);
});

test('APNG inspection checks sequence, frame bounds, data presence and both default-image forms',()=>{
 const first=png(ihdr(),animationControl(),frame(0),idat(),frame(1,{width:2,height:1,x:1}),frameData(2),end());const a=inspectApng(first).animation;assert.equal(a.defaultImageIsFrame,true);assert.deepEqual(a.frames.map(x=>x.source),['IDAT','fdAT']);assert.equal(a.frames[0].delayDenominator,0);
 const poster=png(ihdr(),animationControl(1,3),idat(),frame(0,{width:2,height:1}),frameData(1),end());assert.equal(inspectApng(poster).animation.defaultImageIsFrame,false);
 for(const bytes of [png(ihdr(),animationControl(),frame(1),idat(),frame(2),frameData(3),end()),png(ihdr(),animationControl(),frame(0),idat(),frame(1),frameData(3),end()),png(ihdr(),animationControl(),frame(0),idat(),frame(1,{width:5}),frameData(2),end()),png(ihdr(),animationControl(),frame(0),idat(),frame(1,{blend:2}),frameData(2),end()),png(ihdr(),animationControl(),frame(0),idat(),frame(1),end()),png(ihdr(),animationControl(1),frame(0,{width:2}),idat(),end()),png(ihdr(),idat(),frame(0),frameData(1),end()),png(ihdr(),animationControl(3),frame(0),idat(),frame(1),frameData(2),end())])assert.throws(()=>inspectApng(bytes));
 for(const bytes of [png(ihdr(),idat(),animationControl(1),frame(0),frameData(1),end()),png(ihdr(),animationControl(),animationControl(),frame(0),idat(),frame(1),frameData(2),end()),png(ihdr(),animationControl(1),idat(),frameData(0),end()),png(ihdr(),animationControl(1),frame(0,{disposal:3}),idat(),end()),png(ihdr(),animationControl(0),idat(),end())])assert.throws(()=>inspectApng(bytes));
 assert.throws(()=>inspectApng(first,{maxFrames:1}),/maxFrames/);assert.throws(()=>inspectApng(first,{maxFrames:0}),RangeError);
});

test('PNG and APNG inspect one bounded snapshot and reject shared memory',()=>{
 const make=()=>png(ihdr(),animationControl(1),frame(0),idat(),end());
 for(const inspect of [inspectPng,inspectApng]){
  const bytes=make(),baseline=inspect(bytes);let reads=0;const signal={get aborted(){if(++reads===2)bytes[17]^=1;return false}};
  assert.deepEqual(inspect(bytes,{signal}),baseline);assert.notDeepEqual(bytes,make());
  if(typeof SharedArrayBuffer!=='undefined'){const shared=new Uint8Array(new SharedArrayBuffer(bytes.length));shared.set(make());assert.throws(()=>inspect(shared),/SharedArrayBuffer/);}
 }
 class HostileBytes extends Uint8Array { *[Symbol.iterator](){throw Error('unexpected iterator')} get byteLength(){return 0} get buffer(){return new ArrayBuffer(0)} }
 for(const inspect of [inspectPng,inspectApng]){const hostile=new HostileBytes(make());assert.deepEqual(inspect(hostile),inspect(make()));assert.throws(()=>inspect(hostile,{maxBytes:1}),/maxBytes/);if(typeof SharedArrayBuffer!=='undefined'){const shared=new HostileBytes(new SharedArrayBuffer(make().length));shared.set(make());assert.throws(()=>inspect(shared),/SharedArrayBuffer/);}}
});

test('scanline export validation checks zlib termination, dimensions and filters behind valid chunk CRCs', async () => {
  const strict = { pngDiagnostics: 'reject-scanlines' };
  const make = (compressed, options = { width: 1, height: 1, color: 0 }) =>
    png(ihdr(options), chunk('IDAT', compressed.subarray(0, 3)), chunk('IDAT', compressed.subarray(3)), end());
  async function exportPng(data, options = strict) {
    const deck = createPresentation();
    deck.registerAsset({ id: 'sample', kind: 'image', source: { type: 'bytes', value: data } });
    deck.addSlide({ elements: [{ type: 'image', assetId: 'sample', box: { x: 0, y: 0, width: 40, height: 40 } }] });
    const result = await generatePptx(deck, options);
    return { media: readGeneratedZip(result.bytes).get('ppt/media/sample.png'), warnings: result.warnings };
  }
  const compressed = deflateSync(Buffer.from([0, 75]));
  const checksum = Buffer.from(compressed); checksum[checksum.length - 1] ^= 1;
  const valid = make(compressed);
  assert.deepEqual(await exportPng(valid), { media: valid, warnings: [] });
  const adam7 = make(deflateSync(Buffer.from([0, 1, 0, 2, 0, 3, 0, 4, 5, 6])), { width: 3, height: 2, color: 0, interlace: 1 });
  assert.deepEqual(await exportPng(adam7), { media: adam7, warnings: [] });
  const palette = pixels => png(ihdr({ width: 2, height: 1, color: 3, depth: 2 }),
    chunk('PLTE', Buffer.from([0, 0, 0, 255, 255, 255])), chunk('IDAT', deflateSync(Buffer.from([0, pixels]))), end());
  const indexed = palette(0b00011111); // padding bits are not palette indices
  assert.deepEqual(await exportPng(indexed), { media: indexed, warnings: [] });
  for (const [data, reason] of [
    [make(checksum), /image data validation failed/],
    [make(compressed.subarray(0, -1)), /image data validation failed/],
    [make(Buffer.concat([compressed, Buffer.from([0])])), /image data validation failed/],
    [make(deflateSync(Buffer.from([5, 75]))), /row filter/],
    [make(deflateSync(Buffer.from([0]))), /shorter/],
    [make(deflateSync(Buffer.alloc(8 * 1024 * 1024))), /exceeds the declared dimensions/],
    [make(compressed, { width: 0x7fffffff, height: 0x7fffffff, color: 6 }), /128 MiB/],
    [palette(0b00100000), /palette index/],
  ]) {
    assert.deepEqual((await exportPng(data, { pngDiagnostics: 'reject-structure' })).media, data);
    const result = await exportPng(data);
    assert.equal(result.media, undefined);
    assert.deepEqual(result.warnings.map(w => w.code), ['asset-read-failed', 'image-omitted']);
    assert.match(result.warnings[0].message, reason);
  }

  // The default image is outside the animation; the single fdAT frame
  // has its own smaller geometry and splits the zlib stream across chunks.
  const frameCompressed = deflateSync(Buffer.from([0, 1, 2]));
  const frameChunk = (seq, data) => {
    const number = Buffer.alloc(4); number.writeUInt32BE(seq);
    return chunk('fdAT', Buffer.concat([number, data]));
  };
  const excludedDefault = png(ihdr({ color: 0 }), animationControl(1), chunk('IDAT', deflateSync(Buffer.alloc(10))),
    frame(0, { width: 2, height: 1, x: 1, y: 1 }), frameChunk(1, frameCompressed.subarray(0, 3)),
    frameChunk(2, frameCompressed.subarray(3)), end());
  const checkedExcluded = await exportPng(excludedDefault);
  assert.deepEqual(checkedExcluded.media, excludedDefault);
  assert.deepEqual(checkedExcluded.warnings.map(w => w.code), ['apng-animation-playback-unverified']);

  const animation = await readFile(new URL('../../../tests/fixtures/two-frame-apng.png', import.meta.url));
  const validAnimation = await exportPng(animation);
  assert.deepEqual(validAnimation.media, animation);
  assert.deepEqual(validAnimation.warnings.map(w => w.code), ['apng-animation-playback-unverified']);
  const fdAt = animation.indexOf('fdAT') - 4;
  assert.ok(fdAt > 0);
  const oldLength = animation.readUInt32BE(fdAt);
  const wrongFrame = Buffer.concat([animation.subarray(0, fdAt),
    chunk('fdAT', Buffer.concat([animation.subarray(fdAt + 8, fdAt + 12), deflateSync(Buffer.from([0]))])),
    animation.subarray(fdAt + 12 + oldLength)]);
  assert.deepEqual((await exportPng(wrongFrame, { pngDiagnostics: 'reject-structure' })).media, wrongFrame);
  assert.match((await exportPng(wrongFrame)).warnings[0].message, /shorter/);
});
