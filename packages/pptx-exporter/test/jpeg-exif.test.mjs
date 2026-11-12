import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { inspectJpegExifOrientation as inspect } from '../dist/index.js';
import { inspectJpegExifOrientation as inspectNode } from '../dist/node.js';

const make = (value = 6, little = true, entries = 1) => {
  const tiff = new Uint8Array(8 + 2 + 12 * entries + 4), v = new DataView(tiff.buffer);
  tiff.set(little ? [0x49,0x49] : [0x4d,0x4d]); v.setUint16(2,42,little); v.setUint32(4,8,little); v.setUint16(8,entries,little);
  for(let i=0;i<entries;i++){ const o=10+i*12;v.setUint16(o,0x0112,little);v.setUint16(o+2,3,little);v.setUint32(o+4,1,little);v.setUint16(o+8,value,little); }
  const payload=Uint8Array.from([69,120,105,102,0,0,...tiff]), len=payload.length+2;
  return Uint8Array.from([255,216,255,225,len>>8,len&255,...payload,255,217]);
};

test('reads all eight orientations in both TIFF byte orders and offset views', () => {
  for(const little of [true,false]) for(let orientation=1;orientation<=8;orientation++) {
    const source=make(orientation,little), wrapper=new Uint8Array(source.length+4);wrapper.set(source,2);
    const result={status:'orientation',orientation,stoppedAt:'end'};
    assert.deepEqual(inspect(source),result);assert.deepEqual(inspectNode(wrapper.subarray(2,-2)),result);
  }
});

test('reads independently encoded asymmetric JPEG without mutating the input', async () => {
  // Pillow 12.1.1 generated a 3x2 JPEG with orientation 6; Pillow independently reads tag 274 as 6.
  const source=await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg',import.meta.url)), before=Buffer.from(source);
  assert.deepEqual(inspect(source),{status:'orientation',orientation:6,stoppedAt:'scan'});assert.deepEqual(source,before);
});

test('reports absence conservatively and stops before unexamined image data', () => {
  assert.deepEqual(inspect(Uint8Array.of(255,216,255,217)),{status:'no-exif',stoppedAt:'end'});
  const noTag=make(1,true,0);assert.deepEqual(inspect(noTag),{status:'no-orientation',stoppedAt:'end'});
  assert.deepEqual(inspect(Uint8Array.of(255,216,255,218,0,2,3,4)),{status:'no-exif',stoppedAt:'scan'});
  const xmp=Uint8Array.of(255,216,255,225,0,4,65,66,255,217);assert.deepEqual(inspect(xmp),{status:'no-exif',stoppedAt:'end'});
});

test('rejects corrupt offsets, duplicate or invalid orientation and malformed marker structure', () => {
  const valid=make(), broken=(index,value)=>{const b=valid.slice();b[index]=value;return b};
  for(const input of [make(0),make(9),make(2,true,2),broken(14,0),broken(20,255),broken(24,0),broken(28,2),
    Uint8Array.of(255,216,255,225,0,1),Uint8Array.of(255,216,255,0),Uint8Array.of(255,216,255,225,0,6,69,120,105,102,255,217),Uint8Array.of(255,216,255,225,0,8,69,120,105,102,0,0,255,217),
    Uint8Array.from([...valid.slice(0,-2),...valid.slice(2)])]) assert.throws(()=>inspect(input));
  for(let i=0;i<valid.length;i++) assert.throws(()=>inspect(valid.subarray(0,i)),undefined,`prefix ${i}`);
});

test('enforces limits and pre-abort without decoding the image', () => {
  const input=make();assert.deepEqual(inspect(input,{maxBytes:input.length,maxSegments:2,maxIfdEntries:1}),inspect(input));
  for(const opts of [{maxBytes:input.length-1},{maxSegments:1},{maxIfdEntries:1}]) {
    const data=opts.maxIfdEntries ? make(1,true,2) : input; assert.throws(()=>inspect(data,opts),RangeError);
  }
  for(const key of ['maxBytes','maxSegments','maxIfdEntries']) for(const value of [0,-1,1.5,Infinity,null]) assert.throws(()=>inspect(input,{[key]:value}),RangeError);
  const controller=new AbortController(),reason=new Error('abort');controller.abort(reason);assert.throws(()=>inspect(input,{signal:controller.signal}),e=>e===reason);
});
