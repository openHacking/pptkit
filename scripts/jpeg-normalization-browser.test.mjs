import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';

test('browser JPEG normalization applies all eight EXIF orientations with bounded work', async t => {
  const source = await readFile('tests/fixtures/exif-orientation6.jpg');
  const exifStart = source.indexOf(Buffer.from('Exif\0\0')); assert.ok(exifStart > 4);
  const segmentStart = exifStart - 4, length = source.readUInt16BE(segmentStart + 2);
  const noExif = Buffer.concat([source.subarray(0, segmentStart), source.subarray(segmentStart + 2 + length)]);
  const bundle = await readFile('packages/pptx-exporter/dist/browser/dom.js');
  const adapter = 'data:text/javascript;base64,' + bundle.toString('base64');
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(); const requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const result = await page.evaluate(async ({jpeg,raw,adapter,exifStart}) => {
    const api = await import(adapter);
    async function pixels(blob) {
      const bitmap = await createImageBitmap(blob), canvas = document.createElement('canvas');
      canvas.width = bitmap.width; canvas.height = bitmap.height;
      const ctx = canvas.getContext('2d'); ctx.drawImage(bitmap,0,0); bitmap.close();
      return {width:canvas.width,height:canvas.height,pixels:[...ctx.getImageData(0,0,canvas.width,canvas.height).data]};
    }
    const base = await pixels(new Blob([new Uint8Array(raw)],{type:'image/jpeg'}));
    const rows=[];
    for (let orientation=1;orientation<=8;orientation++) {
      const bytes=new Uint8Array(jpeg); bytes[exifStart+6+18]=0;bytes[exifStart+6+19]=orientation;
      const result=await api.normalizeJpegOrientationInBrowser(bytes,{maxPixels:6});
      const decoded=await pixels(new Blob([result.bytes],{type:'image/png'}));
      const crop=api.orientJpegCrop({left:0,top:0,right:2/3,bottom:1/2},orientation);
      rows.push({orientation,width:result.width,height:result.height,mime:result.mimeType,normalization:result.orientationNormalization,decoded,crop});
    }
    const failures=[];
    for (const opts of [{maxPixels:5},{maxInputBytes:1},{maxOutputBytes:1},{maxPixels:0}]) {
      try {await api.normalizeJpegOrientationInBrowser(new Uint8Array(jpeg),opts);failures.push('accepted');}
      catch(error){failures.push(error.name+':'+error.message)}
    }
    try {await api.normalizeJpegOrientationInBrowser(new Uint8Array(raw));failures.push('accepted');}
    catch(error){failures.push(error.name+':'+error.message)}
    const aborted=new AbortController();aborted.abort('cancelled');
    try{await api.normalizeJpegOrientationInBrowser(new Uint8Array(jpeg),{signal:aborted.signal});failures.push('accepted');}catch(e){failures.push(e)}
    class Hostile extends Uint8Array{get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}
    try{await api.normalizeJpegOrientationInBrowser(new Hostile(jpeg),{maxInputBytes:1});failures.push('accepted')}catch(error){failures.push(error.name+':'+error.message)}
    let shared='unavailable';if(typeof SharedArrayBuffer!=='undefined'){const value=new Hostile(new SharedArrayBuffer(jpeg.length));value.set(jpeg);try{await api.normalizeJpegOrientationInBrowser(value);shared='accepted'}catch(e){shared=e.message}}
    // Source mutation after the async call must not change its captured bytes.
    const mutable=new Uint8Array(jpeg), pending=api.normalizeJpegOrientationInBrowser(mutable);mutable.fill(0);
    const detached=await pending;
    // An uncooperative asynchronous decoder must be released if it resolves after cancellation.
    const original=globalThis.createImageBitmap;let finish,closed=false;
    globalThis.createImageBitmap=()=>new Promise(resolve=>{finish=resolve});
    const controller=new AbortController(), late=api.normalizeJpegOrientationInBrowser(new Uint8Array(jpeg),{signal:controller.signal});
    controller.abort('late-cancel');let lateReason;try{await late}catch(e){lateReason=e}
    finish({close(){closed=true}});await Promise.resolve();await Promise.resolve();globalThis.createImageBitmap=original;
    return {base,rows,failures,shared,detachedOrientation:detached.sourceOrientation,lateReason,closed};
  }, {jpeg:[...source],raw:[...noExif],adapter,exifStart});
  const maps=[[0,1,2,3,4,5],[2,1,0,5,4,3],[5,4,3,2,1,0],[3,4,5,0,1,2],
    [0,3,1,4,2,5],[3,0,4,1,5,2],[5,2,4,1,3,0],[2,5,1,4,0,3]];
  assert.equal(result.base.width,3);assert.equal(result.base.height,2);
  for (const [index,row] of result.rows.entries()) {
    const expected=maps[index].flatMap(i=>result.base.pixels.slice(i*4,i*4+4));
    assert.equal(row.width,index<4?3:2);assert.equal(row.height,index<4?2:3);
    assert.equal(row.decoded.width,row.width);assert.equal(row.decoded.height,row.height);
    assert.equal(row.mime,'image/png');assert.equal(row.normalization,'pixels');
    assert.deepEqual(row.decoded.pixels,expected,`orientation ${index+1}`);
    // The mapped crop selects the original top-left raw pixel after transformation.
    const x = Math.round(row.crop.left * row.width), y = Math.round(row.crop.top * row.height);
    assert.equal(Math.round((1-row.crop.right)*row.width)-x,1);
    assert.equal(Math.round((1-row.crop.bottom)*row.height)-y,1);
    assert.deepEqual(row.decoded.pixels.slice((y*row.width+x)*4,(y*row.width+x+1)*4),result.base.pixels.slice(0,4));
  }
  assert.match(result.failures[0],/maxPixels/);assert.match(result.failures[1],/maxInputBytes/);
  assert.match(result.failures[2],/maxOutputBytes/);assert.match(result.failures[3],/maxPixels/);
  assert.match(result.failures[4],/orientation/);assert.equal(result.failures[5],'cancelled');
  assert.match(result.failures[6],/maxInputBytes/);if(result.shared!=='unavailable')assert.match(result.shared,/SharedArrayBuffer/);
  assert.equal(result.detachedOrientation,6);assert.equal(result.lateReason,'late-cancel');assert.equal(result.closed,true);
  assert.deepEqual(requests,[]);
});
