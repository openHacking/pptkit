import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";
import { pixelPng } from "../../../tests/fixtures/presentations.mjs";
import { readFile } from 'node:fs/promises';
import { probeJpegPrecision } from '../dist/rendering/jpeg-precision-probe.js';
function deck(mimeType, value = new Uint8Array(Buffer.from(pixelPng.split(",")[1], "base64"))) {
  const p = createPresentation();
  p.registerAsset({ id: "bytes", kind: "image", source: { type: "bytes", value }, ...(mimeType ? { mimeType } : {}) });
  p.addSlide().addElement({ type: "image", assetId: "bytes", box: { x: 0, y: 0, width: 80, height: 40 } }); return p;
}
test("byte preview embeds canonical PNG without object URLs or loading", async () => {
  const result = await renderPresentationToSvg(deck());
  assert.deepEqual(result.warnings, []); assert.ok(result.slides[0].svg.includes(pixelPng));
});
test("byte preview reports unsupported/mismatched content and still permits explicit resolver override", async () => {
  for (const p of [deck("image/jpeg"), deck(undefined, Uint8Array.of(1))]) {
    const result = await renderPresentationToSvg(p);
    assert.equal(result.warnings[0].code, "asset-bytes-unsupported"); assert.equal(result.warnings[0].assetId, "bytes");
    assert.ok(!result.slides[0].svg.includes("<image"));
    const override = await renderPresentationToSvg(p, { resolveAsset: () => pixelPng });
    assert.deepEqual(override.warnings, []); assert.ok(override.slides[0].svg.includes(pixelPng));
  }
});

test('non-8-bit JPEG preview warns without changing embedded data or resolver behavior', async () => {
  const eight = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const images = await Promise.all(['jpeg-12bit-sequential.jpg', 'jpeg-12bit-progressive.jpg'].map(name => readFile(new URL(`../../../tests/fixtures/${name}`, import.meta.url))));
  assert.deepEqual((await renderPresentationToSvg(deck('image/jpeg', eight))).warnings, []);
  for (const bytes of images) {
    const presentation = deck('image/jpeg', bytes), result = await renderPresentationToSvg(presentation);
    assert.deepEqual(result.warnings.map(w => [w.code, w.assetId]), [['jpeg-preview-precision-unverified', 'bytes']]);
    assert.equal(result.status, 'rendered-with-warnings'); assert.ok(result.slides[0].svg.includes(bytes.toString('base64')));
    assert.deepEqual((await renderPresentationToSvg(presentation, { resolveAsset: () => pixelPng })).warnings, []);
  }
});

test('JPEG precision probe respects segment lengths and its prefix boundary', () => {
  const marker = Buffer.from([0xff, 0xc0, 0, 8, 12, 0, 1, 0, 1, 1]);
  const jpeg = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 12]), marker, marker]);
  assert.equal(probeJpegPrecision(jpeg.toString('base64')), 12, 'frame-like bytes inside APP payload are not scanned');
  const fake = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 12]), marker]);
  assert.equal(probeJpegPrecision(fake.toString('base64')), undefined);
  const oversized = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0xff, 0xff]), Buffer.alloc(65533), marker]);
  assert.equal(probeJpegPrecision(oversized.toString('base64')), 'incomplete');
  for (const bad of [Buffer.from([0xff,0xd8,0xff,0xda]), Buffer.from([0xff,0xd8,0xff,0xe0,0,1]), Buffer.from([0xff,0xd8,0xff,0xc0,0,9,12])]) assert.equal(probeJpegPrecision(bad.toString('base64')), undefined);
});

test('long JPEG application segment reports a bounded incomplete precision scan', async () => {
  const original = await readFile(new URL('../../../tests/fixtures/jpeg-12bit-progressive.jpg', import.meta.url));
  const longHeader = Buffer.concat([original.subarray(0, 2), Buffer.from([0xff, 0xe1, 0xff, 0xff]), Buffer.alloc(65533), original.subarray(2)]);
  const d = deck('image/jpeg', longHeader), result = await renderPresentationToSvg(d);
  assert.deepEqual(result.warnings.map(w => w.code), ['jpeg-preview-inspection-incomplete']);
  assert.ok(result.slides[0].svg.includes(longHeader.toString('base64')));
  assert.deepEqual((await renderPresentationToSvg(d, { resolveAsset: () => pixelPng })).warnings, []);
  const truncated = Buffer.from([0xff, 0xd8, 0xff, 0xe1, 0xff, 0xff]);
  assert.equal(probeJpegPrecision(truncated.toString('base64')), undefined, 'a short malformed input is not a prefix-budget exhaustion');
});

test('JPEG precision warnings identify each slide and deduplicate repeated uses there', async () => {
  const value = await readFile(new URL('../../../tests/fixtures/jpeg-12bit-progressive.jpg', import.meta.url));
  const d = createPresentation(); d.registerAsset({ id: 'shared', kind: 'image', source: { type: 'bytes', value } });
  for (const id of ['s1', 's2']) {
    const slide = d.addSlide({ id });
    for (let i = 0; i < 2; i++) slide.addElement({ id: `${id}-${i}`, type: 'image', assetId: 'shared', box: { x: i * 20, y: 0, width: 10, height: 10 } });
  }
  const result = await renderPresentationToSvg(d);
  assert.deepEqual(result.warnings.map(w => [w.code, w.slideId, w.assetId]), [
    ['jpeg-preview-precision-unverified', 's1', 'shared'], ['jpeg-preview-precision-unverified', 's2', 'shared'],
  ]);
});

test("byte preview accepts the APNG MIME alias only for PNG signatures", async () => {
  const {readFile}=await import('node:fs/promises'),apng=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url));
  const result=await renderPresentationToSvg(deck(' IMAGE/APNG ',apng));assert.deepEqual(result.warnings.map(w=>w.code),['apng-preview-playback-unverified']);assert.ok(result.slides[0].svg.includes(`data:image/png;base64,${apng.toString('base64')}`));
  const mismatch=await renderPresentationToSvg(deck('image/apng',Uint8Array.of(255,216,255)));assert.ok(mismatch.warnings.some(w=>w.code==='asset-bytes-unsupported'));
});

test("APNG declaration warning is scoped per slide and asset, including resolver output", async () => {
  const d=createPresentation();d.registerAsset({id:'a',kind:'image',mimeType:'image/apng',source:{type:'path',value:'virtual'}});
  for(const id of ['s1','s2']){const slide=d.addSlide({id});for(let i=0;i<2;i++)slide.addElement({id:`${id}-${i}`,type:'image',assetId:'a',box:{x:i*20,y:0,width:10,height:10}});}
  const result=await renderPresentationToSvg(d,{resolveAsset:()=>pixelPng}),warnings=result.warnings.filter(w=>w.code==='apng-preview-playback-unverified');assert.deepEqual(warnings.map(w=>[w.slideId,w.assetId]),[['s1','a'],['s2','a']]);assert.equal(result.status,'rendered-with-warnings');
});

test("opt-in PNG prefix scan reports animation declarations without changing SVG content", async () => {
  const {readFile}=await import('node:fs/promises'),apng=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url)),d=deck('image/png',apng);
  const plain=await renderPresentationToSvg(d),checked=await renderPresentationToSvg(d,{pngAnimationDiagnostics:'warn'});assert.deepEqual(plain.warnings,[]);assert.deepEqual(checked.slides,plain.slides);assert.deepEqual(checked.warnings.map(w=>w.code),['apng-preview-playback-unverified']);
  const staticPng=await readFile(new URL('../../../tests/fixtures/transparent-replacement.png',import.meta.url));assert.deepEqual((await renderPresentationToSvg(deck('image/png',staticPng),{pngAnimationDiagnostics:'warn'})).warnings,[]);
  const truncated=apng.subarray(0,16),result=await renderPresentationToSvg(deck('image/png',truncated),{pngAnimationDiagnostics:'warn'});assert.deepEqual(result.warnings.map(w=>w.code),['png-preview-inspection-incomplete']);
  const oversized=new Uint8Array(8*1024*1024+1);oversized.set([137,80,78,71,13,10,26,10]);const capped=await renderPresentationToSvg(deck('image/png',oversized),{pngAnimationDiagnostics:'warn'});assert.deepEqual(capped.warnings.map(w=>w.code),['png-preview-inspection-incomplete']);
  const chunk=Buffer.from('000000007445587400000000','hex'),many=Buffer.concat([staticPng.subarray(0,33),...Array(100000).fill(chunk)]);const limited=await renderPresentationToSvg(deck('image/png',many),{pngAnimationDiagnostics:'warn'});assert.deepEqual(limited.warnings.map(w=>w.code),['png-preview-inspection-incomplete']);
  const badCrc=Buffer.from(staticPng);badCrc[29]^=1;assert.deepEqual((await renderPresentationToSvg(deck('image/png',badCrc),{pngAnimationDiagnostics:'warn'})).warnings,[]);
  await assert.rejects(renderPresentationToSvg(d,{pngAnimationDiagnostics:null}),/pngAnimationDiagnostics/);
});

test("PNG prefix scan reports resolver and non-byte input as uninspected", async () => {
  const d=createPresentation();d.registerAsset({id:'remote',kind:'image',mimeType:'image/png',source:{type:'url',value:'https://example.invalid/image.png'}});d.addSlide({id:'s',elements:[{id:'a',type:'image',assetId:'remote',box:{x:0,y:0,width:10,height:10}},{id:'b',type:'image',assetId:'remote',box:{x:10,y:0,width:10,height:10}}]});
  for(const options of [{pngAnimationDiagnostics:'warn'},{pngAnimationDiagnostics:'warn',resolveAsset:()=>pixelPng}]){const result=await renderPresentationToSvg(d,options);assert.deepEqual(result.warnings.map(w=>w.code),['png-preview-inspection-incomplete']);}
});


test("asset replacement updates preview and contain geometry while previous output stays detached", async () => {
  const { resolveLayout } = await import("@pptkit/layout");
  const p = createPresentation();
  const first = p.registerAsset({ id: "shared", kind: "image", source: { type: "bytes", value: new Uint8Array(Buffer.from(pixelPng.split(",")[1], "base64")) }, width: 8, height: 4 });
  p.addSlide().addElement({ type: "image", assetId: "shared", fit: "contain", box: { x: 0, y: 0, width: 100, height: 100 } });
  const before = await renderPresentationToSvg(p), geometry = resolveLayout(p);
  const gif = "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  p.replaceAsset("shared", { kind: "image", source: { type: "bytes", encoding: "base64", value: gif }, width: 1, height: 1 });
  const after = await renderPresentationToSvg(p), changed = resolveLayout(p);
  assert.ok(before.slides[0].svg.includes(pixelPng)); assert.ok(after.slides[0].svg.includes(`data:image/gif;base64,${gif}`));
  assert.deepEqual(geometry.slides[0].elements[0].box, { x: 0, y: 25, width: 100, height: 50 });
  assert.deepEqual(changed.slides[0].elements[0].box, { x: 0, y: 0, width: 100, height: 100 });
  assert.equal(first.width, 8); assert.equal(p.slides[0].elements[0].box.height, 100); assert.deepEqual(after.warnings, []);
});
