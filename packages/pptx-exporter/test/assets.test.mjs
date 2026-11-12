import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { generatePptx as generateNode } from "../dist/node.js";
import { pixelPng } from "../../../tests/fixtures/presentations.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const png = () => new Uint8Array(Buffer.from(pixelPng.split(",")[1], "base64"));
function deck(source = { type: "bytes", value: png() }, mimeType) {
  const p = createPresentation();
  p.registerAsset({ id: "picture", kind: "image", source, ...(mimeType ? { mimeType } : {}) });
  const slide = p.addSlide();
  for (let i = 0; i < 2; i++) slide.addElement({ type: "image", assetId: "picture", box: { x: i * 50, y: 0, width: 40, height: 40 } });
  return p;
}
const media = bytes => [...readGeneratedZip(bytes)].filter(([name]) => name.startsWith("ppt/media/"));

test("byte images are exact, share one packaged asset and match Node output without fetch", async () => {
  const original = globalThis.fetch; globalThis.fetch = () => { throw new Error("unexpected fetch"); };
  try {
    const p = deck(), a = await generatePptx(p), b = await generateNode(p);
    assert.deepEqual(a.warnings, []); assert.deepEqual(a.bytes, b.bytes);
    assert.equal(media(a.bytes).length, 1); assert.equal(media(a.bytes)[0][0], "ppt/media/picture.png");
    assert.deepEqual(media(a.bytes)[0][1], Buffer.from(png()));
    assert.match(readGeneratedZip(a.bytes).get("[Content_Types].xml").toString(), /ContentType="image\/png"/);
  } finally { globalThis.fetch = original; }
});

test("byte MIME mismatches and unsupported signatures are located warnings, never mislabeled output", async () => {
  for (const p of [deck(undefined, "image/jpeg"), deck({ type: "bytes", value: Uint8Array.of(1, 2, 3) })]) {
    const result = await generatePptx(p);
    assert.equal(result.status, "generated-with-warnings"); assert.equal(media(result.bytes).length, 0);
    assert.equal(result.warnings[0].code, "asset-read-failed"); assert.equal(result.warnings[0].assetId, "picture");
    assert.equal(result.warnings.filter(w => w.code === "image-omitted").length, 2);
  }
});

test("vector images require explicit rasterization and report the loss once per asset", async () => {
  for (const format of ["emf", "wmf", "svg"]) {
    const source = format === "svg" ? new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>') : Uint8Array.of(1, 2, 3, 4);
    const p = deck({ type: "bytes", value: source }, format === "svg" ? "image/svg+xml" : `image/${format}`);
    for (const generate of [generatePptx, generateNode]) {
      const missing = await generate(p);
      assert.equal(media(missing.bytes).length, 0);
      assert.equal(missing.warnings[0].code, "asset-read-failed");
      const converted = await generate(p, { resolveAsset: () => ({ bytes: png(), mimeType: "image/png", vectorRasterization: format }) });
      assert.deepEqual(media(converted.bytes)[0][1], Buffer.from(png()));
      assert.deepEqual(converted.warnings.map(w => [w.code, w.assetId]), [["asset-vector-rasterized", "picture"]]);
      const capped = await generate(p, { maxAssetBytes: 1, resolveAsset: () => ({ bytes: png(), vectorRasterization: format }) });
      assert.equal(media(capped.bytes).length, 0);
      assert.match(capped.warnings[0].message, /maxAssetBytes/);
      for (const resolved of [
        { bytes: png(), vectorRasterization: format === "emf" ? "wmf" : "emf" },
        { bytes: png(), vectorRasterization: "unknown" },
        { bytes: png(), vectorRasterization: format, animationFallback: "first-frame" },
        { bytes: new TextEncoder().encode("GIF89a"), vectorRasterization: format },
        { bytes: Uint8Array.of(1, 2, 3), vectorRasterization: format },
        { bytes: png(), mimeType: "image/jpeg", vectorRasterization: format },
        { bytes: png(), mimeType: "image/png" },
      ]) {
        const rejected = await generate(p, { resolveAsset: () => resolved });
        assert.equal(media(rejected.bytes).length, 0);
        assert.equal(rejected.warnings[0].code, "asset-read-failed");
      }
    }
    const undeclared = await generatePptx(deck({ type: "path", value: "virtual" }), { resolveAsset: () => ({ bytes: png(), vectorRasterization: format }) });
    assert.equal(undeclared.warnings[0].code, "asset-read-failed");
    assert.equal(media(undeclared.bytes).length, 0);
  }
});

test("vector rasterization keeps authored crop, transform and frame like the same PNG asset", async () => {
  const makeDeck = (source, mimeType) => {
    const p = createPresentation();
    p.registerAsset({ id: "picture", kind: "image", source, mimeType });
    p.addSlide().addElement({ id: "image", type: "image", assetId: "picture", box: { x: 12, y: 23, width: 80, height: 50 }, fit: "crop", crop: { left: .1, right: .2, top: .15, bottom: .05 }, transform: { rotation: 30, flipH: true } });
    return p;
  };
  const source = makeDeck({ type: "path", value: "virtual" }, "image/emf");
  const control = makeDeck({ type: "bytes", value: png() }, "image/png");
  const actual = await generatePptx(source, { resolveAsset: () => ({ bytes: png(), vectorRasterization: "emf" }) });
  const expected = await generatePptx(control);
  const xml = readGeneratedZip(actual.bytes).get("ppt/slides/slide1.xml").toString();
  assert.equal(xml, readGeneratedZip(expected.bytes).get("ppt/slides/slide1.xml").toString());
  assert.match(xml, /<a:srcRect l="10000" t="15000" r="20000" b="5000"\/>/);
  assert.match(xml, /rot="1800000" flipH="1"/);
});

test("SVG pair shares media, creates both owner relationships and bounds the combined bytes", async () => {
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><rect width="8" height="4" fill="red"/></svg>');
  const p = deck({ type: "bytes", value: svg }, "image/svg+xml");
  p.addSlide().addElement({ type: "image", assetId: "picture", box: { x: 0, y: 0, width: 30, height: 20 } });
  for (const generate of [generatePptx, generateNode]) {
    const result = await generate(p, { maxAssetBytes: png().length + svg.length, resolveAsset: () => ({ bytes: png(), mimeType: "image/png", svgPair: svg }) });
    assert.deepEqual(result.warnings.map(w => w.code), ["asset-svg-pair-unverified"]);
    const parts = readGeneratedZip(result.bytes);
    assert.deepEqual(parts.get("ppt/media/picture.png"), Buffer.from(png()));
    assert.deepEqual(parts.get("ppt/media/pptkit-svg-pairs/1.svg"), Buffer.from(svg));
    assert.match(parts.get("[Content_Types].xml").toString(), /Extension="svg" ContentType="image\/svg\+xml"/);
    for (const index of [1, 2]) {
      const slide = parts.get(`ppt/slides/slide${index}.xml`).toString(), rels = parts.get(`ppt/slides/_rels/slide${index}.xml.rels`).toString();
      assert.equal((slide.match(/<asvg:svgBlip/g) ?? []).length, index === 1 ? 2 : 1);
      assert.equal((rels.match(/Target="\.\.\/media\/pptkit-svg-pairs\/1\.svg"/g) ?? []).length, index === 1 ? 2 : 1);
    }
    const capped = await generate(p, { maxAssetBytes: png().length + svg.length - 1, resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
    assert.equal(capped.warnings[0].code, "asset-read-failed"); assert.equal(media(capped.bytes).length, 0);
  }
});

test("SVG pair fails closed for missing or contradictory data without a false success warning", async () => {
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
  const p = deck({ type: "bytes", value: svg }, "image/svg+xml");
  for (const resolved of [
    { bytes: png(), svgPair: new Uint8Array() },
    { bytes: png(), svgPair: [] },
    { bytes: png(), svgPair: new Uint8Array(new SharedArrayBuffer(1)) },
    { bytes: png(), svgPair: svg, vectorRasterization: "svg" },
    { bytes: png(), svgPair: svg, animationFallback: "first-frame" },
    { bytes: new TextEncoder().encode("GIF89a"), svgPair: svg },
    { bytes: Uint8Array.of(1, 2, 3), svgPair: svg },
  ]) {
    const result = await generatePptx(p, { resolveAsset: () => resolved });
    assert.equal(result.warnings[0].code, "asset-read-failed");
    assert.ok(!result.warnings.some(w => w.code === "asset-svg-pair-unverified"));
    assert.equal(media(result.bytes).length, 0);
  }
  const wrongSource = await generatePptx(deck(), { resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
  assert.equal(wrongSource.warnings[0].code, "asset-read-failed");
});

test("paired SVG uses shared parts from slide, layout and master owners", async () => {
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
  const p = createPresentation(); p.registerAsset({ id: "picture", kind: "image", mimeType: "image/svg+xml", source: { type: "bytes", value: svg } });
  const image = id => ({ type: "image", id, assetId: "picture", box: { x: 0, y: 0, width: 20, height: 20 } });
  p.defineSlideMaster({ id: "master", name: "Master", elements: [image("master-image")] });
  p.defineSlideLayout({ id: "layout", name: "Layout", masterId: "master", elements: [image("layout-image")] });
  p.addSlide({ layoutId: "layout", elements: [image("slide-image")] });
  const result = await generatePptx(p, { resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
  const parts = readGeneratedZip(result.bytes);
  assert.equal([...parts.keys()].filter(name => name.endsWith(".svg")).length, 1);
  for (const [owner, index] of [["slides/slide", 1], ["slideLayouts/slideLayout", 2], ["slideMasters/slideMaster", 2]]) {
    const xml = parts.get(`ppt/${owner}${index}.xml`).toString();
    const [directory, name] = owner.split("/");
    const rels = parts.get(`ppt/${directory}/_rels/${name}${index}.xml.rels`).toString();
    assert.match(xml, /<asvg:svgBlip/);
    assert.match(rels, /Target="\.\.\/media\/pptkit-svg-pairs\/1\.svg"/);
  }
});

test("unsafe and case-colliding asset IDs get distinct safe media part names", async () => {
  const { importPptx } = await import("../../pptx-parser/dist/index.js");
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
  const p = createPresentation(), ids = ["../slides/slide1.xml", "A", "a", "pptkit-asset-1", "pptkit-svg-pairs/1"];
  for (const id of ids) p.registerAsset({ id, kind: "image", mimeType: "image/svg+xml", source: { type: "bytes", value: svg } });
  const slide = p.addSlide();
  for (const [i, id] of ids.entries()) slide.addElement({ type: "image", id: `shape-${i}`, assetId: id, box: { x: i * 20, y: 0, width: 20, height: 20 } });
  const result = await generatePptx(p, { resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
  assert.equal(result.warnings.filter(w => w.code === "asset-svg-pair-unverified").length, ids.length);
  const parts = readGeneratedZip(result.bytes), names = [...parts.keys()].filter(name => name.startsWith("ppt/media/"));
  assert.equal(names.length, ids.length * 2);
  assert.equal(new Set(names.map(name => name.toLowerCase())).size, names.length);
  assert.ok(names.every(name => /^ppt\/media\/(?:[A-Za-z0-9._-]+|pptkit-svg-pairs\/[0-9]+\.svg)$/.test(name)));
  assert.equal(parts.has("ppt/slides/slide1.xml"), true);
  const session = await importPptx(result.bytes, { projection: "content" });
  assert.equal(session.document.slides[0].elements.length, ids.length);
  assert.equal(session.document.assets.length, ids.length);
  session.dispose();
  const ordinary = createPresentation();
  for (const id of ids) ordinary.registerAsset({ id, kind: "image", mimeType: "image/png", source: { type: "bytes", value: png() } });
  const page = ordinary.addSlide();
  for (const [i, id] of ids.entries()) page.addElement({ type: "image", id: `ordinary-${i}`, assetId: id, box: { x: i * 20, y: 0, width: 20, height: 20 } });
  const normal = await generatePptx(ordinary);
  const imported = await importPptx(normal.bytes, { projection: "content" });
  assert.equal(imported.document.assets.length, ids.length); assert.equal(imported.document.slides[0].elements.length, ids.length);
  imported.dispose();
});

test("SVG pair snapshots hostile views and aborts pending resolution without publishing", async () => {
  class Hostile extends Uint8Array { get byteLength() { return 0; } get buffer() { return new ArrayBuffer(0); } *[Symbol.iterator]() { throw Error("unexpected iterator"); } }
  const svg = new Hostile(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>'));
  const p = deck({ type: "path", value: "virtual" }, "image/svg+xml");
  const result = await generatePptx(p, { resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
  assert.deepEqual(readGeneratedZip(result.bytes).get("ppt/media/pptkit-svg-pairs/1.svg"), Buffer.from(new Uint8Array(svg)));
  const capped = await generatePptx(p, { maxAssetBytes: png().length + 1, resolveAsset: () => ({ bytes: png(), svgPair: svg }) });
  assert.equal(capped.warnings[0].code, "asset-read-failed");
  const controller = new AbortController(), reason = new Error("cancel pair"); let started, finish;
  const ready = new Promise(resolve => { started = resolve; });
  const pending = generatePptx(p, { signal: controller.signal, resolveAsset: () => { started(); return new Promise(resolve => { finish = resolve; }); } });
  await ready; controller.abort(reason); await assert.rejects(pending, e => e === reason);
  finish({ bytes: png(), svgPair: svg });
});

test("optional resolved-media deduplication shares only exact payloads across source types", async () => {
  const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>');
  const different = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>');
  const p = createPresentation();
  p.registerAsset({ id: "a", kind: "image", mimeType: "image/png", source: { type: "bytes", value: png() } });
  p.registerAsset({ id: "b", kind: "image", mimeType: "image/png", source: { type: "path", value: "virtual" } });
  p.registerAsset({ id: "c", kind: "image", mimeType: "image/svg+xml", source: { type: "url", value: "https://example.invalid/c" } });
  p.registerAsset({ id: "d", kind: "image", mimeType: "image/svg+xml", source: { type: "bytes", value: different } });
  p.registerAsset({ id: "e", kind: "image", mimeType: "image/svg+xml", source: { type: "path", value: "other" } });
  const slide = p.addSlide();
  for (const [i, id] of ["a", "b", "c", "d", "e"].entries()) slide.addElement({ type: "image", assetId: id, box: { x: i * 25, y: 0, width: 20, height: 20 } });
  const calls = [], resolveAsset = asset => { calls.push(asset.id); return ["c", "d", "e"].includes(asset.id) ? { bytes: png(), svgPair: asset.id === "d" ? different : svg } : { bytes: png() }; };
  const normal = await generatePptx(p, { resolveAsset }); assert.equal(media(normal.bytes).length, 8); calls.length = 0;
  const result = await generatePptx(p, { resolveAsset, deduplicateResolvedAssets: true });
  assert.deepEqual(calls, ["a", "b", "c", "d", "e"]);
  assert.equal(media(result.bytes).length, 5);
  assert.deepEqual(result.warnings.map(w => [w.code, w.assetId]), [["asset-svg-pair-unverified", "c"], ["asset-svg-pair-unverified", "d"], ["asset-svg-pair-unverified", "e"]]);
  assert.equal(p.assets.length, 5);
  const parts = readGeneratedZip(result.bytes), rels = parts.get("ppt/slides/_rels/slide1.xml.rels").toString();
  assert.equal((rels.match(/Target="\.\.\/media\/a\.png"/g) ?? []).length, 2);
  assert.equal((rels.match(/Target="\.\.\/media\/pptkit-svg-pairs\/1\.svg"/g) ?? []).length, 2);
  assert.equal((rels.match(/Target="\.\.\/media\/pptkit-svg-pairs\/2\.svg"/g) ?? []).length, 1);
  assert.deepEqual(parts.get("ppt/media/pptkit-svg-pairs/1.svg"), Buffer.from(svg));
  assert.deepEqual(parts.get("ppt/media/pptkit-svg-pairs/2.svg"), Buffer.from(different));
  await assert.rejects(generatePptx(p, { deduplicateResolvedAssets: "yes" }), TypeError);
});

test("resolved-media fingerprints never merge colliding but different bytes", async () => {
  // Equal 32-bit FNV-1a fingerprints, deliberately not valid SVG: source validity is caller-owned.
  const collision = [Uint8Array.of(16,198,154,146,156,197,35,166), Uint8Array.of(5,139,37,161,141,253,183,45)];
  const p = createPresentation(), slide = p.addSlide();
  for (const id of ["one", "two"]) {
    p.registerAsset({ id, kind: "image", mimeType: "image/svg+xml", source: { type: "path", value: id } });
    slide.addElement({ type: "image", assetId: id, box: { x: 0, y: 0, width: 20, height: 20 } });
  }
  const result = await generatePptx(p, { deduplicateResolvedAssets: true, resolveAsset: asset => ({ bytes: png(), svgPair: collision[asset.id === "one" ? 0 : 1] }) });
  const parts = readGeneratedZip(result.bytes);
  assert.equal(media(result.bytes).length, 4);
  assert.deepEqual(parts.get("ppt/media/pptkit-svg-pairs/1.svg"), Buffer.from(collision[0]));
  assert.deepEqual(parts.get("ppt/media/pptkit-svg-pairs/2.svg"), Buffer.from(collision[1]));
});

test("resolved bytes with different declared media types are not deduplicated", async () => {
  const dir = await mkdtemp(join(tmpdir(), "pptkit-media-types-")), file = join(dir, "source");
  try {
    await writeFile(file, png());
    const p = createPresentation(), slide = p.addSlide();
    for (const [id, mimeType] of [["png", "image/png"], ["jpg", "image/jpeg"]]) {
      p.registerAsset({ id, kind: "image", mimeType, source: { type: "path", value: file } });
      slide.addElement({ type: "image", assetId: id, box: { x: 0, y: 0, width: 20, height: 20 } });
    }
    // Default path loading preserves its existing unverified-MIME behavior.
    const result = await generateNode(p, { deduplicateResolvedAssets: true });
    assert.deepEqual(media(result.bytes).map(([name]) => name), ["ppt/media/png.png", "ppt/media/jpg.jpg"]);
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("signature classification is explicit for JPEG/GIF and does not claim full decoding", async () => {
  for (const [value, mimeType, ext] of [[Uint8Array.of(255,216,255), "image/jpeg", "jpg"], [new TextEncoder().encode("GIF89a"), "image/gif", "gif"]]) {
    const result = await generatePptx(deck({ type: "bytes", value }, mimeType));
    assert.equal(media(result.bytes)[0][0], `ppt/media/picture.${ext}`);
    assert.deepEqual(media(result.bytes)[0][1], Buffer.from(value));
  }
});

test("explicit resolver handles paths in browser-neutral export with detached frozen requests", async () => {
  const p = deck({ type: "path", value: "virtual/image.png" }), controller = new AbortController(); let calls = 0;
  const result = await generatePptx(p, { signal: controller.signal, resolveAsset(asset, context) {
    calls++; assert.equal(asset.id, "picture"); assert.equal(context.signal, controller.signal);
    assert.ok(Object.isFrozen(asset)); assert.ok(Object.isFrozen(asset.source)); assert.ok(Object.isFrozen(asset.accessibility));
    assert.throws(() => { asset.source.value = "changed"; }, TypeError);
    return { bytes: png(), mimeType: "image/png" };
  }});
  assert.equal(calls, 1); assert.deepEqual(result.warnings, []); assert.deepEqual(media(result.bytes)[0][1], Buffer.from(png()));
  assert.equal(p.assets[0].source.value, "virtual/image.png");
});

test("resolver omission and failures never fall through to implicit loading", async () => {
  const original = globalThis.fetch; globalThis.fetch = () => { throw new Error("unexpected fetch"); };
  try {
    for (const resolveAsset of [() => undefined, () => { throw new Error("offline"); }, () => ({ bytes: [] }), () => ({ bytes: png(), mimeType: "image/gif" }), () => ({ bytes: new Uint8Array(new SharedArrayBuffer(8)) })]) {
      const result = await generatePptx(deck({ type: "url", value: "https://example.invalid/picture" }), { resolveAsset });
      assert.equal(media(result.bytes).length, 0); assert.equal(result.warnings[0].code, "asset-read-failed");
    }
    const mismatched = await generatePptx(deck(undefined, "image/jpeg"), { resolveAsset: () => ({ bytes: png(), mimeType: "image/png" }) });
    assert.match(mismatched.warnings[0].message, /differs/);
  } finally { globalThis.fetch = original; }
});

test("resolver byte budgets and shared-buffer rejection use the real typed-array view",async()=>{
 class Hostile extends Uint8Array { get byteLength(){return 0} get buffer(){return new ArrayBuffer(0)} *[Symbol.iterator](){throw Error('unexpected iterator')} }
 const input=new Hostile(png()),p=deck({type:'path',value:'virtual'}),ok=await generatePptx(p,{resolveAsset:()=>({bytes:input,mimeType:'image/png'})});assert.deepEqual(media(ok.bytes)[0][1],Buffer.from(png()));
 const capped=await generatePptx(p,{maxAssetBytes:1,resolveAsset:()=>({bytes:input,mimeType:'image/png'})});assert.equal(capped.warnings[0].code,'asset-read-failed');assert.match(capped.warnings[0].message,/maxAssetBytes/);assert.equal(media(capped.bytes).length,0);
 if(typeof SharedArrayBuffer!=='undefined'){const shared=new Hostile(new SharedArrayBuffer(png().length));shared.set(png());const rejected=await generatePptx(p,{resolveAsset:()=>({bytes:shared,mimeType:'image/png'})});assert.equal(rejected.warnings[0].code,'asset-read-failed');assert.match(rejected.warnings[0].message,/SharedArrayBuffer/);assert.equal(media(rejected.bytes).length,0);}
});

test("resolver rejects a shared buffer originating in another JavaScript realm",async()=>{
 const vm=await import('node:vm'),foreign=vm.runInNewContext(`new SharedArrayBuffer(${png().length})`),view=new Uint8Array(foreign);view.set(png());assert.equal(foreign instanceof SharedArrayBuffer,false);
 const result=await generatePptx(deck({type:'path',value:'virtual'}),{resolveAsset:()=>({bytes:view,mimeType:'image/png'})});assert.equal(result.warnings[0].code,'asset-read-failed');assert.match(result.warnings[0].message,/SharedArrayBuffer/);assert.equal(media(result.bytes).length,0);
});

test("per-asset budgets cover byte, resolver, streamed URL and Node path loads", async () => {
  const dir = await mkdtemp(join(tmpdir(), "pptkit-byte-assets-")), file = join(dir, "test.png"); await writeFile(file, png());
  try {
    const paths = [() => generatePptx(deck(), { maxAssetBytes: 8 }), () => generatePptx(deck(), { maxAssetBytes: 8, resolveAsset: () => ({ bytes: png() }) }), () => generatePptx(deck({ type: "url", value: pixelPng }), { maxAssetBytes: 8 }), () => generateNode(deck({ type: "path", value: file }), { maxAssetBytes: 8 })];
    for (const run of paths) { const result = await run(); assert.equal(media(result.bytes).length, 0); assert.match(result.warnings[0].message, /maxAssetBytes/); }
    for (const generate of [generatePptx, generateNode]) assert.deepEqual((await generate(deck(), { maxAssetBytes: png().length })).warnings, []);
    assert.deepEqual((await generateNode(deck({ type: "path", value: file }))).warnings, []);
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("invalid limits/resolvers and pre-abort reject before resource work", async () => {
  for (const value of [0, -1, 1.5, Infinity, NaN]) await assert.rejects(generatePptx(deck(), { maxAssetBytes: value }), RangeError);
  await assert.rejects(generatePptx(deck(), { resolveAsset: true }), TypeError);
  const controller = new AbortController(), reason = new Error("stop"); controller.abort(reason); let calls = 0;
  await assert.rejects(generatePptx(deck(), { signal: controller.signal, resolveAsset: () => { calls++; return { bytes: png() }; } }), e => e === reason);
  assert.equal(calls, 0);
});

test("abort interrupts uncooperative resolver waiting and later resolution cannot complete export", async () => {
  const controller = new AbortController(), reason = new Error("cancelled"); let finish, started;
  const began = new Promise(resolve => { started = resolve; });
  const pending = generatePptx(deck(), { signal: controller.signal, resolveAsset: () => { started(); return new Promise(resolve => { finish = resolve; }); } });
  await began; controller.abort(reason);
  await assert.rejects(pending, e => e === reason);
  finish({ bytes: png() }); await new Promise(resolve => setTimeout(resolve, 0));
  assert.equal((await generatePptx(deck())).status, "generated");
});

test("document and options are captured before waiting on a resolver", async () => {
  const p = deck(), options = {}; let finish;
  options.resolveAsset = () => new Promise(resolve => { finish = resolve; });
  const pending = generatePptx(p, options);
  p.addSlide(); options.resolveAsset = () => { throw new Error("changed callback"); }; options.maxAssetBytes = 1;
  finish({ bytes: png() }); const result = await pending;
  assert.equal(result.slideCount, 1); assert.deepEqual(result.warnings, []);
});

test("stream budgets cancel oversized readers instead of consuming the remaining response", async () => {
  const original = globalThis.fetch; let cancelled = false;
  globalThis.fetch = async () => new Response(new ReadableStream({ start(c) { c.enqueue(new Uint8Array(20)); }, cancel() { cancelled = true; } }), { headers: { "content-type": "image/png" } });
  try {
    const result = await generatePptx(deck({ type: "url", value: "https://example.invalid/stream" }), { maxAssetBytes: 10 });
    assert.equal(cancelled, true); assert.match(result.warnings[0].message, /maxAssetBytes/);
  } finally { globalThis.fetch = original; }
});

test("abort cancels a stalled response reader and is not converted to an omission warning", async () => {
  const original = globalThis.fetch, controller = new AbortController(); let started, cancelled = false;
  const began = new Promise(resolve => { started = resolve; });
  globalThis.fetch = async () => new Response(new ReadableStream({ start() { started(); }, cancel() { cancelled = true; } }));
  try {
    const pending = generatePptx(deck({ type: "url", value: "https://example.invalid/stream" }), { signal: controller.signal });
    await began; await new Promise(resolve => setTimeout(resolve, 0)); controller.abort();
    await assert.rejects(pending, { name: "AbortError" });
    await new Promise(resolve => setTimeout(resolve, 0)); assert.equal(cancelled, true);
  } finally { globalThis.fetch = original; }
});


test("resolved bytes are copied before subsequent asset work can mutate their original buffer", async () => {
  const p = deck(), shared = png(), expected = Buffer.from(shared);
  p.registerAsset({ id: "second", kind: "image", source: { type: "path", value: "second" } });
  const result = await generatePptx(p, { resolveAsset: (asset) => {
    if (asset.id === "picture") return { bytes: shared };
    shared.fill(0); return { bytes: png() };
  }});
  assert.deepEqual(media(result.bytes)[0][1], expected);
});

test("URL stream chunks are detached even when a producer reuses its buffer", async () => {
  const original = globalThis.fetch, bytes = png(), reused = new Uint8Array(1); let offset = 0;
  globalThis.fetch = async () => new Response(new ReadableStream({ pull(controller) {
    if (offset === bytes.length) { controller.close(); return; }
    reused[0] = bytes[offset++]; controller.enqueue(reused);
  } }, { highWaterMark: 0 }), { headers: { "content-type": "image/png" } });
  try {
    const result = await generatePptx(deck({ type: "url", value: "https://example.invalid/reused" }));
    assert.deepEqual(media(result.bytes)[0][1], Buffer.from(bytes));
  } finally { globalThis.fetch = original; }
});

test("byte and data-URL inputs produce identical package contents", async () => {
  assert.deepEqual((await generatePptx(deck())).bytes, (await generatePptx(deck({ type: "url", value: pixelPng }))).bytes);
});


test("shared asset replacement changes media and relationships without leaving stale package parts", async () => {
  const p = deck(), before = await generatePptx(p);
  const gif = new Uint8Array(Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64"));
  p.replaceAsset("picture", { kind: "image", source: { type: "bytes", value: gif }, mimeType: "image/gif", width: 1, height: 1 });
  const after = await generatePptx(p), entries = readGeneratedZip(after.bytes);
  assert.deepEqual(after.warnings, []); assert.equal(media(after.bytes).length, 1);
  assert.equal(media(after.bytes)[0][0], "ppt/media/picture.gif"); assert.deepEqual(media(after.bytes)[0][1], Buffer.from(gif));
  assert.equal(entries.has("ppt/media/picture.png"), false);
  const relationships = entries.get("ppt/slides/_rels/slide1.xml.rels").toString();
  assert.equal((relationships.match(/Target="..\/media\/picture.gif"/g) ?? []).length, 2);
  assert.match(entries.get("[Content_Types].xml").toString(), /ContentType="image\/gif"/);
  assert.equal(media(before.bytes)[0][0], "ppt/media/picture.png");
  for (const element of p.slides[0].elements) p.slides[0].removeElement(element.id);
  assert.deepEqual(p.removeUnusedAssets().map(asset => asset.id), ["picture"]);
  assert.equal(media((await generatePptx(p)).bytes).length, 0);
});

test("byte content deduplication yields a single media entry without conflating explicit identities", async () => {
  const p = deck(), input = { kind: "image", source: { type: "bytes", value: png() } };
  assert.equal(p.registerAsset(input, { deduplicate: "content" }).id, "picture");
  assert.equal(media((await generatePptx(p)).bytes).length, 1);
  p.registerAsset({ ...input, id: "explicit-other" }, { deduplicate: "content" });
  assert.equal(media((await generatePptx(p)).bytes).length, 2);
  assert.deepEqual(p.removeUnusedAssets().map(asset => asset.id), ["explicit-other"]);
  assert.equal(media((await generatePptx(p)).bytes).length, 1);
});

test("opt-in GIF diagnostics report animation once per packaged asset without changing bytes", async () => {
  const gif = Uint8Array.from(Buffer.from("R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7", "base64"));
  const p = deck({ type: "bytes", value: gif });
  const plain = await generatePptx(p), checked = await generatePptx(p, { gifDiagnostics: "warn" });
  const native = await generateNode(p, { gifDiagnostics: "warn" });
  assert.deepEqual(await generatePptx(p, { gifDiagnostics: "reject-lzw" }), checked);
  assert.deepEqual(plain.warnings, []); assert.deepEqual(plain.bytes, checked.bytes);
  assert.deepEqual(checked.bytes, native.bytes); assert.deepEqual(checked.warnings, native.warnings);
  assert.equal(checked.status, "generated-with-warnings");
  assert.deepEqual(checked.warnings.map(({ code, assetId }) => [code, assetId]), [["gif-animation-playback-unverified", "picture"]]);
  assert.deepEqual(media(checked.bytes)[0][1], Buffer.from(gif));
  assert.equal(media(checked.bytes).length, 1);
});

test("opt-in PNG diagnostics report APNG and preserve bytes once per shared asset", async () => {
  const apng=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url)),p=deck({type:'bytes',value:apng});
  const plain=await generatePptx(p),checked=await generatePptx(p,{pngDiagnostics:'warn'}),node=await generateNode(p,{pngDiagnostics:'warn'});
  assert.deepEqual(plain.warnings,[]);assert.deepEqual(checked.bytes,plain.bytes);assert.deepEqual(node,checked);assert.deepEqual(checked.warnings.map(w=>[w.code,w.assetId]),[['apng-animation-playback-unverified','picture']]);assert.deepEqual(media(checked.bytes)[0][1],apng);
});

test("declared image/apng is packaged as PNG with diagnostics across byte, path and resolver inputs", async () => {
  const apng=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url)),path=new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url).pathname;
  const documents=[await generatePptx(deck({type:'bytes',value:apng},' IMAGE/APNG '),{pngDiagnostics:'warn'}),await generateNode(deck({type:'path',value:path},'image/apng'),{pngDiagnostics:'warn'}),await generatePptx(deck({type:'path',value:'virtual'},'image/apng'),{pngDiagnostics:'warn',resolveAsset:()=>({bytes:apng,mimeType:'image/png'})})];
  for(const result of documents){assert.deepEqual(result.warnings.map(w=>w.code),['apng-animation-playback-unverified']);assert.deepEqual(media(result.bytes)[0][1],apng);assert.match(readGeneratedZip(result.bytes).get('[Content_Types].xml').toString(),/Extension="png" ContentType="image\/png"/);}
  const mismatch=await generatePptx(deck({type:'bytes',value:Uint8Array.of(255,216,255)},'image/apng'));assert.equal(mismatch.warnings[0].code,'asset-read-failed');assert.equal(media(mismatch.bytes).length,0);
});

test("URL image/apng content type is normalized for package without changing existing URL trust", async () => {
  const bytes=await readFile(new URL('../../../tests/fixtures/two-frame-apng.png',import.meta.url)),original=globalThis.fetch;globalThis.fetch=async()=>new Response(bytes,{status:200,headers:{'content-type':'image/apng'}});
  try{const result=await generatePptx(deck({type:'url',value:'https://example.invalid/a.apng'}),{pngDiagnostics:'warn'});assert.deepEqual(result.warnings.map(w=>w.code),['apng-animation-playback-unverified']);assert.match(readGeneratedZip(result.bytes).get('[Content_Types].xml').toString(),/Extension="png" ContentType="image\/png"/);assert.deepEqual(media(result.bytes)[0][1],bytes);}finally{globalThis.fetch=original;}
});

test("PNG diagnostics distinguish static, malformed and unresolved structures without removing media", async () => {
  const staticPng=await readFile(new URL('../../../tests/fixtures/transparent-replacement.png',import.meta.url));assert.deepEqual((await generatePptx(deck({type:'bytes',value:staticPng}),{pngDiagnostics:'warn'})).warnings,[]);
  const unknown=Buffer.concat([staticPng.subarray(0,33),Buffer.from('0000000041424344db1720a5','hex'),staticPng.subarray(33)]),unknownResult=await generatePptx(deck({type:'bytes',value:unknown}),{pngDiagnostics:'warn'});assert.deepEqual(unknownResult.warnings.map(w=>w.code),['png-inspection-incomplete']);assert.deepEqual(media(unknownResult.bytes)[0][1],unknown);
  const incomplete=staticPng.subarray(0,16),result=await generatePptx(deck({type:'bytes',value:incomplete}),{pngDiagnostics:'warn'});assert.deepEqual(result.warnings.map(w=>w.code),['png-inspection-incomplete']);assert.deepEqual(media(result.bytes)[0][1],incomplete);
  const resolved=await generatePptx(deck({type:'path',value:'virtual'}),{pngDiagnostics:'warn',resolveAsset:()=>({bytes:incomplete,mimeType:'image/png'})});assert.deepEqual(resolved.warnings.map(w=>w.code),['png-inspection-incomplete']);assert.deepEqual(media(resolved.bytes)[0][1],incomplete);
  for(const value of [null,'invalid',3])await assert.rejects(generatePptx(deck(),{pngDiagnostics:value}),/pngDiagnostics/);
});

test("PNG diagnostics propagate cancellation during asynchronous asset resolution", async () => {
  const controller=new AbortController();let finish;const pending=generatePptx(deck({type:'path',value:'virtual'}),{pngDiagnostics:'warn',signal:controller.signal,resolveAsset:()=>new Promise(resolve=>{finish=resolve})});
  await new Promise(resolve=>setImmediate(resolve));controller.abort(new Error('png-cancelled'));finish({bytes:png(),mimeType:'image/png'});await assert.rejects(pending,/png-cancelled/);
});

test("GIF diagnostics distinguish static input and incomplete inspection without omitting media", async () => {
  const staticGif = Uint8Array.from(Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64"));
  const staticResult = await generatePptx(deck({ type: "bytes", value: staticGif }), { gifDiagnostics: "warn" });
  assert.deepEqual(staticResult.warnings, []);
  const incomplete = new TextEncoder().encode("GIF89a");
  const bad = await generatePptx(deck({ type: "bytes", value: incomplete }), { gifDiagnostics: "warn" });
  assert.deepEqual(bad.warnings.map(w => w.code), ["gif-inspection-incomplete"]);
  assert.deepEqual(media(bad.bytes)[0][1], Buffer.from(incomplete));
  const resolved = await generatePptx(deck({ type: "path", value: "virtual" }), {
    gifDiagnostics: "warn", resolveAsset: () => ({ bytes: incomplete, mimeType: "image/gif" }),
  });
  assert.deepEqual(resolved.warnings.map(w => w.code), ["gif-inspection-incomplete"]);
  assert.deepEqual(media(resolved.bytes)[0][1], Buffer.from(incomplete));
  await assert.rejects(generatePptx(deck(), { gifDiagnostics: "invalid" }), /gifDiagnostics/);
  await assert.rejects(generateNode(deck(), { gifDiagnostics: null }), /gifDiagnostics/);
});

test("GIF diagnostics report interactive controls and reserved disposal even on one frame", async () => {
  const header = [...Buffer.from("GIF89a"), 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255];
  const frame = [0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0x44, 0];
  for (const flags of [2, 16, 28, 0xe0]) {
    const gif = Uint8Array.from([...header, 0x21, 0xf9, 4, flags, 0, 0, 0, 0, ...frame, 0x3b]);
    const plain = await generatePptx(deck({ type: "bytes", value: gif }));
    const checked = await generatePptx(deck({ type: "bytes", value: gif }), { gifDiagnostics: "warn" });
    assert.deepEqual(plain.warnings, []);
    assert.deepEqual(checked.warnings.map(w => w.code), ["gif-inspection-incomplete"]);
    assert.deepEqual(checked.bytes, plain.bytes);
    assert.deepEqual(media(checked.bytes)[0][1], Buffer.from(gif));
  }
});

test("GIF inspection budget exhaustion is reported as unknown rather than static or omitted", async () => {
  const header = [...Buffer.from("GIF89a"), 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255];
  const frame = [0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0x44, 0];
  const gif = new Uint8Array(header.length + frame.length * 10001 + 1);
  gif.set(header); for (let i = 0; i < 10001; i++) gif.set(frame, header.length + i * frame.length);
  gif[gif.length - 1] = 0x3b;
  const result = await generatePptx(deck({ type: "bytes", value: gif }), { gifDiagnostics: "warn" });
  assert.deepEqual(result.warnings.map(w => w.code), ["gif-inspection-incomplete"]);
  assert.match(result.warnings[0].message, /maxFrames/);
  assert.deepEqual(media(result.bytes)[0][1], Buffer.from(gif));
});

test("resolver-reported first-frame conversion emits an asset-level loss warning", async () => {
  const originalGif = Uint8Array.from(Buffer.from("R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7", "base64"));
  // This PNG is frame 0 decoded independently with Pillow 12.1.1.
  const firstFrame = Uint8Array.from(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR4nGP8zwACTGCSAQANHQEDgslx/wAAAABJRU5ErkJggg==", "base64"));
  const p = deck({ type: "bytes", value: originalGif }); let calls = 0;
  const resolveAsset = async () => { calls++; return { bytes: firstFrame, mimeType: "image/png", animationFallback: "first-frame" }; };
  const a = await generatePptx(p, { resolveAsset }), b = await generateNode(p, { resolveAsset, gifDiagnostics: "warn" });
  assert.equal(calls, 2); assert.deepEqual(a.bytes, b.bytes); assert.deepEqual(a.warnings, b.warnings);
  assert.equal(a.status, "generated-with-warnings"); assert.equal(a.warnings.length, 1);
  assert.equal(a.warnings[0].code, "asset-animation-flattened"); assert.equal(a.warnings[0].assetId, "picture");
  assert.deepEqual(media(a.bytes)[0][1], Buffer.from(firstFrame)); assert.equal(media(a.bytes).length, 1);
  const noReport = await generatePptx(p, { resolveAsset: () => ({ bytes: firstFrame, mimeType: "image/png" }) });
  assert.deepEqual(noReport.warnings, []); assert.deepEqual(noReport.bytes, a.bytes);
});

test("invalid or incompatible fallback declarations fail as asset reads without false loss claims", async () => {
  const gif = Uint8Array.from(Buffer.from("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7", "base64"));
  for (const resolved of [
    { bytes: png(), animationFallback: null }, { bytes: png(), animationFallback: "unknown" },
    { bytes: gif, mimeType: "image/gif", animationFallback: "first-frame" },
  ]) {
    const result = await generatePptx(deck({ type: "path", value: "virtual" }), { resolveAsset: () => resolved });
    assert.equal(media(result.bytes).length, 0);
    assert.equal(result.warnings.filter(w => w.code === "asset-read-failed").length, 1);
    assert.equal(result.warnings.filter(w => w.code === "image-omitted").length, 2);
    assert.ok(!result.warnings.some(w => w.code === "asset-animation-flattened"));
  }
  const declaredGif = deck({ type: "path", value: "virtual.gif" }, "image/gif");
  const mismatch = await generatePptx(declaredGif, { resolveAsset: () => ({ bytes: png(), mimeType: "image/png", animationFallback: "first-frame" }) });
  assert.match(mismatch.warnings[0].message, /MIME/);
  assert.ok(!mismatch.warnings.some(w => w.code === "asset-animation-flattened"));
});

test("aborted first-frame resolution rejects instead of publishing a fallback warning", async () => {
  const controller = new AbortController(), reason = new Error("cancel fallback");
  let finish, started;
  const ready = new Promise(resolve => { started = resolve; });
  const pending = generatePptx(deck({ type: "path", value: "virtual" }), { signal: controller.signal,
    resolveAsset: () => { started(); return new Promise(resolve => { finish = resolve; }); } });
  await ready; controller.abort(reason);
  await assert.rejects(pending, error => error === reason);
  finish({ bytes: png(), animationFallback: "first-frame" });
});

test("GIF diagnostics do not treat uninspected graphic or extension blocks as static", async () => {
  const header = [...Buffer.from("GIF89a"), 1, 0, 1, 0, 0x80, 0, 0, 0, 0, 0, 255, 255, 255];
  const frame = [0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 1, 0x44, 0];
  const plainText = [0x21, 1, 12, ...Array(12).fill(0), 1, 65, 0];
  const unknownApp = [0x21, 0xff, 11, ...Buffer.from("UNKNOWN1234"), 1, 5, 0];
  const unknownExtension = [0x21, 0xee, 1, 42, 0];
  const comment = [0x21, 0xfe, 1, 65, 0];
  for (const [extra, expected] of [[plainText, ["gif-inspection-incomplete"]], [unknownApp, ["gif-inspection-incomplete"]],
    [unknownExtension, ["gif-inspection-incomplete"]], [comment, []], [[...plainText, ...frame], ["gif-animation-playback-unverified", "gif-inspection-incomplete"]]]) {
    const gif = Uint8Array.from([...header, ...extra, ...frame, 0x3b]);
    const result = await generatePptx(deck({ type: "bytes", value: gif }), { gifDiagnostics: "warn" });
    assert.deepEqual(result.warnings.map(w => w.code), expected);
    assert.deepEqual(media(result.bytes)[0][1], Buffer.from(gif));
  }
});

test("opt-in JPEG orientation diagnostics preserve bytes and warn once per shared asset", async () => {
  const jpeg = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const p = deck({ type: 'bytes', value: jpeg });
  const normal = await generatePptx(p), checked = await generatePptx(p, { jpegOrientationDiagnostics: 'warn' });
  const native = await generateNode(p, { jpegOrientationDiagnostics: 'warn' });
  assert.deepEqual(normal.warnings, []); assert.deepEqual(normal.bytes, checked.bytes); assert.deepEqual(checked.bytes, native.bytes);
  assert.deepEqual(checked.warnings, native.warnings); assert.equal(checked.status, 'generated-with-warnings');
  assert.deepEqual(checked.warnings.map(w => [w.code, w.assetId]), [['jpeg-exif-orientation-unverified','picture']]);
  assert.match(checked.warnings[0].message, /orientation 6/); assert.equal(media(checked.bytes).length, 1);
  assert.deepEqual(media(checked.bytes)[0][1], jpeg);
});

test("JPEG diagnostic distinguishes default orientation and incomplete inspection without discarding output", async () => {
  const jpeg = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const standard = Buffer.from(jpeg), offset = standard.indexOf(Buffer.from('Exif\0\0'));
  assert.ok(offset >= 0); standard[offset + 6 + 18] = 0; standard[offset + 6 + 19] = 1;
  assert.deepEqual((await generatePptx(deck({ type: 'bytes', value: standard }), { jpegOrientationDiagnostics: 'warn' })).warnings, []);
  for (let orientation = 2; orientation <= 8; orientation++) {
    const modified = Buffer.from(standard); modified[offset + 6 + 19] = orientation;
    const output = await generatePptx(deck({ type: 'bytes', value: modified }), { jpegOrientationDiagnostics: 'warn' });
    assert.equal(output.warnings[0].code, 'jpeg-exif-orientation-unverified');
    assert.match(output.warnings[0].message, new RegExp(`orientation ${orientation}`));
    assert.deepEqual(media(output.bytes)[0][1], modified);
  }
  const partial = Uint8Array.of(255,216,255,225,0,6,69,120,105,102,255,217);
  const result = await generatePptx(deck({ type: 'bytes', value: partial }), { jpegOrientationDiagnostics: 'warn' });
  assert.deepEqual(result.warnings.map(w => w.code), ['jpeg-exif-inspection-incomplete']);
  assert.deepEqual(media(result.bytes)[0][1], Buffer.from(partial));
  const resolved = await generatePptx(deck({ type: 'path', value: 'virtual' }), {
    jpegOrientationDiagnostics: 'warn', resolveAsset: () => ({ bytes: jpeg, mimeType: 'image/jpeg' }),
  });
  assert.deepEqual(resolved.warnings.map(w => w.code), ['jpeg-exif-orientation-unverified']);
  for (const invalid of [null, 'invalid', 3]) await assert.rejects(generatePptx(deck(), { jpegOrientationDiagnostics: invalid }), /jpegOrientationDiagnostics/);
});

test("JPEG diagnostic cancellation never degrades into an inspection warning", async () => {
  const controller = new AbortController(), reason = new Error('cancel jpeg');
  let finish, started; const ready = new Promise(resolve => { started = resolve; });
  const pending = generatePptx(deck({ type: 'path', value: 'virtual' }), { jpegOrientationDiagnostics: 'warn', signal: controller.signal,
    resolveAsset: () => { started(); return new Promise(resolve => { finish = resolve; }); } });
  await ready; controller.abort(reason); await assert.rejects(pending, error => error === reason);
  finish({ bytes: await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url)) });
});

test('resolver-attested orientation normalization allows JPEG-to-PNG and reports changed media', async () => {
  const jpeg = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const normalized = await readFile(new URL('../../../tests/fixtures/exif-orientation6-normalized.png', import.meta.url));
  const p = deck({ type: 'bytes', value: jpeg }, 'image/jpeg');
  const resolveAsset = () => ({ bytes: normalized, mimeType: 'image/png', orientationNormalization: 'pixels' });
  const result = await generatePptx(p, { resolveAsset, jpegOrientationDiagnostics: 'warn' });
  const node = await generateNode(p, { resolveAsset, jpegOrientationDiagnostics: 'warn' });
  assert.deepEqual(result.bytes, node.bytes); assert.deepEqual(result.warnings, node.warnings);
  assert.deepEqual(result.warnings.map(w => [w.code,w.assetId]), [['asset-orientation-normalized','picture']]);
  assert.equal(result.status, 'generated-with-warnings'); assert.equal(media(result.bytes).length, 1);
  assert.equal(media(result.bytes)[0][0], 'ppt/media/picture.png'); assert.deepEqual(media(result.bytes)[0][1], normalized);
  assert.deepEqual(Buffer.from(p.assets[0].source.value, 'base64'), jpeg);
  const slide = readGeneratedZip(result.bytes).get('ppt/slides/slide1.xml').toString();
  assert.equal((slide.match(/<p:pic>/g) ?? []).length, 2);
  const missingDeclaration = await generatePptx(p, { resolveAsset: () => ({ bytes: normalized, mimeType: 'image/png' }) });
  assert.equal(missingDeclaration.warnings[0].code, 'asset-read-failed');
});

test('orientation normalization rejects incompatible and contradictory resolver claims', async () => {
  const jpeg = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const p = deck({ type: 'bytes', value: jpeg }, 'image/jpeg');
  for (const resolved of [
    { bytes: png(), orientationNormalization: null },
    { bytes: png(), orientationNormalization: 'unknown' },
    { bytes: jpeg, mimeType: 'image/jpeg', orientationNormalization: 'pixels' },
    { bytes: Uint8Array.from(Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7','base64')), orientationNormalization: 'pixels' },
    { bytes: png(), orientationNormalization: 'pixels', animationFallback: 'first-frame' },
    { bytes: png(), mimeType: 'image/jpeg', orientationNormalization: 'pixels' },
  ]) {
    const result = await generatePptx(p, { resolveAsset: () => resolved });
    assert.equal(media(result.bytes).length, 0); assert.equal(result.warnings[0].code, 'asset-read-failed');
    assert.equal(result.warnings.filter(w => w.code === 'image-omitted').length, 2);
    assert.ok(!result.warnings.some(w => w.code === 'asset-orientation-normalized'));
  }
  const notJpeg = deck({ type:'path',value:'virtual.png' }, 'image/png');
  const result = await generatePptx(notJpeg, { resolveAsset: () => ({ bytes: png(), orientationNormalization: 'pixels' }) });
  assert.match(result.warnings[0].message, /JPEG source MIME/);
});

test('orientation-normalized JPEG must not retain a non-default EXIF tag', async () => {
  const original = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const normalized = Buffer.from(original), offset = normalized.indexOf(Buffer.from('Exif\0\0'));
  normalized[offset + 6 + 18] = 0; normalized[offset + 6 + 19] = 1;
  const result = await generatePptx(deck({ type:'bytes',value:original }, 'image/jpeg'), {
    resolveAsset: () => ({ bytes: normalized, mimeType:'image/jpeg', orientationNormalization:'pixels' }),
  });
  assert.equal(result.warnings[0].code, 'asset-orientation-normalized');
  assert.deepEqual(media(result.bytes)[0][1], normalized);
});

test('aborting deferred orientation normalization does not publish a transformed result', async () => {
  const controller = new AbortController(), reason = new Error('cancel orientation');
  let finish, started; const ready = new Promise(resolve => { started = resolve; });
  const pending = generatePptx(deck({type:'path',value:'virtual'}, 'image/jpeg'), { signal:controller.signal,
    resolveAsset: () => { started(); return new Promise(resolve => { finish = resolve; }); } });
  await ready; controller.abort(reason); await assert.rejects(pending, error => error === reason);
  finish({bytes:png(), mimeType:'image/png', orientationNormalization:'pixels'});
});

test("strict external URL loading matches signatures, declared MIME, and concrete HTTP MIME", async () => {
  const source = { type: "url", value: "https://example.invalid/wrong.gif" };
  const strict = { externalImageValidation: "signature" };
  const previous = globalThis.fetch;
  let responseMime, responseBytes = png();
  globalThis.fetch = async () => new Response(responseBytes, { headers: responseMime ? { "content-type": responseMime } : {} });
  try {
    for (responseMime of [undefined, "application/octet-stream", " IMAGE/APNG ; x=y", "image/png"]) {
      const result = await generatePptx(deck(source), strict);
      assert.deepEqual(result.warnings, []);
      assert.deepEqual(media(result.bytes), [["ppt/media/picture.png", Buffer.from(png())]]);
    }
    responseMime = "image/jpeg";
    const badHeader = await generatePptx(deck(source, "image/png"), strict);
    assert.equal(media(badHeader.bytes).length, 0);
    assert.match(badHeader.warnings[0].message, /HTTP Content-Type image\/jpeg does not match image\/png/);
    assert.equal(badHeader.warnings.filter(w => w.code === "image-omitted").length, 2);
    responseMime = "image/png";
    const badDeclaration = await generatePptx(deck(source, "image/gif"), strict);
    assert.equal(media(badDeclaration.bytes).length, 0);
    assert.match(badDeclaration.warnings[0].message, /Declared MIME image\/gif/);
    responseBytes = new TextEncoder().encode("<html>not an image</html>");
    const badBytes = await generatePptx(deck(source), strict);
    assert.equal(media(badBytes.bytes).length, 0);
    assert.match(badBytes.warnings[0].message, /unsupported image signature/);
    assert.equal(media((await generatePptx(deck(source))).bytes).length, 1, "legacy default remains unchanged");
  } finally { globalThis.fetch = previous; }
});

test("strict Node path loading uses the signature without changing budget or resolver behavior", async () => {
  const directory = await mkdtemp(join(tmpdir(), "pptkit-asset-signature-"));
  try {
    const file = join(directory, "picture.jpg");
    await writeFile(file, png());
    const source = { type: "path", value: file }, p = deck(source);
    const strict = { externalImageValidation: "signature" };
    assert.equal(media((await generateNode(p)).bytes)[0][0], "ppt/media/picture.jpg");
    assert.equal(media((await generateNode(p, strict)).bytes)[0][0], "ppt/media/picture.png");
    const mismatch = await generateNode(deck(source, "image/jpeg"), strict);
    assert.equal(media(mismatch.bytes).length, 0);
    assert.match(mismatch.warnings[0].message, /Declared MIME image\/jpeg/);
    const capped = await generateNode(p, { ...strict, maxAssetBytes: 8 });
    assert.equal(media(capped.bytes).length, 0);
    assert.match(capped.warnings[0].message, /maxAssetBytes/);
    await assert.rejects(generateNode(p, { externalImageValidation: "headers" }), /externalImageValidation/);
    const resolved = await generateNode(deck({ type: "path", value: join(directory, "missing") }),
      { ...strict, resolveAsset: () => ({ bytes: png() }) });
    assert.deepEqual(resolved.warnings, []);
    assert.deepEqual(media(resolved.bytes)[0][1], Buffer.from(png()));
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test("reject-structure omits damaged PNG/APNG media and keeps valid controls unchanged", async () => {
  const options = { pngDiagnostics: "reject-structure" };
  const still = await readFile(new URL("../../../tests/fixtures/transparent-replacement.png", import.meta.url));
  const animated = await readFile(new URL("../../../tests/fixtures/two-frame-apng.png", import.meta.url));
  const bytes = value => deck({ type: "bytes", value });
  const crc = Buffer.from(still);
  crc[crc.length - 1] ^= 1;
  const unknown = Buffer.concat([still.subarray(0, 33), Buffer.from("0000000041424344db1720a5", "hex"), still.subarray(33)]);
  const emptyIdat = Buffer.concat([still.subarray(0, 33), Buffer.from("000000004944415435af061e", "hex"), still.subarray(-12)]);
  const animationControl = animated.indexOf("acTL") - 4;
  assert.ok(animationControl >= 33);
  const wrongFrameCount = Buffer.from(animated);
  Buffer.from("000000086163544c0000000300000000ceedbac0", "hex").copy(wrongFrameCount, animationControl);

  for (const [data, reason] of [
    [crc, /CRC for IEND/], [unknown, /unsupported critical chunk: ABCD/],
    [emptyIdat, /IDAT has no compressed data/], [wrongFrameCount, /frame count/],
    [still.subarray(0, 16), /truncated chunk/],
  ]) {
    const p = bytes(data);
    const lenient = await generatePptx(p, { pngDiagnostics: "warn" });
    assert.deepEqual(media(lenient.bytes)[0][1], data);
    const strict = await generatePptx(p, options);
    assert.equal(media(strict.bytes).length, 0);
    assert.deepEqual(strict.warnings.map(w => [w.code, w.assetId]),
      [["asset-read-failed", "picture"], ["image-omitted", "picture"], ["image-omitted", "picture"]]);
    assert.match(strict.warnings[0].message, reason);
  }
  const strictStill = await generatePptx(bytes(still), options);
  assert.deepEqual(strictStill, await generatePptx(bytes(still)));
  const strictAnimated = await generatePptx(bytes(animated), options);
  assert.deepEqual(strictAnimated, await generatePptx(bytes(animated), { pngDiagnostics: "warn" }));
  assert.deepEqual(strictAnimated.warnings.map(w => w.code), ["apng-animation-playback-unverified"]);

  const url = { type: "url", value: `data:image/gif;base64,${still.toString("base64")}` };
  const mismatched = await generatePptx(deck(url), options);
  assert.equal(media(mismatched.bytes).length, 0);
  assert.match(mismatched.warnings[0].message, /PNG signature does not match/);
  const resolved = await generatePptx(deck({ type: "path", value: "virtual" }),
    { ...options, resolveAsset: () => ({ bytes: crc }) });
  assert.equal(media(resolved.bytes).length, 0);
  assert.match(resolved.warnings[0].message, /CRC for IEND/);
});

test("GIF LZW validation tracks code widths, deferred clears, palettes and complete frames", async () => {
  const strict = { gifDiagnostics: "reject-lzw" };
  function frame({ width = 1, height = 1, screenWidth = width, screenHeight = height, palette = 2,
    localPalette = 0, minimum = 2, codes, extra = [], control = [], split = 255 }) {
    const colorCode = n => Math.log2(n) - 1;
    const word = n => [n & 255, n >>> 8];
    const bytes = [...Buffer.from("GIF89a"), ...word(screenWidth), ...word(screenHeight),
      palette ? 0x80 | colorCode(palette) : 0, 0, 0, ...Array(palette * 3).fill(0),
      ...control, 0x2c, 0, 0, 0, 0, ...word(width), ...word(height),
      localPalette ? 0x80 | colorCode(localPalette) : 0, ...Array(localPalette * 3).fill(0), minimum];
    const compressed = [];
    let bits = 0, used = 0;
    for (const [code, size] of codes) {
      bits |= code << used; used += size;
      while (used >= 8) { compressed.push(bits & 255); bits >>>= 8; used -= 8; }
    }
    if (used) compressed.push(bits);
    compressed.push(...extra);
    for (let i = 0; i < compressed.length; i += split) {
      const block = compressed.slice(i, i + split);
      bytes.push(block.length, ...block);
    }
    bytes.push(0, 0x3b);
    return Buffer.from(bytes);
  }
  const pack = (gif, options = strict) => generatePptx(deck({ type: "bytes", value: gif }), options);
  const validCodes = [[4, 3], [0, 3], [5, 3]];
  const special = frame({ width: 6, codes: [[4, 3], [0, 3], [6, 3], [6, 3], [4, 4], [1, 3], [5, 3]], split: 1 });
  const literalWidth = count => count < 255 ? 9 : count < 767 ? 10 : count < 1791 ? 11 : 12;
  const deferred = frame({ width: 5000, palette: 256, minimum: 8, codes: [[256, 9],
    ...Array.from({ length: 5000 }, (_, i) => [i % 256, literalWidth(i)]), [257, 12]] });
  for (const gif of [special, deferred, frame({ palette: 0, localPalette: 2, codes: validCodes }),
    frame({ codes: [[0, 3], [5, 3]] })]) {
    const result = await pack(gif);
    assert.deepEqual(result.warnings, []);
    assert.deepEqual(media(result.bytes)[0][1], gif);
  }
  for (const [gif, reason] of [
    [frame({ codes: [[4, 3], [6, 3], [5, 3]] }), /invalid dictionary code/],
    [frame({ codes: [[4, 3], [0, 3]] }), /missing end code/],
    [frame({ codes: [[4, 3], [0, 3], [1, 3], [5, 3]] }), /exceeds frame dimensions/],
    [frame({ width: 2, codes: validCodes }), /does not match frame dimensions/],
    [frame({ codes: [[4, 3], [2, 3], [5, 3]] }), /palette index/],
    [frame({ palette: 0, codes: validCodes }), /no local or global color table/],
    [frame({ codes: validCodes, control: [0x21, 0xf9, 4, 1, 0, 0, 2, 0] }), /transparent index/],
    [frame({ codes: validCodes, extra: [0] }), /bytes after end code/],
    [frame({ width: 65535, height: 65535, codes: validCodes }), /128 MiB/],
  ]) {
    const warningOnly = await pack(gif, { gifDiagnostics: "warn" });
    assert.deepEqual(media(warningOnly.bytes)[0][1], gif);
    const result = await pack(gif);
    assert.equal(media(result.bytes).length, 0);
    assert.deepEqual(result.warnings.map(w => w.code), ["asset-read-failed", "image-omitted", "image-omitted"]);
    assert.match(result.warnings[0].message, reason);
  }
});

test("JPEG structure mode preserves standalone baseline, multiscan and progressive images", async () => {
  const validation = { jpegValidation: "structure" };
  for (const name of ["exif-orientation6", "jpeg-icc-srgb", "jpeg-sequential-multiscan", "jpeg-progressive-components",
    "jpeg-12bit-sequential", "jpeg-12bit-progressive"]) {
    const bytes = await readFile(new URL(`../../../tests/fixtures/${name}.jpg`, import.meta.url));
    const p = deck({ type: "bytes", value: bytes });
    assert.deepEqual(await generatePptx(p, validation), await generatePptx(p), name);
  }
  await assert.rejects(generatePptx(deck(), { jpegValidation: "decode" }), /jpegValidation/);
});

test("JPEG structure mode omits truncated or inconsistent markers without changing the default", async () => {
  const original = await readFile(new URL("../../../tests/fixtures/exif-orientation6.jpg", import.meta.url));
  const segment = marker => {
    const offset = original.indexOf(Buffer.from([0xff, marker]), 2);
    assert.ok(offset > 0);
    return offset;
  };
  const sof = segment(0xc0), sos = segment(0xda), dqt = segment(0xdb), dht = segment(0xc4);
  const change = (offset, value) => { const bytes = Buffer.from(original); bytes[offset] = value; return bytes; };
  const beforeScan = original.subarray(0, sos);
  const scanHeader = original.subarray(sos, sos + 2 + original.readUInt16BE(sos + 2));
  // Marker-framing control only, not intended to be an entropy-decodable image.
  const restart = (number, dri = true) => Buffer.concat([beforeScan,
    dri ? Buffer.from([0xff, 0xdd, 0, 4, 0, 1]) : Buffer.alloc(0), scanHeader,
    Buffer.from([0x41, 0xff, number, 0xff, 0, 0x42, 0xff, 0xd9])]);
  const pack = (bytes, options) => generatePptx(deck({ type: "bytes", value: bytes }), options);
  const validation = { jpegValidation: "structure" };
  assert.deepEqual(media((await pack(restart(0xd0), validation)).bytes)[0][1], restart(0xd0));
  const progressive = await readFile(new URL("../../../tests/fixtures/jpeg-progressive-components.jpg", import.meta.url));
  const firstScan = progressive.indexOf(Buffer.from([0xff, 0xda]));
  const firstHeader = progressive.subarray(firstScan, firstScan + 2 + progressive.readUInt16BE(firstScan + 2));
  const repeatedScan = Buffer.concat([progressive.subarray(0, -2), firstHeader, Buffer.from([0x11, 0xff, 0xd9])]);
  for (const [bytes, reason] of [
    [original.subarray(0, -2), /missing EOI/],
    [Buffer.concat([original, Buffer.of(0)]), /bytes after EOI/],
    [change(sof + 6, 0), /invalid or duplicate frame/],
    [change(sof + 1, 0xc3), /outside the supported profile/],
    [change(dqt + 5, 0), /zero quantization value/],
    [change(dht + 5, 2), /Huffman table/],
    [change(sos + 5, 0xff), /unknown scan component/],
    [restart(0xd1), /restart marker sequence/],
    [restart(0xd0, false), /restart marker sequence/],
    [repeatedScan, /out-of-order coefficient scan/],
  ]) {
    assert.equal(media((await pack(bytes)).bytes).length, 1);
    const result = await pack(bytes, validation);
    assert.equal(media(result.bytes).length, 0, String(reason));
    assert.deepEqual(result.warnings.map(w => w.code), ["asset-read-failed", "image-omitted", "image-omitted"]);
    assert.match(result.warnings[0].message, reason);
  }
});
