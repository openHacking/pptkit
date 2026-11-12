import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "../dist/index.js";

test("asset registration deduplicates immutable assets", () => {
  const presentation = createPresentation();
  const first = presentation.registerAsset({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
    accessibility: { description: "Hero" },
  });
  const second = presentation.registerAsset({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    dedupeKey: "hero",
    accessibility: { description: "Hero" },
  });
  assert.equal(first, second);
  assert.equal(presentation.assets.length, 1);
  assert.throws(() => { first.source.value = "changed"; }, TypeError);
});

test("asset registration rejects conflicting metadata", () => {
  const presentation = createPresentation();
  presentation.registerAsset({ kind: "image", id: "logo", source: { type: "path", value: "./logo.png" } });
  assert.throws(
    () => presentation.registerAsset({ kind: "image", id: "logo", source: { type: "path", value: "./other.png" } }),
    /already registered with different metadata/,
  );
});


test("byte sources snapshot offset Buffer inputs into immutable serializable canonical state", () => {
  const p = createPresentation(), buffer = Buffer.from([99, 0, 255, 128, 1, 99]);
  const a = p.registerAsset({ id: "bytes", kind: "image", source: { type: "bytes", value: buffer.subarray(1, 5) } });
  buffer.fill(0);
  assert.deepEqual(a.source, { type: "bytes", encoding: "base64", value: "AP+AAQ==" });
  assert.throws(() => { a.source.value = "AA=="; }, TypeError);
  const ir = normalizePresentation(p), restored = JSON.parse(JSON.stringify(ir));
  assert.deepEqual(restored.assets[0].source, a.source);
  ir.assets[0].source.value = "AA==";
  assert.equal(a.source.value, "AP+AAQ==");
});

test("raw and canonical byte sources deduplicate consistently without aliasing", () => {
  const p = createPresentation(), value = Uint8Array.of(1, 2, 3);
  const a = p.registerAsset({ kind: "image", dedupeKey: "bytes", source: { type: "bytes", value } });
  assert.equal(p.registerAsset({ kind: "image", dedupeKey: "bytes", source: { type: "bytes", encoding: "base64", value: "AQID" } }), a);
  assert.throws(() => p.registerAsset({ id: a.id, kind: "image", source: { type: "bytes", value: Uint8Array.of(2) } }), /different metadata/);
  assert.equal(p.assets.length, 1);
});

test("byte encoding handles chunk and padding boundaries without argument overflow", () => {
  for (const length of [1, 2, 3, 4, 24575, 24576, 24577, 49153, 100000]) {
    const bytes = Uint8Array.from({ length }, (_, i) => i % 256);
    const a = createPresentation().registerAsset({ kind: "image", source: { type: "bytes", value: bytes } });
    assert.equal(a.source.value, Buffer.from(bytes).toString("base64"));
  }
});

test("invalid, empty, shared and noncanonical byte inputs fail before registration", () => {
  const p = createPresentation();
  for (const value of ["", "A", "AAA", "A===", "AA==\n", "AB==", "AAB=", "_w==", "====", "!!!!"]) {
    assert.throws(() => p.registerAsset({ kind: "image", source: { type: "bytes", encoding: "base64", value } }), TypeError);
  }
  for (const value of [new Uint8Array(), [], new Uint16Array([1]), new Uint8Array(new SharedArrayBuffer(1)), null]) {
    assert.throws(() => p.registerAsset({ kind: "image", source: { type: "bytes", value } }), TypeError);
  }
  assert.throws(() => p.registerAsset({ kind: "image", source: { type: "bytes", value: "AA==" } }), TypeError);
  assert.equal(p.assets.length, 0);
  assert.ok(p.registerAsset({ kind: "image", source: { type: "bytes", value: Uint8Array.of(0) } }));
});

test("raw byte registration rejects disguised shared buffers and captures the value only once",()=>{
 class Hostile extends Uint8Array { get buffer(){return new ArrayBuffer(0)} *[Symbol.iterator](){throw Error('unexpected iterator')} }
 const p=createPresentation(),ordinary=new Hostile([1,2,3]);let reads=0;const source={type:'bytes',get value(){reads++;return ordinary}};
 const asset=p.registerAsset({id:'hostile',kind:'image',source});assert.equal(reads,1);assert.equal(asset.source.value,'AQID');ordinary.fill(0);assert.equal(asset.source.value,'AQID');
 if(typeof SharedArrayBuffer!=='undefined'){const shared=new Hostile(new SharedArrayBuffer(3));shared.set([1,2,3]);assert.throws(()=>p.registerAsset({id:'shared',kind:'image',source:{type:'bytes',value:shared}}),/SharedArrayBuffer/);assert.equal(p.assets.length,1);}
});

test("raw byte registration rejects shared buffers from another JavaScript realm",async()=>{
 const vm=await import('node:vm'),foreign=vm.runInNewContext('new SharedArrayBuffer(3)'),view=new Uint8Array(foreign);assert.equal(foreign instanceof SharedArrayBuffer,false);assert.throws(()=>createPresentation().registerAsset({kind:'image',source:{type:'bytes',value:view}}),/SharedArrayBuffer/);
});

test('atomic asset batch preserves order, content deduplication and immutable snapshots', () => {
  const p = createPresentation(), bytes = Uint8Array.of(1, 2, 3), previous = p.assets;
  const result = p.registerAssets([
    { asset: { kind: 'image', source: { type: 'bytes', value: bytes }, width: 3, height: 1 }, options: { deduplicate: 'content' } },
    { asset: { kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'AQID' }, width: 3, height: 1 }, options: { deduplicate: 'content' } },
    { asset: { id: 'explicit', kind: 'image', source: { type: 'bytes', value: bytes }, width: 3, height: 1 } },
  ]);
  assert.equal(result.length, 3); assert.equal(result[0], result[1]); assert.notEqual(result[0], result[2]);
  assert.deepEqual(result.map(a => a.id), ['asset-1', 'asset-1', 'explicit']); assert.deepEqual(p.assets.map(a => a.id), ['asset-1', 'explicit']);
  assert.equal(previous.length, 0); bytes.fill(0); assert.equal(result[0].source.value, 'AQID'); assert.ok(Object.isFrozen(result)); assert.ok(Object.isFrozen(result[0]));
  const empty = p.registerAssets([]); assert.deepEqual(empty, []); assert.ok(Object.isFrozen(empty));
});

test('asset batch failure leaves registry and allocator unchanged and can retry', () => {
  const p = createPresentation(), original = p.registerAsset({ id: 'fixed', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  const good = { asset: { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(2) } } };
  assert.throws(() => p.registerAssets([good, { asset: { id: 'fixed', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(3) } } }]), /different metadata/);
  assert.deepEqual(p.assets, [original]); assert.equal(p.registerAssets([good])[0].id, 'asset-1');
  for (const input of [null, {}, [null], [{ asset: good.asset, extra: true }], [{ asset: good.asset, options: { deduplicate: 'invalid' } }]]) {
    const before = p.assets; assert.throws(() => p.registerAssets(input), TypeError); assert.deepEqual(p.assets, before);
  }
});

test('asset batch detects reentrant mutations and does not overwrite them', () => {
  const p = createPresentation(); p.registerAsset({ id: 'old', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } });
  const source = { type: 'bytes', get value() { p.replaceAsset('old', { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(9) } }); return Uint8Array.of(2); } };
  assert.throws(() => p.registerAssets([{ asset: { kind: 'image', source } }]), /changed during batch/);
  assert.deepEqual(p.assets.map(a => a.id), ['old']); assert.equal(p.getAsset('old').source.value, 'CQ==');
  const reentrant = { get asset() { p.registerAssets([{ asset: { id: 'nested', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(5) } } }]); return { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(6) } }; } };
  assert.throws(() => p.registerAssets([reentrant]), /changed during batch/);
  assert.deepEqual(p.assets.map(a => a.id), ['old', 'nested']); assert.equal(p.registerAssets([{ asset: { kind: 'image', source: { type: 'bytes', value: Uint8Array.of(7) } } }])[0].id, 'asset-1');
});

test('single registration remains visible when its getter commits a nested batch', () => {
  const p = createPresentation();
  const input = { id: 'outer', kind: 'image', get source() {
    p.registerAssets([{ asset: { id: 'nested', kind: 'image', source: { type: 'bytes', value: Uint8Array.of(1) } } }]);
    return { type: 'bytes', value: Uint8Array.of(2) };
  } };
  const outer = p.registerAsset(input);
  assert.deepEqual(p.assets.map(a => a.id), ['nested', 'outer']); assert.equal(p.getAsset('outer'), outer);
});
