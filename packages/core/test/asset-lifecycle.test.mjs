import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation } from "../dist/index.js";
const source = (n = 1) => ({ type: "bytes", value: Uint8Array.of(n, 2, 3) });
const image = (id, assetId = "shared") => ({ type: "image", id, assetId, box: { x: 0, y: 0, width: 40, height: 20 } });
const group = (id, children) => ({ type: "group", id, children, box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 } });
function fixture() {
  const p = createPresentation();
  p.registerAsset({ id: "shared", kind: "image", source: source(), width: 8, height: 4 });
  p.addSlide({ id: "one" }).addElement(image("direct"));
  p.addSlide({ id: "two", hidden: true }).addElement(group("outer", [group("inner", [image("nested")])]));
  p.defineSlideLayout({ id: "unused-layout", name: "Reusable", elements: [group("layout-group", [image("layout-image")])] });
  return p;
}

test("reference inventory includes hidden slides, nested groups and unused layouts in stable order", () => {
  const p = fixture(), references = p.getAssetReferences("shared");
  assert.deepEqual(references, [
    { assetId: "shared", elementId: "direct", slideId: "one", path: "slides.0.elements.0.assetId" },
    { assetId: "shared", elementId: "nested", slideId: "two", path: "slides.1.elements.0.children.0.children.0.assetId" },
    { assetId: "shared", elementId: "layout-image", layoutId: "unused-layout", path: "layouts.0.elements.0.children.0.assetId" },
  ]);
  assert.ok(Object.isFrozen(references)); assert.ok(references.every(Object.isFrozen));
  p.removeSlide("one"); assert.equal(references.length, 3); assert.equal(p.getAssetReferences("shared").length, 2);
  assert.deepEqual(p.getAssetReferences("unknown"), []);
});

test("reference queries can locate dangling image uses without requiring registered assets", () => {
  const p = createPresentation(); p.addSlide().addElement(image("dangling", "missing"));
  assert.equal(p.getAssetReferences("missing")[0].elementId, "dangling");
  assert.throws(() => p.removeAsset("missing"), /Unknown asset/);
});

test("replacement preserves identity, ordering, references and existing immutable snapshots", () => {
  const p = fixture(), old = p.getAsset("shared"), oldAssets = p.assets, oldIR = normalizePresentation(p), before = p.getAssetReferences("shared");
  p.registerAsset({ id: "later", kind: "image", source: source() });
  const replacement = { kind: "image", source: source(4), width: 4, height: 8, accessibility: { description: "New image" } };
  const updated = p.replaceAsset("shared", replacement); replacement.source.value.fill(0); replacement.accessibility.description = "caller edit";
  assert.equal(updated.id, "shared"); assert.notEqual(updated, old); assert.equal(p.assets[0], updated);
  assert.equal(old.source.value, "AQID"); assert.equal(oldAssets[0], old); assert.equal(oldIR.assets[0].source.value, "AQID");
  assert.equal(updated.source.value, "BAID"); assert.equal(updated.accessibility.description, "New image");
  assert.deepEqual(p.getAssetReferences("shared"), before); assert.deepEqual(validatePresentation(p), []);
  assert.throws(() => { updated.source.value = "AAAA"; }, TypeError);
  assert.deepEqual(p.assets.map(a => a.id), ["shared", "later"]);
});

test("replacement is full metadata replacement rather than implicit partial merge", () => {
  const p = fixture();
  const updated = p.replaceAsset("shared", { kind: "image", source: source(9) });
  assert.equal(updated.width, undefined); assert.equal(updated.height, undefined); assert.equal(updated.accessibility, undefined);
});

test("failed replacement leaves existing registry and references unchanged", () => {
  const p = fixture(), before = normalizePresentation(p), old = p.getAsset("shared");
  for (const replacement of [{ kind: "image", source: { type: "bytes", value: new Uint8Array() } }, { id: "changed", kind: "image", source: source() }, { kind: "image", source: source(), accessibility: { description: () => {} } }]) {
    assert.throws(() => p.replaceAsset("shared", replacement));
    assert.equal(p.getAsset("shared"), old); assert.deepEqual(normalizePresentation(p), before);
  }
  assert.throws(() => p.replaceAsset("unknown", { kind: "image", source: source() }), /Unknown asset/);
  assert.ok(p.replaceAsset("shared", { kind: "image", source: source(8) }));
});

test("replacement does not overwrite a reentrant replacement and tolerates unrelated index movement", () => {
  const p = createPresentation();
  p.registerAsset({ id: "first", kind: "image", source: source() });
  p.registerAsset({ id: "target", kind: "image", source: source() });
  const updated = p.replaceAsset("target", { kind: "image", get source() { if (p.getAsset("first")) p.removeAsset("first"); return source(4); } });
  assert.equal(p.assets.length, 1); assert.equal(p.assets[0], updated);
  assert.throws(() => p.replaceAsset("target", { kind: "image", get source() { p.replaceAsset("target", { kind: "image", source: source(7) }); return source(9); } }), /changed while capturing/);
  assert.equal(p.getAsset("target").source.value, "BwID");
});

test("referenced removal rejects without changing registry even for unused layout definitions", () => {
  const p = fixture(), snapshot = p.assets;
  assert.throws(() => p.removeAsset("shared"), /slides.0.elements.0.assetId/);
  p.removeSlide("one"); p.removeSlide("two");
  assert.throws(() => p.removeAsset("shared"), /layouts.0.elements.0.children.0.assetId/);
  assert.deepEqual(p.assets, snapshot); assert.deepEqual(p.removeUnusedAssets(), []);
});

test("last known use removal permits explicit asset deletion and later identity reuse", () => {
  const p = createPresentation(), a = p.registerAsset({ id: "shared", kind: "image", source: source() });
  const s = p.addSlide(); s.addElement(group("outer", [image("child")])); s.removeElement("outer");
  assert.equal(p.removeAsset("shared"), a); assert.equal(p.getAsset("shared"), undefined);
  assert.throws(() => p.removeAsset("shared"), /Unknown asset/);
  const b = p.registerAsset({ id: "shared", kind: "image", source: source(7) });
  assert.notEqual(a, b); assert.equal(a.source.value, "AQID"); assert.deepEqual(validatePresentation(p), []);
});

test("unused cleanup preserves known references and explicit external retention, returning immutable ordered removals", () => {
  const p = fixture();
  for (const id of ["spare-one", "external", "spare-two"]) p.registerAsset({ id, kind: "image", source: source() });
  const before = p.assets;
  assert.throws(() => p.removeAsset("external", { retainAssetIds: ["external"] }), /explicitly retained/);
  const removed = p.removeUnusedAssets({ retainAssetIds: ["external", "external"] });
  assert.deepEqual(removed.map(a => a.id), ["spare-one", "spare-two"]); assert.ok(Object.isFrozen(removed));
  assert.deepEqual(p.assets.map(a => a.id), ["shared", "external"]); assert.equal(before.length, 4);
  assert.deepEqual(p.removeUnusedAssets({ retainAssetIds: ["external"] }), []);
  assert.deepEqual(p.removeUnusedAssets().map(a => a.id), ["external"]);
  assert.deepEqual(validatePresentation(p), []);
});

test("unknown retention IDs and invalid options fail cleanup atomically", () => {
  const p = fixture(); p.registerAsset({ id: "spare", kind: "image", source: source() }); const before = p.assets;
  for (const options of [null, [], { retainAssetIDs: ["spare"] }, { retainAssetIds: ["missing"] }, { retainAssetIds: "spare" }, { retainAssetIds: [null] }]) {
    assert.throws(() => p.removeUnusedAssets(options)); assert.deepEqual(p.assets, before);
    assert.throws(() => p.removeAsset("spare", options)); assert.deepEqual(p.assets, before);
  }
});

test("content deduplication is explicit, byte-only, metadata-aware and independent of property order", () => {
  const p = createPresentation();
  const a = p.registerAsset({ kind: "image", source: source(), width: 8, height: 4, accessibility: { decorative: false, description: "Same" } });
  const b = p.registerAsset({ height: 4, width: 8, source: { type: "bytes", encoding: "base64", value: "AQID" }, accessibility: { description: "Same", decorative: false }, kind: "image" }, { deduplicate: "content" });
  assert.equal(a, b);
  assert.notEqual(p.registerAsset({ kind: "image", source: source(), width: 8, height: 4, accessibility: { decorative: false, description: "Same" } }), a);
  assert.notEqual(p.registerAsset({ kind: "image", source: source(), width: 4, height: 8 }, { deduplicate: "content" }), a);
  const explicit = p.registerAsset({ id: "explicit", kind: "image", source: source(), width: 8, height: 4, accessibility: { decorative: false, description: "Same" } }, { deduplicate: "content" });
  assert.equal(explicit.id, "explicit"); assert.notEqual(explicit, a);
  assert.throws(() => p.registerAsset({ kind: "image", source: source() }, { deduplicate: "hash" }), /deduplication/);
});

test("content deduplication never fetches URL/path sources or conflates metadata", () => {
  const p = createPresentation();
  for (const type of ["url", "path"]) {
    const input = { kind: "image", source: { type, value: "same" } };
    assert.notEqual(p.registerAsset(input, { deduplicate: "content" }), p.registerAsset(input, { deduplicate: "content" }));
  }
  for (const extra of [{ mimeType: "image/png" }, { mimeType: "image/jpeg" }, { dedupeKey: "x" }, { accessibility: { description: "a" } }, { accessibility: { description: "b" } }]) p.registerAsset({ kind: "image", source: source(), ...extra }, { deduplicate: "content" });
  assert.equal(p.assets.length, 9);
});

test("replacement/removal update future content matches without stale indexes", () => {
  const p = createPresentation(), a = p.registerAsset({ kind: "image", source: source() });
  p.replaceAsset(a.id, { kind: "image", source: source(2) });
  const b = p.registerAsset({ kind: "image", source: source() }, { deduplicate: "content" });
  assert.notEqual(b.id, a.id); p.removeAsset(b.id);
  const c = p.registerAsset({ kind: "image", source: source() }, { deduplicate: "content" });
  assert.notEqual(c.id, b.id); assert.equal(p.assets.length, 2);
});

test("re-registering equivalent known metadata ignores object property insertion order", () => {
  const p = createPresentation();
  const a = p.registerAsset({ id: "stable", kind: "image", source: { type: "url", value: "image" }, accessibility: { description: "a", decorative: false } });
  assert.equal(p.registerAsset({ accessibility: { decorative: false, description: "a" }, source: { value: "image", type: "url" }, kind: "image", id: "stable" }), a);
});


test("registration and replacement evaluate source getters once per call", () => {
  const p = createPresentation(); let reads = 0;
  const input = { kind: "image", get source() { reads++; return source(); } };
  const a = p.registerAsset(input); assert.equal(reads, 1);
  p.replaceAsset(a.id, input); assert.equal(reads, 2);
});


test("large alternating cleanup keeps registry order and remains stable across repeated passes", () => {
  const p = createPresentation(), retainAssetIds = [];
  for (let i = 0; i < 2000; i++) {
    const id = `asset-${i}`; p.registerAsset({ id, kind: "image", source: source() });
    if (i % 2 === 0) retainAssetIds.push(id);
  }
  const removed = p.removeUnusedAssets({ retainAssetIds });
  assert.equal(removed.length, 1000); assert.equal(removed[0].id, "asset-1"); assert.equal(removed.at(-1).id, "asset-1999");
  assert.deepEqual(p.assets.map(a => a.id), retainAssetIds);
  assert.deepEqual(p.removeUnusedAssets({ retainAssetIds }), []);
  assert.equal(p.removeUnusedAssets().length, 1000); assert.equal(p.assets.length, 0);
});
