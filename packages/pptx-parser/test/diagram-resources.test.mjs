import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx } from "../dist/index.js";
import { diagramResourcesFixture, DSP, CACHE_REL, CACHE_TYPE } from "./diagram-resources-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const inspect = f => inspectPptx(f.bytes());
const cacheOf = r => r.slides[0].diagrams[0].cache;

test("both dialects resolve the data extension against the owning SLIDE and retain raw reference provenance", async () => {
  for (const strict of [false, true]) {
    const f = diagramResourcesFixture(strict);
    f.parts["deck/diagrams/trap.xml"] = f.cacheXml; f.types["deck/diagrams/trap.xml"] = CACHE_TYPE;
    f.parts["deck/diagrams/_rels/data.xml.rels"] = relationships([{ Id: "cache", Type: CACHE_REL, Target: "trap.xml" }, f.rel("data-image", "image", "../media/shared.png")]);
    const r = await inspect(f), cache = cacheOf(r);
    assert.equal(cache.status, "located"); assert.equal(cache.part, "/deck/diagrams/cache.xml"); assert.equal(cache.dataPart, "/deck/diagrams/data.xml");
    assert.deepEqual(cache.references, [{ elementIndex: 4, relationshipId: "cache", minimumVersion: "urn:test:version" }]);
    assert.equal(cache.freshness, "unknown");
    assert.deepEqual(r.diagramInventory.cacheCandidates.map(c => [c.relationshipId, c.status, c.referencedBy]), [["cache", "located", [r.slides[0].diagrams[0].elementIndex]]]);
    assert.ok(!r.diagramInventory.parts.some(p => p.part.endsWith("trap.xml")));
  }
});

test("multiple diagrams, two pages and outside incoming owners form a shared reference graph without copied payloads", async () => {
  const f = diagramResourcesFixture(); f.setTree(f.frame("2") + f.frame("3"));
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"]; f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/z.xml.rels"];
  f.parts["deck/outside.xml"] = '<outside/>'; f.types["deck/outside.xml"] = "application/xml";
  f.parts["deck/_rels/outside.xml.rels"] = relationships([{ Id: "outside", Type: "urn:unknown-owner", Target: "diagrams/cache.xml" }]);
  const r = await inspect(f), inventory = r.diagramInventory;
  assert.equal(inventory.coverage, "direct-source-associations");
  const cache = inventory.parts.find(p => p.part === "/deck/diagrams/cache.xml"), data = inventory.parts.find(p => p.part === "/deck/diagrams/data.xml");
  assert.equal(cache.owners.length, 4); assert.equal(data.owners.length, 4);
  assert.equal(new Set(cache.owners.map(o => `${o.slidePart}:${o.elementIndex}`)).size, 4);
  assert.equal(cache.incomingRelationships.length, 3); assert.ok(cache.incomingRelationships.some(e => e.source === "/deck/outside.xml"));
  const image = inventory.parts.find(p => p.part === "/deck/media/shared.png");
  assert.deepEqual(image.roles, ["image"]); assert.deepEqual(image.owners, []); assert.equal(image.incomingRelationships.length, 3);
  assert.deepEqual(cache.imageRelationships.map(i => [i.relationshipId, i.status, i.references.length]), [["cache-image", "located", 1], ["unreferenced-image", "located", 0]]);
  assert.equal(data.imageRelationships[0].references[0].attribute, "embed");
  assert.equal(inventory.parts.filter(p => p.part === image.part).length, 1);
});

test("unassigned caches remain candidates with zero inferred owners, even on pages with no diagrams", async () => {
  const f = diagramResourcesFixture();
  f.parts["deck/diagrams/other.xml"] = f.cacheXml; f.types["deck/diagrams/other.xml"] = CACHE_TYPE;
  f.relations.push({ Id: "other", Type: CACHE_REL, Target: "diagrams/other.xml" }); f.setRelations();
  const r = await inspect(f);
  assert.deepEqual(r.diagramInventory.cacheCandidates.map(c => c.referencedBy.length), [1, 0]);
  assert.equal(r.diagramInventory.parts.find(p => p.part.endsWith("other.xml")).owners.length, 0);
  f.setTree(""); const empty = await inspect(f);
  assert.equal(empty.slides[0].diagrams, undefined); assert.deepEqual(empty.diagramInventory.cacheCandidates.map(c => c.referencedBy), [[], []]);
});

test("a shared data part resolves its same relationship ID separately for each owning slide", async () => {
  for (const strict of [false, true]) {
    const f = diagramResourcesFixture(strict);
    f.parts["deck/a.xml"] = f.parts["deck/z.xml"];
    f.parts["deck/diagrams/second.xml"] = f.cacheXml; f.types["deck/diagrams/second.xml"] = CACHE_TYPE;
    f.parts["deck/_rels/a.xml.rels"] = relationships(f.relations.map(r => r.Id === "cache" ? { ...r, Target: "diagrams/second.xml" } : r));
    const r = await inspect(f);
    assert.deepEqual(r.slides.map(s => s.diagrams[0].cache.part), ["/deck/diagrams/cache.xml", "/deck/diagrams/second.xml"]);
    for (const part of ["cache.xml", "second.xml"]) assert.equal(r.diagramInventory.parts.find(p => p.part.endsWith(part)).owners.length, 1);
    assert.equal(r.diagramInventory.parts.find(p => p.part.endsWith("data.xml")).owners.length, 2);
  }
});

test("shared resources stay deduplicated as diagrams and raw image references grow", async () => {
  const f = diagramResourcesFixture();
  f.setTree(Array.from({ length: 160 }, (_, i) => f.frame(String(i + 2))).join(""));
  f.parts["deck/diagrams/cache.xml"] = f.cacheXml.replace('</dsp:drawing>', `${'<a:blip r:embed="cache-image"/>'.repeat(240)}</dsp:drawing>`);
  const r = await inspect(f), inventory = r.diagramInventory;
  assert.equal(inventory.parts.length, 6); // Four components, one cache, one shared image.
  assert.equal(inventory.parts.find(p => p.part.endsWith("cache.xml")).owners.length, 160);
  assert.equal(inventory.parts.find(p => p.part.endsWith("cache.xml")).imageRelationships[0].references.length, 241);
  assert.equal(inventory.cacheCandidates[0].referencedBy.length, 160);
  assert.equal(inventory.parts.find(p => p.part.endsWith("shared.png")).incomingRelationships.length, 3);
});

test("cache reference failures distinguish wrong role, external, fragment, content type, root and absent declaration", async () => {
  for (const strict of [false, true]) for (const [expected, change] of [
    ["missing-relationship", f => { f.setData('<dsp:dataModelExt relId="unknown"/>'); }],
    ["wrong-relationship-type", f => { f.relations.at(-1).Type = `${f.R}/image`; }],
    ["external", f => { Object.assign(f.relations.at(-1), { Target: "https://invalid.example/cache", TargetMode: "External" }); }],
    ["fragment-target", f => { f.relations.at(-1).Target += "#fragment"; }],
    ["wrong-content-type", f => { f.types["deck/diagrams/cache.xml"] = "drawingml.diagramData"; }],
    ["wrong-root", f => { f.parts["deck/diagrams/cache.xml"] = `<drawing xmlns="urn:wrong"/>`; }],
  ]) {
    const f = diagramResourcesFixture(strict); change(f); f.setRelations(); const r = await inspect(f);
    assert.equal(cacheOf(r).status, expected); assert.equal(cacheOf(r).freshness, "unknown");
    assert.ok(r.diagramInventory.cacheCandidates.every(c => c.referencedBy.length === 0));
  }
});

test("missing, ambiguous, opaque and MCE data references never borrow a same-page candidate", async () => {
  for (const [status, change] of [
    ["missing-reference", f => f.setData("")],
    ["missing-reference", f => f.setData('<dsp:dataModelExt/>')],
    ["missing-reference", f => f.setData('<dsp:dataModelExt r:relId="cache"/>')],
    ["missing-reference", f => f.setData('<dsp:dataModelExt relId="cache"/>', "urn:unknown-extension")],
    ["ambiguous-reference", f => f.setData('<dsp:dataModelExt relId="cache"/><dsp:dataModelExt relId="cache"/>')],
    ["source-only", f => { f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:extLst>', '<d:extLst xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'); }],
    ["unresolved-data", f => { f.setTree(f.frame().replace('r:dm="dm"', 'r:dm="absent"')); }],
  ]) {
    const f = diagramResourcesFixture(); change(f); const r = await inspect(f);
    assert.equal(cacheOf(r).status, status); assert.deepEqual(r.diagramInventory.cacheCandidates[0].referencedBy, []);
  }
  const nested = diagramResourcesFixture(); nested.setData('<dsp:unknown><dsp:dataModelExt relId="cache"/></dsp:unknown>');
  assert.equal(cacheOf(await inspect(nested)).status, "missing-reference");
});

test("raw blip reports retain unknown/inactive branches and classify missing, wrong, external and invalid image targets", async () => {
  const f = diagramResourcesFixture();
  f.parts["deck/diagrams/_rels/cache.xml.rels"] = relationships([
    f.rel("cache-image", "image", "../media/shared.png"), { ...f.rel("external", "image", "https://invalid.example/image"), TargetMode: "External" },
    f.rel("wrong-role", "hyperlink", "../media/shared.png"), f.rel("wrong-type", "image", "data.xml"), f.rel("fragment", "image", "../media/shared.png#x"),
  ]);
  f.parts["deck/diagrams/cache.xml"] = f.cacheXml.replace('</dsp:drawing>', '<a:extLst><a:ext uri="urn:opaque"><a:blip r:embed="missing" r:link="external"/><a:blip r:embed="wrong-role"/><a:blip r:embed="wrong-type"/><a:blip r:embed="fragment"/></a:ext></a:extLst></dsp:drawing>');
  const cache = (await inspect(f)).diagramInventory.parts.find(p => p.part.endsWith("cache.xml"));
  assert.deepEqual(Object.fromEntries(cache.imageRelationships.map(i => [i.relationshipId, i.status])), { "cache-image": "located", external: "external", "wrong-type": "wrong-content-type", fragment: "fragment-target", missing: "missing-relationship", "wrong-role": "wrong-relationship-type" });
  assert.equal(cache.imageRelationships.find(i => i.relationshipId === "external").references[0].attribute, "link");
  assert.ok(cache.imageRelationships.find(i => i.relationshipId === "missing").references[0].elementIndex > 1);
});

test("resource output is detached, bounded and no-op import retains all raw source associations", async () => {
  for (const projection of ["structure", "content"]) {
    const f = diagramResourcesFixture(), bytes = f.bytes(), session = await importPptx(bytes, { projection });
    try {
      assert.equal(session.report.inspection.slides[0].diagrams[0].cache.status, "located");
      const prep = await session.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready");
      assert.deepEqual((await session.save(prep.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
  const f = diagramResourcesFixture(), r = await inspect(f);
  r.diagramInventory.parts[0].owners.push({ slidePart: "/fake", elementIndex: 1, role: "data" });
  assert.ok(!(await inspect(f)).diagramInventory.parts[0].owners.some(o => o.slidePart === "/fake"));
  await assert.rejects(inspectPptx(f.bytes(), { limits: { maxRelationships: 1 } }), { code: "package-limit-exceeded" });
  await assert.rejects(inspectPptx(f.bytes(), { signal: AbortSignal.abort() }), { code: "operation-aborted" });
});
