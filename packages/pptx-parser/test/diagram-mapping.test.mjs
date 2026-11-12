import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx } from "../dist/index.js";
import { diagramDataFixture } from "./diagram-data-fixture.mjs";
import { DSP, CACHE_TYPE, CACHE_REL } from "./diagram-resources-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const inspect = async f => (await inspectPptx(f.bytes())).diagramInventory;
const setShapes = (f, xml) => { f.parts["deck/diagrams/cache.xml"] = `<dsp:drawing xmlns:dsp="${DSP}"><dsp:spTree>${xml}</dsp:spTree></dsp:drawing>`; };
const sp = id => `<dsp:sp${id === undefined ? "" : ` modelId="${id}"`}/>`;

test("both dialects expose normalized content/display/cache paths and retain repeated visual references", async () => {
  for (const strict of [false, true]) {
    const f = diagramDataFixture(strict); setShapes(f, sp("6") + sp("+05") + sp("5"));
    const inv = await inspect(f), m = inv.dataModels[0], cache = inv.cacheMappings[0];
    assert.equal(inv.cacheMappings.length, 1); assert.equal(cache.coverage, "explicit-model-references"); assert.equal(cache.status, "inspected");
    assert.deepEqual(cache.shapes.map(s => s.model.targetElementIndex), [m.points[6].elementIndex, m.points[5].elementIndex, m.points[5].elementIndex]);
    assert.equal(cache.shapes[1].model.value, "+05"); assert.deepEqual(cache.shapes.map(s => s.elementIndex), [3, 4, 5]);
    for (const mapping of m.presentationMappings) {
      assert.equal(mapping.agreement, "same-target"); assert.equal(mapping.targetElementIndex, m.points[1].elementIndex);
      assert.deepEqual(mapping.sources.map(s => s.kind), ["presAssocID", "presOf"]);
      assert.ok(mapping.sources.every(s => s.reference.targetElementIndex === m.points[1].elementIndex));
    }
  }
});

test("property/connection evidence keeps differing and unresolved targets without selecting a text owner", async () => {
  const f = diagramDataFixture();
  f.setModel(f.point("1") + f.point("2") + f.point("5", "pres", '<d:prSet presAssocID="1"/>') + f.point("6", "pres") + f.point("7", "pres") + f.point("8", "pres", '<d:prSet presAssocID="99"/>'),
    f.connection("10", "2", "5", "", "presOf") + f.connection("11", "1", "6", "", "presOf") + f.connection("12", "1", "6", "", "presOf")
      + f.connection("13", "1", "1", "", "presOf") + f.connection("14", "1", "99", "", "presOf"));
  const m = (await inspect(f)).dataModels[0];
  assert.deepEqual(m.presentationMappings.map(p => p.agreement), ["different-targets", "same-target", "none", "unresolved"]);
  assert.equal(m.presentationMappings[0].targetElementIndex, undefined);
  assert.deepEqual(m.presentationMappings[0].sources.map(s => s.reference.value), ["1", "2"]);
  assert.equal(m.presentationMappings[1].sources.length, 2); assert.equal(m.presentationMappings[1].targetElementIndex, m.points[0].elementIndex);
  assert.equal(m.presentationMappings[2].sources.length, 0); assert.equal(m.presentationMappings[3].sources[0].reference.status, "missing");
  assert.equal(m.connections.length, 5); // Nonpresentation/unresolved destinations stay in the source graph only.
});

test("cache model references distinguish absent, unknown, missing, wrong-kind and ambiguous IDs", async () => {
  const f = diagramDataFixture(); f.setModel(f.points + f.point("05"), f.connections);
  setShapes(f, [undefined, "bad", "99", "10", "5", "6"].map(sp).join(""));
  const inv = await inspect(f), mapping = inv.cacheMappings[0];
  assert.equal(mapping.status, "partial"); assert.deepEqual(mapping.shapes.map(s => s.model.status), ["absent", "unrecognized", "missing", "wrong-kind", "ambiguous", "resolved"]);
  assert.equal(mapping.findings.length, 5); assert.ok(mapping.findings.every(f => f.attribute === "modelId"));
  assert.ok(mapping.shapes.slice(0, 5).every(s => s.model.targetElementIndex === undefined));
  f.setModel(f.points, f.connections + f.connection("5", "0", "1"));
  assert.equal((await inspect(f)).cacheMappings[0].shapes[4].model.status, "ambiguous");
});

test("only recognized nested shape-tree paths map; opaque raw shapes remain inventory without invented usage", async () => {
  const f = diagramDataFixture();
  setShapes(f, sp("5") + `<dsp:grpSp>${sp("6")}<dsp:grpSp>${sp("1")}</dsp:grpSp></dsp:grpSp><dsp:extLst>${sp("99")}</dsp:extLst>`);
  const inv = await inspect(f);
  assert.deepEqual(inv.cacheMappings[0].shapes.map(s => s.model.value), ["5", "6", "1"]);
  assert.equal(inv.parts.find(p => p.part.endsWith("cache.xml")).modelReferences.length, 4);
  assert.equal(inv.cacheMappings[0].findings.length, 0);
});

test("MCE, ambiguous trees and source-only data do not become successful empty inferred mappings", async () => {
  for (const change of [
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('<dsp:spTree>', '<dsp:spTree xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'); },
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('<dsp:spTree>', '<dsp:spTree/><dsp:spTree>'); },
    f => { f.parts["deck/diagrams/cache.xml"] = `<dsp:drawing xmlns:dsp="${DSP}"/>`; },
    f => f.setData(),
  ]) {
    const f = diagramDataFixture(); change(f); const mapping = (await inspect(f)).cacheMappings[0];
    assert.equal(mapping.status, "source-only"); assert.deepEqual(mapping.shapes, []); assert.equal(mapping.findings[0].code, "source-only");
  }
  const f = diagramDataFixture(); setShapes(f, ""); const empty = (await inspect(f)).cacheMappings[0];
  assert.equal(empty.status, "inspected"); assert.deepEqual(empty.shapes, []); // An explicit empty tree is source evidence only.
});

test("shared data/cache pairs deduplicate across owners while different slide-local caches stay separate", async () => {
  const f = diagramDataFixture(); f.setTree(f.frame("2") + f.frame("3"));
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"]; f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/z.xml.rels"];
  let inv = await inspect(f); assert.equal(inv.cacheMappings.length, 1); assert.equal(inv.dataModels.length, 1);
  f.parts["deck/diagrams/second.xml"] = f.parts["deck/diagrams/cache.xml"]; f.types["deck/diagrams/second.xml"] = CACHE_TYPE;
  f.parts["deck/_rels/a.xml.rels"] = relationships(f.relations.map(r => r.Id === "cache" ? { ...r, Target: "diagrams/second.xml" } : r));
  inv = await inspect(f); assert.equal(inv.cacheMappings.length, 2); assert.equal(inv.dataModels.length, 1);
  assert.deepEqual(new Set(inv.cacheMappings.map(m => m.cachePart)), new Set(["/deck/diagrams/cache.xml", "/deck/diagrams/second.xml"]));
  f.parts["deck/diagrams/data2.xml"] = f.parts["deck/diagrams/data.xml"].replace('modelId="5"', 'modelId="55"'); f.types["deck/diagrams/data2.xml"] = "drawingml.diagramData";
  f.parts["deck/_rels/a.xml.rels"] = relationships(f.relations.map(r => r.Id === "dm" ? { ...r, Target: "diagrams/data2.xml" } : r));
  inv = await inspect(f); assert.equal(inv.cacheMappings.length, 2);
  assert.equal(inv.cacheMappings.find(m => m.dataPart.endsWith("data2.xml")).shapes[0].model.status, "missing");
  assert.equal(inv.cacheMappings.find(m => m.dataPart.endsWith("data.xml")).shapes[0].model.status, "resolved");
});

test("unassigned or missing caches never form a Cartesian mapping with data resources", async () => {
  const f = diagramDataFixture();
  f.parts["deck/diagrams/unassigned.xml"] = f.parts["deck/diagrams/cache.xml"]; f.types["deck/diagrams/unassigned.xml"] = CACHE_TYPE;
  f.relations.push({ Id: "unused", Type: CACHE_REL, Target: "diagrams/unassigned.xml" }); f.setRelations();
  const inv = await inspect(f); assert.equal(inv.cacheMappings.length, 1); assert.equal(inv.cacheCandidates.length, 2);
  f.setData(""); assert.equal((await inspect(f)).cacheMappings, undefined);
});

test("mapping budgets aggregate failures and shared uses; results stay detached and no-op preserves bytes", async () => {
  const f = diagramDataFixture(); setShapes(f, sp("99") + sp("98"));
  await assert.rejects(inspectPptx(f.bytes(), { diagramLimits: { maxFindings: 1 } }), e => e.code === "package-limit-exceeded" && e.entry === "/deck/diagrams/cache.xml");
  setShapes(f, sp("5"));
  await inspectPptx(f.bytes(), { diagramLimits: { maxWorkUnits: 120 } });
  f.setTree(Array.from({ length: 200 }, (_, i) => f.frame(String(i + 2))).join(""));
  await assert.rejects(inspectPptx(f.bytes(), { diagramLimits: { maxWorkUnits: 120 } }), { code: "package-limit-exceeded" });
  const inv = await inspect(f); assert.equal(inv.cacheMappings.length, 1);
  inv.cacheMappings[0].shapes[0].model.status = "changed";
  assert.equal((await inspect(f)).cacheMappings[0].shapes[0].model.status, "resolved");
  const bytes = f.bytes(), session = await importPptx(bytes, { projection: "structure" });
  try { const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes)); } finally { session.dispose(); }
});
