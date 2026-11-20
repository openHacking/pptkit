import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx, compareOpcPackages } from "../dist/index.js";
import { diagramTextFixture } from "./diagram-text-fixture.mjs";
import { CACHE_TYPE, CACHE_REL } from "./diagram-resources-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
const option = { diagramTextComparison: "direct-paragraphs" };
const inspect = async f => (await inspectPptx(f.bytes(), option)).diagramInventory;

test("explicit both-dialect observations retain exact paragraphs, soft breaks, empty text and raw source locations", async () => {
  for (const strict of [false, true]) {
    const f = diagramTextFixture(strict), original = await inspectPptx(f.bytes()); assert.equal(original.diagramInventory.textComparisons, undefined);
    const texts = ["Leading & < 😀\r", "Line\nbreak", "", "e\u0301"];
    f.setText(texts); const inv = await inspect(f), result = inv.textComparisons[0], shape = result.shapes[0];
    assert.equal(result.dataPart, "/deck/diagrams/data.xml"); assert.equal(result.cachePart, "/deck/diagrams/cache.xml"); assert.equal(result.freshness, "unknown");
    assert.equal(shape.status, "equal"); assert.deepEqual(shape.sourceParagraphs, texts); assert.deepEqual(shape.cacheParagraphs, texts); assert.equal(shape.nonemptyText, true);
    assert.equal(shape.cacheElementIndex, inv.cacheMappings[0].shapes[0].elementIndex); assert.equal(shape.presentationElementIndex, inv.dataModels[0].points[2].elementIndex); assert.equal(shape.sourcePointElementIndex, inv.dataModels[0].points[1].elementIndex);
    f.setText(["", " \t"]); const blank = (await inspect(f)).textComparisons[0].shapes[0]; assert.equal(blank.status, "equal"); assert.equal(blank.nonemptyText, false);
  }
});

test("data-only, cache-only and paragraph/Unicode changes are different observations without a freshness verdict", async () => {
  for (const [source, cache] of [[["Fresh data"], ["Old cache"]], [["Old data"], ["Fresh cache"]], [["A\nB"], ["A", "B"]], [["é"], ["e\u0301"]], [["a "], ["a"]]]) {
    const f = diagramTextFixture(); f.setText(source, cache); const result = (await inspect(f)).textComparisons[0]; assert.equal(result.shapes[0].status, "different"); assert.equal(result.freshness, "unknown"); assert.deepEqual(result.shapes[0].sourceParagraphs, source); assert.deepEqual(result.shapes[0].cacheParagraphs, cache);
  }
});

test("ambiguous/non-content associations, absent/opaque text and fields never become guessed equality", async () => {
  const changes = [
    f => f.setText(["A"], ["A"], { association: '<d:prSet presAssocID="3"/>', extraPoints: '<d:pt modelId="3"/>' }),
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('modelId="2"', 'modelId="0"'); },
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace(/<dsp:txBody>[\s\S]*?<\/dsp:txBody>/, ''); },
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('<a:r>', '<a:fld>').replace('</a:r>', '</a:fld>'); },
    f => { f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<a:r>', '<a:fld>').replace('</a:r>', '</a:fld>'); },
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('<a:r>', '<a:r>opaque'); },
    f => f.setText(["A"], ["A"], { cacheExtra: '<a:unknown/>' }),
    f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('</a:p>', '<a:pPr/><a:pPr/></a:p>'); },
  ];
  for (const change of changes) { const f = diagramTextFixture(); change(f); const s = (await inspect(f)).textComparisons[0].shapes[0]; assert.equal(s.status, "unavailable"); assert.ok(s.reason); assert.equal(s.sourceParagraphs, undefined); }
  const direct = diagramTextFixture(); direct.parts["deck/diagrams/cache.xml"] = direct.parts["deck/diagrams/cache.xml"].replace('modelId="2"', 'modelId="1"'); const s = (await inspect(direct)).textComparisons[0].shapes[0]; assert.equal(s.status, "equal"); assert.equal(s.presentationElementIndex, undefined);
});

test("non-XML whitespace outside cache text leaves does not produce guessed equality", async () => {
  for (const strict of [false, true]) for (const anchor of ["<dsp:txBody>", "<a:p>", "<a:r>"]) {
    const f = diagramTextFixture(strict);
    f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace(anchor, `${anchor}\u00a0`);
    const result = (await inspect(f)).textComparisons[0].shapes[0];
    assert.equal(result.status, "unavailable", anchor);
    assert.equal(result.sourceParagraphs, undefined);
  }
});

test("partial graphs and MCE stay source-only; cache empty/missing and unassigned candidates do not invent comparisons", async () => {
  for (const change of [f => f.setText(["A"], ["A"], { sourceExtra: '<a:unknown/>' }), f => { f.parts["deck/diagrams/cache.xml"] = f.parts["deck/diagrams/cache.xml"].replace('<dsp:spTree>', '<dsp:spTree xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'); }]) {
    const f = diagramTextFixture(); change(f); const r = (await inspect(f)).textComparisons[0]; assert.equal(r.status, "source-only"); assert.deepEqual(r.shapes, []);
  }
  const empty = diagramTextFixture(); empty.setCache(""); assert.deepEqual((await inspect(empty)).textComparisons[0].shapes, []);
  empty.setData(""); assert.equal((await inspect(empty)).textComparisons, undefined);
});

test("shared data/cache owners deduplicate while distinct pairs stay scoped", async () => {
  const f = diagramTextFixture(); f.setTree(f.previewFrame + f.previewFrame.replace('id="2"', 'id="3"'));
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"]; f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/z.xml.rels"];
  assert.equal((await inspect(f)).textComparisons.length, 1);
  f.parts["deck/diagrams/second.xml"] = f.parts["deck/diagrams/cache.xml"].replace('Source &amp; 😀', 'Other'); f.types["deck/diagrams/second.xml"] = CACHE_TYPE;
  f.parts["deck/_rels/a.xml.rels"] = relationships(f.relations.map(r => r.Id === "cache" ? { ...r, Target: "diagrams/second.xml" } : r));
  let inv = await inspect(f); assert.deepEqual(new Set(inv.textComparisons.map(p => p.shapes[0].status)), new Set(["equal", "different"]));
  f.parts["deck/diagrams/unassigned.xml"] = f.parts["deck/diagrams/cache.xml"]; f.types["deck/diagrams/unassigned.xml"] = CACHE_TYPE;
  f.relations.push({ Id: "unused", Type: CACHE_REL, Target: "diagrams/unassigned.xml" }); f.setRelations(); inv = await inspect(f); assert.equal(inv.textComparisons.length, 2);
});

test("character comparison and disagreement findings consume aggregate graph budgets with part provenance", async () => {
  const f = diagramTextFixture(); f.setText(["x".repeat(2000)], ["x".repeat(2000)]);
  await inspectPptx(f.bytes(), { diagramLimits: { maxWorkUnits: 1000 } });
  await assert.rejects(inspectPptx(f.bytes(), { ...option, diagramLimits: { maxWorkUnits: 1000 } }), e => e.code === "package-limit-exceeded" && ["/deck/diagrams/data.xml", "/deck/diagrams/cache.xml"].includes(e.entry));
  f.setText(["New"], ["Old"]); f.setCache(f.parts["deck/diagrams/cache.xml"].match(/<dsp:sp modelId=[\s\S]*?<\/dsp:sp>/)[0].repeat(3));
  await assert.rejects(inspectPptx(f.bytes(), { ...option, diagramLimits: { maxFindings: 2 } }), e => e.code === "package-limit-exceeded" && e.entry === "/deck/diagrams/cache.xml");
  assert.equal((await inspect(f)).textComparisons[0].shapes.length, 3);
});

test("option capture and supported unrelated saves preserve reports/payloads without new cache editing authority", async () => {
  const f = diagramTextFixture(), bytes = f.bytes(), config = { projection: "content", editSaveProfile: "text-runs-v3", diagramPreview: "cache", ...option };
  const pending = importPptx(bytes, config); config.diagramTextComparison = "invalid"; const session = await pending;
  try {
    const comparison = session.report.inspection.diagramInventory.textComparisons; assert.equal(comparison[0].shapes[0].status, "equal");
    const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "text"), control = structuredClone(slide.elements[index]); control.content[0].runs[0].text = "Unrelated control"; slide.removeElement(control.id); slide.insertElement(index, control);
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); const saved = (await session.save(prepared.plan)).bytes;
    assert.deepEqual((await compareOpcPackages(bytes, saved)).differences.map(d => d.part), ["/deck/z.xml"]);
    assert.deepEqual((await inspectPptx(saved, option)).diagramInventory.textComparisons, comparison);
  } finally { session.dispose(); }
  for (const value of [true, false, "cache", null]) { await assert.rejects(inspectPptx(bytes, { diagramTextComparison: value }), TypeError); await assert.rejects(importPptx(bytes, { projection: "structure", diagramTextComparison: value }), TypeError); }
  const controller = new AbortController(); controller.abort(); await assert.rejects(inspectPptx(bytes, { ...option, signal: controller.signal }), { code: "operation-aborted" });
});
