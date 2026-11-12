import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx, DEFAULT_DIAGRAM_LIMITS } from "../dist/index.js";
import { diagramDataFixture } from "./diagram-data-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const model = async f => (await inspectPptx(f.bytes())).diagramInventory.dataModels[0];
test("both dialects distinguish content, assistant, transition and presentation points with typed references", async () => {
  for (const strict of [false, true]) {
    const f = diagramDataFixture(strict), m = await model(f);
    assert.equal(m.status, "inspected"); assert.equal(m.coverage, "identities-references-direct-text"); assert.deepEqual(m.findings, []);
    assert.deepEqual(m.points.map(p => p.kind), ["doc", "node", "asst", "parTrans", "sibTrans", "pres", "pres"]);
    assert.deepEqual(m.connections.map(c => c.kind), ["parOf", "presOf", "presOf"]);
    assert.equal(m.connections[0].source.targetElementIndex, m.points[0].elementIndex); assert.equal(m.connections[0].destination.targetElementIndex, m.points[1].elementIndex);
    assert.equal(m.connections[0].parentTransition.targetElementIndex, m.points[3].elementIndex); assert.equal(m.points[3].connection.targetElementIndex, m.connections[0].elementIndex);
    assert.deepEqual(m.points.slice(5).map(p => p.presentationAssociation.targetElementIndex), [m.points[1].elementIndex, m.points[1].elementIndex]);
    assert.equal(m.connections[0].sourceOrder, "2"); assert.equal(m.connections[0].sourceOrderValue, 2);
    assert.deepEqual(m.connections.slice(1).map(c => c.sourceOrderValue), [1, 0]); assert.equal(m.connections[1].presentationId, "urn:test:layout");
    assert.deepEqual(m.points[5].properties, [{ namespaceUri: "", local: "presAssocID", value: "1" }, { namespaceUri: "", local: "presStyleLbl", value: "node0" }, { namespaceUri: "", local: "custAng", value: "0" }]);
  }
});

test("source text retains paragraph/run/field/break identity, empty paragraphs and lexical levels", async () => {
  const f = diagramDataFixture(), p = (await model(f)).points[1].text;
  assert.equal(p[0].level, "02"); assert.deepEqual(p[0].runs.map(r => [r.kind, r.text]), [["text", "A & B 😀"], ["break", "\n"], ["field", "7"]]);
  assert.equal(p[0].runs[2].fieldId, "field-1"); assert.equal(p[0].runs[2].fieldType, "slidenum"); assert.deepEqual(p[1].runs, []);
  assert.ok(p[0].runs.every(r => r.elementIndex > p[0].elementIndex));
  f.setModel(f.point("0", "doc", f.richText.replace('<a:t>A &amp; B 😀</a:t>', '<a:t>bad<a:child/></a:t>')), "");
  const malformed = await model(f); assert.equal(malformed.status, "partial"); assert.deepEqual(malformed.points[0].text[0].runs.map(r => r.kind), ["break", "field"]);
});

test("numeric aliases are resolved by value; recognized GUIDs and canonical collisions stay separate", async () => {
  const f = diagramDataFixture(), guid = "{ABCDEF01-2345-6789-ABCD-EF0123456789}";
  f.setModel(f.point("+01") + f.point(guid), f.connection("-2", "1", guid));
  const m = await model(f); assert.equal(m.connections[0].source.status, "resolved"); assert.equal(m.connections[0].destination.status, "resolved"); assert.equal(m.points[0].modelId, "+01");
  f.setModel(f.point("1") + f.point("01"), f.connection("2", "1", "01"));
  const dup = await model(f); assert.ok(dup.points.every(p => p.identityStatus === "duplicate")); assert.equal(dup.connections[0].source.status, "ambiguous");
  f.setModel(f.point("1"), f.connection("1", "1", "1")); assert.equal((await model(f)).connections[0].identityStatus, "duplicate");
  f.setModel(f.point("bad") + f.point(guid.toLowerCase()), f.connection("3", "2147483648", guid.toLowerCase()));
  const unsupported = await model(f); assert.ok(unsupported.points.every(p => p.identityStatus === "unrecognized")); assert.equal(unsupported.connections[0].source.status, "unrecognized");
});

test("missing/ambiguous/wrong-kind endpoints, unknown enums and invalid ordering are reported without repairing the graph", async () => {
  const f = diagramDataFixture();
  f.setModel(f.point("0", "future", "", 'cxnId="0"') + '<d:pt/>' + f.point("1", "parTrans", "", 'cxnId="0"'), f.connection("10", "10", "999", 'parTransId="missing"', "future").replace('srcOrd="0"', 'srcOrd="-1"').replace(' destOrd="0"', ""));
  const m = await model(f); assert.equal(m.status, "partial"); assert.equal(m.points[0].kind, "unrecognized"); assert.equal(m.points[1].identityStatus, "missing");
  assert.equal(m.points[0].connection.status, "uninterpreted"); assert.equal(m.points[2].connection.status, "wrong-kind"); assert.equal(m.connections[0].source.status, "wrong-kind"); assert.equal(m.connections[0].destination.status, "missing");
  assert.equal(m.connections[0].sourceOrder, "-1"); assert.equal(m.connections[0].sourceOrderValue, undefined); assert.equal(m.connections[0].destinationOrderValue, undefined);
  assert.ok(m.findings.some(f => f.attribute === "srcOrd")); assert.ok(m.findings.some(f => f.attribute === "destOrd"));
});

test("data cycles and source order stay literal; the inspector does not infer a layout or repair a hierarchy", async () => {
  const f = diagramDataFixture(); f.setModel(f.point("0", "doc") + f.point("1", "node", "", 'cxnId="0"'), f.connection("2", "0", "1") + f.connection("3", "1", "0"));
  const m = await model(f); assert.equal(m.status, "inspected"); assert.equal(m.connections.length, 2); assert.equal(m.points[1].connection.status, "uninterpreted");
  assert.deepEqual(m.connections.map(c => [c.source.value, c.destination.value]), [["0", "1"], ["1", "0"]]);
});

test("source identity/order boundaries and direct-list scope do not fabricate nested or missing targets", async () => {
  const f = diagramDataFixture();
  f.setModel(f.point("-2147483648") + f.point("2147483647") + '<d:future>' + f.point("99") + '</d:future>',
    f.connection("-1", " -2147483648 ", "+2147483647").replace('srcOrd="0"', 'srcOrd="+4294967295"')
      .replace('destOrd="0"', 'destOrd="4294967296"') + '<d:cxn modelId="1" srcOrd="0" destOrd="0"/>');
  const m = await model(f); assert.equal(m.points.length, 2);
  assert.equal(m.connections[0].source.status, "resolved"); assert.equal(m.connections[0].destination.status, "resolved");
  assert.equal(m.connections[0].sourceOrderValue, 4294967295); assert.equal(m.connections[0].destinationOrderValue, undefined);
  assert.equal(m.connections[1].source.status, "absent"); assert.equal(m.connections[1].destination.status, "absent");
  assert.ok(m.findings.some(f => f.code === "uninterpreted-structure"));
  assert.ok(m.findings.some(f => f.attribute === "srcId")); assert.ok(m.findings.some(f => f.attribute === "destId"));
  for (const body of ['<d:ptLst/>', '<d:ptLst/><d:cxnLst/><d:cxnLst/>']) {
    f.setData(); f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:extLst>', body + '<d:extLst>');
    const result = await model(f);
    assert.equal(result.status, body === '<d:ptLst/>' ? "inspected" : "source-only");
    assert.deepEqual(result.connections, []);
  }
});

test("in-flight cancellation rejects without returning a partial graph and a fresh invocation can retry", async () => {
  const f = diagramDataFixture(); f.setModel(Array.from({ length: 3000 }, (_, i) => f.point(String(i))).join(""), "");
  const controller = new AbortController(), pending = inspectPptx(f.bytes(), { signal: controller.signal });
  const timer = setTimeout(() => controller.abort(), 10);
  try { await assert.rejects(pending, { code: "operation-aborted" }); } finally { clearTimeout(timer); }
  assert.equal((await model(f)).points.length, 3000);
});

test("diagram work/findings limits are cumulative, configurable, captured before await and never truncate success", async () => {
  const f = diagramDataFixture();
  assert.ok(Object.isFrozen(DEFAULT_DIAGRAM_LIMITS));
  await assert.rejects(inspectPptx(f.bytes(), { diagramLimits: { maxWorkUnits: 1 } }), e => e.code === "package-limit-exceeded" && e.entry === "/deck/diagrams/data.xml");
  const config = { diagramLimits: { maxWorkUnits: 100 } }, pending = inspectPptx(f.bytes(), config); config.diagramLimits.maxWorkUnits = 1;
  assert.equal((await pending).diagramInventory.dataModels[0].status, "inspected");
  f.setTree(f.frame() + f.frame("3", 'r:dm="dm2" r:lo="lo" r:cs="cs" r:qs="qs"'));
  f.parts["deck/diagrams/data2.xml"] = f.parts["deck/diagrams/data.xml"]; f.types["deck/diagrams/data2.xml"] = "drawingml.diagramData";
  f.parts["deck/_rels/z.xml.rels"] = relationships([...f.relations, f.rel("dm2", "diagramData", "diagrams/data2.xml")]);
  await assert.rejects(inspectPptx(f.bytes(), { diagramLimits: { maxWorkUnits: 100 } }), { code: "package-limit-exceeded" });
  const invalid = diagramDataFixture(); invalid.setModel('<d:pt/><d:pt/>', "");
  await assert.rejects(inspectPptx(invalid.bytes(), { diagramLimits: { maxFindings: 1 } }), { code: "package-limit-exceeded" });
  await assert.rejects(importPptx(invalid.bytes(), { projection: "structure", diagramLimits: { maxFindings: 1 } }), { code: "import-limit-exceeded" });
  for (const diagramLimits of [{ maxFindings: 0 }, { maxWorkUnits: NaN }, { unknown: 1 }]) await assert.rejects(inspectPptx(f.bytes(), { diagramLimits }));
});

test("MCE/missing/duplicate lists remain source-only and unsupported source children are located", async () => {
  for (const change of [
    f => { f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:ptLst>', '<d:ptLst xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'); },
    f => { f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace('<d:ptLst>', '<d:ptLst/><d:ptLst>'); },
    f => f.setData(),
  ]) { const f = diagramDataFixture(); change(f); const m = await model(f); assert.equal(m.status, "source-only"); assert.deepEqual(m.points, []); assert.deepEqual(m.connections, []); }
  const f = diagramDataFixture(); f.setModel('<d:unknown/>' + f.point("0", "node", '<d:future/>'), ""); const m = await model(f);
  assert.equal(m.status, "partial"); assert.equal(m.findings.filter(f => f.code === "uninterpreted-structure").length, 2);
});

test("shared data is read once; repeated cache model references stay distinct rather than duplicate data IDs", async () => {
  const f = diagramDataFixture(); f.setTree(Array.from({ length: 80 }, (_, i) => f.frame(String(i + 2))).join(""));
  f.parts["deck/diagrams/cache.xml"] = f.cacheXml.replace('modelId="node-1"', 'modelId="5"').replace('</dsp:spTree>', '<dsp:sp modelId="5"/></dsp:spTree>');
  const r = await inspectPptx(f.bytes()); assert.equal(r.diagramInventory.dataModels.length, 1); assert.equal(r.diagramInventory.dataModels[0].points.length, 7);
  const cache = r.diagramInventory.parts.find(p => p.part.endsWith("cache.xml")); assert.deepEqual(cache.modelReferences.map(m => m.modelId), ["5", "5"]); assert.equal(new Set(cache.modelReferences.map(m => m.elementIndex)).size, 2);
});

test("graph reports are detached, no-op retains source bytes and existing budgets/cancellation remain effective", async () => {
  const f = diagramDataFixture(), bytes = f.bytes(), m = await model(f); m.points[0].modelId = "changed"; assert.equal((await model(f)).points[0].modelId, "0");
  for (const projection of ["structure", "content"]) {
    const s = await importPptx(bytes, { projection });
    try { assert.equal(s.report.inspection.diagramInventory.dataModels[0].points.length, 7); const prep = await s.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready"); assert.deepEqual((await s.save(prep.plan)).bytes, new Uint8Array(bytes)); } finally { s.dispose(); }
  }
  await assert.rejects(inspectPptx(bytes, { xmlLimits: { maxNodes: 5 } }), { code: "xml-limit-exceeded" });
  await assert.rejects(inspectPptx(bytes, { signal: AbortSignal.abort() }), { code: "operation-aborted" });
});
