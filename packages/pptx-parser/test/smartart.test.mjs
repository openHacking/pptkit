import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, importPptx } from "../dist/index.js";
import { smartArtFixture } from "./smartart-fixture.mjs";

const first = async f => (await inspectPptx(f.bytes())).slides[0].diagrams;

test("both dialects inventory four roles by relationship, preserving uninterpreted layout identity", async () => {
  for (const strict of [false, true]) {
    const f = smartArtFixture(strict), bytes = f.bytes(), copy = new Uint8Array(bytes), result = await inspectPptx(bytes), [d] = result.slides[0].diagrams;
    assert.equal(d.identityStatus, "unique"); assert.equal(d.drawingId, "2"); assert.equal(d.name, "Diagram 2"); assert.deepEqual(d.groupElementIndices, []);
    assert.ok(d.elementIndex > 1);
    assert.deepEqual(d.components.map(c => [c.role, c.relationshipId, c.status, c.part]), [
      ["data", "dm", "located", "/deck/diagrams/data.xml"], ["layout", "lo", "located", "/deck/diagrams/layout.xml"],
      ["colors", "cs", "located", "/deck/diagrams/colors.xml"], ["style", "qs", "located", "/deck/diagrams/style.xml"],
    ]);
    assert.equal(d.components[1].layoutIdentifier, "urn:pptkit:test:unknown-layout");
    assert.equal(result.slides[1].diagrams, undefined); assert.deepEqual(new Uint8Array(bytes), copy);
    // Returned data cannot modify the source or another inspection.
    d.components[0].status = "external";
    assert.equal((await first(f))[0].components[0].status, "located");
  }
});

test("nested groups retain document order and shared components without collapsing diagrams", async () => {
  for (const strict of [false, true]) {
    const f = smartArtFixture(strict);
    f.setTree(f.frame("2") + `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3"/></p:nvGrpSpPr>${f.frame("4")}<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5"/></p:nvGrpSpPr>${f.frame("6")}</p:grpSp></p:grpSp>`);
    const diagrams = await first(f);
    assert.deepEqual(diagrams.map(d => d.drawingId), ["2", "4", "6"]);
    assert.deepEqual(diagrams.map(d => d.groupElementIndices.length), [0, 1, 2]);
    assert.equal(diagrams[1].groupElementIndices[0], diagrams[2].groupElementIndices[0]);
    assert.ok(diagrams[2].groupElementIndices[1] > diagrams[2].groupElementIndices[0]);
    assert.equal(new Set(diagrams.map(d => d.components[0].part)).size, 1);
    assert.ok(diagrams.every(d => d.identityStatus === "unique"));
  }
});

test("identity ambiguity includes ordinary shapes, groups and numeric lexical aliases", async () => {
  const f = smartArtFixture();
  f.setTree(f.frame("02") + '<p:sp><p:nvSpPr><p:cNvPr id="+2"/></p:nvSpPr></p:sp>' + f.frame("1") + f.frame("4294967296") + f.frame("bad") + f.frame("7").replace(' id="7"', ""));
  const diagrams = await first(f);
  assert.deepEqual(diagrams.map(d => d.identityStatus), ["ambiguous", "ambiguous", "invalid", "invalid", "missing"]);
  assert.equal(diagrams[0].drawingId, "02"); assert.equal(new Set(diagrams.map(d => d.elementIndex)).size, 5);
});

test("each role distinguishes absent IDs, unknown IDs, wrong role, external, fragment, type and root", async () => {
  for (const [roleIndex, attr] of ["dm", "lo", "cs", "qs"].entries()) {
    for (const [status, change] of [
      ["missing-reference", f => f.setTree(f.frame().replace(`r:${attr}="${attr}"`, ""))],
      ["missing-relationship", f => f.setTree(f.frame().replace(`r:${attr}="${attr}"`, `r:${attr}="unknown"`))],
      ["wrong-relationship-type", f => { f.relations[roleIndex + 1].Type = `${f.R}/image`; }],
      ["external", f => { Object.assign(f.relations[roleIndex + 1], { Target: "https://invalid.example/no-fetch", TargetMode: "External" }); }],
      ["fragment-target", f => { f.relations[roleIndex + 1].Target += "#x"; }],
      ["wrong-content-type", f => { const name = f.relations[roleIndex + 1].Target; f.types[`deck/${name}`] = "theme"; }],
      ["wrong-root", f => { f.parts[`deck/${f.relations[roleIndex + 1].Target}`] = `<d:wrong xmlns:d="${f.D}"/>`; }],
    ]) {
      const f = smartArtFixture(); change(f); f.setRelations();
      const [d] = await first(f);
      assert.equal(d.components[roleIndex].status, status, `${attr}: ${status}`);
      assert.ok(d.components.filter((_, i) => i !== roleIndex).every(c => c.status === "located"));
    }
  }
});

test("discovery ignores fake namespaces, other graphicData, arbitrary descendants and opaque extensions", async () => {
  const f = smartArtFixture();
  f.setTree(f.frame("2", "", "urn:other") + `<p:extLst><p:ext uri="opaque">${f.frame("3")}</p:ext></p:extLst><p:sp>${f.frame("4")}</p:sp>` + f.frame("5").replaceAll("d:relIds", "a:relIds"));
  const [d] = await first(f); assert.equal(d.drawingId, "5"); assert.ok(d.components.every(c => c.status === "missing-reference"));
});

test("MCE selected branch uses raw source identity and never borrows inactive component references", async () => {
  for (const strict of [false, true]) {
    const f = smartArtFixture(strict);
    f.setTree(`<mc:AlternateContent><mc:Choice Requires="future">${f.frame("9")}</mc:Choice><mc:Fallback>${f.frame("8", 'r:dm="dm"')}</mc:Fallback></mc:AlternateContent>`, 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:future"');
    const [d] = await first(f); assert.equal(d.drawingId, "8"); assert.equal(d.components[0].status, "located");
    assert.ok(d.components.slice(1).every(c => c.status === "missing-reference"));
    const plain = smartArtFixture(strict); plain.setTree(plain.frame("8", 'r:dm="dm"'));
    assert.ok(d.elementIndex > (await first(plain))[0].elementIndex);
  }
});

test("package-invalid declared targets, duplicate relIds and resource/cancellation failures remain explicit", async () => {
  const f = smartArtFixture(); f.relations[1].Target = "absent.xml"; f.setRelations();
  await assert.rejects(inspectPptx(f.bytes()), e => e.code === "package-invalid");
  const duplicate = smartArtFixture(); duplicate.setTree(duplicate.frame().replace('</a:graphicData>', '<d:relIds/></a:graphicData>'));
  await assert.rejects(inspectPptx(duplicate.bytes()), e => e.code === "presentation-invalid" && e.entry === "/deck/z.xml");
  const valid = smartArtFixture();
  await assert.rejects(inspectPptx(valid.bytes(), { xmlLimits: { maxNodes: 2 } }), { code: "xml-limit-exceeded" });
  await assert.rejects(inspectPptx(valid.bytes(), { signal: AbortSignal.abort() }), { code: "operation-aborted" });
});

test("both import projections report diagrams without fabricating editable nodes and preserve exact no-op", async () => {
  for (const strict of [false, true]) for (const projection of ["structure", "content"]) {
    const f = smartArtFixture(strict), bytes = f.bytes(), session = await importPptx(bytes, { projection });
    try {
      assert.equal(session.report.inspection.slides[0].diagrams[0].components[0].status, "located");
      assert.equal(session.document.slides[0].elements.length, 0);
      const prep = await session.prepareSave({ mode: "no-op" }); assert.equal(prep.status, "ready");
      assert.deepEqual((await session.save(prep.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});
