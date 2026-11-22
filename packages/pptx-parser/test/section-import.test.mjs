import assert from "node:assert/strict";
import test from "node:test";
import { importPptx } from "../dist/index.js";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
import { section, withSections, SECTION_IDS, SECTION_URI } from "./sections-fixture.mjs";

const open = (fixture, projection = "structure", options = {}) => importPptx(encodeStructure(fixture), { projection, ...options });
const equalBytes = (a, b) => assert.equal(Buffer.compare(Buffer.from(a), Buffer.from(b)), 0);

test("opt-in section collections retain names, empty entries and stable source membership", async () => {
  for (const strict of [false, true]) for (const projection of ["structure", "content"]) {
    const body = section([300], `id="${SECTION_IDS[0]}" name="same"`) + section([], 'name=""') + section([256], `id="${SECTION_IDS[2]}" name="same"`) + section();
    const fixture = withSections(structureFixture(strict), body), bytes = encodeStructure(fixture);
    const session = await importPptx(bytes, { projection, sectionProjection: "collection-v1" });
    assert.deepEqual(session.document.sections, [
      { id: "source-section-1", name: "same", slideIds: ["source-slide-300"] },
      { id: "source-section-2", name: "", slideIds: [] },
      { id: "source-section-3", name: "same", slideIds: ["source-slide-256"] },
      { id: "source-section-4", slideIds: [] },
    ]);
    assert.deepEqual(session.document.slides.map(s => s.section), ["same", "same"]);
    const sectionIdentities = session.report.sectionSources;
    assert.deepEqual(sectionIdentities.map(identity => [identity.id, identity.source.sourceObjectId]), [["source-section-1", SECTION_IDS[0]], ["source-section-2", undefined], ["source-section-3", SECTION_IDS[2]], ["source-section-4", undefined]]);
    assert.ok(session.report.findings.some(f => f.feature === "section.native-collection" && f.operation === "import" && f.tier === "editable"));
    assert.ok(session.report.findings.some(f => f.feature === "section.native-collection" && f.operation === "edit-save" && f.tier === "unsupported"));
    assert.ok(session.report.diagnostics.some(d => d.code === "section-collection-projection"));
    const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
    equalBytes((await session.save(ready.plan)).bytes, bytes);
    session.document.setSections([{ id: "other", slideIds: session.document.slides.map(s => s.id) }]);
    assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    session.dispose();
  }
});

test("collection-v2 keeps native sections independent of legacy slide labels", async () => {
  for (const strict of [false, true]) for (const projection of ["structure", "content"]) {
    const fixture = withSections(structureFixture(strict)), bytes = encodeStructure(fixture);
    const session = await importPptx(bytes, { projection, sectionProjection: "collection-v2" });
    try {
      assert.deepEqual(session.document.sections.map(s => [s.name, s.slideIds]), [["Opening", ["source-slide-300"]], ["Empty", []], ["Closing", ["source-slide-256"]]]);
      assert.deepEqual(session.document.slides.map(s => s.section), [undefined, undefined]);
      assert.ok(session.report.findings.some(f => f.feature === "section.label" && f.operation === "import" && f.tier === "unsupported"));
      assert.equal(session.report.sectionSources.length, 3);
      const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
      equalBytes((await session.save(ready.plan)).bytes, bytes);
      session.document.setSections(session.document.sections.map((s, index) => index === 0 ? { ...s, name: "Renamed" } : s));
      assert.deepEqual(session.document.slides.map(s => s.section), [undefined, undefined]);
      assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { session.dispose(); }
  }
});

test("collection projection retains partial and out-of-order membership without changing default imports", async () => {
  for (const [body, members, exportTier] of [
    [section([300]), [["source-slide-300"]], "unsupported"],
    [section([256]) + section([300]), [["source-slide-256"], ["source-slide-300"]], "unsupported"],
    [section([256, 300]), [["source-slide-300", "source-slide-256"]], "editable"],
  ]) {
    const fixture = withSections(structureFixture(), body);
    const projected = await open(fixture, "structure", { sectionProjection: "collection-v1" });
    assert.deepEqual(projected.document.sections.map(s => s.slideIds), members);
    assert.ok(projected.report.findings.filter(f => f.feature === "section.native-collection" && f.operation === "export").every(f => f.tier === exportTier));
    projected.dispose();
    const session = await open(fixture); assert.equal(session.document.sections, undefined); session.dispose();
  }
  const absent = await open(structureFixture(), "content", { sectionProjection: "collection-v1" });
  assert.equal(absent.document.sections, undefined); absent.dispose();
  await assert.rejects(open(structureFixture(), "content", { sectionProjection: "invalid" }), TypeError);
  const compatible = await open(structureFixture(), "content", { sectionProjection: "collection-v1", editSaveProfile: "text-runs-v1" }); compatible.dispose();
});

test("collection projection keeps unknown section extensions in source inspection", async () => {
  const nested = '<s:extLst><p:ext uri="urn:future"><u:data xmlns:u="urn:future"/></p:ext></s:extLst>';
  const fixture = withSections(structureFixture(), section([300, 256], 'name="Known"', nested));
  const source = encodeStructure(fixture), session = await importPptx(source, { projection: "structure", sectionProjection: "collection-v1" });
  assert.deepEqual(session.document.sections, [{ id: "source-section-1", name: "Known", slideIds: ["source-slide-300", "source-slide-256"] }]);
  assert.equal(session.report.inspection.sectionList.sections[0].extensions[0].status, "uninterpreted");
  const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
  equalBytes((await session.save(ready.plan)).bytes, source); session.dispose();
});

test("native member names project by stable slide identity in both dialects and projections", async () => {
  for (const strict of [false, true]) for (const projection of ["structure", "content"]) {
    const session = await open(withSections(structureFixture(strict)), projection);
    assert.deepEqual(session.document.slides.map(slide => [slide.id, slide.section]), [["source-slide-300", "Opening"], ["source-slide-256", "Closing"]]);
    assert.deepEqual(validatePresentation(session.document), []);
    const native = session.report.findings.filter(f => f.feature === "section.native-collection");
    assert.equal(native.length, 21);
    assert.deepEqual(native.filter(f => f.operation === "no-op-save").map(f => f.tier), ["preserved", "preserved", "preserved"]);
    assert.ok(native.filter(f => f.operation !== "no-op-save").every(f => f.tier === "unsupported"));
    const labels = session.report.findings.filter(f => f.feature === "section.label");
    assert.equal(labels.length, 14);
    assert.deepEqual(labels.filter(f => f.operation === "import").map(f => [f.slideId, f.source.sourceObjectId, f.source.relationshipId]), [
      ["source-slide-300", SECTION_IDS[0], "z"], ["source-slide-256", SECTION_IDS[2], "a"],
    ]);
    assert.equal(labels.find(f => f.operation === "import").source.path, `presentation/extLst/ext[@uri="${SECTION_URI}"]/sectionLst/section[1]/sldIdLst/sldId[1]`);
    assert.ok(labels.filter(f => f.operation === "export").every(f => f.tier === "fallback" && /not as a native/.test(f.reason)));
    assert.ok(session.report.diagnostics.some(d => d.code === "section-label-projection-only"));
    assert.ok(!session.report.identities.some(identity => identity.kind === "section"));
    session.dispose();
  }
});

test("absent collection, empty source section, absent name and explicit empty name remain distinct", async () => {
  const absent = await open(structureFixture());
  assert.deepEqual(absent.document.slides.map(s => s.section), [undefined, undefined]);
  assert.ok(!absent.report.findings.some(f => f.feature.startsWith("section.")));
  assert.ok(!absent.report.diagnostics.some(d => d.code === "section-label-projection-only"));
  absent.dispose();
  const empty = await open(withSections(structureFixture(), section([], 'name="Empty"')));
  assert.deepEqual(empty.document.slides.map(s => s.section), [undefined, undefined]);
  assert.equal(empty.report.inspection.sectionList.sections.length, 1);
  assert.equal(empty.report.findings.filter(f => f.feature === "section.native-collection").length, 7);
  assert.ok(!empty.report.findings.some(f => f.feature === "section.label"));
  empty.dispose();
  const names = await open(withSections(structureFixture(), section([300]) + section([256], 'name=""')));
  assert.deepEqual(names.document.slides.map(s => s.section), [undefined, ""]);
  assert.ok(names.report.findings.some(f => f.feature === "section.label" && f.operation === "import" && f.tier === "unsupported" && f.slideId === "source-slide-300"));
  assert.ok(names.report.findings.some(f => f.feature === "section.label" && f.operation === "import" && f.tier === "editable" && f.slideId === "source-slide-256"));
  names.dispose();
});

test("duplicate decoded names do not merge source identities or create a Core section collection", async () => {
  const label = "中文 & \n";
  const session = await open(withSections(structureFixture(), section([300], `id="${SECTION_IDS[0]}" name="中文 &amp; &#xA;"`) + section([256], `id="${SECTION_IDS[1]}" name="中文 &amp; &#xA;"`)));
  assert.deepEqual(session.document.slides.map(s => s.section), [label, label]);
  assert.deepEqual(session.report.inspection.sectionList.sections.map(s => s.id), SECTION_IDS.slice(0, 2));
  const copy = session.document.duplicateSlide(session.document.slides[0].id);
  assert.equal(copy.section, label);
  assert.equal(session.report.inspection.sectionList.sections[0].slides.length, 1);
  session.dispose();
});

test("partial, reversed and noncontiguous source membership never rewrites presentation order", async () => {
  const reversed = await open(withSections(structureFixture(), section([256], 'name="Second"') + section([300], 'name="First"')));
  assert.deepEqual(reversed.document.slides.map(s => [s.id, s.section]), [["source-slide-300", "First"], ["source-slide-256", "Second"]]);
  assert.equal(reversed.report.diagnostics.filter(d => d.code === "section-order").length, 2);
  reversed.dispose();
  const partial = await open(withSections(structureFixture(), section([256], 'name="Only"')));
  assert.deepEqual(partial.document.slides.map(s => s.section), [undefined, "Only"]);
  assert.ok(partial.report.diagnostics.some(d => d.code === "unsectioned-slide" && d.source.sourceObjectId === "300"));
  partial.dispose();
  const noncontiguous = await open(withSections(structureFixture(), section([256, 300], 'name="One"')));
  assert.deepEqual(noncontiguous.document.slides.map(s => s.section), ["One", "One"]);
  assert.deepEqual(noncontiguous.report.inspection.sectionList.sections[0].slides.map(s => s.id), [256, 300]);
  noncontiguous.dispose();
});

test("reorder, duplicate and deletion retain labels but invalidate no-op until exact semantic undo", async () => {
  const fixture = withSections(structureFixture()), source = encodeStructure(fixture), session = await open(fixture);
  const original = normalizePresentation(session.document), first = session.document.slides[0].id;
  const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  session.document.moveSlide(first, 1);
  assert.deepEqual(session.document.slides.map(s => s.section), ["Closing", "Opening"]);
  await assert.rejects(session.save(prepared.plan), e => e.code === "save-snapshot-stale");
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  session.document.moveSlide(first, 0);
  const copy = session.document.duplicateSlide(first); assert.equal(copy.section, "Opening");
  session.document.removeSlide(copy.id);
  assert.deepEqual(normalizePresentation(session.document), original);
  equalBytes((await session.save(prepared.plan)).bytes, source);
  assert.deepEqual(session.report.inspection.sectionList.sections.map(s => s.slides.map(v => v.id)), [[300], [], [256]]);
  session.dispose();
});

test("section capability rows respect the shared report budget and failed imports remain retryable", async () => {
  const fixture = withSections(structureFixture()), baseline = await open(fixture);
  const maxReportFindings = baseline.report.findings.length;
  baseline.dispose();
  await assert.rejects(open(fixture, "structure", { sessionLimits: { maxReportFindings: maxReportFindings - 1 } }), e => e.code === "import-limit-exceeded");
  const recovered = await open(fixture, "structure", { sessionLimits: { maxReportFindings } });
  assert.equal(recovered.report.findings.length, maxReportFindings);
  recovered.dispose();
});

test("section projection checkpoints cover empty and member-bearing collections", async () => {
  const { projectSections } = await import("../dist/import/project-sections.js");
  const inspection = { sectionList: { part: "/ppt/presentation.xml", extensionUri: SECTION_URI, sections: [{ name: "Empty", slides: [] }, { name: "Member", slides: [{ id: 256, relationshipId: "rId1" }] }] } };
  const stop = new Error("cancelled"), rows = [];
  let checkpoints = 0;
  await assert.rejects(projectSections(inspection, (...row) => rows.push(row), async () => { if (++checkpoints === 3) throw stop; }), error => error === stop);
  assert.equal(checkpoints, 3);
  assert.equal(rows.length, 14);
  const recovered = await projectSections(inspection, () => {}, async () => {});
  assert.deepEqual([...recovered.labels], [[256, "Member"]]);
});
