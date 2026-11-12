import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptxTiming, inspectPptx, importPptx } from "../dist/index.js";
import { timingBytes, timingFixture } from "./timing-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";
import { animationSpecimens, nativeTimingSpecimen } from "../../../scripts/research/native-timing-fixture.mjs";
const first = report => report.slides[0];

test("both dialects inventory all four native specimens with exact source order, lexical attributes and located drawing/time references", async () => {
  for (const strict of [false, true]) for (const spec of animationSpecimens) {
    const bytes = timingBytes({ strict, name: spec.name, timing: nativeTimingSpecimen(spec.name, 2).replace('</p:timing>', '<p:tn val="5"/></p:timing>') }), result = await inspectPptxTiming(bytes), slide = first(result);
    assert.equal(result.schemaVersion, "pptkit-native-timing-inventory-v1"); assert.equal(result.dialect, strict ? "strict" : "transitional"); assert.deepEqual(result.slides.map(s => s.slideId), [300, 256]);
    const behavior = slide.nodes.find(n => n.role === "behavior"); assert.equal(behavior.local, spec.tag);
    assert.deepEqual(Object.fromEntries(behavior.attributes.map(a => [a.local, a.value])), spec.attributes);
    const duration = slide.nodes.find(n => n.role === "time-node" && n.rawId === "5"); assert.equal(duration.attributes.find(a => a.local === "dur").value, "500");
    assert.ok(duration.parentTimeNodeElementIndex); assert.ok(duration.source.path.includes('Q{'));
    const target = slide.nodes.find(n => n.role === "drawing-target"); assert.equal(target.targetKind, "whole-object"); assert.equal(target.drawingReference.status, "located");
    assert.equal(slide.drawings.find(d => d.source.elementIndex === target.drawingReference.candidates[0].source.elementIndex).name, "Object 2");
    assert.equal(slide.nodes.find(n => n.role === "time-reference").timeReference.candidates[0].source.elementIndex, duration.source.elementIndex);
    assert.ok(slide.nodes.every((n, i) => i === 0 || n.source.elementIndex > slide.nodes[i - 1].source.elementIndex));
    assert.equal(result.nativePlayback, "not-performed"); assert.equal(result.authoredMapping, "not-performed");
    assert.ok(!Object.hasOwn((await inspectPptx(bytes)).slides[0], "timing"), "existing default structure shape stays unchanged");
  }
});

test("unknown namespaces, conditions and text/subobject/build targets retain raw values without acquiring Core mapping", async () => {
  const timing = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" repeatCount="1500" autoRev="1"><p:stCondLst><p:cond evt="onClick" delay="indefinite"><p:tn val="999"/></p:cond></p:stCondLst><p:childTnLst><x:future xmlns:x="urn:future" x:code="A&amp;B">left<![CDATA[中😀]]><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt></x:future></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="2" grpId="1"/></p:bldLst></p:timing>';
  const r = await inspectPptxTiming(timingBytes({ timing })), slide = first(r), future = slide.nodes.find(n => n.local === "future");
  assert.equal(future.role, "unknown"); assert.equal(future.text, "left中😀"); assert.deepEqual(future.attributes.find(a => a.local === "code"), { namespaceUri: "urn:future", local: "code", value: "A&B" });
  assert.equal(slide.nodes.find(n => n.role === "time-reference").timeReference.status, "missing");
  assert.equal(slide.nodes.find(n => n.role === "drawing-target").targetKind, "subobject"); assert.equal(slide.nodes.find(n => n.local === "bldP").targetKind, "uninterpreted");
  assert.ok(r.findings.some(f => f.code === "unknown-timing-node")); assert.ok(r.findings.some(f => f.code === "subobject-target")); assert.ok(r.findings.some(f => f.code === "missing-reference"));
  assert.equal(slide.nodes.find(n => n.role === "condition").attributes.find(a => a.local === "delay").value, "indefinite");
});

test("drawing/time namespaces remain separate and duplicates/numeric lexical aliases never choose a unique owner", async () => {
  for (const drawings of [shape({ id: "2" }) + shape({ id: "2" }), shape({ id: "2" }) + shape({ id: "02" })]) {
    const timing = '<p:timing><p:cTn id="2"/><p:cTn id="2"/><p:tn val="2"/><p:spTgt spid="2"/><p:spTgt spid="999"/><p:spTgt/></p:timing>';
    const r = await inspectPptxTiming(timingBytes({ drawings, timing })), s = first(r);
    assert.equal(s.nodes.find(n => n.role === "time-reference").timeReference.status, "ambiguous");
    const refs = s.nodes.filter(n => n.drawingReference).map(n => n.drawingReference); assert.deepEqual(refs.map(r => r.status), ["ambiguous", "missing", "missing"]);
    assert.equal(refs[0].candidates.length, 2); assert.ok(r.findings.some(f => f.code === "duplicate-time-id"));
    if (drawings.includes('id="02"')) assert.equal(refs[0].candidates[1].match, "numeric-alias");
  }
  const raw = await inspectPptxTiming(timingBytes({ drawings: shape({ id: "02" }), timing: '<p:timing><p:cTn id="str"/><p:tn val="str"/><p:spTgt spid="2"/></p:timing>' }));
  assert.ok(first(raw).nodes.filter(n => n.timeReference || n.drawingReference).every(n => (n.timeReference ?? n.drawingReference).status === "source-only"));
});

test("MCE raw-only branches and opaque extensions retain identities but do not pollute effective reference candidates", async () => {
  const mc = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"';
  const drawings = `<mc:AlternateContent><mc:Choice Requires="x">${shape({ id: "99" })}</mc:Choice><mc:Fallback>${shape({ id: "2" })}</mc:Fallback></mc:AlternateContent>`;
  const timing = `<mc:AlternateContent><mc:Choice Requires="x">${nativeTimingSpecimen("fade-in", 99)}</mc:Choice><mc:Fallback>${nativeTimingSpecimen("fade-out", 2)}</mc:Fallback></mc:AlternateContent><p:extLst><p:ext uri="future">${nativeTimingSpecimen("spin", 2)}</p:ext></p:extLst>`;
  const r = await inspectPptxTiming(timingBytes({ drawings, timing, rootAttributes: mc })), s = first(r);
  assert.deepEqual(s.nodes.filter(n => n.role === "timing").map(n => n.view), ["raw-only", "effective", "raw-only"]);
  const effective = s.nodes.find(n => n.drawingReference && n.view === "effective"); assert.equal(effective.drawingReference.status, "located"); assert.equal(effective.drawingReference.candidates.length, 1);
  assert.ok(s.nodes.filter(n => n.drawingReference && n.view === "raw-only").every(n => n.drawingReference.status === "source-only" && n.drawingReference.candidates.length === 0));
  assert.ok(!r.findings.some(f => f.code === "duplicate-time-id")); assert.equal(s.drawings.find(d => d.rawId === "99").view, "raw-only");
});

test("shared global work/record/text/finding quotas have exact boundaries and invocation captures permit independent retry", async () => {
  const bytes = timingBytes({ timing: nativeTimingSpecimen("fade-in", 2).replace('</p:timing>', '<x:future xmlns:x="urn:future"/><x:another xmlns:x="urn:future"/></p:timing>') }), baseline = await inspectPptxTiming(bytes);
  const exact = { maxWorkUnits: baseline.workUnits, maxRecords: baseline.recordCount, maxTextCharacters: baseline.textCharacters, maxFindings: baseline.findings.length };
  assert.deepEqual(await inspectPptxTiming(bytes, { timingLimits: exact }), baseline);
  for (const key of Object.keys(exact)) await assert.rejects(inspectPptxTiming(bytes, { timingLimits: { ...exact, [key]: exact[key] - 1 } }), { code: "package-limit-exceeded" });
  const options = { timingLimits: { maxRecords: 100000 } }, mutable = Uint8Array.from(bytes), pending = inspectPptxTiming(mutable, options); options.timingLimits.maxRecords = 1; mutable.fill(0); assert.deepEqual(await pending, baseline);
  for (const timingLimits of [false, { unknown: 1 }, { maxWorkUnits: 0 }, { maxRecords: Infinity }]) await assert.rejects(inspectPptxTiming(bytes, { timingLimits }));
  let invoked = false; await assert.rejects(inspectPptxTiming(bytes, { timingLimits: { get maxRecords() { invoked = true; return 1; } } }), TypeError); assert.equal(invoked, false);
  const abort = new AbortController(); abort.abort(); await assert.rejects(inspectPptxTiming(bytes, { signal: abort.signal }), { code: "operation-aborted" });
  assert.deepEqual(await inspectPptxTiming(bytes), baseline);
});

test("namespace aliases, UTF16, immutable public records and multiple timing roots remain explicit", async () => {
  const f = timingFixture({ timing: nativeTimingSpecimen("fade-in", 2) + nativeTimingSpecimen("fade-out", 2) });
  f.parts["deck/z.xml"] = Buffer.concat([Buffer.from([255, 254]), Buffer.from(f.parts["deck/z.xml"].replaceAll('<p:', '<z:').replaceAll('</p:', '</z:').replaceAll('xmlns:p=', 'xmlns:z='), "utf16le")]);
  const r = await inspectPptxTiming(encodeStructure(f)); assert.ok(r.findings.some(f => f.code === "multiple-effective-timing"));
  const s = first(r); assert.ok(Object.isFrozen(r) && Object.isFrozen(s.nodes) && Object.isFrozen(s.nodes[0].source)); assert.throws(() => { s.nodes[0].attributes.push({}); }, TypeError);
  assert.ok(s.nodes.every(n => n.source.path.includes(`Q{${f.P}}`))); assert.ok(s.nodes.every(n => !("children" in n)));
});

test("inspection adds no authoring/edit authority and three unrelated source saves retain timing exactly", async () => {
  let bytes = timingBytes(); const baseline = await inspectPptxTiming(bytes);
  for (let i = 0; i < 3; i++) {
    const session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    try {
      assert.equal(session.document.slides[0].eventTimeline, undefined);
      const slide = session.document.slides[0], changed = structuredClone(slide.elements[0]); changed.text.content[0].runs[0].text = `Round ${i}`; slide.replaceElement(changed.id, changed);
      const p = await session.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready"); bytes = (await session.save(p.plan)).bytes;
      assert.deepEqual(await inspectPptxTiming(bytes), baseline);
    } finally { session.dispose(); }
  }
});

test("time/drawing identities are part-local while report budgets accumulate across multiple slides", async () => {
  const single = timingFixture(), one = await inspectPptxTiming(encodeStructure(single));
  const both = timingFixture(); both.parts["deck/a.xml"] = both.parts["deck/z.xml"];
  const bytes = encodeStructure(both), report = await inspectPptxTiming(bytes);
  assert.equal(report.slides[0].nodes.length, report.slides[1].nodes.length);
  assert.ok(!report.findings.some(f => f.code === "duplicate-time-id" || f.code === "duplicate-drawing-id"));
  assert.ok(report.slides.every(s => s.nodes.filter(n => n.drawingReference).every(n => n.drawingReference.status === "located")));
  await assert.rejects(inspectPptxTiming(bytes, { timingLimits: { maxRecords: one.recordCount } }), { code: "package-limit-exceeded" });
  assert.deepEqual(await inspectPptxTiming(bytes), report);
});
