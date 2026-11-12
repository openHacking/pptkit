import test from "node:test";
import assert from "node:assert/strict";
import { importPptx, inspectPptx, PptxImportError, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { metadataFixture, encodeMetadata, custom, GUID } from "./metadata-fixture.mjs";
import { withSections } from "./sections-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { execFileSync } from "node:child_process";

const options = { projection: "structure" };
const input = () => new Uint8Array(encodeMetadata(metadataFixture()));
const equalBytes = (a, b) => assert.equal(Buffer.compare(Buffer.from(a), Buffer.from(b)), 0, "source bytes differ");
const open = (bytes = input(), extra = {}) => importPptx(bytes, { ...options, ...extra });
const code = (expected) => (error) => { assert.ok(error instanceof PptxImportError); assert.equal(error.code, expected); return true; };
async function plan(session) { const value = await session.prepareSave({ mode: "no-op" }); assert.equal(value.status, "ready"); return value.plan; }
function frozen(value) {
  if (!value || typeof value !== "object") return;
  assert.ok(Object.isFrozen(value)); for (const child of Object.values(value)) frozen(child);
}

test("explicit structural import maps exact dimensions, source order, visibility and selected metadata without fabricated content", async () => {
  for (const strict of [false, true]) {
    const f = withSections(metadataFixture(strict));
    f.add("core", '<d:title>中文 &amp; title</d:title><d:creator>Author</d:creator><d:subject>Subject</d:subject><d:description>Description</d:description><d:language>zh-CN</d:language><c:revision>9007199254740993</c:revision><c:keywords>a,b</c:keywords>');
    f.add("extended", '<Company>Company</Company>');
    const session = await open(encodeMetadata(f));
    assert.deepEqual(session.document.size, { width: 720, height: 540, unit: "pt" });
    assert.deepEqual(session.document.metadata, { title: "中文 & title", author: "Author", subject: "Subject", description: "Description", language: "zh-CN", keywords: ["a,b"], company: "Company" });
    assert.ok(session.report.findings.some((f) => f.feature === "metadata.revision" && f.tier === "unsupported" && f.source.path === "coreProperties/revision"));
    assert.deepEqual(session.document.slides.map((s) => [s.id, s.hidden, s.elements.length, s.section, s.notes]), [["source-slide-300", true, 0, "Opening", undefined], ["source-slide-256", false, 0, "Closing", undefined]]);
    assert.deepEqual(validatePresentation(session.document), []);
    assert.equal(normalizePresentation(session.document).irVersion, 2);
    assert.equal(session.report.inspection.sectionList.sections.length, 3);
    assert.deepEqual(session.report.identities.filter((i) => i.kind === "slide").map((i) => i.source), [
      { part: "/deck/z.xml", sourceObjectId: "300", relationshipId: "z" }, { part: "/deck/a.xml", sourceObjectId: "256", relationshipId: "a" },
    ]);
    assert.ok(session.report.diagnostics.some((d) => d.code === "structural-projection-only"));
    assert.ok(session.report.findings.some((f) => f.feature === "source-part-content" && f.operation === "preview" && f.tier === "unsupported"));
    session.dispose();
  }
});

test("structural import maps safe integer revision and reports mixed keyword segments", async () => {
  const f = metadataFixture();
  f.add("core", '<c:revision>42</c:revision><c:keywords><c:value xml:lang="zh">中文</c:value>,<c:value xml:lang="en">English</c:value></c:keywords>');
  const session = await open(encodeMetadata(f));
  assert.deepEqual(session.document.metadata, { keywords: ["中文,English"], revision: 42 });
  assert.ok(session.report.findings.some((finding) => finding.feature === "metadata.revision" && finding.tier === "editable"));
  assert.ok(session.report.findings.some((finding) => finding.feature === "metadata.keywords" && finding.tier === "editable"));
  assert.ok(session.report.findings.some((finding) => finding.feature === "metadata.keyword-segments" && finding.tier === "unsupported"));
  session.dispose();
});

test("a single keyword value wrapper still reports source segment structure", async () => {
  const f = metadataFixture();
  f.add("core", "<c:keywords><c:value>one</c:value></c:keywords>");
  const session = await open(encodeMetadata(f));
  assert.deepEqual(session.document.metadata.keywords, ["one"]);
  assert.ok(session.report.findings.some(finding => finding.feature === "metadata.keyword-segments" && finding.tier === "unsupported"));
  assert.deepEqual(session.report.inspection.metadata.core.properties[0].keywordSegments, [{ kind: "value", value: "one" }]);
  session.dispose();
});

test("empty keyword language remains visible as an unsupported source annotation", async () => {
  for (const strict of [false, true]) {
    for (const [attribute, expected] of [["", false], [' xml:lang=""', true], [' xml:lang="zh-CN"', true]]) {
      const f = metadataFixture(strict); f.add("core", `<c:keywords${attribute}/>`);
      const bytes = encodeMetadata(f), session = await open(bytes);
      try {
        assert.deepEqual(session.document.metadata.keywords, [""]);
        assert.equal(session.report.findings.some(finding => finding.feature === "metadata.keyword-segments" && finding.tier === "unsupported"), expected);
        assert.equal(session.report.inspection.metadata.core.properties[0].language, expected ? attribute.includes("zh-CN") ? "zh-CN" : "" : undefined);
        const saved = await session.save(await plan(session));
        equalBytes(saved.bytes, bytes);
      } finally { session.dispose(); }
    }
  }
});

test("revision projection preserves only canonical safe nonnegative integer lexicals", async () => {
  for (const [lexical, expected] of [["0", 0], ["9007199254740991", Number.MAX_SAFE_INTEGER], ["9007199254740992", undefined], ["+42", undefined], ["042", undefined], [" 42 ", undefined], ["-1", undefined], ["1e2", undefined], ["", undefined]]) {
    const f = metadataFixture(); f.add("core", `<c:revision>${lexical}</c:revision>`);
    const session = await open(encodeMetadata(f));
    assert.equal(session.document.metadata.revision, expected, lexical);
    assert.equal(session.report.inspection.metadata.core.properties[0].value, lexical);
    assert.equal(session.report.findings.find(finding => finding.feature === "metadata.revision").tier, expected === undefined ? "unsupported" : "editable");
    session.dispose();
  }
});

test("unmapped metadata properties report source identity and remain in no-op output", async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("core", '<d:title>Title</d:title><t:created xsi:type="t:W3CDTF">2020-01-02T03:04:05+08:00</t:created><c:category/>');
    f.add("extended", '<Company>Mapped</Company><Manager>Alice</Manager><Slides>2</Slides>');
    f.add("custom", custom(2, "lpwstr", "", 'name="Empty"') + custom(3, "vector", '<v:lpstr>opaque</v:lpstr>', 'name="Complex"'));
    const bytes = encodeMetadata(f), session = await open(bytes);
    try {
      const findings = session.report.findings.filter(item => item.tier === "unsupported" && item.feature.startsWith("metadata."));
      assert.deepEqual(findings.map(item => [item.feature, item.source.path, item.source.sourceObjectId]), [
        ["metadata.core.category", "coreProperties/category", undefined],
        ["metadata.extended.Manager", "Properties/Manager", undefined], ["metadata.extended.Slides", "Properties/Slides", undefined],
        ["metadata.custom-property", "Properties/property", `${GUID}:3`],
      ]);
      assert.deepEqual(session.document.metadata, { title: "Title", createdAt: "2020-01-01T19:04:05.000Z", company: "Mapped", customProperties: [{ name: "Empty", type: "string", value: "" }] });
      assert.equal(session.report.inspection.metadata.core.properties[1].value, "2020-01-02T03:04:05+08:00");
      assert.equal(session.report.inspection.metadata.custom.properties[0].value.value, "");
      assert.equal(session.report.inspection.metadata.custom.properties[1].value.status, "uninterpreted");
      equalBytes((await session.save(await plan(session))).bytes, bytes);
    } finally { session.dispose(); }
  }
});

test("metadata absence, empty source strings and normalization defaults remain distinct", async () => {
  for (const strict of [false, true]) {
    const missing = metadataFixture(strict), absent = await open(encodeMetadata(missing));
    try {
      assert.deepEqual(absent.document.metadata, {});
      assert.deepEqual(normalizePresentation(absent.document).metadata, { title: "PPTKit Presentation", author: "PPTKit", language: "en-US", keywords: [], revision: 1 });
      assert.equal(absent.report.inspection.metadata.core, undefined);
    } finally { absent.dispose(); }
    const f = metadataFixture(strict);
    f.add("core", '<d:title/><d:creator/><d:language/><d:subject/><d:description/><c:keywords/><c:revision>0</c:revision>');
    f.add("extended", '<Company/>');
    const bytes = encodeMetadata(f), present = await open(bytes);
    try {
      assert.deepEqual(present.document.metadata, { title: "", author: "", language: "", subject: "", description: "", keywords: [""], revision: 0, company: "" });
      assert.deepEqual(normalizePresentation(present.document).metadata, { title: "", author: "", language: "", subject: "", description: "", keywords: [""], revision: 0, company: "" });
      equalBytes((await present.save(await plan(present))).bytes, bytes);
    } finally { present.dispose(); }
  }
});

test("custom property projection maps only the explicit loss-aware subset", async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("custom", [
      custom(2, "lpwstr", "hello", 'name="text"'), custom(3, "bool", "1", 'name="flag"'),
      custom(4, "i4", "-2147483648", 'name="count"'), custom(5, "r8", "-0", 'name="zero"'),
      custom(6, "filetime", "2020-01-02T03:04:05.000Z", 'name="date"'),
      custom(7, "i4", "+1", 'name="noncanonical"'), custom(8, "r8", "1.00", 'name="precision"'),
      custom(9, "filetime", "2020-01-02T03:04:05+08:00", 'name="other-date"'),
      custom(10, "lpwstr", "linked", 'name="link" linkTarget="x"'), custom(11, "lpwstr", "anonymous"),
    ].join(""));
    const bytes = encodeMetadata(f), session = await open(bytes);
    try {
      assert.deepEqual(session.document.metadata.customProperties, [
        { name: "text", type: "string", value: "hello" }, { name: "flag", type: "boolean", value: true },
        { name: "count", type: "integer", value: -2147483648 }, { name: "zero", type: "real", value: -0 },
        { name: "date", type: "date", value: "2020-01-02T03:04:05.000Z" },
        { name: "other-date", type: "date", value: "2020-01-01T19:04:05.000Z" },
      ]);
      assert.deepEqual(session.report.findings.filter(f => f.feature === "metadata.custom-property").map(f => f.tier), ["editable", "editable", "editable", "editable", "editable", "unsupported", "unsupported", "editable", "unsupported", "unsupported"]);
      equalBytes((await session.save(await plan(session))).bytes, bytes);
    } finally { session.dispose(); }
  }
});

test("custom source OOXML escape-like text is retained in inspection without unsafe projection", async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("custom", custom(2, "lpwstr", "_x0041_", 'name="value"') + custom(3, "lpwstr", "plain", 'name="_x0042_"') + custom(4, "lpwstr", "_x005F_x0041_", 'name="escaped"'));
    const source = encodeMetadata(f), session = await open(source);
    try {
      assert.deepEqual(session.document.metadata.customProperties, [{ name: "escaped", type: "string", value: "_x0041_" }]);
      assert.deepEqual(session.report.inspection.metadata.custom.properties.map(p => p.value.value), ["_x0041_", "plain", "_x005F_x0041_"]);
      assert.deepEqual(session.report.findings.filter(f => f.feature === "metadata.custom-property").map(f => f.tier), ["unsupported", "unsupported", "editable"]);
      equalBytes((await session.save(await plan(session))).bytes, source);
    } finally { session.dispose(); }
  }
});

test("overlapping decoded escape-like values remain source-only", async () => {
  const f = metadataFixture(); f.add("custom", custom(2, "lpwstr", "_x005F_x005F_x0041_", 'name="test"'));
  const source = encodeMetadata(f), session = await open(source);
  try {
    assert.deepEqual(session.document.metadata.customProperties, []);
    assert.equal(session.report.findings.find(f => f.feature === "metadata.custom-property").tier, "unsupported");
    equalBytes((await session.save(await plan(session))).bytes, source);
  } finally { session.dispose(); }
});

test("decoded duplicate custom names are both excluded from Core", async () => {
  const f = metadataFixture();
  f.add("custom", custom(2, "lpwstr", "one", 'name="_x005F_x0041_"') + custom(3, "lpwstr", "two", 'name="_x005f_x0041_"'));
  const session = await open(encodeMetadata(f));
  try {
    assert.deepEqual(session.document.metadata.customProperties, []);
    assert.deepEqual(session.report.findings.filter(f => f.feature === "metadata.custom-property").map(f => f.tier), ["unsupported", "unsupported"]);
  } finally { session.dispose(); }
});

test("unmapped metadata findings obey the session report budget atomically", async () => {
  const f = metadataFixture(); f.add("custom", custom(2, "lpwstr", "one") + custom(3, "lpwstr", "two"));
  const bytes = encodeMetadata(f), baseline = await open(bytes);
  const limit = baseline.report.findings.length;
  baseline.dispose();
  const exact = await open(bytes, { sessionLimits: { maxReportFindings: limit } });
  try { assert.equal(exact.report.findings.length, limit); } finally { exact.dispose(); }
  await assert.rejects(open(bytes, { sessionLimits: { maxReportFindings: limit - 1 } }), code("import-limit-exceeded"));
});

test("every source part receives distinct operation findings and immutable detached provenance", async () => {
  const a = await open(), b = await open();
  for (const part of a.report.inspection.package.parts) {
    const findings = a.report.findings.filter((f) => f.source.part === part.name && f.feature.startsWith("source-part"));
    assert.equal(findings.length, 7);
    assert.equal(findings.find((f) => f.operation === "no-op-save").tier, "preserved");
    assert.ok(findings.filter((f) => f.operation !== "no-op-save").every((f) => f.tier === "unsupported"));
  }
  assert.equal(new Set(a.report.findings.map((f) => f.id)).size, a.report.findings.length);
  frozen(a.report);
  assert.notEqual(a.report, b.report); assert.notEqual(a.report.inspection, b.report.inspection);
  assert.throws(() => { a.report.findings[0].tier = "editable"; }, TypeError);
  assert.throws(() => { a.document = b.document; }, TypeError);
  a.dispose(); b.dispose();
});

test("no-op plans return byte-identical source snapshots with independently owned outputs and no rebasing", async () => {
  const bytes = input(), original = new Uint8Array(bytes), importing = open(bytes);
  bytes.fill(0);
  const session = await importing, p = await plan(session), q = await plan(session);
  assert.equal(p.byteLength, original.length); assert.equal(p.report.assurance, "original-bytes"); frozen(p.report);
  const first = await session.save(p), second = await session.save(p, { acceptLossIds: [] });
  equalBytes(first.bytes, original); equalBytes(second.bytes, original);
  assert.notEqual(first.bytes.buffer, second.bytes.buffer); first.bytes.fill(1);
  equalBytes((await session.save(q)).bytes, original);
  assert.equal((await inspectPptx(second.bytes)).slides.length, 2);
  assert.throws(() => { p.byteLength = 0; }, TypeError);
  p.dispose(); q.dispose(); session.dispose(); equalBytes(second.bytes, original);
});

test("opaque binary parts, cycles, unrelated extensions and external targets survive exact-byte save without interpretation", async () => {
  const f = metadataFixture();
  f.parts["unknown/a.bin"] = Uint8Array.from([0, 255, 17, 88]); f.parts["unknown/b.bin"] = Uint8Array.from([5, 4, 3]);
  f.metadataTypes["unknown/a.bin"] = f.metadataTypes["unknown/b.bin"] = "application/octet-stream";
  f.parts["unknown/_rels/a.bin.rels"] = relationships([{ Id: "other", Type: "urn:future", Target: "b.bin" }, { Id: "link", Type: "urn:external", TargetMode: "External", Target: "https://example.invalid/resource" }]);
  f.parts["unknown/_rels/b.bin.rels"] = relationships([{ Id: "back", Type: "urn:future", Target: "a.bin" }]);
  const bytes = encodeMetadata(f), session = await open(bytes);
  const p = await plan(session); equalBytes((await session.save(p)).bytes, bytes);
  assert.ok(session.report.findings.some((f) => f.source.part === "/unknown/a.bin" && f.operation === "no-op-save"));
  session.dispose();
});

test("same normalized state is reusable; semantic edits block no-op and stale prepared plans", async () => {
  const session = await open(), p = await plan(session), baseline = normalizePresentation(session.document);
  session.document.moveSlide("source-slide-300", 1);
  const prepared = await session.prepareSave({ mode: "no-op" });
  assert.equal(prepared.status, "blocked"); assert.equal(prepared.report.assurance, "none"); assert.equal(prepared.report.diagnostics[0].code, "save-document-changed");
  await assert.rejects(session.save(p), code("save-snapshot-stale"));
  session.document.moveSlide("source-slide-300", 0);
  assert.deepEqual(normalizePresentation(session.document), baseline); assert.ok((await session.save(p)).bytes.length);
  session.document.addSlide({ id: "new" });
  await assert.rejects(session.save(p), code("save-snapshot-stale"));
  session.document.removeSlide("new"); assert.ok((await session.save(p)).bytes.length);
  session.dispose();
});

test("edit and rebuild modes block explicitly without plans, bytes or a misleading output assurance", async () => {
  const session = await open();
  for (const mode of ["edit", "rebuild"]) {
    const result = await session.prepareSave({ mode });
    assert.deepEqual(Object.keys(result).sort(), ["report", "status"]);
    assert.equal(result.status, "blocked"); assert.equal(result.report.mode, mode); assert.equal(result.report.assurance, "none");
    assert.equal(result.report.diagnostics[0].code, "save-mode-unsupported"); frozen(result);
  }
  assert.ok((await session.save(await plan(session))).bytes.length); session.dispose();
});

test("cross-session, forged, prototype-derived and cloned plans cannot pass ownership checks", async () => {
  const a = await open(), b = await open(), p = await plan(a);
  assert.equal(a.document.id, b.document.id); // Equal contents/IDs do not convey source ownership.
  await assert.rejects(b.save(p), code("plan-session-mismatch"));
  for (const fake of [{}, { ...p }, Object.create(p), null, new Proxy(p, {})]) await assert.rejects(a.save(fake), code("plan-session-mismatch"));
  assert.ok((await a.save(p)).bytes.length); a.dispose(); b.dispose();
});

test("loss-free consent requires exactly no IDs and cannot be overridden by generic flags", async () => {
  const session = await open(), p = await plan(session);
  for (const consent of [{ acceptLossIds: ["foreign"] }, { acceptLossIds: ["a", "a"] }, { acceptLossIds: "" }, { acceptLossIds: null }, { allowLossy: true }, { acceptLossIds: [], extra: true }, null, []]) await assert.rejects(session.save(p, consent), code("save-loss-consent-mismatch"));
  assert.ok((await session.save(p)).bytes.length); assert.ok((await session.save(p, { acceptLossIds: [] })).bytes.length); session.dispose();
});

test("plan disposal is idempotent, releases capacity and cannot invalidate independent plans or outputs", async () => {
  const session = await open(), p = await plan(session), q = await plan(session), output = await session.save(p);
  p.dispose(); p.dispose();
  await assert.rejects(session.save(p), code("plan-disposed")); assert.ok((await session.save(q)).bytes.length);
  equalBytes(output.bytes, input()); session.dispose();
});

test("session disposal is idempotent, revokes plans, leaves authoring state independent and prevents serialization", async () => {
  const session = await open(), document = session.document, p = await plan(session);
  assert.throws(() => JSON.stringify(session), code("session-not-serializable"));
  assert.throws(() => JSON.stringify({ p }), code("session-not-serializable"));
  assert.throws(() => structuredClone(session)); assert.throws(() => structuredClone(p));
  session.dispose(); session.dispose(); p.dispose();
  await assert.rejects(session.prepareSave({ mode: "no-op" }), code("session-disposed"));
  await assert.rejects(session.save(p), code("session-disposed"));
  document.addSlide({ id: "after-disposal" }); assert.deepEqual(validatePresentation(document), []);
  assert.equal(document.slides.length, 3); frozen(session.report);
});

test("mutation while preparing produces a stale blocked result and leaves the baseline usable after undo", async () => {
  const session = await open(), pending = session.prepareSave({ mode: "no-op" });
  session.document.addSlide({ id: "during-prepare" });
  const result = await pending; assert.equal(result.status, "blocked"); assert.equal(result.report.diagnostics[0].code, "save-snapshot-stale");
  session.document.removeSlide("during-prepare"); assert.ok((await plan(session)).byteLength); session.dispose();
});

test("mid-preparation cancellation/disposal revoke in-flight output and capacity is recovered", async () => {
  for (const action of ["abort", "dispose"]) {
    const session = await open(undefined, { sessionLimits: { maxLivePlans: 1 } }), controller = new AbortController();
    const pending = session.prepareSave({ mode: "no-op", signal: controller.signal });
    // First prepare yield reserves output; this timer runs during the copy yield.
    setTimeout(() => action === "abort" ? controller.abort() : session.dispose(), 0);
    await assert.rejects(pending, code(action === "abort" ? "operation-aborted" : "session-disposed"));
    if (action === "abort") assert.ok((await plan(session)).byteLength);
    session.dispose();
  }
});

test("mutations during output copying are rechecked and never expose stale candidate bytes", async () => {
  const session = await open(undefined, { sessionLimits: { maxLivePlans: 1 } });
  const pending = session.prepareSave({ mode: "no-op" });
  setTimeout(() => session.document.addSlide({ id: "copy-edit" }), 0);
  const result = await pending; assert.equal(result.status, "blocked"); assert.equal(result.report.diagnostics[0].code, "save-snapshot-stale");
  session.document.removeSlide("copy-edit"); assert.ok((await plan(session)).byteLength); session.dispose();
});

test("invalid Core edits block preparation and cannot be saved through previously prepared plans", async () => {
  const session = await open(), p = await plan(session);
  session.document.slides[0].addElement({ type: "text", id: "invalid", text: "invalid", box: { x: 0, y: 0, width: -1, height: 10 } });
  const result = await session.prepareSave({ mode: "no-op" });
  assert.equal(result.status, "blocked"); assert.equal(result.report.diagnostics[0].code, "save-invalid-document");
  assert.ok(result.report.diagnostics.some((d) => d.elementId === "invalid" && d.documentPath && d.source === undefined));
  await assert.rejects(session.save(p), code("save-invalid-document"));
  session.document.slides[0].removeElement("invalid"); assert.ok((await session.save(p)).bytes.length); session.dispose();
});

test("output budgets apply before allocation, include concurrent reservations and recover after disposal", async () => {
  const bytes = input();
  const small = await open(bytes, { sessionLimits: { maxPreparedOutputBytes: bytes.length - 1 } });
  await assert.rejects(plan(small), code("save-limit-exceeded")); small.dispose();
  const session = await open(bytes, { sessionLimits: { maxRetainedOutputBytes: bytes.length, maxLivePlans: 2 } });
  const results = await Promise.allSettled([session.prepareSave({ mode: "no-op" }), session.prepareSave({ mode: "no-op" })]);
  assert.equal(results.filter((r) => r.status === "fulfilled").length, 1);
  assert.equal(results.find((r) => r.status === "rejected").reason.code, "save-limit-exceeded");
  const ready = results.find((r) => r.status === "fulfilled").value.plan; ready.dispose();
  const next = await plan(session); assert.ok((await session.save(next)).bytes.length); session.dispose();
  const count = await open(bytes, { sessionLimits: { maxLivePlans: 1 } });
  const one = await plan(count); await assert.rejects(plan(count), code("save-limit-exceeded")); one.dispose(); assert.ok((await plan(count)).byteLength); count.dispose();
});

test("caller option mutation does not alter captured import limits, source signal or save mode", async () => {
  const source = input(), controller = new AbortController(), opts = { projection: "structure", sessionLimits: { maxLivePlans: 1 }, zipLimits: { maxInputBytes: source.length }, signal: controller.signal };
  const importing = importPptx(source, opts);
  opts.sessionLimits.maxLivePlans = 0; opts.zipLimits.maxInputBytes = 1; opts.projection = "other"; opts.signal = AbortSignal.abort();
  const session = await importing, saving = { mode: "no-op", signal: controller.signal }, pending = session.prepareSave(saving);
  saving.mode = "rebuild"; saving.signal = AbortSignal.abort();
  const result = await pending; assert.equal(result.status, "ready"); assert.equal(result.plan.report.mode, "no-op"); session.dispose();
});

test("aborted imports/preparations produce typed errors and do not affect other sessions", async () => {
  await assert.rejects(open(input(), { signal: AbortSignal.abort() }), code("operation-aborted"));
  const controller = new AbortController(), pending = open(input(), { signal: controller.signal }); controller.abort();
  await assert.rejects(pending, code("operation-aborted"));
  const session = await open(); await assert.rejects(session.prepareSave({ mode: "no-op", signal: AbortSignal.abort() }), code("operation-aborted"));
  assert.ok((await plan(session)).byteLength); session.dispose();
});

test("corrupt envelopes, bounded XML failures and invalid input never create a usable session", async () => {
  await assert.rejects(open(new Uint8Array([0, 1, 2])), (e) => code("import-invalid-package")(e) && e.cause.code === "archive-invalid");
  await assert.rejects(open(input(), { zipLimits: { maxInputBytes: 1 } }), code("import-limit-exceeded"));
  await assert.rejects(open(input(), { limits: { maxXmlNodes: 1 } }), code("import-limit-exceeded"));
  await assert.rejects(open(input(), { sessionLimits: { maxReportFindings: 1 } }), code("import-limit-exceeded"));
  const f = metadataFixture(); f.parts["deck/main.data"] = '<invalid';
  await assert.rejects(open(encodeMetadata(f)), (e) => code("import-invalid-package")(e) && e.cause.code === "xml-invalid");
  await assert.rejects(importPptx([], options), TypeError);
  await assert.rejects(open(new Uint8Array(new SharedArrayBuffer(16))), code("import-unsupported"));
});

test("macro/template/slideshow/signed policy and absent dimensions reject independently of filename", async () => {
  for (const type of ["application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml", "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml", "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"]) {
    const f = metadataFixture(); f.metadataTypes["deck/main.data"] = type;
    await assert.rejects(open(encodeMetadata(f)), code("import-unsupported"));
  }
  for (const type of ["application/vnd.openxmlformats-package.digital-signature-origin", "application/vnd.ms-office.vbaProject"]) {
    const f = metadataFixture(); f.parts["policy.bin"] = Uint8Array.from([1]); f.metadataTypes["policy.bin"] = type;
    await assert.rejects(open(encodeMetadata(f)), code("import-unsupported"));
  }
  const f = metadataFixture(); f.parts["deck/main.data"] = f.parts["deck/main.data"].replace(/<p:sldSz[^>]+\/>/, "");
  await assert.rejects(open(encodeMetadata(f)), code("import-unsupported"));
});

test("empty effective slide order stays empty rather than becoming an invented page", async () => {
  const f = metadataFixture(); f.parts["deck/main.data"] = f.parts["deck/main.data"].replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/, '<p:sldIdLst/>');
  const bytes = encodeMetadata(f), session = await open(bytes); assert.equal(session.document.slides.length, 0);
  equalBytes((await session.save(await plan(session))).bytes, bytes); session.dispose();
});

test("invalid projection/options/limits and ambiguous save modes are rejected explicitly", async () => {
  for (const opts of [undefined, null, {}, [], { projection: "semantic" }]) await assert.rejects(importPptx(input(), opts), TypeError);
  for (const value of [0, -1, 1.5, Infinity, NaN, undefined]) await assert.rejects(open(input(), { sessionLimits: { maxLivePlans: value } }), RangeError);
  await assert.rejects(open(input(), { sessionLimits: [] }), TypeError);
  const session = await open(); for (const opts of [undefined, null, {}, [], { mode: "save" }]) await assert.rejects(session.prepareSave(opts), TypeError);
  assert.ok(Object.isFrozen(DEFAULT_SESSION_LIMITS)); session.dispose();
});

test("multi-chunk prepared outputs retain every binary byte and failed copying releases reservations", async () => {
  const f = metadataFixture(), binary = new Uint8Array(2 * 1024 * 1024 + 17); let state = 123456789;
  for (let i = 0; i < binary.length; i++) { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; binary[i] = state & 255; }
  f.parts["large.bin"] = binary; f.metadataTypes["large.bin"] = "application/octet-stream";
  const bytes = encodeMetadata(f), session = await open(bytes, { sessionLimits: { maxLivePlans: 1, maxRetainedOutputBytes: bytes.length } });
  const controller = new AbortController(), pending = session.prepareSave({ mode: "no-op", signal: controller.signal });
  setTimeout(() => controller.abort(), 0); await assert.rejects(pending, code("operation-aborted"));
  const prepared = await plan(session); assert.ok(prepared.byteLength > 2 * 1024 * 1024);
  equalBytes((await session.save(prepared)).bytes, bytes); session.dispose();
});

test("option getters cannot change the document or lifecycle after the final save validation", async () => {
  const session = await open(), p = await plan(session);
  await assert.rejects(session.save(p, { get acceptLossIds() { session.document.addSlide({ id: "getter-edit" }); return []; } }), code("save-snapshot-stale"));
  session.document.removeSlide("getter-edit");
  const q = await plan(session);
  await assert.rejects(session.save(q, { get acceptLossIds() { q.dispose(); return []; } }), code("plan-disposed"));
  await assert.rejects(session.save(p, { get acceptLossIds() { session.dispose(); return []; } }), code("session-disposed"));
});

test("cyclic and non-record custom data cannot hang or bypass normalized freshness checks", async () => {
  const session = await open(), p = await plan(session), cyclic = {}; cyclic.self = cyclic;
  for (const data of [cyclic, new Map([["hidden", "value"]]), { value: new Date(0) }, { value: 1n }]) {
    session.document.addSlide({ id: "unsupported-state", customData: data });
    const prepared = await session.prepareSave({ mode: "no-op" });
    assert.equal(prepared.status, "blocked"); assert.equal(prepared.report.diagnostics[0].code, "save-invalid-document");
    await assert.rejects(session.save(p), code("save-invalid-document"));
    session.document.removeSlide("unsupported-state");
  }
  assert.ok((await session.save(p)).bytes.length); session.dispose();
});

test("normalized snapshot budgets reject oversized edits and recover without losing prepared plans", async () => {
  await assert.rejects(open(input(), { sessionLimits: { maxSnapshotValues: 1 } }), code("import-limit-exceeded"));
  const session = await open(undefined, { sessionLimits: { maxSnapshotValues: 2000 } }), p = await plan(session);
  session.document.addSlide({ id: "large-state", customData: { values: Array(3000).fill("value") } });
  await assert.rejects(session.prepareSave({ mode: "no-op" }), code("save-limit-exceeded"));
  await assert.rejects(session.save(p), code("save-limit-exceeded"));
  session.document.removeSlide("large-state"); assert.ok((await session.save(p)).bytes.length); session.dispose();
});

test("offset Buffer views snapshot only their selected bytes", async () => {
  const bytes = input(), padded = Buffer.concat([Buffer.alloc(13, 255), Buffer.from(bytes), Buffer.alloc(7, 255)]);
  const importing = open(padded.subarray(13, 13 + bytes.length)); padded.fill(0);
  const session = await importing; equalBytes((await session.save(await plan(session))).bytes, bytes); session.dispose();
});

test("shared acyclic subgraphs remain bounded rather than expanding exponentially during freshness comparison", () => {
  // A separate process bounds a future complexity regression even if it blocks
  // the JavaScript event loop (an ordinary test timer could not interrupt it).
  const index = new URL("../dist/index.js", import.meta.url).href, fixture = new URL("./metadata-fixture.mjs", import.meta.url).href;
  const output = execFileSync(process.execPath, ["--input-type=module", "-e", `
    import { importPptx } from ${JSON.stringify(index)};
    import { metadataFixture, encodeMetadata } from ${JSON.stringify(fixture)};
    let shared = { leaf: "value" };
    for (let i = 0; i < 60; i++) shared = { left: shared, right: shared };
    const session = await importPptx(encodeMetadata(metadataFixture()), { projection: "structure" });
    session.document.addSlide({ id: "dag", customData: shared });
    const result = await session.prepareSave({ mode: "no-op" });
    if (result.status !== "blocked" || result.report.diagnostics[0].code !== "save-document-changed") throw new Error("DAG state was rejected or ignored");
    session.document.removeSlide("dag");
    if ((await session.prepareSave({ mode: "no-op" })).status !== "ready") throw new Error("session failed to recover");
    session.dispose(); console.log("bounded DAG comparison passed");
  `], { timeout: 10000, encoding: "utf8", maxBuffer: 65536 });
  assert.match(output, /bounded DAG comparison passed/);
});

test("normalized comparison keeps optional fields distinct, ignores key insertion order and enforces its own work budget", async () => {
  const { equalState } = await import("../dist/import/state.js");
  assert.equal(equalState({ a: 1, b: [2, 3] }, { b: [2, 3], a: 1 }, 100), true);
  assert.equal(equalState({ optional: undefined }, {}, 100), false);
  assert.equal(equalState([1, 2], [2, 1], 100), false);
  assert.equal(equalState({ value: -0 }, { value: 0 }, 100), false);
  assert.throws(() => equalState({ values: Array(100).fill(1) }, { values: Array(100).fill(1) }, 50), code("save-limit-exceeded"));
});


test("comparison reserves child work before expanding wide records", async () => {
  const { equalState } = await import("../dist/import/state.js");
  let reads = 0;
  const left = {}, right = {};
  for (let i = 0; i < 1000; i++) {
    Object.defineProperty(left, `key${i}`, { enumerable: true, get() { reads++; return i; } });
    right[`key${i}`] = i;
  }
  assert.throws(() => equalState(left, right, 10), code("save-limit-exceeded"));
  assert.equal(reads, 0, "over-budget child values must not be expanded");
  assert.equal(equalState({ a: 1, b: 2 }, { a: 1, b: 2 }, 3), true);
  assert.throws(() => equalState({ a: 1, b: 2 }, { a: 1, b: 2 }, 2), code("save-limit-exceeded"));
});

test("session limit option getters are captured once before asynchronous import", async () => {
  let reads = 0;
  const session = await importPptx(input(), {
    projection: "structure",
    get sessionLimits() {
      reads++;
      if (reads > 1) throw new Error("sessionLimits was read repeatedly");
      return { maxLivePlans: 1 };
    },
  });
  assert.equal(reads, 1);
  const prepared = await plan(session);
  await assert.rejects(session.prepareSave({ mode: "no-op" }), code("save-limit-exceeded"));
  prepared.dispose();
  assert.equal((await session.prepareSave({ mode: "no-op" })).status, "ready");
  session.dispose();
});


test("Core asset lifecycle changes invalidate plans and exact cleanup restores the imported baseline", async () => {
  const bytes = input(), session = await open(bytes), prepared = await plan(session);
  session.document.registerAsset({ id: "new-asset", kind: "image", source: { type: "bytes", value: Uint8Array.of(1, 2, 3) } });
  await assert.rejects(session.save(prepared), code("save-snapshot-stale"));
  session.document.replaceAsset("new-asset", { kind: "image", source: { type: "bytes", value: Uint8Array.of(3, 2, 1) } });
  assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
  assert.equal(session.document.removeUnusedAssets({ retainAssetIds: ["new-asset"] }).length, 0);
  assert.equal(session.document.removeUnusedAssets().length, 1);
  equalBytes((await session.save(prepared)).bytes, bytes);
  session.dispose();
});
