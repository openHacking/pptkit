import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync, existsSync } from "node:fs";
import { describeImportCapabilities, DEFAULT_IMPORT_CAPABILITY_LIMITS, importPptx } from "../dist/index.js";
import { diagramTextFixture } from "./diagram-text-fixture.mjs";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
const operations = ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"];
const finding = (values = {}) => ({ id: "a", feature: "drawing.text.mapped-fields", operation: "import", tier: "editable", source: { part: "/deck/slide.xml", path: "/sp[1]", sourceObjectId: "2" }, reason: "Explicit field only", ...values });
const input = (...findings) => ({ projection: "content", findings });
const normalized = report => { const { workUnits, ...result } = report; return result; };

test("seven explicit operation rows preserve exact subject/provenance, missing and conflicting observations", () => {
  const f = finding(), r = describeImportCapabilities(input(f, finding({ id: "b", operation: "preview", tier: "fallback" }), finding({ id: "c", tier: "unsupported", reason: "Unmapped neighbor" })));
  assert.equal(r.schemaVersion, "pptkit-import-capabilities-v1"); assert.equal(r.nativeVerification, "not-performed"); assert.equal(r.candidateSaveValidation, "not-performed");
  assert.equal(r.totalSubjects, 1); assert.equal(r.returnedFindings, 3); assert.equal(r.coverage, "reported-source-findings");
  const rows = r.subjects[0].operations; assert.deepEqual(rows.map(v => v.operation), operations);
  assert.equal(rows[0].status, "unreported"); assert.equal(rows[0].tier, undefined); assert.deepEqual(rows[0].observations, []);
  assert.equal(rows[1].status, "conflicting"); assert.equal(rows[1].tier, undefined); assert.deepEqual(rows[1].observations.map(v => v.id), ["a", "c"]);
  assert.equal(rows[3].status, "reported"); assert.equal(rows[3].tier, "fallback"); assert.deepEqual(r.subjects[0].source, f.source);
  assert.equal(r.subjects[0].references.artifacts[0].kind, "contract"); assert.equal(r.totals[1].conflicting, 1);
  const repeated = describeImportCapabilities(input(f, structuredClone(f))); assert.equal(repeated.subjects[0].operations[1].status, "reported"); assert.equal(repeated.subjects[0].operations[1].observations.length, 2);
});

test("all location/feature/Core identity fields participate without separator collisions or label matching", () => {
  const base = finding(); const variants = [base, ...["part", "path", "sourceObjectId", "relationshipId", "elementIndex"].map(key => finding({ source: { ...base.source, [key]: key === "elementIndex" ? 1 : key === "part" ? "/other.xml" : "other" } })), finding({ slideId: "s" }), finding({ elementId: "e" }), finding({ feature: "other" }), finding({ source: { ...base.source, path: "" } }), finding({ source: { part: base.source.part } })];
  const r = describeImportCapabilities(input(...variants)); assert.equal(r.totalSubjects, variants.length);
  const collision = describeImportCapabilities(input(finding({ feature: "a|b", source: { part: "/c" } }), finding({ feature: "a", source: { part: "/b|c" } })));
  assert.equal(collision.totalSubjects, 2); assert.ok(collision.subjects.every(s => s.references === undefined));
});

test("seeded shuffled observations obey independent per-subject operation oracle and stable rows", () => {
  let state = 0x5e2a37; const random = () => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state / 2 ** 32; };
  const subjects = Array.from({ length: 24 }, (_, i) => ({ feature: `future.${String(i).padStart(2, "0")}`, source: { part: `/slides/${i % 3}.xml`, path: `/item[${i + 1}]`, elementIndex: i + 1 }, elementId: `e-${i}` }));
  const rows = [];
  for (const subject of subjects) for (const operation of operations) for (let k = 0, n = Math.floor(random() * 4); k < n; k++) rows.push(finding({ ...subject, id: `f-${rows.length}`, operation, tier: ["editable", "preserved", "fallback", "unsupported"][Math.floor(random() * 4)], reason: `Reason ${rows.length}` }));
  const expected = describeImportCapabilities(input(...rows));
  for (let round = 0; round < 40; round++) {
    const shuffled = [...rows]; for (let i = shuffled.length - 1; i > 0; i--) { const k = Math.floor(random() * (i + 1)); [shuffled[k], shuffled[i]] = [shuffled[i], shuffled[k]]; }
    const report = describeImportCapabilities(input(...shuffled)); assert.deepEqual(normalized(report), normalized(expected));
    for (const subject of report.subjects) for (const result of subject.operations) {
      const observed = rows.filter(v => v.feature === subject.feature && v.operation === result.operation), tiers = new Set(observed.map(v => v.tier));
      assert.equal(result.status, !observed.length ? "unreported" : tiers.size > 1 ? "conflicting" : "reported");
      assert.equal(result.tier, tiers.size === 1 ? [...tiers][0] : undefined); assert.equal(result.observations.length, observed.length);
    }
  }
});

test("exact filters and whole-subject truncation retain complete global matched counts", () => {
  const rows = [finding(), finding({ feature: "unknown" }), finding({ feature: "unknown", source: { part: "/x" }, elementId: "control" })];
  const full = describeImportCapabilities(input(...rows)), limited = describeImportCapabilities(input(...rows), { limits: { maxSubjects: 1 } });
  assert.equal(limited.truncated, true); assert.equal(limited.totalSubjects, 3); assert.equal(limited.returnedSubjects, 1); assert.equal(limited.subjects[0].operations.length, 7); assert.deepEqual(limited.totals, full.totals);
  const zero = describeImportCapabilities(input(...rows), { limits: { maxSubjects: 0 } }); assert.deepEqual(zero.subjects, []); assert.equal(zero.matchedFindings, 3); assert.equal(zero.returnedFindings, 0); assert.equal(zero.truncated, true);
  const selected = describeImportCapabilities(input(...rows), { filter: { feature: "unknown", part: "/x", elementId: "control" } }); assert.equal(selected.totalSubjects, 1); assert.equal(selected.totalFindings, 3); assert.equal(selected.matchedFindings, 1);
  const missing = describeImportCapabilities(input(...rows), { filter: { slideId: "absent" } }); assert.equal(missing.totalSubjects, 0); assert.equal(missing.truncated, false); assert.equal(missing.nativeVerification, "not-performed");
});

test("future fields are ignored without executing accessors; known accessors and malformed enums/locations reject", () => {
  let calls = 0; const future = { get future() { calls++; throw new Error("must not run"); } }, f = { ...finding(), ...{ source: { ...finding().source } } };
  Object.defineProperty(f, "future", { get() { calls++; throw new Error("must not run"); } });
  const r = describeImportCapabilities({ ...input(f), ...{ future: 1 } }, { future, filter: { future } }); assert.equal(calls, 0); assert.equal(r.totalSubjects, 1);
  const known = finding(); Object.defineProperty(known, "reason", { get() { calls++; return "bad"; } }); assert.throws(() => describeImportCapabilities(input(known)), TypeError); assert.equal(calls, 0);
  const list = []; Object.defineProperty(list, "0", { get() { calls++; return finding(); }, configurable: true }); assert.throws(() => describeImportCapabilities({ projection: "content", findings: list }), TypeError); assert.equal(calls, 0);
  for (const invalid of [{ operation: "play" }, { tier: "supported" }, { source: { part: "relative" } }, { source: { part: "/", elementIndex: 0 } }, { source: { part: "/", path: null } }, { feature: "" }, { reason: 1 }, { id: "" }]) assert.throws(() => describeImportCapabilities(input(finding(invalid))), TypeError);
  assert.throws(() => describeImportCapabilities({ projection: "content", findings: Array(1) }), TypeError);
  assert.throws(() => describeImportCapabilities({ projection: "future", findings: [] }), TypeError);
  assert.throws(() => describeImportCapabilities(input(), { limits: { future: 1 } }), TypeError);
});

test("aggregate finding/text/work budgets fail exactly, including repeated long-prefix sorting; retry remains detached", () => {
  const simple = input({ id: "a", feature: "f", reason: "r", operation: "import", tier: "editable", source: { part: "/" } });
  assert.equal(describeImportCapabilities(simple, { limits: { maxTextCodeUnits: 4 } }).totalSubjects, 1);
  assert.throws(() => describeImportCapabilities(simple, { limits: { maxTextCodeUnits: 3 } }), e => e.code === "import-limit-exceeded");
  const original = describeImportCapabilities(simple); assert.equal(describeImportCapabilities(simple, { limits: { maxWorkUnits: original.workUnits } }).workUnits, original.workUnits);
  assert.throws(() => describeImportCapabilities(simple, { limits: { maxWorkUnits: original.workUnits - 1 } }), e => e.code === "import-limit-exceeded");
  assert.throws(() => describeImportCapabilities(input(finding(), finding()), { limits: { maxFindings: 1 } }), e => e.code === "import-limit-exceeded");
  const long = Array.from({ length: 40 }, (_, i) => finding({ id: String(i), feature: "x".repeat(2000) + i }));
  assert.throws(() => describeImportCapabilities(input(...long), { limits: { maxWorkUnits: 10000 } }), e => e.code === "import-limit-exceeded");
  assert.equal(describeImportCapabilities(input(...long)).totalSubjects, 40);
  // Counts-only skips retained-observation sorting. Reusing that measured budget
  // with a retained subject isolates long shared-prefix comparison work.
  const shared = input(...Array.from({ length: 100 }, (_, i) => finding({ id: "same", reason: "r".repeat(2000) + (100 - i) })));
  const countsOnly = describeImportCapabilities(shared, { limits: { maxSubjects: 0 } });
  assert.throws(() => describeImportCapabilities(shared, { limits: { maxSubjects: 1, maxWorkUnits: countsOnly.workUnits + 1000 } }), e => e.code === "import-limit-exceeded");
  assert.equal(describeImportCapabilities(shared).returnedFindings, 100);
  assert.throws(() => describeImportCapabilities(input(finding({ feature: "outside", tier: "future" })), { filter: { feature: "not-matched" } }), TypeError);
  for (const [key, value] of [["maxSubjects", -1], ["maxFindings", 0], ["maxWorkUnits", Infinity], ["maxTextCodeUnits", 1.5]]) assert.throws(() => describeImportCapabilities(simple, { limits: { [key]: value } }), RangeError);
  const detached = describeImportCapabilities(simple); simple.findings[0].source.part = "/changed"; assert.equal(detached.subjects[0].source.part, "/");
  assert.throws(() => { detached.subjects[0].operations[0].status = "reported"; }, TypeError); assert.ok(Object.isFrozen(DEFAULT_IMPORT_CAPABILITY_LIMITS));
});

test("reference hints resolve to actual repository tasks/artifacts without promoting native validation", () => {
  const tasks = readFileSync(new URL("../../../docs/planning/advanced-capability-todos.md", import.meta.url), "utf8") + readFileSync(new URL("../../../docs/planning/implementation-backlog.md", import.meta.url), "utf8");
  const features = ["source-part", "metadata.title", "presentation-dimensions", "first-slide-number", "slide-order-and-visibility", "section.label", "slide.transition", "smartart.cache-preview", "chart.data", "table.cells", "notes.body", "drawing.text", "master-artwork.shape", "artwork.background.unmapped-property"];
  const r = describeImportCapabilities(input(...features.map(feature => finding({ feature }))));
  for (const subject of r.subjects) {
    assert.ok(subject.references); for (const task of subject.references.tasks) assert.ok(tasks.includes(task), task);
    for (const artifact of subject.references.artifacts) assert.ok(existsSync(new URL("../../../" + artifact.path, import.meta.url)), artifact.path);
    assert.equal(subject.operations[1].tier, "editable"); assert.equal(subject.operations[6].status, "unreported");
  }
  assert.equal(r.nativeVerification, "not-performed"); assert.throws(() => r.subjects[0].references.tasks.push("forged"), TypeError);
  const similar = describeImportCapabilities(input(finding({ feature: 'artwork.backgroundish' }), finding({ feature: 'first-slide-number-extra' })));
  assert.ok(similar.subjects.every(subject => subject.references === undefined));
});

test("actual starting-number findings gain references without changing their operations", async () => {
  for (const strict of [false, true]) {
    const fixture = structureFixture(strict);
    fixture.parts['deck/main.data'] = fixture.parts['deck/main.data'].replace('<p:presentation ', '<p:presentation firstSlideNum="17" ');
    const session = await importPptx(encodeStructure(fixture), { projection: 'structure' });
    try {
      const original = session.report.findings.filter(f => f.feature === 'first-slide-number');
      assert.equal(original.length, 1);
      const view = describeImportCapabilities(session.report, { filter: { feature: 'first-slide-number' } });
      assert.equal(view.subjects.length, 1);
      assert.deepEqual(view.subjects[0].references.tasks, ['AUT-08.1.3']);
      assert.equal(view.subjects[0].operations[1].tier, original[0].tier);
      assert.equal(view.subjects[0].operations[6].status, 'unreported');
      assert.equal(view.nativeVerification, 'not-performed');
    } finally { session.dispose(); }
  }
});

test("real content/cache source reports remain unchanged and capability transport grants no save authority", async () => {
  for (const strict of [false, true]) {
    const f = diagramTextFixture(strict), bytes = f.bytes(), session = await importPptx(bytes, { projection: "content", diagramPreview: "cache", editSaveProfile: "text-runs-v3" });
    try {
      const before = JSON.stringify(session.report), r = describeImportCapabilities(session.report, { filter: { feature: "smartart.cache-preview" } });
      assert.equal(r.totalSubjects, 1); assert.equal(r.subjects[0].operations[6].tier, "unsupported"); assert.equal(r.subjects[0].operations[5].tier, "preserved");
      assert.equal(JSON.stringify(session.report), before); const p = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(p.plan)).bytes, new Uint8Array(bytes));
      assert.equal(describeImportCapabilities(session.report, { filter: { feature: "not-present" } }).matchedFindings, 0);
    } finally { session.dispose(); }
  }
});
