import assert from "node:assert/strict";
import test from "node:test";
import { inspectOpcClosure, DEFAULT_CLOSURE_LIMITS, inspectOpcPackage, PackageInspectionError } from "../dist/index.js";
import { closureBytes, closureParts, closureTypes, sourceAlternates, alternateParts, MC } from "./closure-fixture.mjs";
import { contentTypes, pack, relationships, REL } from "./opc-fixture.mjs";
import { zip } from "./zip-fixture.mjs";

const inspect = (parts = closureParts(), options = {}) => inspectOpcClosure(closureBytes(parts), { roots: ["/a.xml"], ...options });
const names = (result) => result.parts.map((part) => part.name);
const union = ["/_rels/a.xml.rels", "/_rels/b.xml.rels", "/a.xml", "/b.xml", "/shared.bin"];

test("source closure follows internal shared/cyclic/self references and implicitly includes owned relationship sets", async () => {
  const result = await inspect();
  assert.equal(result.profile, "source-relationships-v1"); assert.deepEqual(result.roots, ["/a.xml"]);
  assert.deepEqual(names(result), union);
  assert.equal(result.relationships.length, 6); assert.equal(result.sourceRelationshipCount, 8);
  assert.equal(result.effectiveRelationshipCount, 8);
  assert.ok(result.relationships.every((r) => r.matchesEffectiveRelationship));
  assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: true });
  assert.deepEqual(result.findings, []);
  assert.equal(result.relationships.find((r) => r.id === "self").fragment, "fragment");
  assert.equal(result.relationships.find((r) => r.id === "external").targetPart, undefined);
  const parts = closureParts();
  const expectedSize = Buffer.byteLength(closureTypes()) + union.reduce((sum, name) => sum + Buffer.byteLength(parts[name.slice(1)]), 0);
  assert.equal(result.expandedBytes, expectedSize);
  assert.deepEqual(result.controlStream, { name: "/[Content_Types].xml", expandedBytes: Buffer.byteLength(closureTypes()) });
  assert.equal(result.parts.find((p) => p.name === "/a.xml").reason, "root");
  assert.equal(result.parts.find((p) => p.name === "/_rels/a.xml.rels").reason, "relationship-set");
  assert.deepEqual(result.parts.find((p) => p.name === "/b.xml").via, { source: "/a.xml", relationshipId: "b", elementIndex: 2 });
});

test("outside incoming owners and unreached parts are explicit rather than deletion authorization", async () => {
  const result = await inspect();
  assert.deepEqual(result.incomingRelationships.map((r) => [r.source, r.id, r.targetPart]), [
    ["/", "root", "/a.xml"], ["/outside.xml", "shared", "/shared.bin"],
  ]);
  assert.deepEqual(result.unreachedParts, ["/_rels/.rels", "/_rels/outside.xml.rels", "/orphan.bin", "/outside.xml"]);
  assert.ok(!Object.hasOwn(result, "canDelete"));
  const packageRoot = await inspect(closureParts(), { roots: ["/"] });
  assert.ok(names(packageRoot).includes("/_rels/.rels"));
  assert.ok(packageRoot.unreachedParts.includes("/orphan.bin"));
  assert.ok(packageRoot.incomingRelationships.some((r) => r.source === "/outside.xml"));
});

test("multiple and case-equivalent roots deduplicate without losing root ownership", async () => {
  const result = await inspect(closureParts(), { roots: ["/B.XML", "/a.xml", "/A.XML", "/shared.bin"] });
  assert.deepEqual(result.roots, ["/a.xml", "/b.xml", "/shared.bin"]);
  assert.deepEqual(names(result), union);
  assert.equal(new Set(names(result)).size, result.parts.length);
  for (const name of result.roots) {
    const part = result.parts.find((p) => p.name === name);
    assert.equal(part.reason, "root"); assert.equal(part.via, undefined);
  }
});

test("MCE source union includes dependencies from unselected branches as well as effective fallback", async () => {
  const result = await inspect(alternateParts());
  assert.deepEqual(names(result), ["/_rels/a.xml.rels", "/a.xml", "/active.bin", "/shadow.bin"]);
  assert.equal(result.sourceRelationshipCount, 2); assert.equal(result.effectiveRelationshipCount, 1);
  assert.deepEqual(result.relationships.map((r) => [r.id, r.matchesEffectiveRelationship, r.elementIndex]), [
    ["hidden", false, 4], ["selected", true, 6],
  ]);
  assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: true });
  assert.deepEqual(result.unreachedParts, []);
});

test("mutually exclusive source declarations with the same ID retain separate targets and locations", async () => {
  const parts = alternateParts('<Relationship Id="selected" Type="urn:asset" Target="shadow.bin"/>');
  const result = await inspect(parts);
  assert.deepEqual(result.relationships.map((r) => r.id), ["selected", "selected"]);
  assert.equal(new Set(result.relationships.map((r) => r.elementIndex)).size, 2);
  assert.equal(result.relationships.filter((r) => r.matchesEffectiveRelationship).length, 1);
  assert.ok(names(result).includes("/shadow.bin"));
  assert.deepEqual(result.findings, []);
  const identical = alternateParts('<Relationship Id="selected" Type="urn:asset" Target="active.bin"/>');
  assert.ok((await inspect(identical)).relationships.every((r) => r.matchesEffectiveRelationship), "value match does not imply selected branch identity");
});

test("coexisting inactive duplicate IDs are diagnosed without confusing mutually exclusive MCE branches", async () => {
  const choice = '<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"/><Relationship Id="hidden" Type="urn:asset" Target="active.bin"/>';
  const result = await inspect(alternateParts(choice));
  assert.ok(names(result).includes("/shadow.bin"));
  assert.equal(result.coverage.outgoingComplete, false);
  assert.equal(result.findings.filter((f) => f.code === "ambiguous-source-identity").length, 1);
  const parts = alternateParts('<Relationship Id="selected" Type="urn:asset" Target="shadow.bin"/>');
  parts["_rels/a.xml.rels"] = parts["_rels/a.xml.rels"].replace("<mc:AlternateContent>", '<Relationship Id="selected" Type="urn:asset" Target="shadow.bin"/><mc:AlternateContent>').replace('Id="selected" Type="urn:asset" Target="active.bin"', 'Id="fallback" Type="urn:asset" Target="active.bin"');
  const unconditional = await inspect(parts);
  assert.ok(unconditional.findings.some((f) => f.code === "ambiguous-source-identity"));
});

test("missing targets in inactive source markup downgrade coverage while valid effective inspection still passes", async () => {
  const parts = alternateParts('<Relationship Id="hidden" Type="urn:asset" Target="missing.bin"/>');
  const result = await inspect(parts);
  assert.equal((await inspectOpcPackage(closureBytes(parts))).relationships.length, 1);
  assert.equal(result.coverage.outgoingComplete, false);
  assert.equal(result.coverage.incomingComplete, true);
  assert.equal(result.findings[0].code, "missing-source-target");
  assert.equal(result.findings[0].source, "/a.xml");
  assert.equal(result.relationships.find((r) => r.id === "hidden").targetPart, undefined);
  assert.ok(result.unreachedParts.includes("/shadow.bin"));
});

test("missing selected targets still reject the package instead of returning a reassuring partial closure", async () => {
  const parts = alternateParts();
  delete parts["active.bin"];
  await assert.rejects(inspect(parts), (e) => e instanceof PackageInspectionError && e.code === "package-invalid");
});

test("invalid inactive fields still retain parseable dependencies and explain incomplete coverage", async () => {
  for (const choice of [
    '<Relationship Type="urn:asset" Target="shadow.bin"/>',
    '<Relationship Id="1invalid" Type="urn:asset" Target="shadow.bin"/>',
    '<Relationship Id="hidden" Target="shadow.bin"/>',
    '<Relationship Id="hidden" Type="bad type" Target="shadow.bin"/>',
  ]) {
    const result = await inspect(alternateParts(choice));
    assert.ok(names(result).includes("/shadow.bin"), choice);
    assert.equal(result.coverage.outgoingComplete, false);
    assert.ok(result.findings.some((f) => f.code === "invalid-source-relationship"));
  }
  for (const choice of [
    '<Relationship Id="hidden" Type="urn:asset" Target="../escape.bin"/>',
    '<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin" TargetMode="unknown"/>',
    '<Relationship Id="hidden" Type="urn:asset" Target="https://example.invalid/internal"/>',
    '<Relationship Id="hidden" Type="urn:asset" Target="_rels/a.xml.rels"/>',
  ]) {
    const result = await inspect(alternateParts(choice));
    assert.equal(result.coverage.outgoingComplete, false, choice);
    assert.ok(result.findings.some((f) => f.code === "invalid-source-relationship"));
  }
});

test("opaque ignored markup and attributes cannot claim complete source relationships", async () => {
  const parts = alternateParts();
  parts["_rels/a.xml.rels"] = sourceAlternates('<x:opaque><Relationship Id="nested" Type="urn:asset" Target="shadow.bin"/></x:opaque>');
  const result = await inspect(parts);
  assert.ok(names(result).includes("/shadow.bin"));
  assert.equal(result.coverage.outgoingComplete, false);
  assert.ok(result.findings.some((f) => f.code === "opaque-source-markup"));
  const withAttribute = await inspect(alternateParts('<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin" x:semantic="unknown"/>'));
  assert.equal(withAttribute.coverage.outgoingComplete, false);
  assert.equal(withAttribute.findings[0].code, "opaque-source-markup");
});

test("opaque outside source markup affects incoming coverage without hiding local outgoing proof", async () => {
  const parts = closureParts();
  parts["_rels/outside.xml.rels"] = `<Relationships xmlns="${REL}" xmlns:mc="${MC}" xmlns:x="urn:future" mc:Ignorable="x"><x:opaque/><Relationship Id="shared" Type="urn:asset" Target="shared.bin"/></Relationships>`;
  const result = await inspect(parts);
  assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: false });
  assert.equal(result.findings[0].source, "/outside.xml");
  assert.ok(result.incomingRelationships.some((r) => r.source === "/outside.xml"));
});

test("inactive nested relationship markup and unknown control text produce located findings", async () => {
  for (const choice of [
    '<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"><Relationship Id="nested" Type="urn:x" Target="active.bin"/></Relationship>',
    'unknown source text<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"/>',
    '<Relationships><Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"/></Relationships>',
  ]) {
    const result = await inspect(alternateParts(choice));
    assert.equal(result.coverage.outgoingComplete, false, choice);
    assert.ok(result.findings.length > 0);
    assert.ok(result.findings.every((f) => f.relationshipPart === "/_rels/a.xml.rels" && f.elementIndex > 0));
  }
});

test("Unicode logical names and relationship IDs retain namespace-independent source identity", async () => {
  const parts = {
    "dir/%E4%B8%AD.xml": "<r/>", "dir/payload%20one.bin": Buffer.from([1]),
    "dir/_rels/%E4%B8%AD.xml.rels": relationships([{ Id: "名称", Type: "urn:asset", Target: "payload%20one.bin" }]),
  };
  const result = await inspectOpcClosure(pack(parts, closureTypes()), { roots: ["/dir/中.xml"] });
  assert.deepEqual(result.roots, ["/dir/中.xml"]);
  assert.ok(names(result).includes("/dir/payload%20one.bin"));
  assert.equal(result.relationships[0].id, "名称");
});

test("external schemes, relative references and fragments are inventoried without loading", async () => {
  const original = globalThis.fetch; let calls = 0;
  globalThis.fetch = () => { calls++; throw new Error("network"); };
  try {
    const parts = { "a.xml": "<r/>", "_rels/a.xml.rels": relationships([
      { Id: "relative", Type: "urn:url", Target: "../other.pptx#x", TargetMode: "External" },
      { Id: "file", Type: "urn:url", Target: "file:///not-loaded", TargetMode: "External" },
      { Id: "script", Type: "urn:url", Target: "javascript:never()", TargetMode: "External" },
    ]) };
    const result = await inspect(parts);
    assert.equal(result.relationships.length, 3); assert.equal(calls, 0);
    assert.ok(result.relationships.every((r) => !r.targetPart && r.targetMode === "External"));
  } finally { globalThis.fetch = original; }
});

test("empty/binary/disconnected OPC packages have explicit root and unreached behavior", async () => {
  const empty = await inspectOpcClosure(pack({}, contentTypes([])), { roots: ["/"] });
  assert.deepEqual(empty.parts, []); assert.equal(empty.expandedBytes, Buffer.byteLength(contentTypes([])));
  const binary = await inspect({ "a.bin": Buffer.from([1]) }, { roots: ["/a.bin"] });
  assert.deepEqual(names(binary), ["/a.bin"]); assert.deepEqual(binary.relationships, []);
  await assert.rejects(inspect({}, { roots: ["/missing.xml"] }), { code: "package-invalid" });
  await assert.rejects(inspect(closureParts(), { roots: ["/_rels/.rels"] }), TypeError);
});

test("roots/limits/input are snapshotted before yielding and reports are deeply immutable", async () => {
  const bytes = Buffer.from(closureBytes()), options = { roots: ["/a.xml"], closureLimits: { maxRoots: 5 }, limits: { maxRelationships: 100 } };
  const promise = inspectOpcClosure(bytes, options);
  bytes.fill(0); options.roots[0] = "/missing.xml"; options.closureLimits.maxRoots = 0; options.limits.maxRelationships = 0;
  const result = await promise;
  assert.deepEqual(result.roots, ["/a.xml"]);
  for (const object of [result, result.roots, result.parts, result.parts[0], result.parts.find((p) => p.via).via, result.relationships, result.relationships[0], result.incomingRelationships, result.unreachedParts, result.findings, result.coverage, result.controlStream]) assert.ok(Object.isFrozen(object));
  assert.throws(() => { result.parts[0].name = "/mutated"; }, TypeError);
  assert.ok(Object.isFrozen(DEFAULT_CLOSURE_LIMITS));
});

test("exact closure part/expanded byte/work budgets pass and one-less limits reject", async () => {
  const result = await inspect();
  for (const [limit, value] of [["maxClosureParts", result.parts.length], ["maxClosureExpandedBytes", result.expandedBytes], ["maxWorkUnits", result.workUnits]]) {
    const exact = await inspect(closureParts(), { closureLimits: { [limit]: value } });
    assert.deepEqual(names(exact), names(result), limit);
    await assert.rejects(inspect(closureParts(), { closureLimits: { [limit]: value - 1 } }), { code: "package-limit-exceeded" }, limit);
  }
  await assert.rejects(inspect({}, { roots: ["/"], closureLimits: { maxClosureExpandedBytes: 1 } }), { code: "package-limit-exceeded" });
});

test("source relationship count is budgeted independently of effective MCE relationship count", async () => {
  const parts = alternateParts();
  assert.equal((await inspect(parts, { closureLimits: { maxSourceRelationships: 2 }, limits: { maxRelationships: 1 } })).sourceRelationshipCount, 2);
  await assert.rejects(inspect(parts, { closureLimits: { maxSourceRelationships: 1 } }), { code: "package-limit-exceeded" });
  const repeated = '<Relationship Id="same" Type="urn:asset" Target="shadow.bin"/>'.repeat(200);
  await assert.rejects(inspect(alternateParts(repeated), { closureLimits: { maxSourceRelationships: 50 } }), { code: "package-limit-exceeded" });
});

test("finding and raw root limits reject rather than truncating a seemingly complete graph", async () => {
  const parts = alternateParts('<x:one/><x:two/>');
  const result = await inspect(parts, { closureLimits: { maxFindings: 2 } });
  assert.equal(result.findings.length, 2);
  await assert.rejects(inspect(parts, { closureLimits: { maxFindings: 1 } }), { code: "package-limit-exceeded" });
  await assert.rejects(inspect(closureParts(), { roots: ["/a.xml", "/A.XML"], closureLimits: { maxRoots: 1 } }), { code: "package-limit-exceeded" });
});

test("output and first-discovery provenance are stable across ZIP orders/compression and root order", async () => {
  const entries = [{ name: "[Content_Types].xml", body: closureTypes() }, ...Object.entries(closureParts()).map(([name, body]) => ({ name, body }))];
  const bytes = zip(entries.toReversed().map((e) => ({ ...e, method: 0 })), { reverseCentral: true, comment: "different envelope" }).bytes;
  const a = await inspect(closureParts(), { roots: ["/a.xml", "/b.xml"] });
  const b = await inspectOpcClosure(bytes, { roots: ["/b.xml", "/a.xml"] });
  assert.deepEqual(a, b);
});

test("invalid roots/options/budgets and malformed package inputs fail without a partial result", async () => {
  const bytes = closureBytes();
  for (const options of [undefined, null, [], {}, { roots: [] }, { roots: "a" }, { roots: [1] }, { roots: ["/a.xml"], closureLimits: [] }, { roots: ["/a.xml"], closureLimits: null }]) await assert.rejects(inspectOpcClosure(bytes, options), TypeError);
  for (const root of ["relative.xml", "/../a.xml", "/a.xml#x"]) await assert.rejects(inspectOpcClosure(bytes, { roots: [root] }), { code: "package-invalid" });
  for (const key of Object.keys(DEFAULT_CLOSURE_LIMITS)) for (const value of [0, -1, .5, Infinity, NaN, undefined, "1", Number.MAX_SAFE_INTEGER + 1]) {
    await assert.rejects(inspectOpcClosure(bytes, { roots: ["/"], closureLimits: { [key]: value } }), RangeError);
  }
  await assert.rejects(inspectOpcClosure(new Uint8Array([0, 1]), { roots: ["/"] }), { code: "archive-invalid" });
  await assert.rejects(inspectOpcClosure(bytes, { roots: ["/"], limits: { maxXmlNodes: 1 } }), { code: "package-limit-exceeded" });
});

test("pre-abort/in-flight cancellation yield no inventory and independent retry succeeds", async () => {
  const bytes = closureBytes();
  const controller = new AbortController(); controller.abort();
  await assert.rejects(inspectOpcClosure(bytes, { roots: ["/"], signal: controller.signal }), { code: "operation-aborted" });
  const live = new AbortController();
  const parts = alternateParts('<Relationship Id="hidden" Type="urn:asset" Target="shadow.bin"/>'.repeat(1000));
  const promise = inspect(parts, { signal: live.signal });
  setTimeout(() => live.abort(), 0);
  await assert.rejects(promise, { code: "operation-aborted" });
  assert.equal((await inspect()).coverage.outgoingComplete, true);
});

test("independent graph-model oracle verifies closure, relationship-set inclusion and incoming boundaries", async () => {
  let seed = 0x75a09;
  const random = (n) => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  for (let sample = 0; sample < 32; sample++) {
    const count = 4 + random(9), model = Array.from({ length: count }, () => []);
    const parts = {};
    for (let i = 0; i < count; i++) {
      parts[`n${i}.bin`] = Buffer.from([i]);
      const length = random(5);
      for (let j = 0; j < length; j++) model[i].push({ id: `e${j}`, target: random(count) });
      if (length || random(2)) parts[`_rels/n${i}.bin.rels`] = relationships(model[i].map((e) => ({ Id: e.id, Type: "urn:model", Target: `n${e.target}.bin` })));
    }
    const roots = new Set([random(count), random(count)]);
    // Expected results come directly from the generated graph model, not Parser inventory.
    const reached = new Set(roots), pending = [...roots];
    while (pending.length) for (const edge of model[pending.pop()]) {
      if (!reached.has(edge.target)) { reached.add(edge.target); pending.push(edge.target); }
    }
    const expectedParts = [...reached].flatMap((i) => [`/n${i}.bin`, ...Object.hasOwn(parts, `_rels/n${i}.bin.rels`) ? [`/_rels/n${i}.bin.rels`] : []]).sort();
    const expectedIncoming = model.flatMap((list, i) => reached.has(i) ? [] : list.filter((e) => reached.has(e.target)).map((e) => `${i}/${e.id}/${e.target}`)).sort();
    const result = await inspectOpcClosure(pack(parts, closureTypes()), { roots: [...roots].map((i) => `/n${i}.bin`) });
    assert.deepEqual(names(result), expectedParts, `sample ${sample}`);
    assert.deepEqual(result.incomingRelationships.map((e) => `${e.source.slice(2, -4)}/${e.id}/${e.targetPart.slice(2, -4)}`).sort(), expectedIncoming, `incoming ${sample}`);
    assert.deepEqual(result.coverage, { outgoingComplete: true, incomingComplete: true });
  }
});

test("long cyclic dependency chains and deep inactive XML use bounded iterative traversal", async () => {
  const parts = {}, length = 300;
  for (let i = 0; i < length; i++) {
    parts[`n${i}.bin`] = Buffer.from([i % 256]);
    parts[`_rels/n${i}.bin.rels`] = relationships([{ Id: "next", Type: "urn:chain", Target: `n${(i + 1) % length}.bin` }]);
  }
  const result = await inspectOpcClosure(pack(parts, closureTypes()), { roots: ["/n0.bin"] });
  assert.equal(result.parts.length, length * 2); assert.equal(result.relationships.length, length);
  assert.deepEqual(result.unreachedParts, []);
  const depth = 300;
  const deep = alternateParts("<x:opaque>".repeat(depth) + '<Relationship Id="deep" Type="urn:asset" Target="shadow.bin"/>' + "</x:opaque>".repeat(depth));
  const nested = await inspect(deep, { xmlLimits: { maxDepth: depth + 10 } });
  assert.ok(names(nested).includes("/shadow.bin")); assert.equal(nested.findings.length, depth);
  assert.equal(nested.coverage.outgoingComplete, false);
});
