import assert from "node:assert/strict";
import test from "node:test";
import { compareOpcPackages, DEFAULT_COMPARISON_LIMITS, PackageInspectionError } from "../dist/index.js";
import { basicParts, basicTypes, contentTypes, pack, relationships } from "./opc-fixture.mjs";
import { crc32, zip } from "./zip-fixture.mjs";

const xmlOptions = { profile: "opc-xml-v1" };
const deck = (xml = "<r/>", extra = {}) => pack({ "a.xml": xml, ...extra });
const compareXml = (a, b, options = {}) => compareOpcPackages(deck(a), deck(b), { ...xmlOptions, ...options });
const changed = (result, code, part) => result.differences.some((finding) => finding.code === code && (part === undefined || finding.part === part));

test("OPC payload comparison ignores ZIP ordering, compression, metadata, comments and empty directories", async () => {
  const entries = [{ name: "[Content_Types].xml", body: basicTypes() }, ...Object.entries(basicParts()).map(([name, body]) => ({ name, body }))];
  const a = zip(entries.map((entry) => ({ ...entry, method: 0 }))).bytes;
  const b = zip([{ name: "empty/", body: "", method: 0 }, ...entries.toReversed()], { comment: "different ZIP comment", reverseCentral: true }).bytes;
  // DOS date/time are irrelevant container metadata.
  b.writeUInt16LE(1, 10);
  const result = await compareOpcPackages(a, b);
  assert.equal(result.equal, true); assert.equal(result.profile, "opc-payload-v1");
  assert.equal(result.differenceCount, 0); assert.equal(result.truncated, false);
  assert.deepEqual(result.differences, []); assert.deepEqual(result.before, result.after);
  assert.equal(result.before.parts, 4); assert.equal(result.before.relationships, 2);
  assert.ok(result.workUnits > 0); assert.ok(Object.isFrozen(DEFAULT_COMPARISON_LIMITS));
});

test("every payload participates, including unreferenced binary resources and extensionless XML", async () => {
  const types = contentTypes([{ Extension: "bin", ContentType: "application/octet-stream" }], [{ PartName: "/xml", ContentType: "application/custom+xml" }]);
  const a = pack({ xml: "<r/>", "orphan.bin": Buffer.from([0, 1, 2]) }, types);
  const b = pack({ xml: "<r/>", "orphan.bin": Buffer.from([0, 1, 3]) }, types);
  const result = await compareOpcPackages(a, b, xmlOptions);
  assert.equal(result.equal, false);
  assert.deepEqual(result.differences, [{ code: "payload-changed", part: "/orphan.bin", byteOffset: 2 }]);
  const c = pack({ xml: "<r changed='1'/>", "orphan.bin": Buffer.from([0, 1, 2]) }, types);
  assert.ok(changed(await compareOpcPackages(a, c, xmlOptions), "xml-changed", "/xml"));
});

test("binary equality never relies on equal lengths or CRC32 collisions", async () => {
  const left = Buffer.from("plumless"), right = Buffer.from("buckeroo");
  assert.equal(left.length, right.length); assert.equal(crc32(left), crc32(right));
  const types = contentTypes([{ Extension: "bin", ContentType: "application/octet-stream" }]);
  for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
    const result = await compareOpcPackages(pack({ "a.bin": left }, types), pack({ "a.bin": right }, types), { profile });
    assert.equal(result.equal, false);
    assert.deepEqual(result.differences, [{ code: "payload-changed", part: "/a.bin", byteOffset: 0 }]);
  }
});

test("payload comparison reports the first differing byte and prefix length without exposing content", async () => {
  const types = contentTypes([{ Extension: "bin", ContentType: "application/octet-stream" }]);
  for (const [a, b, offset] of [[[], [1], 0], [[1], [], 0], [[1, 2], [1, 2, 3], 2], [[1, 2], [1, 3], 1]]) {
    const result = await compareOpcPackages(pack({ "a.bin": Buffer.from(a) }, types), pack({ "a.bin": Buffer.from(b) }, types));
    assert.deepEqual(result.differences, [{ code: "payload-changed", part: "/a.bin", byteOffset: offset }]);
  }
});

test("added/removed parts and changed content types are explicit and deterministically ordered", async () => {
  const before = pack({ "z.xml": "<r/>", "a.bin": "data" }, basicTypes());
  const after = pack({ "m.xml": "<r/>", "a.bin": "data" }, contentTypes([{ Extension: "xml", ContentType: "application/xml" }, { Extension: "bin", ContentType: "application/x-new" }]));
  const result = await compareOpcPackages(before, after);
  assert.deepEqual(result.differences.map(({ code, part }) => [code, part]), [
    ["payload-changed", "/[Content_Types].xml"], ["content-type-changed", "/a.bin"], ["part-added", "/m.xml"], ["part-removed", "/z.xml"],
  ]);
  const swapped = await compareOpcPackages(after, before);
  assert.ok(changed(swapped, "part-removed", "/m.xml"));
  assert.ok(changed(swapped, "part-added", "/z.xml"));
});

test("control stream declarations remain significant even when unused or mapping to equal effective types", async () => {
  const a = pack({ "a.xml": "<r/>" });
  const b = pack({ "a.xml": "<r/>" }, contentTypes([{ Extension: "xml", ContentType: "application/xml" }, { Extension: "unused", ContentType: "application/x-unused" }]));
  for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
    const result = await compareOpcPackages(a, b, { profile });
    assert.equal(result.equal, false); assert.equal(result.differences[0].part, "/[Content_Types].xml");
  }
});

test("relationship source-local identities, exact external targets, fragments and shared cycles are compared", async () => {
  const parts = {
    "a.xml": "<a/>", "b.xml": "<b/>",
    "_rels/.rels": relationships([{ Id: "same", Type: "urn:root", Target: "a.xml" }]),
    "_rels/a.xml.rels": relationships([{ Id: "same", Type: "urn:internal", Target: "b.xml#one" }, { Id: "external", Type: "urn:url", Target: "https://example.invalid/old", TargetMode: "External" }]),
    "_rels/b.xml.rels": relationships([{ Id: "same", Type: "urn:cycle", Target: "a.xml" }]),
  };
  const modified = { ...parts, "_rels/a.xml.rels": relationships([
    { Id: "same", Type: "urn:internal", Target: "b.xml#two" },
    { Id: "external", Type: "urn:url", Target: "https://example.invalid/new", TargetMode: "External" },
    { Id: "added", Type: "urn:shared", Target: "b.xml" },
  ]) };
  const result = await compareOpcPackages(pack(parts), pack(modified), xmlOptions);
  assert.equal(result.equal, false); assert.equal(result.before.relationships, 4); assert.equal(result.after.relationships, 5);
  assert.deepEqual(result.differences.filter((d) => d.relationshipId).map((d) => [d.code, d.relationshipId, d.part]), [
    ["relationship-added", "added", "/_rels/a.xml.rels"],
    ["relationship-changed", "external", "/_rels/a.xml.rels"],
    ["relationship-changed", "same", "/_rels/a.xml.rels"],
  ]);
  const reverse = await compareOpcPackages(pack(modified), pack(parts));
  assert.ok(reverse.differences.some((d) => d.code === "relationship-removed" && d.relationshipId === "added"));
});

test("graph-equivalent internal target spelling is conservative source change, never silently normalized", async () => {
  const a = basicParts(), b = basicParts();
  b["doc/_rels/main.xml.rels"] = relationships([{ Id: "data", Type: "urn:data", Target: "data.bin" }]);
  const result = await compareOpcPackages(pack(a, basicTypes()), pack(b, basicTypes()), xmlOptions);
  assert.ok(changed(result, "xml-changed", "/doc/_rels/main.xml.rels"));
  assert.ok(changed(result, "relationship-changed", "/doc/_rels/main.xml.rels"));
});

test("XML profile permits quote/attribute order/entity/CDATA/empty-element representation differences only when selected", async () => {
  const a = '<r b="2" a="A&amp;B"><t>A&amp;B😀</t><empty/></r>';
  const b = "<r a='A&#38;B' b='2'><t><![CDATA[A&B]]>&#x1F600;</t><empty></empty></r>";
  assert.equal((await compareXml(a, b)).equal, true);
  assert.equal((await compareOpcPackages(deck(a), deck(b))).equal, false);
  assert.equal((await compareXml("<r/>", "<r><![CDATA[]]></r>")).equal, true);
});

test("XML encoding and line-ending normalization have explicit representation semantics", async () => {
  const a = deck('<?xml version="1.0" encoding="utf-8"?><r a="one&#xA;two">A\r\nB</r>');
  const utf16 = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('<?xml version="1.0" encoding="utf-16"?><r a="one&#10;two">A\nB</r>', "utf16le")]);
  const b = pack({ "a.xml": utf16 });
  assert.equal((await compareOpcPackages(a, b, xmlOptions)).equal, true);
  assert.equal((await compareOpcPackages(a, b)).equal, false);
});

test("XML child order, indentation, text spaces, values, namespace bindings and QName prefixes remain significant", async () => {
  const pairs = [
    ["<r><a/><b/></r>", "<r><b/><a/></r>"],
    ["<r><a/></r>", "<r> <a/></r>"],
    ["<r>A B</r>", "<r>A  B</r>"],
    ['<r v="1"/>', '<r v="1.0"/>'],
    ['<r xmlns:q="urn:a" value="q:x"/>', '<r xmlns:q="urn:b" value="q:x"/>'],
    ['<q:r xmlns:q="urn:a"/>', '<p:r xmlns:p="urn:a"/>'],
    ['<r xmlns="urn:a"/>', '<r xmlns="urn:b"/>'],
    ['<r xmlns:q="urn:unused"/>', '<r/>'],
    ['<r xml:space="preserve"> </r>', '<r xml:space="default"> </r>'],
    ["<r>\u00a0</r>", "<r> </r>"],
    ["<r>é</r>", "<r>e\u0301</r>"],
  ];
  for (const [a, b] of pairs) {
    const result = await compareXml(a, b);
    assert.equal(result.equal, false, `${a} / ${b}`);
    assert.equal(result.differences[0].code, "xml-changed"); assert.match(result.differences[0].path, /^\$/);
  }
});

test("prolog/epilog comments and processing instructions and standalone declarations cannot disappear", async () => {
  for (const [a, b] of [
    ["<!--keep--><r/>", "<r/>"], ["<r/><!--keep-->", "<r/>"],
    ["<?app preserve?><r/>", "<?app changed?><r/>"], ["<r/><?app preserve?>", "<r/>"],
    ["<r><!--keep--></r>", "<r/>"],
    ['<?xml version="1.0" standalone="yes"?><r/>', '<?xml version="1.0" standalone="no"?><r/>'],
    ['<?xml version="1.0" standalone="no"?><r/>', "<r/>"],
  ]) assert.equal((await compareXml(a, b)).equal, false, a);
  assert.equal((await compareXml('<!--same--><?app same?><r/><!--end-->', '<!--same--><?app same?><r></r><!--end-->')).equal, true);
});

test("unselected MCE branches and unknown extensions remain part of source equality", async () => {
  const a = '<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:new"><mc:AlternateContent><mc:Choice Requires="x"><x:opaque hidden="one"/></mc:Choice><mc:Fallback><known/></mc:Fallback></mc:AlternateContent></r>';
  const result = await compareXml(a, a.replace('hidden="one"', 'hidden="two"'));
  assert.equal(result.equal, false); assert.equal(result.differences[0].code, "xml-changed");
});

test("logical part spelling and content-type spelling stay exact without non-ASCII or case folding", async () => {
  const result = await compareOpcPackages(deck("<r/>"), pack({ "A.xml": "<r/>" }));
  assert.ok(changed(result, "part-added", "/A.xml")); assert.ok(changed(result, "part-removed", "/a.xml"));
  const type = await compareOpcPackages(deck(), pack({ "a.xml": "<r/>" }, contentTypes([{ Extension: "xml", ContentType: "APPLICATION/XML" }])), xmlOptions);
  assert.ok(changed(type, "content-type-changed", "/a.xml"));
});

test("malformed identical inputs and corrupt candidates reject rather than returning equality or partial results", async () => {
  const broken = Buffer.from("not a zip");
  await assert.rejects(compareOpcPackages(broken, broken), (e) => e instanceof PackageInspectionError && e.code === "archive-invalid");
  const good = deck();
  const invalids = [
    [deck("<r>"), "xml-invalid"],
    [pack({ "a.xml": "<r/>", "_rels/.rels": relationships([{ Id: "r", Type: "urn:x", Target: "missing.xml" }]) }), "package-invalid"],
    [pack({ "a.xml": "<r/>", "_rels/.rels": relationships([{ Id: "r", Type: "urn:x", Target: "a.xml" }, { Id: "r", Type: "urn:x", Target: "a.xml" }]) }), "package-invalid"],
    [deck("<!DOCTYPE r><r/>"), "xml-unsupported"],
  ];
  for (const [bytes, code] of invalids) {
    await assert.rejects(compareOpcPackages(good, bytes), { code });
    await assert.rejects(compareOpcPackages(bytes, good, xmlOptions), { code });
  }
});

test("comparison captures both input buffers and options before yielding, including aliased buffers", async () => {
  const a = Buffer.from(deck()), b = Buffer.from(deck("<changed/>"));
  const options = { profile: "opc-xml-v1", comparisonLimits: { maxDifferences: 1 }, zipLimits: { maxEntries: 100 } };
  const promise = compareOpcPackages(a, b, options);
  a.fill(0); b.fill(0); options.profile = "bad"; options.comparisonLimits.maxDifferences = 0; options.zipLimits.maxEntries = 0;
  const result = await promise;
  assert.equal(result.equal, false); assert.equal(result.profile, "opc-xml-v1");
  const shared = Buffer.from(deck()); const same = compareOpcPackages(shared, shared); shared.fill(0);
  assert.equal((await same).equal, true);
});

test("bounded reports keep counting differences after truncation, with immutable detached output", async () => {
  const a = pack({ "a.xml": "<a/>", "b.xml": "<b/>", "c.xml": "<c/>" });
  const b = pack({ "a.xml": "<aa/>", "b.xml": "<bb/>", "c.xml": "<cc/>" });
  const result = await compareOpcPackages(a, b, { ...xmlOptions, comparisonLimits: { maxDifferences: 1 } });
  assert.equal(result.equal, false); assert.equal(result.differenceCount, 3);
  assert.equal(result.differences.length, 1); assert.equal(result.truncated, true);
  assert.equal(result.differences[0].part, "/a.xml");
  for (const object of [result, result.before, result.after, result.differences, result.differences[0]]) assert.ok(Object.isFrozen(object));
  assert.throws(() => { result.differences[0].part = "changed"; }, TypeError);
  // Findings truncation must not skip validation of later candidate parts.
  await assert.rejects(compareOpcPackages(a, pack({ "a.xml": "<aa/>", "b.xml": "<broken>" }), { comparisonLimits: { maxDifferences: 1 } }), { code: "xml-invalid" });
});

test("combined expanded-byte and exact work budgets are enforced without a partial report", async () => {
  const bytes = deck(), expanded = Buffer.byteLength(contentTypes()) + 4;
  assert.equal((await compareOpcPackages(bytes, bytes, { comparisonLimits: { maxTotalExpandedBytes: expanded * 2 } })).equal, true);
  await assert.rejects(compareOpcPackages(bytes, bytes, { comparisonLimits: { maxTotalExpandedBytes: expanded * 2 - 1 } }), { code: "archive-limit-exceeded" });
  await assert.rejects(compareOpcPackages(bytes, bytes, { comparisonLimits: { maxTotalExpandedBytes: expanded } }), { code: "package-limit-exceeded" });
  for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
    const result = await compareOpcPackages(bytes, bytes, { profile });
    assert.equal((await compareOpcPackages(bytes, bytes, { profile, comparisonLimits: { maxWorkUnits: result.workUnits } })).equal, true);
    await assert.rejects(compareOpcPackages(bytes, bytes, { profile, comparisonLimits: { maxWorkUnits: result.workUnits - 1 } }), { code: "package-limit-exceeded" });
  }
});

test("inspection limits apply to both inputs and external references cause no fetch", async () => {
  const a = deck(), b = pack(basicParts(), basicTypes());
  await assert.rejects(compareOpcPackages(a, b, { zipLimits: { maxEntries: 2 } }), { code: "archive-limit-exceeded" });
  await assert.rejects(compareOpcPackages(a, b, { limits: { maxXmlNodes: 1 } }), { code: "package-limit-exceeded" });
  const external = pack({ ...basicParts(), "_rels/.rels": relationships([{ Id: "r", Type: "urn:url", Target: "https://example.invalid/never", TargetMode: "External" }]) }, basicTypes());
  const original = globalThis.fetch; let calls = 0;
  globalThis.fetch = () => { calls++; throw new Error("network"); };
  try { assert.equal((await compareOpcPackages(external, external, xmlOptions)).equal, true); assert.equal(calls, 0); }
  finally { globalThis.fetch = original; }
});

test("invalid profile, options, input kinds and nonfinite/unsafe/nonpositive budgets reject", async () => {
  const bytes = deck();
  for (const input of [undefined, null, [], "zip", new ArrayBuffer(8)]) await assert.rejects(compareOpcPackages(bytes, input), TypeError);
  for (const options of [null, [], "options", { profile: "unknown" }, { comparisonLimits: [] }, { comparisonLimits: null }, { limits: [] }]) await assert.rejects(compareOpcPackages(bytes, bytes, options), TypeError);
  for (const key of Object.keys(DEFAULT_COMPARISON_LIMITS)) for (const value of [0, -1, .5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, undefined, "1"]) {
    await assert.rejects(compareOpcPackages(bytes, bytes, { comparisonLimits: { [key]: value } }), RangeError);
  }
});

test("pre-abort, in-flight cancellation and a later independent comparison have explicit outcomes", async () => {
  const bytes = deck("<r>" + "<x/>".repeat(1000) + "</r>");
  const aborted = new AbortController(); aborted.abort();
  await assert.rejects(compareOpcPackages(bytes, bytes, { signal: aborted.signal }), { code: "operation-aborted" });
  const controller = new AbortController();
  const promise = compareOpcPackages(bytes, bytes, { ...xmlOptions, signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(promise, { code: "operation-aborted" });
  assert.equal((await compareOpcPackages(bytes, bytes, xmlOptions)).equal, true);
});

test("deep XML comparison uses an iterative traversal under caller-selected XML depth budgets", async () => {
  const depth = 600, source = "<r>".repeat(depth) + "value" + "</r>".repeat(depth);
  const result = await compareXml(source, source, { xmlLimits: { maxDepth: depth } });
  assert.equal(result.equal, true);
  const changedResult = await compareXml(source, source.replace("value", "changed"), { xmlLimits: { maxDepth: depth } });
  assert.equal(changedResult.equal, false);
});
