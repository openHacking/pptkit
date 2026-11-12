import assert from "node:assert/strict";
import test from "node:test";
import { inspectOpcPackage, inspectXmlPart, PackageInspectionError, DEFAULT_OPC_LIMITS } from "../dist/index.js";
import { zip } from "./zip-fixture.mjs";
import { pack, basicParts, basicTypes, contentTypes, relationships, CT, REL, REL_TYPE } from "./opc-fixture.mjs";
const inspect = (parts = basicParts(), types = basicTypes(), options) => inspectOpcPackage(pack(parts, types), options);
const rejects = (parts, code = "package-invalid", types = basicTypes(), options) => assert.rejects(inspect(parts, types, options), (e) => e instanceof PackageInspectionError && e.code === code);
const rootRel = (target, rest = {}) => relationships([{ Id: "r", Type: "urn:root", Target: target, ...rest }]);

test("OPC inventory resolves control streams, implicit relationship content types and targets", async () => {
  const result = await inspect();
  assert.equal(result.archive.entries.length, 5); assert.equal(result.parts.length, 4);
  assert.equal(result.xmlTotals.parts, 4); assert.equal(result.relationships.length, 2);
  assert.equal(result.relationships[0].source, "/"); assert.equal(result.relationships[0].targetPart, "/doc/main.xml");
  assert.equal(result.relationships[1].source, "/doc/main.xml"); assert.equal(result.relationships[1].targetPart, "/doc/data.bin");
  assert.equal(result.parts.find((p) => p.name === "/doc/data.bin").xml, undefined);
  assert.ok(result.parts.filter((p) => p.relationshipPart).every((p) => p.contentType === REL_TYPE));
  assert.equal(result.parts.find((p) => p.name === "/doc/main.xml").xml.root.namespaceUri, "urn:document");
  assert.ok(Object.isFrozen(DEFAULT_OPC_LIMITS));
});

test("OPC is not a PPTX claim; empty, binary and unknown XML packages have explicit inventories", async () => {
  const empty = await inspect({}, contentTypes([])); assert.equal(empty.parts.length, 0); assert.equal(empty.xmlTotals.parts, 1);
  const binary = await inspect({ "payload.bin": Buffer.from([1, 2]) }, contentTypes([{ Extension: "bin", ContentType: "application/octet-stream" }]));
  assert.equal(binary.xmlTotals.parts, 1); assert.equal(binary.relationships.length, 0);
  const xml = await inspect({ "unknown.custom": "<unknown xmlns='urn:unrecognized'/>" }, contentTypes([{ Extension: "custom", ContentType: "application/vendor.custom+xml" }]));
  assert.equal(xml.parts[0].xml.root.namespaceUri, "urn:unrecognized");
});

test("Default extension and Override names are ASCII-case-insensitive; overrides take precedence", async () => {
  const types = contentTypes([{ Extension: "XML", ContentType: "application/octet-stream" }], [{ PartName: "/DOC/MAIN.xml", ContentType: "application/custom+xml;charset=\"utf-8\"" }]);
  const result = await inspect({ "doc/main.xml": "<root/>", "unused.XML": Buffer.from([0, 1]) }, types);
  assert.equal(result.parts[0].xml.root.local, "root"); assert.equal(result.parts[1].xml, undefined);
  assert.equal(result.contentTypes.overrides[0].partName, "/doc/main.xml");
  await inspect({ "name": Buffer.from([1]) }, contentTypes([], [{ PartName: "/name", ContentType: "application/octet-stream" }]));
  await inspect({}, contentTypes([{ Extension: "x~+!", ContentType: "application/x-test" }]));
});

test("bad or missing content-type declarations reject without a partial inventory", async () => {
  await assert.rejects(inspectOpcPackage(zip([{ name: "data", body: "x" }]).bytes), { code: "package-invalid", entry: "[Content_Types].xml" });
  for (const types of [
    `<wrong xmlns="${CT}"/>`, `<Types/>`, `<Types xmlns="${CT}" bad="1"/>`,
    contentTypes([{ Extension: "xml", ContentType: "application/xml" }, { Extension: "XML", ContentType: "application/xml" }]),
    contentTypes([], [{ PartName: "/absent", ContentType: "application/xml" }]),
    contentTypes([], [{ PartName: "/doc/main.xml", ContentType: "application/xml" }, { PartName: "/DOC/main.XML", ContentType: "application/xml" }]),
    contentTypes([{ Extension: "a.b", ContentType: "application/xml" }]),
    contentTypes([{ Extension: "xml", ContentType: "not-a-type" }]),
    `<Types xmlns="${CT}"><Default Extension="xml" ContentType="application/xml"><child/></Default></Types>`,
    `<Types xmlns="${CT}"><Other/></Types>`,
    `<Types xmlns="${CT}">not whitespace<Default Extension="xml" ContentType="application/xml"/></Types>`,
  ]) await rejects(basicParts(), "package-invalid", types);
  await rejects({ "without-extension": "x" }, "package-invalid", contentTypes());
});

test("MIME parameter syntax is checked and OPC relationship types cannot have parameters", async () => {
  for (const type of ['application/x-custom;key="a;b\\\"c"', "application/x-custom; key=token", "APPLICATION/X-CUSTOM+XML"]) {
    const body = type.toLowerCase().endsWith("+xml") ? "<r/>" : "bytes";
    await inspect({ "a.data": body }, contentTypes([{ Extension: "data", ContentType: type }]));
  }
  for (const type of [' application/xml', 'application/xml ', 'application/xml;', 'application/xml; key=', 'application/xml;key="unterminated', 'application /xml']) await rejects({ "a.xml": "<r/>" }, "package-invalid", contentTypes([{ Extension: "xml", ContentType: type }]));
  await rejects({ "_rels/.rels": relationships() }, "package-invalid", contentTypes([{ Extension: "rels", ContentType: `${REL_TYPE};charset=utf-8` }]));
});

test("relative, absolute, fragment, shared and cyclic internal targets resolve without traversal recursion", async () => {
  const parts = { "d/a.xml": "<a/>", "d/b.xml": "<b/>", "_rels/.rels": rootRel("/D/a.XML"), "d/_rels/a.xml.rels": relationships([
    { Id: "self", Type: "urn:same", Target: "#section" },
    { Id: "b", Type: "urn:other", Target: "../d/./b.xml#frag%20value" },
    { Id: "shared", Type: "urn:other", Target: "/d/b.xml" },
  ]), "d/_rels/b.xml.rels": rootRel("a.xml") };
  const result = await inspect(parts);
  assert.equal(result.relationships.find((r) => r.id === "self").targetPart, "/d/a.xml");
  assert.equal(result.relationships.find((r) => r.id === "b").fragment, "frag%20value");
  assert.equal(result.relationships.at(-1).targetPart, "/d/a.xml");
  assert.equal(result.relationships.length, 5);
});

test("internal target traversal, schemes, malformed escaping, directories and missing targets reject", async () => {
  for (const target of ["../escape.xml", "../../d/a.xml", "https://example.invalid/a", "//host/a", "doc//main.xml", "doc/", "doc/..", "doc/main.xml?query", "doc/%6dain.xml", "doc/%2Fmain.xml", "doc/main.xml%", "missing.xml", "#fragment", "[Content_Types].xml", "_rels/.rels", "doc/main.xml#one#two"]) await rejects({ ...basicParts(), "_rels/.rels": rootRel(target) });
});

test("external references retain spelling and never resolve against the source or fetch", async () => {
  const original = globalThis.fetch; let calls = 0;
  globalThis.fetch = () => { calls++; throw new Error("Network access"); };
  try {
    const targets = ["https://example.invalid/a?q=1&b=2#x", "../other.pptx", "file:///nonexistent/example", "mailto:a@example.invalid", "javascript:alert(1)"];
    const result = await inspect({ "_rels/.rels": relationships(targets.map((Target, i) => ({ Id: `r${i}`, Type: "urn:external", Target, TargetMode: "External" }))) });
    assert.deepEqual(result.relationships.map((r) => r.target), targets);
    assert.ok(result.relationships.every((r) => r.targetPart === undefined)); assert.equal(calls, 0);
    for (const target of ["bad space", "bad%", "bad\\path", "1bad:scheme"]) await rejects({ "_rels/.rels": rootRel(target, { TargetMode: "External" }) });
  } finally { globalThis.fetch = original; }
});

test("part URI mapping supports encoded Unicode, escaped spaces and exact non-ASCII identity", async () => {
  const result = await inspect({ "ppt/%E4%B8%AD.xml": "<r/>", "ppt/a%20b.xml": "<s/>", "_rels/.rels": relationships([{ Id: "r", Type: "urn:x", Target: "ppt/中.xml" }, { Id: "s", Type: "urn:x", Target: "ppt/a%20b.xml" }]) });
  assert.equal(result.parts[0].name, "/ppt/中.xml"); assert.equal(result.parts[0].entryName, "ppt/%E4%B8%AD.xml");
  assert.equal(result.relationships[0].targetPart, "/ppt/中.xml"); assert.equal(result.relationships[1].targetPart, "/ppt/a%20b.xml");
  await rejects({ "ppt/中.xml": "<r/>" }, "package-unsupported");
  for (const name of ["a./x.xml", "a space.xml", "%61.xml", "a.%78ml"]) await rejects({ [name]: "<r/>" });
  await rejects({ "%C3%89.xml": "<r/>", "_rels/.rels": rootRel("é.xml") }); // ASCII case matching must not fold É to é.
});

test("relationship IDs use source-local xsd:ID identity with XML whitespace collapse", async () => {
  const result = await inspect({ ...basicParts(), "_rels/.rels": rootRel("doc/main.xml", { Id: "  名称  " }) });
  assert.equal(result.relationships[0].id, "名称");
  await inspect({ ...basicParts(), "_rels/.rels": rootRel("doc/main.xml", { Id: "data" }) }); // Same id in a different set is valid.
  for (const Id of ["1invalid", "a:b", "two words", "\u00a0id\u00a0"]) await rejects({ "_rels/.rels": rootRel("a.xml", { Id }), "a.xml": "<r/>" });
  await rejects({ "a.xml": "<r/>", "_rels/.rels": relationships([{ Id: "a", Type: "urn:x", Target: "a.xml" }, { Id: " a ", Type: "urn:x", Target: "a.xml" }]) });
});

test("relationship name/type/source agreements and forbidden endpoints are checked", async () => {
  await rejects({ "_rels/absent.xml.rels": relationships() });
  await rejects({ "_rels/.rels": relationships(), "_rels/_rels/.rels.rels": relationships() });
  await rejects({ "_rels/.rels": relationships() }, "package-invalid", contentTypes([{ Extension: "rels", ContentType: "application/xml" }]));
  await rejects({ "ordinary.xml": relationships() }, "package-invalid", contentTypes([{ Extension: "xml", ContentType: REL_TYPE }]));
  await rejects({ "a.xml": "<r/>", "_rels/a.xml.rels": relationships(), "_rels/.rels": rootRel("_rels/a.xml.rels") });
});

test("relationship control markup rejects invalid attributes and unprocessed extensions", async () => {
  for (const value of [
    `<wrong xmlns="${REL}"/>`, `<Relationships/>`, `<Relationships xmlns="${REL}" xml:base="elsewhere"/>`,
    `<Relationships xmlns="${REL}"><Wrong/></Relationships>`,
    `<Relationships xmlns="${REL}"><Relationship Id="r" Type="urn:x" Target="a.xml" bad="yes"/></Relationships>`,
    relationships([{ Id: "r", Type: "urn:x", Target: "a.xml", TargetMode: "external" }]),
    relationships([{ Id: "r", Target: "a.xml" }]),
  ]) await rejects({ "a.xml": "<r/>", "_rels/.rels": value });
  for (const value of [`<Relationships xmlns="${REL}" xmlns:m="urn:m" m:Ignorable="m"/>`, `<Relationships xmlns="${REL}"><x:Extension xmlns:x="urn:x"/></Relationships>`]) await rejects({ "_rels/.rels": value }, "package-unsupported");
  const text = `<Relationships xmlns="${REL}"><Relationship Id="r" Type="urn:x" Target="a.xml">opaque character content</Relationship></Relationships>`;
  assert.equal((await inspect({ "a.xml": "<r/>", "_rels/.rels": text })).relationships.length, 1);
});

test("all declared XML parts are checked, including unreferenced parts and extensionless XML", async () => {
  await rejects({ ...basicParts(), "unreferenced.xml": "<broken>" }, "xml-invalid");
  await rejects({ "part": "<broken>" }, "xml-invalid", contentTypes([], [{ PartName: "/part", ContentType: "application/custom+xml" }]));
  await rejects({ "a.xml": '<!DOCTYPE a><a/>' }, "xml-unsupported");
  await rejects({ "a.xml": "<a/>" }, "xml-invalid", "<Types>");
});

test("aggregate XML budgets count the types stream and enforce actual cumulative parsing", async () => {
  const parts = basicParts(), types = basicTypes(), result = await inspect(parts, types);
  const xmls = [types, ...Object.entries(parts).filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).map(([,value]) => value)];
  const summaries = await Promise.all(xmls.map((value) => inspectXmlPart(Buffer.from(value))));
  for (const metric of ["characters", "nodes", "textCharacters"]) assert.equal(result.xmlTotals[metric], summaries.reduce((sum, value) => sum + value[metric], 0));
  const limits = { maxXmlParts: result.xmlTotals.parts, maxXmlCharacters: result.xmlTotals.characters, maxXmlNodes: result.xmlTotals.nodes, maxXmlTextCharacters: result.xmlTotals.textCharacters, maxRelationships: result.relationships.length };
  await inspect(parts, types, { limits });
  for (const [name, value] of Object.entries(limits)) await rejects(parts, "package-limit-exceeded", types, { limits: { [name]: value - 1 } });
  await rejects(parts, "xml-limit-exceeded", types, { xmlLimits: { maxDepth: 1 } });
  await rejects(parts, "archive-limit-exceeded", types, { zipLimits: { maxEntries: 1 } });
  // The first huge unfinished text run must hit the aggregate source limit during parsing.
  await rejects({ "large.xml": "<r>" + "x".repeat(10000) }, "package-limit-exceeded", contentTypes(), { limits: { maxXmlCharacters: 2048 } });
});

test("options and bytes are captured before yielding, failures do not poison subsequent calls", async () => {
  const bytes = pack(basicParts(), basicTypes()), limits = { maxRelationships: 2 }, xmlLimits = { maxDepth: 64 }, options = { limits, xmlLimits };
  const pending = inspectOpcPackage(bytes, options); bytes.fill(0); limits.maxRelationships = 1; xmlLimits.maxDepth = 1;
  assert.equal((await pending).relationships.length, 2);
  const c = new AbortController(), abortOptions = { signal: c.signal };
  const cancelled = inspect(basicParts(), basicTypes(), abortOptions); abortOptions.signal = new AbortController().signal; c.abort();
  await assert.rejects(cancelled, { code: "operation-aborted" });
  assert.equal((await inspect()).relationships.length, 2);
});

test("configuration errors are not mistaken for package errors", async () => {
  for (const options of [null, [], { limits: null }, { xmlLimits: [] }, { zipLimits: false }]) await assert.rejects(inspect(basicParts(), basicTypes(), options), TypeError);
  for (const value of [0, -1, NaN, Infinity, 1.5]) await assert.rejects(inspect(basicParts(), basicTypes(), { limits: { maxRelationships: value } }), RangeError);
});

test("large cyclic relationship graphs are inventoried iteratively and asynchronously cancellable", async () => {
  const count = 600, parts = { "_rels/.rels": rootRel("nodes/0.xml") };
  for (let i = 0; i < count; i++) {
    parts[`nodes/${i}.xml`] = "<r/>";
    parts[`nodes/_rels/${i}.xml.rels`] = rootRel(`${(i + 1) % count}.xml`);
  }
  const result = await inspect(parts);
  assert.equal(result.relationships.length, count + 1);
  assert.equal(result.relationships.at(-1).targetPart, "/nodes/0.xml");
  const controller = new AbortController(), pending = inspect(parts, basicTypes(), { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, { code: "operation-aborted" });
});

test("control stream and central-directory ordering do not change source/target meaning", async () => {
  const parts = basicParts(), types = basicTypes();
  const records = [...Object.entries(parts).reverse().map(([name, body]) => ({ name, body })), { name: "[Content_Types].xml", body: types }, { name: "empty/", body: "", method: 0 }];
  const result = await inspectOpcPackage(zip(records, { reverseCentral: true }).bytes);
  const expected = await inspect();
  const sort = (rows) => [...rows].sort((a, b) => a.source.localeCompare(b.source) || a.id.localeCompare(b.id));
  assert.deepEqual(sort(result.relationships), sort(expected.relationships));
  assert.equal(result.parts.length, 4); assert.equal(result.xmlTotals.parts, 4);
});

test("interleaved and trash physical items are explicitly unsupported rather than treated as normal parts", async () => {
  await rejects({ "doc.xml/[0].last.piece": "<r/>" }, "package-unsupported");
  await rejects({ "[trash]/0000.dat": Buffer.from([1]) }, "package-unsupported");
  await assert.rejects(inspectOpcPackage(zip([{ name: "[Content_Types].xml/[0].last.piece", body: contentTypes() }]).bytes), { code: "package-unsupported" });
});

test("OPC-defined XML encoding declarations follow the narrower packaging contract", async () => {
  const encode = (xml, encoding, big = false) => {
    const bytes = Buffer.concat([Buffer.from([255, 254]), Buffer.from(`<?xml version="1.0" encoding="${encoding}"?>${xml}`, "utf16le")]);
    return big ? bytes.swap16() : bytes;
  };
  for (const big of [false, true]) {
    await inspect({ "_rels/.rels": encode(relationships(), "UTF-16", big) }, encode(contentTypes(), "UTF-16", big));
    await rejects({}, "package-invalid", encode(contentTypes(), big ? "UTF-16BE" : "UTF-16LE", big));
    await rejects({ "_rels/.rels": encode(relationships(), big ? "UTF-16BE" : "UTF-16LE", big) });
  }
  const custom = encode("<custom/>", "UTF-16LE");
  await inspect({ "custom.xml": custom }); // Ordinary XML still uses the generic XML profile.
  await rejects({ "core.xml": custom }, "package-invalid", contentTypes([{ Extension: "xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml" }]));
});
