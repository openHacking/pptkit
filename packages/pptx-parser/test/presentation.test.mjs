import test from "node:test";
import assert from "node:assert/strict";
import { inspectPresentationPackage, PackageInspectionError } from "../dist/index.js";
import { pack, contentTypes, relationships } from "./opc-fixture.mjs";
const namespaces = { strict: "http://purl.oclc.org/ooxml/presentationml/main", transitional: "http://schemas.openxmlformats.org/presentationml/2006/main" };
const relTypes = { strict: "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument", transitional: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" };
function fixture({ dialect = "transitional", kind = "presentation", macro = false, path = "unusual/deck.data", type, root, entries, extra = {} } = {}) {
  type ??= macro ? `application/vnd.ms-powerpoint.${kind}.macroEnabled.main+xml` : `application/vnd.openxmlformats-officedocument.presentationml.${kind}.main+xml`;
  return pack({ [path]: root ?? `<q:presentation xmlns:q="${namespaces[dialect]}"/>`, "_rels/.rels": relationships(entries ?? [{ Id: "main", Type: relTypes[dialect], Target: path }]), ...extra }, contentTypes(undefined, [{ PartName: `/${path}`, ContentType: type }]));
}
const rejects = (bytes, code, entry) => assert.rejects(inspectPresentationPackage(bytes), (error) => {
  assert.ok(error instanceof PackageInspectionError); assert.equal(error.code, code);
  if (entry !== undefined) assert.equal(error.entry, entry);
  return true;
});

test("identifies all six main content types and both entry dialects without filename assumptions", async () => {
  for (const dialect of Object.keys(namespaces)) for (const [kind, ext] of [["presentation", "ppt"], ["template", "pot"], ["slideshow", "pps"]]) for (const macro of [false, true]) {
    const value = await inspectPresentationPackage(fixture({ dialect, kind, macro }));
    assert.equal(value.mainPart, "/unusual/deck.data"); assert.equal(value.mainRelationshipId, "main");
    assert.equal(value.dialect, dialect); assert.equal(value.kind, kind); assert.equal(value.macroEnabled, macro);
    assert.equal(value.suggestedExtension, `${ext}${macro ? "m" : "x"}`);
    assert.equal(value.package.parts.length, 2);
  }
});

test("namespace URI and relationship type determine dialect, not prefixes, paths or unused declarations", async () => {
  const root = `<presentation xmlns="${namespaces.strict}" xmlns:p="${namespaces.transitional}"/>`;
  const value = await inspectPresentationPackage(fixture({ dialect: "strict", path: "ppt/presentation.xml", root }));
  assert.equal(value.dialect, "strict");
  await rejects(fixture({ dialect: "strict", root: `<p:presentation xmlns:p="${namespaces.transitional}"/>` }), "presentation-invalid", "unusual/deck.data");
  await rejects(fixture({ root: `<presentation xmlns="${namespaces.strict}"/>` }), "presentation-invalid");
  await rejects(fixture({ root: `<p:presentation xmlns:p="${namespaces.transitional.toUpperCase()}"/>` }), "presentation-invalid");
});

test("requires one package entry even when candidates share the same target", async () => {
  const one = { Id: "a", Type: relTypes.transitional, Target: "unusual/deck.data" };
  await rejects(fixture({ entries: [] }), "presentation-invalid", "_rels/.rels");
  await rejects(fixture({ entries: [one, { ...one, Id: "b" }] }), "presentation-invalid");
  await rejects(fixture({ entries: [one, { ...one, Id: "b", Type: relTypes.strict }] }), "presentation-invalid");
  await rejects(fixture({ entries: [{ ...one, Type: one.Type.toUpperCase() }] }), "presentation-invalid");
  await rejects(fixture({ entries: [], extra: { "unusual/_rels/deck.data.rels": relationships([{ ...one, Target: "deck.data" }]) } }), "presentation-invalid");
});

test("external and fragment entry targets reject; ordinary external references remain inert", async () => {
  for (const entry of [
    { Target: "https://example.invalid/deck.pptx", TargetMode: "External" },
    { Target: "unusual/deck.data#page" }, { Target: "unusual/deck.data#" },
  ]) await rejects(fixture({ entries: [{ Id: "main", Type: relTypes.transitional, ...entry }] }), "presentation-invalid");
  const value = await inspectPresentationPackage(fixture({ entries: [
    { Id: "main", Type: relTypes.transitional, Target: "unusual/deck.data" },
    { Id: "link", Type: "urn:custom", Target: "file:///not-loaded", TargetMode: "External" },
  ] }));
  assert.equal(value.package.relationships[1].target, "file:///not-loaded");
});

test("unsupported main types and invalid roots never masquerade as presentations", async () => {
  for (const type of ["application/xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml; charset=utf-8"]) {
    await rejects(fixture({ type }), "presentation-unsupported", "unusual/deck.data");
  }
  for (const root of ["<presentation/>", `<p:sld xmlns:p="${namespaces.transitional}"/>`, `<p:Presentation xmlns:p="${namespaces.transitional}"/>`]) await rejects(fixture({ root }), "presentation-invalid");
  const value = await inspectPresentationPackage(fixture({ type: "APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.PRESENTATIONML.PRESENTATION.MAIN+XML" }));
  assert.equal(value.kind, "presentation");
});

test("classification intentionally does not claim schema validity, MCE processing or macro absence", async () => {
  const root = `<p:presentation xmlns:p="${namespaces.transitional}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><mc:AlternateContent><mc:Choice Requires="x"><x:anything/></mc:Choice><mc:Fallback/></mc:AlternateContent><p:unknown/></p:presentation>`;
  const value = await inspectPresentationPackage(fixture({ root }));
  assert.equal(value.dialect, "transitional");
  assert.ok(value.package.parts[0].xml.namespaceUris.includes("urn:unknown"));
  assert.equal(value.macroEnabled, false);
  assert.equal(value.package.parts[0].xml.elements, 6);
});

test("retains OPC failure codes, budgets, input capture and cancellation", async () => {
  const bytes = fixture(), source = Uint8Array.from(bytes), controller = new AbortController();
  const limits = { maxRelationships: 10 }, options = { limits, signal: controller.signal };
  const pending = inspectPresentationPackage(source, options);
  source.fill(0); limits.maxRelationships = 1; options.signal = AbortSignal.abort();
  assert.equal((await pending).mainPart, "/unusual/deck.data");
  await assert.rejects(inspectPresentationPackage(bytes, { limits: { maxXmlNodes: 1 } }), { code: "package-limit-exceeded" });
  await assert.rejects(inspectPresentationPackage(bytes, { signal: AbortSignal.abort() }), { code: "operation-aborted" });
  await assert.rejects(inspectPresentationPackage(bytes.slice(0, 20)), { code: "archive-invalid" });
  const value = await inspectPresentationPackage(bytes); value.package.parts.length = 0;
  assert.equal((await inspectPresentationPackage(bytes)).package.parts.length, 2);
});

test("ignores nested officeDocument relationships and unrelated parts", async () => {
  const value = await inspectPresentationPackage(fixture({ extra: {
    "decoy.xml": "<unrelated/>",
    "unusual/_rels/deck.data.rels": relationships([{ Id: "nested", Type: relTypes.strict, Target: "../decoy.xml" }]),
  } }));
  assert.equal(value.mainPart, "/unusual/deck.data"); assert.equal(value.package.parts.length, 4);
});
