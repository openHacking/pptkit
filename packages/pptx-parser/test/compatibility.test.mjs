import test from "node:test";
import assert from "node:assert/strict";
import { inspectMarkupCompatibility, inspectOpcPackage } from "../dist/index.js";
import { readXml } from "../dist/xml/read-xml.js";
import { captureConfiguration, MCE } from "../dist/compatibility/configuration.js";
import { processCompatibility } from "../dist/compatibility/process.js";
import { pack, REL } from "./opc-fixture.mjs";

const bytes = (s) => new TextEncoder().encode(s);
const doc = (body, attributes = "") => `<r xmlns="urn:base" xmlns:mc="${MCE}" xmlns:x="urn:x" xmlns:y="urn:y" ${attributes}>${body}</r>`;
const inspect = (xml, options = {}) => inspectMarkupCompatibility(bytes(xml), { understoodNamespaces: ["urn:base"], ...options });
const reject = (xml, code = "mce-invalid", options = {}) => assert.rejects(inspect(xml, { partName: "fixture.xml", ...options }), (error) => {
  assert.equal(error.code, code); assert.equal(error.entry, "fixture.xml"); return true;
});
async function project(xml, options = {}) {
  const input = await readXml(bytes(xml), {}, true), before = structuredClone(input.root);
  const result = await processCompatibility(input.root, captureConfiguration({ understoodNamespaces: ["urn:base"], ...options }));
  assert.deepEqual(input.root, before, "projection must not mutate the source tree");
  return { ...result, input: input.root };
}
function textContent(nodes) {
  return nodes.map((node) => node.kind === "element" ? textContent(node.children) : ["text", "cdata"].includes(node.kind) ? node.value : "").join("");
}

test("selects first fully understood Choice with stable source identity, not best or last choice", async () => {
  const xml = doc('<mc:AlternateContent><mc:Choice Requires="x y"><x:n>both</x:n></mc:Choice><mc:Choice Requires="x"><x:n>first</x:n></mc:Choice><mc:Choice Requires="x"><x:n>last</x:n></mc:Choice><mc:Fallback><n>fallback</n></mc:Fallback></mc:AlternateContent>');
  for (const [understoodNamespaces, expected, selected] of [
    [["urn:base"], "fallback", 9], [["urn:base", "urn:x"], "first", 5], [["urn:base", "urn:x", "urn:y"], "both", 3],
  ]) {
    const value = await project(xml, { understoodNamespaces });
    assert.equal(textContent(value.nodes), expected);
    assert.equal(value.report.decisions[0].selectedElementIndex, selected);
    assert.equal(value.report.output.elements, 2);
  }
});

test("root AlternateContent may produce zero, one or multiple roots", async () => {
  const root = `<mc:AlternateContent xmlns:mc="${MCE}" xmlns:x="urn:x"><mc:Choice Requires="x"><x:a/><x:b/></mc:Choice></mc:AlternateContent>`;
  assert.equal((await inspect(root)).output.roots.length, 0);
  assert.deepEqual((await inspect(root, { understoodNamespaces: ["urn:x"] })).output.roots.map((r) => r.local), ["a", "b"]);
  const withoutMatch = await inspect(doc('<mc:AlternateContent><mc:Choice Requires="x"><x:n/></mc:Choice></mc:AlternateContent>'));
  assert.equal(withoutMatch.output.elements, 1);
  assert.equal(withoutMatch.decisions[0].selectedElementIndex, undefined);
});

test("ignorable element/attribute removal and ProcessContent preserve selected mixed content in order", async () => {
  const xml = doc('A<x:drop><n>lost</n></x:drop>B<x:wrap x:flag="v">C<![CDATA[D]]><n>E</n></x:wrap>F', 'mc:Ignorable="x" mc:ProcessContent="x:wrap" x:flag="hidden"');
  const value = await project(xml);
  assert.equal(textContent(value.nodes), "ABCDEF");
  assert.deepEqual(value.report.decisions.map((d) => d.action), ["ignored-attribute", "ignored-element", "unwrapped-element"]);
  assert.equal(value.report.output.textCharacters, 6);
  assert.equal(value.report.output.elements, 2);
  assert.ok(value.nodes[0].attributes.every((a) => a.namespaceUri !== MCE && a.namespaceUri !== "urn:x"));
  const known = await project(xml, { understoodNamespaces: ["urn:base", "urn:x"] });
  assert.equal(textContent(known.nodes), "AlostBCDEF");
  assert.equal(known.report.decisions.length, 0);
});

test("URI-based inheritance handles aliases, prefix rebinding, additive declarations and scope exit", async () => {
  const xml = doc('<a xmlns:z="urn:x" mc:ProcessContent="z:wrap"><x:wrap><n>A</n></x:wrap><b xmlns:x="urn:y" mc:Ignorable="x"><x:gone/></b></a><x:wrap>B</x:wrap>', 'mc:Ignorable="x"');
  const value = await project(xml);
  assert.equal(textContent(value.nodes), "A");
  assert.deepEqual(value.report.decisions.map((d) => d.action), ["unwrapped-element", "ignored-element", "ignored-element"]);
  await reject(doc('<a xmlns:x="urn:y"><x:mustFail/></a>', 'mc:Ignorable="x"'), "mce-unsupported");
  assert.equal((await inspect(doc('<x:a>A</x:a><x:b>B</x:b>', 'mc:Ignorable="x" mc:ProcessContent="x:*"'))).output.textCharacters, 2);
});

test("MustUnderstand is checked on active ordinary/unwrapped/selected nodes, not discarded content", async () => {
  for (const xml of [doc("", 'mc:MustUnderstand="x"'), doc('<x:w mc:MustUnderstand="y"/>', 'mc:Ignorable="x" mc:ProcessContent="x:w"'), doc('<mc:AlternateContent mc:MustUnderstand="y"><mc:Choice Requires="x"/><mc:Fallback/></mc:AlternateContent>'), doc('<mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback mc:MustUnderstand="y"/></mc:AlternateContent>')]) await reject(xml, "mce-unsupported");
  const xml = doc('<x:drop mc:MustUnderstand="y"/><mc:AlternateContent><mc:Choice Requires="x" mc:MustUnderstand="y"><y:unknown/></mc:Choice><mc:Fallback><n>ok</n></mc:Fallback></mc:AlternateContent>', 'mc:Ignorable="x"');
  assert.equal(textContent((await project(xml)).nodes), "ok");
  assert.equal((await inspect(doc("", 'mc:MustUnderstand="x"'), { understoodNamespaces: ["urn:base", "urn:x"] })).output.elements, 1);
});

test("opaque application extension islands bypass MCE and unknown namespaces with detached copies", async () => {
  const xml = doc('<x:ext mc:MustUnderstand="y"><mc:Choice/><y:n mc:Unknown="uninterpreted">opaque</y:n></x:ext>', 'mc:Ignorable="x"');
  const result = await project(xml, { extensionElements: [{ namespaceUri: "urn:x", local: "ext" }] });
  assert.equal(textContent(result.nodes), "opaque");
  assert.deepEqual(result.report.decisions.map((d) => d.action), ["opaque-extension"]);
  assert.equal(result.report.output.elements, 4);
  result.nodes[0].children[0].children.length = 0;
  assert.equal(result.input.children[0].children.length, 2);
});

test("unknown active namespaces mismatch while no-namespace understanding is explicit", async () => {
  await reject(doc('<x:n/>'), "mce-unsupported");
  await reject(doc('<n x:attr="v"/>'), "mce-unsupported");
  await reject('<r/>', "mce-unsupported", { understandNoNamespace: false });
  await reject(doc('<n attr="v"/>'), "mce-unsupported", { understandNoNamespace: false });
  assert.equal((await inspect('<r xml:lang="en"/>')).output.elements, 1);
  assert.equal((await inspect(doc('<n/>', 'mc:Ignorable="" mc:ProcessContent=" \t " mc:MustUnderstand=""'))).output.elements, 2);
});

test("prefix/QName lists are validated including discarded branches, Unicode names and XML whitespace", async () => {
  for (const attributes of ['mc:Ignorable="missing"', 'mc:Ignorable="mc"', 'mc:MustUnderstand="x:y"', 'mc:Ignorable="x&#160;y"', 'mc:ProcessContent="x:a"', 'mc:Ignorable="x" mc:ProcessContent="x:a:b"', 'mc:Ignorable="x" mc:ProcessContent="*:a"', 'mc:Ignorable="x" mc:ProcessContent="x:"']) await reject(doc("", attributes));
  await reject(doc('<mc:AlternateContent><mc:Choice Requires="missing"/><mc:Fallback/></mc:AlternateContent>'));
  const unicode = `<r xmlns:mc="${MCE}" xmlns:新="urn:new" mc:Ignorable="新" mc:ProcessContent="新:内容"><新:内容>ok</新:内容></r>`;
  assert.equal((await inspect(unicode)).output.textCharacters, 2);
});

test("malformed structural MCE is rejected independently of whether a branch is selected", async () => {
  for (const body of [
    '<mc:Choice Requires="x"/>', '<mc:Fallback/>', '<mc:AlternateContent/>', '<mc:AlternateContent><mc:Fallback/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires=""/></mc:AlternateContent>', '<mc:AlternateContent><mc:Choice mc:Requires="x"/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback/><mc:Choice Requires="y"/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback/><mc:Fallback/></mc:AlternateContent>',
    '<mc:AlternateContent bad="1"><mc:Choice Requires="x"/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires="x" xml:lang="en"/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires="x" bad="1"/></mc:AlternateContent>',
    '<mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback Requires="y"/></mc:AlternateContent>',
    '<mc:AlternateContent>text<mc:Choice Requires="x"/></mc:AlternateContent>',
  ]) await reject(doc(body), body.includes('mc:Requires') ? "mce-unsupported" : "mce-invalid");
});

test("future AlternateContent children must be ignorable and actually ignored", async () => {
  const body = '<mc:AlternateContent><x:future/><mc:Choice Requires="y"/><mc:Fallback><n>fallback</n></mc:Fallback></mc:AlternateContent>';
  assert.equal(textContent((await project(doc(body, 'mc:Ignorable="x"'))).nodes), "fallback");
  await reject(doc(body));
  await reject(doc(body, 'mc:Ignorable="x"'), "mce-unsupported", { understoodNamespaces: ["urn:base", "urn:x"] });
  await reject(doc(body, 'mc:Ignorable="x" mc:ProcessContent="x:*"'), "mce-unsupported");
  assert.equal((await inspect(doc(body.replace('<x:future/>', '<x:future mc:Ignorable="x"/>')))).output.elements, 2);
});

test("unwrapping rejects inherited-context-changing XML attributes and unknown MCE policy rejects explicitly", async () => {
  for (const name of ["base", "lang", "space"]) await reject(doc(`<x:w xml:${name}="value"/>`, 'mc:Ignorable="x" mc:ProcessContent="x:w"'));
  for (const name of ["PreserveElements", "PreserveAttributes", "Unknown"]) await reject(doc("", `mc:${name}="x:a"`), "mce-unsupported");
  await reject(doc('<mc:Future/>'), "mce-unsupported");
});

test("XML limits apply to the complete source including discarded and opaque branches", async () => {
  const xml = doc('<x:drop>' + '<n/>'.repeat(100) + '</x:drop>', 'mc:Ignorable="x"');
  await reject(xml, "xml-limit-exceeded", { limits: { maxNodes: 50 } });
  await reject(xml, "xml-limit-exceeded", { limits: { maxNodes: 50 }, extensionElements: [{ namespaceUri: "urn:x", local: "drop" }] });
  await reject('<!DOCTYPE r><r/>', "xml-unsupported");
});

test("configuration and bytes are captured, failures isolate calls, and cancellation retains signal identity", async () => {
  const xml = bytes(doc('<x:a/>')), source = Uint8Array.from(xml), understoodNamespaces = ["urn:base", "urn:x"];
  const controller = new AbortController(), options = { understoodNamespaces, signal: controller.signal };
  const pending = inspectMarkupCompatibility(source, options);
  source.fill(0); understoodNamespaces.length = 0; options.signal = AbortSignal.abort();
  assert.equal((await pending).output.elements, 2);
  await assert.rejects(inspectMarkupCompatibility(xml, { understoodNamespaces: ["urn:base"], signal: AbortSignal.abort() }), { code: "operation-aborted" });
  const results = await Promise.allSettled([inspectMarkupCompatibility(xml, { understoodNamespaces: ["urn:base"] }), inspectMarkupCompatibility(xml, { understoodNamespaces: ["urn:base", "urn:x"] })]);
  assert.equal(results[0].status, "rejected"); assert.equal(results[1].status, "fulfilled");
});

test("iterative traversal handles deep trees and asynchronous cancellation during MCE processing", async () => {
  const deep = '<r>'.repeat(1200) + 'v' + '</r>'.repeat(1200);
  const value = await inspect(deep, { limits: { maxDepth: 1300 } });
  assert.equal(value.output.elements, 1200); assert.equal(value.output.textCharacters, 1);
  const parsed = await readXml(bytes(doc('<n/>'.repeat(3000))), {}, true);
  const controller = new AbortController();
  const pending = processCompatibility(parsed.root, captureConfiguration({ understoodNamespaces: ["urn:base"] }), "cancel.xml", controller.signal);
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, { code: "operation-aborted", entry: "cancel.xml" });
});

test("invalid application configurations are rejected before processing", async () => {
  for (const options of [undefined, null, [], {}, { understoodNamespaces: "urn:x" }, { understoodNamespaces: [""] }, { understoodNamespaces: [MCE] }, { understoodNamespaces: [42] }, { understoodNamespaces: [], understandNoNamespace: 1 }, { understoodNamespaces: [], extensionElements: [{}] }, { understoodNamespaces: [], extensionElements: [{ namespaceUri: MCE, local: "Choice" }] }, { understoodNamespaces: [], extensionElements: [{ namespaceUri: "urn:x", local: "a:b" }] }]) await assert.rejects(inspectMarkupCompatibility(bytes('<r/>'), options), TypeError);
});

test("OPC consumes effective MCE relationships, reports decisions, and budgets all source markup", async () => {
  const rels = `<Relationships xmlns="${REL}" xmlns:mc="${MCE}" xmlns:x="urn:future" mc:Ignorable="x"><x:ignored/><mc:AlternateContent><mc:Choice Requires="x"><Relationship Id="future" Type="urn:data" Target="absent.xml"/></mc:Choice><mc:Fallback><Relationship Id="current" Type="urn:data" Target="data.xml"/></mc:Fallback></mc:AlternateContent></Relationships>`;
  const input = pack({ "_rels/.rels": rels, "data.xml": '<data/>' });
  const result = await inspectOpcPackage(input);
  assert.equal(result.relationships.length, 1); assert.equal(result.relationships[0].id, "current");
  const part = result.parts.find((p) => p.relationshipPart);
  assert.equal(part.xml.elements, 7); assert.equal(part.markupCompatibility.output.elements, 2);
  assert.deepEqual(part.markupCompatibility.decisions.map((d) => d.action), ["ignored-element", "alternate-content"]);
  await assert.rejects(inspectOpcPackage(input, { limits: { maxXmlNodes: 5 } }), { code: "package-limit-exceeded" });
  await assert.rejects(inspectOpcPackage(pack({ "_rels/.rels": rels.replace('mc:Ignorable="x"', 'mc:MustUnderstand="x"'), "data.xml": '<data/>' })), { code: "mce-unsupported" });
});

test("nested alternatives match a fixed truth table across independent capability combinations", async () => {
  const xml = doc('<mc:AlternateContent><mc:Choice Requires="x"><mc:AlternateContent><mc:Choice Requires="y"><n>XY</n></mc:Choice><mc:Fallback><n>X</n></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><mc:AlternateContent><mc:Choice Requires="z"><n>Z</n></mc:Choice><mc:Fallback><n>F</n></mc:Fallback></mc:AlternateContent></mc:Fallback></mc:AlternateContent>', 'xmlns:z="urn:z"');
  const expected = ["F", "X", "F", "XY", "Z", "X", "Z", "XY"];
  for (let mask = 0; mask < 8; mask++) {
    const understoodNamespaces = ["urn:base", ...["urn:x", "urn:y", "urn:z"].filter((_, bit) => mask & (1 << bit))];
    const result = await project(xml, { understoodNamespaces });
    assert.equal(textContent(result.nodes), expected[mask]);
    assert.equal(result.report.decisions.length, 2, "only alternatives reached in the effective projection are reported");
  }
});
