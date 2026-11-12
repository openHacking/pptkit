import test from "node:test";
import assert from "node:assert/strict";
import { inspectPptx } from "../dist/index.js";
import { metadataFixture, encodeMetadata, custom, DT, MC, GUID } from "./metadata-fixture.mjs";
const read = (f, options) => inspectPptx(encodeMetadata(f), options);
async function invalid(key, body, change = () => {}, code = "presentation-invalid") {
  const f = metadataFixture(), part = f.add(key, body); change(f, part);
  await assert.rejects(read(f), (e) => { assert.equal(e.code, code); assert.ok(e.entry); return true; });
}

test("absent metadata stays absent; arbitrary filenames and MIME case resolve through package relationships", async () => {
  assert.deepEqual((await read(metadataFixture())).metadata, {});
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    for (const key of ["core", "extended", "custom"]) f.metadataTypes[f.add(key, "")] = f.metadataTypes[`inventory/${key}.data`].toUpperCase();
    const { metadata } = await read(f);
    for (const key of ["core", "extended", "custom"]) assert.deepEqual(metadata[key], { part: `/inventory/${key}.data`, relationshipId: key, properties: [] });
  }
});

test("all fifteen core properties preserve order, decoded whitespace and lexical dates", async () => {
  const names = ["category", "contentStatus", "lastModifiedBy", "lastPrinted", "revision", "version"], dc = ["creator", "description", "identifier", "language", "subject", "title"];
  const f = metadataFixture();
  f.add("core", names.map((n) => `<c:${n}> value &amp; <![CDATA[<text>]]></c:${n}>`).join("") + dc.map((n) => `<d:${n}/>`).join("") + '<t:created xsi:type="t:W3CDTF">2026-10-23T12:15:00+08:00</t:created><t:modified xsi:type="t:W3CDTF"/><c:keywords/>');
  const props = (await read(f)).metadata.core.properties;
  assert.deepEqual(props.map((p) => p.name), [...names, ...dc, "created", "modified", "keywords"]);
  assert.equal(props[0].value, " value & <text>"); assert.equal(props[12].value, "2026-10-23T12:15:00+08:00"); assert.equal(props[13].value, "");
});

test("keyword mixed content preserves language inheritance, resets, boundaries and aggregate text", async () => {
  const f = metadataFixture();
  f.add("core", '<c:keywords xml:lang="zh"> 前<c:value>中文</c:value>,<c:value xml:lang="en">word</c:value><c:value xml:lang="">x</c:value><![CDATA[ 后]]></c:keywords>');
  const property = (await read(f)).metadata.core.properties[0];
  assert.equal(property.value, " 前中文,wordx 后");
  assert.deepEqual(property.keywordSegments, [
    { kind: "text", value: " 前", language: "zh" }, { kind: "value", value: "中文", language: "zh" },
    { kind: "text", value: ",", language: "zh" }, { kind: "value", value: "word", language: "en" },
    { kind: "value", value: "x", language: "" }, { kind: "text", value: " 后", language: "zh" },
  ]);
});

test("core date QName uses namespace identity including default namespace and local rebinding", async () => {
  for (const attrs of [`xmlns:alias="${DT}" xsi:type=" alias:W3CDTF "`, `xmlns="${DT}" xsi:type="W3CDTF"`, `xmlns:t="urn:wrong" xmlns:q="${DT}" xsi:type="q:W3CDTF"`]) {
    const f = metadataFixture(); f.add("core", `<q:created xmlns:q="${DT}" ${attrs}>2026</q:created>`.replace(`xmlns:q="${DT}" xmlns:t`, "xmlns:t"));
    assert.equal((await read(f)).metadata.core.properties[0].value, "2026");
  }
});

test("core rejects duplicate/foreign properties, forbidden attributes and nested scalar markup", async () => {
  for (const body of ['<d:title/><d:title/>', '<c:title/>', '<c:unknown/>', '<d:title xml:lang="en"/>', '<d:title xsi:type="t:W3CDTF"/>', '<d:title><c:value/></d:title>', '<c:keywords><d:value/></c:keywords>', '<c:keywords><c:value extra="x"/></c:keywords>', 'text<d:title/>']) await invalid("core", body);
  await invalid("core", "", (f, p) => { f.parts[p] = f.core("", 'extra="x"'); });
});

test("core date types require a bound W3CDTF QName; MCE markup is forbidden", async () => {
  for (const attrs of ['', 'xsi:type="other:W3CDTF"', 'xsi:type="d:W3CDTF"', 'xsi:type="t:Other"', 'xsi:type="t:W3CDTF:bad"', 'xsi:type="t:W3CDTF" xml:lang="en"']) await invalid("core", `<t:created ${attrs}/>`);
  await invalid("core", "", (f, p) => { f.parts[p] = f.core("", `xmlns:mc="${MC}" xmlns:u="urn:unknown" mc:Ignorable="u"`); });
  const f = metadataFixture(); f.add("core", ""); f.parts["inventory/core.data"] = f.core("", `xmlns:mc="${MC}"`);
  assert.deepEqual((await read(f)).metadata.core.properties, []);
});

test("extended properties distinguish all scalar fields and opaque structured wrappers in both dialects", async () => {
  const strings = ["Template", "Manager", "Company", "PresentationFormat", "HyperlinkBase", "Application", "AppVersion"];
  const ints = ["Pages", "Words", "Characters", "Lines", "Paragraphs", "Slides", "Notes", "TotalTime", "HiddenSlides", "MMClips", "CharactersWithSpaces", "DocSecurity"];
  const bools = ["ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged"];
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("extended", strings.map((n) => `<${n}> A </${n}>`).join("") + ints.map((n) => `<${n}> -0001 </${n}>`).join("") + bools.map((n) => `<${n}>1</${n}>`).join("") + ['HeadingPairs', 'TitlesOfParts', 'HLinks'].map((n) => `<${n}><v:vector baseType="lpstr" size="1"><v:lpstr>X</v:lpstr></v:vector></${n}>`).join("") + '<DigSig><v:blob>AA==</v:blob></DigSig>');
    const properties = (await read(f)).metadata.extended.properties;
    assert.equal(properties.length, 27);
    assert.deepEqual(properties.slice(0, 7).map((p) => p.value), strings.map(() => ({ status: "scalar", kind: "string", value: " A " })));
    assert.deepEqual(properties.slice(7, 19).map((p) => p.value), ints.map(() => ({ status: "scalar", kind: "integer", value: " -0001 " })));
    assert.deepEqual(properties.slice(-4).map((p) => p.value.status), Array(4).fill("uninterpreted"));
    assert.equal(properties.at(-1).value.local, "blob");
  }
});

test("extended scalars and wrappers reject malformed shape, out-of-range integers and invalid booleans", async () => {
  for (const body of ['<Unknown/>', '<Company/><Company/>', '<Company extra="x"/>', '<Company><Company/></Company>', '<Slides>2147483648</Slides>', '<Slides>-2147483649</Slides>', '<Slides>1.0</Slides>', '<Slides/>', '<SharedDoc>TRUE</SharedDoc>', '<SharedDoc>\u00a0true</SharedDoc>', '<HeadingPairs/>', '<HeadingPairs><vector/></HeadingPairs>', '<DigSig><v:vector/></DigSig>', '<TitlesOfParts>x<v:vector/></TitlesOfParts>']) await invalid("extended", body);
});

test("custom values preserve source lexical types, optional names and link targets without object-key collisions", async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add("custom", custom(2, "lpwstr", " __proto__ &amp; ", 'name="__proto__" linkTarget="local-reference"') + custom(3, "ui8", "18446744073709551615", 'name="constructor"') + custom(4, "i8", "-9223372036854775808") + custom(5, "filetime", "2026-10-23T23:59:59.1234567+08:00") + custom(6, "bool", " false ") + custom(7, "decimal", "+000.2500") + custom(8, "r8", "1.234567890123456789e300") + custom(9, "empty") + custom(10, "null"));
    const values = (await read(f)).metadata.custom.properties;
    assert.equal(values[0].name, "__proto__"); assert.equal(values[0].linkTarget, "local-reference"); assert.equal(values[0].value.value, " __proto__ & ");
    assert.equal(values[1].value.value, "18446744073709551615"); assert.equal(values[2].value.value, "-9223372036854775808"); assert.ok(!("name" in values[2]));
    assert.deepEqual(values.map((p) => p.value.kind), ["string", "integer", "integer", "date", "boolean", "decimal", "real", "empty", "null"]);
    assert.equal(JSON.parse(JSON.stringify(values))[1].value.value, "18446744073709551615");
  }
});

test("all integer variant widths check exact signed and unsigned bounds without precision loss", async () => {
  for (const [type, bits, unsigned] of [["i1", 8], ["i2", 16], ["i4", 32], ["i8", 64], ["int", 32], ["ui1", 8, true], ["ui2", 16, true], ["ui4", 32, true], ["ui8", 64, true], ["uint", 32, true]]) {
    const upper = 1n << BigInt(unsigned ? bits : bits - 1), min = unsigned ? 0n : -upper, max = upper - 1n;
    const f = metadataFixture(); f.add("custom", custom(2, type, min) + custom(3, type, max) + custom(4, type, "+0000"));
    assert.deepEqual((await read(f)).metadata.custom.properties.map((p) => p.value.value), [String(min), String(max), "+0000"]);
    for (const bad of [min - 1n, max + 1n, "1e1", "2.0", "", "9".repeat(100)]) await invalid("custom", custom(2, type, bad));
  }
});

test("custom identity validates int32 pid, GUID, optional unique name and fmtid/pid pair", async () => {
  for (const body of [custom(2, "lpstr") + custom("+02", "lpstr"), custom(2, "lpstr", "", 'name="same"') + custom(3, "lpstr", "", 'name="same"'), custom("2147483648", "lpstr"), custom(2, "lpstr").replace(GUID, "bad"), custom(2, "lpstr").replace('pid="2"', ""), custom(2, "lpstr").replace(`fmtid="${GUID}"`, ""), custom(2, "lpstr", "", 'extra="x"')]) await invalid("custom", body);
  const f = metadataFixture(); f.add("custom", custom(-2147483648, "lpstr") + custom(2147483647, "lpstr") + custom(2, "lpstr") + custom(2, "lpstr").replace(GUID, "{00000000-0000-0000-0000-000000000000}"));
  assert.equal((await read(f)).metadata.custom.properties.length, 4);
});

test("custom opaque variants are explicitly uninterpreted and never flattened into scalar text", async () => {
  const types = ["vector", "array", "blob", "oblob", "cy", "error", "stream", "ostream", "storage", "ostorage", "vstream", "clsid"];
  const f = metadataFixture(); f.add("custom", types.map((type, i) => custom(i + 2, type, type === "vector" ? '<v:lpstr>not flattened</v:lpstr>' : "opaque")).join(""));
  const props = (await read(f)).metadata.custom.properties;
  assert.deepEqual(props.map((p) => p.value), types.map((local) => ({ status: "uninterpreted", local, namespaceUri: f.VT })));
});

test("custom rejects unknown or missing types, bad scalar lexemes and nested scalar content", async () => {
  for (const [type, value] of [["variant", ""], ["constructor", ""], ["bool", "yes"], ["decimal", "1e3"], ["r8", "+INF"], ["r4", "Infinity"], ["null", "x"], ["empty", "x"], ["lpstr", "<v:lpstr/>"]]) await invalid("custom", custom(2, type, value));
  for (const body of [custom(2, "lpstr").replace('<v:lpstr></v:lpstr>', ""), custom(2, "lpstr").replace('</property>', '<v:i4>1</v:i4></property>'), custom(2, "lpstr").replaceAll('v:lpstr', 'lpstr')]) await invalid("custom", body);
  const f = metadataFixture(); f.add("custom", ["NaN", "INF", "-INF", ".1", "1.", "-0", "1e-300"].map((v, i) => custom(i + 2, "r8", v)).join(""));
  assert.equal((await read(f)).metadata.custom.properties.length, 7);
});

test("metadata root relationships enforce singleton, type, dialect and whole internal targets", async () => {
  for (const key of ["core", "extended", "custom"]) {
    for (const change of [
      (f) => f.rootRelations.pop(),
      (f) => f.rootRelations.push({ ...f.rootRelations.at(-1), Id: "duplicate" }),
      (f) => { f.rootRelations.at(-1).Target += "#fragment"; },
      (f) => { Object.assign(f.rootRelations.at(-1), { Target: "https://example.invalid/metadata", TargetMode: "External" }); },
      (f, p) => { f.metadataTypes[p] = "application/xml"; },
      (f, p) => { f.parts["inventory/other.xml"] = f.parts[p]; f.metadataTypes["inventory/other.xml"] = f.metadataTypes[p]; },
      (f, p) => { f.parts[p] = '<wrong xmlns="urn:wrong"/>'; },
    ]) await invalid(key, "", change);
  }
  await invalid("extended", "", (f) => { f.rootRelations.at(-1).Type = "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties"; });
  await invalid("custom", "", (f, p) => { f.parts[p] = f.properties("extended", ""); });
});

test("metadata MCE uses explicit namespaces, reports projection and rejects nonignorable unknown content", async () => {
  const f = metadataFixture(); f.add("extended", "");
  f.parts["inventory/extended.data"] = f.properties("extended", '<u:unused/><mc:AlternateContent><mc:Choice Requires="u"><Company>wrong</Company></mc:Choice><mc:Fallback><Company>chosen</Company></mc:Fallback></mc:AlternateContent>', `xmlns:mc="${MC}" xmlns:u="urn:unknown" mc:Ignorable="u"`);
  const result = await read(f);
  assert.equal(result.metadata.extended.properties[0].value.value, "chosen"); assert.ok(result.markupCompatibility.some((r) => r.part === "/inventory/extended.data"));
  await invalid("extended", '<u:other xmlns:u="urn:unknown"/>');
});

test("metadata result mutation cannot affect concurrent or subsequent inspections", async () => {
  const f = metadataFixture(); f.add("core", '<d:title>original</d:title>');
  const bytes = encodeMetadata(f), results = await Promise.all([inspectPptx(bytes), inspectPptx(bytes)]);
  results[0].metadata.core.properties[0].value = "mutated";
  assert.equal(results[1].metadata.core.properties[0].value, "original");
  assert.equal((await inspectPptx(bytes)).metadata.core.properties[0].value, "original");
});

test("metadata uses source XML budgets including ignored branches and leaves later reads usable", async () => {
  const f = metadataFixture(); f.add("custom", custom(2, "lpstr", "x".repeat(5000)));
  await assert.rejects(read(f, { limits: { maxXmlTextCharacters: 100 } }), (e) => e.code === "package-limit-exceeded");
  assert.equal((await read(f)).metadata.custom.properties[0].value.value.length, 5000);
  f.parts["inventory/custom.data"] = f.properties("custom", `<mc:AlternateContent><mc:Choice Requires="u">${custom(2, "lpstr", "x".repeat(5000))}</mc:Choice><mc:Fallback>${custom(2, "lpstr", "chosen")}</mc:Fallback></mc:AlternateContent>`, `xmlns:mc="${MC}" xmlns:u="urn:unknown"`);
  await assert.rejects(read(f, { limits: { maxXmlTextCharacters: 100 } }), (e) => e.code === "package-limit-exceeded");
  assert.equal((await read(f)).metadata.custom.properties[0].value.value, "chosen");
});

test("large custom inventories and mixed keywords yield and honor cancellation inside metadata traversal", async () => {
  // Inject the same checkpoint contract as structureContext, after XML decoding, so this
  // specifically proves cancellation in metadata traversal rather than ZIP/XML parsing.
  const { readXml } = await import("../dist/xml/read-xml.js");
  const { customProperties } = await import("../dist/metadata/properties.js");
  const { coreProperties } = await import("../dist/metadata/core.js");
  const f = metadataFixture(), body = Array.from({ length: 300 }, (_, i) => custom(i + 2, "lpstr", `value${i}`)).join("");
  const { root } = await readXml(new TextEncoder().encode(f.properties("custom", body)), {}, true);
  let checkpoints = 0;
  const stop = new Error("cancel metadata");
  await assert.rejects(customProperties(root, f.VT, "/custom.xml", async () => { if (++checkpoints === 130) throw stop; }), (e) => e === stop);
  assert.equal(checkpoints, 130);
  const props = await customProperties(root, f.VT, "/custom.xml", async () => {});
  assert.equal(props.length, 300); assert.equal(props.at(-1).value.value, "value299");
  const core = await readXml(new TextEncoder().encode(f.core(`<c:keywords>${'<c:value>word</c:value>'.repeat(300)}</c:keywords>`)), {}, true);
  checkpoints = 0;
  await assert.rejects(coreProperties(core.root, "/core.xml", async () => { if (++checkpoints === 130) throw stop; }), (e) => e === stop);
  assert.equal((await coreProperties(core.root, "/core.xml", async () => {}))[0].keywordSegments.length, 300);
});
