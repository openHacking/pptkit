import assert from "node:assert/strict";
import test from "node:test";
import { inspectXmlPart, DEFAULT_XML_LIMITS, PackageInspectionError } from "../dist/index.js";
import { readXml } from "../dist/xml/read-xml.js";
const bytes = (xml) => Buffer.from(xml, "utf8");
const inspect = (xml, options) => inspectXmlPart(bytes(xml), options);
const rejects = (xml, code = "xml-invalid", options) => assert.rejects(inspect(xml, options), (error) => error instanceof PackageInspectionError && error.code === code);
function utf16(xml, bigEndian = false) {
  const body = Buffer.from(xml, "utf16le"), result = Buffer.concat([Buffer.from([255, 254]), body]);
  if (bigEndian) result.swap16();
  return result;
}

test("namespace-aware XML inventory counts mixed content without claiming OOXML validity", async () => {
  const xml = '<p:root xmlns:p="urn:p" id="a"><child/>A&amp;B<![CDATA[<>]]><!--note--><?pi body?></p:root>';
  const result = await inspect(xml);
  assert.deepEqual(result.root, { name: "p:root", prefix: "p", local: "root", namespaceUri: "urn:p" });
  assert.equal(result.encoding, "utf-8"); assert.equal(result.characters, xml.length);
  assert.equal(result.inputBytes, Buffer.byteLength(xml)); assert.equal(result.nodes, 6);
  assert.equal(result.elements, 2); assert.equal(result.attributes, 2); assert.equal(result.maxDepth, 2);
  assert.equal(result.textCharacters, 19); // namespace URI, id, text, CDATA, comment, PI body
  assert.deepEqual(result.namespaceUris, ["http://www.w3.org/2000/xmlns/", "urn:p"]);
  assert.ok(Object.isFrozen(DEFAULT_XML_LIMITS));
});

test("private tree resolves inherited/rebound names, keeps attribute namespaces and mixed order", async () => {
  const { root } = await readXml(bytes('<r xmlns="urn:a" xmlns:x="urn:x" x:id="1" id="2">a<x:c xmlns:x="urn:y"/>b<![CDATA[c]]><!--d--><?p e?></r>'), {}, true);
  assert.equal(root.namespaceUri, "urn:a");
  assert.equal(root.attributes.find((a) => a.name === "id").namespaceUri, "");
  assert.equal(root.attributes.find((a) => a.name === "x:id").namespaceUri, "urn:x");
  assert.deepEqual(root.children.map((child) => child.kind), ["text", "element", "text", "cdata", "comment", "processing-instruction"]);
  assert.equal(root.children[1].namespaceUri, "urn:y");
  assert.equal((await readXml(bytes("<r/>"))).root, undefined);
  const inherited = await readXml(bytes('<r xmlns="urn:r"><a/><b xmlns=""/></r>'), {}, true);
  assert.deepEqual(inherited.root.children.map((element) => element.namespaceUri), ["urn:r", ""]);
});

test("UTF-8 BOM and both BOM-marked UTF-16 byte orders produce equivalent inventories", async () => {
  const content = '<r name="中文">😀\r\n文字&#x41;</r>';
  const utf8 = await inspectXmlPart(Buffer.concat([Buffer.from([239, 187, 191]), bytes(content)]));
  for (const big of [false, true]) {
    const value = await inspectXmlPart(utf16(content, big));
    assert.deepEqual(value.root, utf8.root); assert.equal(value.textCharacters, utf8.textCharacters);
    assert.equal(value.characters, content.length); assert.equal(value.encoding, big ? "utf-16be" : "utf-16le");
    await inspectXmlPart(utf16('<?xml version="1.0" encoding="UTF-16"?><r/>', big));
    await inspectXmlPart(utf16(`<?xml version="1.0" encoding="UTF-16${big ? "BE" : "LE"}"?><r/>`, big));
  }
  assert.equal(utf8.characters, content.length);
  await inspect('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><r/>');
});

test("encoding errors and unsupported declarations never use replacement decoding", async () => {
  for (const data of [Buffer.from([0xc0, 0xaf]), Buffer.from([0xed, 0xa0, 0x80]), Buffer.from([0xff]), Buffer.concat([utf16("<r/>"), Buffer.from([0])]), utf16("<r>\ud800</r>")]) await assert.rejects(inspectXmlPart(data), (e) => e.code === "xml-invalid");
  for (const data of [Buffer.from([0xff, 0xfe, 0, 0]), Buffer.from([0, 0, 0xfe, 0xff]), Buffer.from("<r/>", "utf16le")]) await assert.rejects(inspectXmlPart(data), (e) => e.code === "xml-unsupported");
  for (const encoding of ["UTF-16", "ISO-8859-1", "windows-1252"]) await rejects(`<?xml version="1.0" encoding="${encoding}"?><r/>`, "xml-unsupported");
  await assert.rejects(inspectXmlPart(utf16('<?xml version="1.0" encoding="UTF-16BE"?><r/>')), (e) => e.code === "xml-unsupported");
  await rejects('<?xml version="1.1"?><r/>', "xml-unsupported");
});

test("malformed structure, names, references, XML characters and declarations reject", async () => {
  const invalid = ["", " ", "text", "<r>", "<r></x>", "<r/><s/>", "<r/>tail", "<1r/>", '<r a="1" a="2"/>', "<r a=1/>", "<r>&unknown;</r>", "<r>&#0;</r>", "<r>&#xD800;</r>", "<r>&#x110000;</r>", "<r>\0</r>", "<r>\u0001</r>", "<r>\ufffe</r>", "<r>]]></r>", "<r><!--a--b--></r>", "<r><![CDATA[x</r>", '<r a="<"/>', ' <?xml version="1.0"?><r/>', '<?xml encoding="UTF-8"?><r/>', '<r><?xml version="1.0"?></r>'];
  for (const xml of invalid) await rejects(xml);
  await inspect('<r>&lt;&gt;&amp;&apos;&quot;&#9;&#10;&#13;&#x1F600;</r>');
});

test("namespace well-formedness rejects unbound/reserved prefixes and expanded-attribute duplicates", async () => {
  for (const xml of ['<x:r/>', '<r x:a="1"/>', '<r xmlns:xml="urn:wrong"/>', '<r xmlns:xmlns="urn:wrong"/>', '<r xmlns:x=""/>', '<r xmlns:x="urn:a" xmlns:y="urn:a" x:a="1" y:a="2"/>', '<a:b:c xmlns:a="urn:a"/>', '<xmlns:r/>']) await rejects(xml);
  await inspect('<__proto__ xmlns:constructor="urn:constructor" constructor:toString="ok"/>');
});

test("DTD and external/user entities reject while inert PI/comment text never causes retrieval", async () => {
  const original = globalThis.fetch;
  let calls = 0;
  globalThis.fetch = () => { calls++; throw new Error("Unexpected network request"); };
  try {
    for (const xml of ['<!DOCTYPE r><r/>', '<!DOCTYPE r SYSTEM "https://example.invalid/dtd"><r/>', '<!DOCTYPE r [<!ENTITY x "expanded">]><r>&x;</r>', '<!DOCTYPE r [<!ENTITY x SYSTEM "file:///etc/passwd">]><r>&x;</r>', '<!DOCTYPE r [<!ENTITY % e SYSTEM "https://example.invalid/e">%e;]><r/>']) await rejects(xml, "xml-unsupported");
    await inspect('<?xml-stylesheet href="https://example.invalid/style"?><r><!--<!DOCTYPE r>--><![CDATA[<!ENTITY x SYSTEM "file:///ignored">]]></r>');
    assert.equal(calls, 0);
  } finally { globalThis.fetch = original; }
});

test("all budgets enforce declared units and boundary equality", async () => {
  const xml = '<r a="x"><c/>ab</r>', result = await inspect(xml);
  const exact = { maxInputBytes: bytes(xml).length, maxCharacters: xml.length, maxNodes: result.nodes, maxDepth: 2, maxAttributesPerElement: 1, maxTextCharacters: 3 };
  await inspect(xml, { limits: exact });
  for (const name of ["maxInputBytes", "maxCharacters", "maxNodes", "maxDepth", "maxTextCharacters"]) await rejects(xml, "xml-limit-exceeded", { limits: { [name]: exact[name] - 1 } });
  await rejects('<r xmlns:a="urn:a" a:x="1"/>', "xml-limit-exceeded", { limits: { maxAttributesPerElement: 1 } });
  await inspect("<r/>", { limits: { maxPendingCharacters: 3 } });
  await rejects("<root/>", "xml-limit-exceeded", { limits: { maxPendingCharacters: 2 } });
});

test("unfinished and completed large tokens are bounded without waiting for a closing delimiter", async () => {
  for (const xml of ["<" + "a".repeat(5000), '<r a="' + "a".repeat(5000), "<r><!--" + "a".repeat(5000), "<r><![CDATA[" + "a".repeat(5000), "<r><?p " + "a".repeat(5000), "<!DOCTYPE r [" + "a".repeat(5000), "<r>" + "a".repeat(5000)]) await rejects(xml, "xml-limit-exceeded", { limits: { maxPendingCharacters: 128 } });
  await rejects('<r a="' + "a".repeat(512) + '"/>', "xml-limit-exceeded", { limits: { maxPendingCharacters: 128 } });
  await rejects('<r>' + '<a/>'.repeat(100) + '</r>', "xml-limit-exceeded", { limits: { maxNodes: 100 } });
  await rejects('<r>'.repeat(100) + '</r>'.repeat(100), "xml-limit-exceeded", { limits: { maxDepth: 99 } });
});

test("malformed configurations fail before parsing", async () => {
  for (const options of [null, [], { limits: null }, { limits: [] }, { partName: 1 }]) await assert.rejects(inspect("<r/>", options), TypeError);
  for (const value of [0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) await assert.rejects(inspect("<r/>", { limits: { maxNodes: value } }), RangeError);
  await assert.rejects(inspectXmlPart("<r/>"), TypeError);
});

test("bytes, limits, signal identity and diagnostic label are captured before yielding", async () => {
  const input = bytes("<r><a/></r>"), limits = { maxNodes: 2 }, options = { limits, partName: "original.xml" };
  const pending = inspectXmlPart(input, options); input.fill(0); limits.maxNodes = 1; options.partName = "changed.xml";
  assert.equal((await pending).elements, 2);
  const controller = new AbortController(), abortOptions = { signal: controller.signal, partName: "original.xml" };
  const aborted = inspect("<r/>", abortOptions); abortOptions.signal = new AbortController().signal; abortOptions.partName = "changed.xml"; controller.abort();
  await assert.rejects(aborted, (e) => e.code === "operation-aborted" && e.entry === "original.xml");
});

test("abort works before parsing and during large flat documents; errors leave later calls usable", async () => {
  const early = new AbortController(); early.abort();
  await assert.rejects(inspect("<r/>", { signal: early.signal }), (e) => e.code === "operation-aborted");
  const controller = new AbortController();
  const pending = inspect('<r>' + '<c/>'.repeat(20000) + '</r>', { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, (e) => e.code === "operation-aborted");
  assert.equal((await inspect("<r/>")).elements, 1);
});

test("errors identify the part and a detached immutable XML cursor location", async () => {
  await assert.rejects(inspect('<r>\n<unclosed>\n</r>', { partName: "ppt/slides/slide1.xml" }), (error) => {
    assert.equal(error.code, "xml-invalid"); assert.equal(error.entry, "ppt/slides/slide1.xml");
    assert.equal(error.location.line, 3); assert.ok(error.location.column >= 1); assert.ok(error.location.offset > 0);
    assert.ok(Object.isFrozen(error.location)); return true;
  });
});

test("stream chunk boundaries preserve UTF characters, entities, namespace names and normalized text", async () => {
  for (const offset of [1017, 1018, 1019, 1020, 1021, 1022, 1023, 1024]) {
    const source = '<r>' + 'x'.repeat(offset - 3) + '😀&#x1F600;\r\n<![CDATA[中]]><p:c xmlns:p="urn:p"/></r>';
    const result = await readXml(bytes(source), {}, true);
    assert.equal(result.root.children[0].value, 'x'.repeat(offset - 3) + '😀😀\n');
    assert.equal(result.inspection.nodes, 4); assert.equal(result.root.children[2].namespaceUri, "urn:p");
  }
});


test("pending-source budgets use event cursors and never double-count completed chunks", async () => {
  const xml = '<r>' + '<c/>'.repeat(2000) + '</r>';
  const result = await inspect(xml, { limits: { maxPendingCharacters: 4 } });
  assert.equal(result.elements, 2001);
  assert.equal(result.characters, xml.length);
});


test("encoding BOM removal does not permit a second leading BOM", async () => {
  await assert.rejects(inspectXmlPart(Buffer.concat([Buffer.from([239, 187, 191]), bytes("\ufeff<r/>")])), (e) => e.code === "xml-invalid");
  for (const big of [false, true]) await assert.rejects(inspectXmlPart(utf16("\ufeff<r/>", big)), (e) => e.code === "xml-invalid");
  await inspect("<r>\ufeff</r>"); // A non-leading U+FEFF is ordinary character content.
});

test("XML attribute and end-of-line normalization retains numeric references distinctly", async () => {
  const { root } = await readXml(bytes('<r a="x\t\r\ny&#x9;&#xA;&#xD;">x\r\ny\rz</r>'), {}, true);
  assert.equal(root.attributes[0].value, "x  y\t\n\r");
  assert.equal(root.children[0].value, "x\ny\nz");
});

test("concurrent inspections isolate namespaces, failure state and cancellation", async () => {
  const controller = new AbortController();
  const large = '<a:r xmlns:a="urn:cancel">' + '<a:c/>'.repeat(12000) + '</a:r>';
  const operations = [
    inspect('<a:r xmlns:a="urn:first"><a:c/></a:r>'),
    inspect('<a:r xmlns:a="urn:second"><a:c/></a:r>'),
    inspect('<r><c/></r>', { limits: { maxDepth: 1 } }),
    inspect(large, { signal: controller.signal }),
  ];
  setTimeout(() => controller.abort(), 0);
  const results = await Promise.allSettled(operations);
  assert.equal(results[0].value.root.namespaceUri, "urn:first");
  assert.equal(results[1].value.root.namespaceUri, "urn:second");
  assert.equal(results[2].reason.code, "xml-limit-exceeded");
  assert.equal(results[3].reason.code, "operation-aborted");
  await rejects('<a:r/>');
});
