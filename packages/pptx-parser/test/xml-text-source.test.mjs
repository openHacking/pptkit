import assert from "node:assert/strict";
import test from "node:test";
import { captureXmlTextSource, DEFAULT_XML_TEXT_EDIT_LIMITS } from "../dist/preservation/xml-text-source.js";
import { compareOpcPackages, importPptx, inspectOpcClosure } from "../dist/index.js";
import { contentFixture, body, shape } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { zip } from "./zip-fixture.mjs";
import { createHash } from "node:crypto";

const utf8 = (text) => Buffer.from(text);
const string = (bytes) => Buffer.from(bytes).toString();
const select = (source, expectedText, local = "t", namespaceUri = "urn:a") => {
  const element = source.elements.find((e) => e.local === local && e.namespaceUri === namespaceUri && e.text === expectedText);
  assert.ok(element, `Missing ${local} with ${JSON.stringify(expectedText)}`);
  return { elementIndex: element.elementIndex, namespaceUri, local, expectedText };
};
const edit = (source, from, to, local, ns) => ({ ...select(source, from, local, ns), text: to });
const code = (name) => (error) => { assert.equal(error.code, name); return true; };
const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
const repack = (parts) => zip([...parts].map(([name, bytes]) => ({ name, body: bytes }))).bytes;

test("character-only replacement retains every lexical byte outside the selected content", async () => {
  const before = '<?xml version="1.0" standalone="yes"?>\r\n<?keep raw?>\r\n<a:r xmlns:a="urn:a" xmlns:x="urn:future" x:raw = \'A &gt; B\'><!--prefix--><a:t xml:space = "preserve">old &amp; text</a:t><x:effect a:role="a:t"><![CDATA[<a:t>old & text</a:t>]]></x:effect><!--suffix--></a:r>\r\n<?end ok?>';
  const source = await captureXmlTextSource(utf8(before));
  const result = await source.replaceText([edit(source, "old & text", "new < & >\r\n🙂")]);
  const expected = before.replace('>old &amp; text</a:t>', '>new &lt; &amp; &gt;&#xD;\n🙂</a:t>');
  assert.equal(string(result.bytes), expected);
  assert.equal(result.changes.length, 1);
  const [{ sourceStart, sourceEnd }] = result.changes;
  assert.equal(before.slice(sourceStart, sourceEnd), "old &amp; text");
  assert.equal(string(result.bytes).slice(0, sourceStart), before.slice(0, sourceStart));
  assert.ok(Object.isFrozen(result.changes) && Object.isFrozen(result.changes[0]));
});

test("element indices include wrappers and namespace identity disambiguates identical lexical text", async () => {
  const text = '<r xmlns="urn:a" xmlns:x="urn:other"><t>same</t><x:t>same</x:t><box xmlns="urn:other"><t>same</t></box><t>same</t></r>';
  const source = await captureXmlTextSource(utf8(text));
  assert.deepEqual(source.elements.map((e) => e.elementIndex), [1, 2, 3, 4, 5, 6]);
  const result = await source.replaceText([{ elementIndex: 6, namespaceUri: "urn:a", local: "t", expectedText: "same", text: "last" }]);
  assert.equal(string(result.bytes), text.slice(0, -12) + "last</t></r>");
  await assert.rejects(source.replaceText([{ elementIndex: 3, namespaceUri: "urn:a", local: "t", expectedText: "same", text: "wrong" }]), code("xml-edit-conflict"));
});

test("raw parent identities distinguish equal text under separate runs and inactive branches", async () => {
  const xml = '<root xmlns:a="urn:a" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><a:p><!--kept--><a:r><a:t>same</a:t></a:r><a:r><a:t>same</a:t></a:r><mc:AlternateContent><mc:Choice Requires="a"><a:r><a:t>same</a:t></a:r></mc:Choice><mc:Fallback><a:r><a:t>same</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p></root>';
  const source = await captureXmlTextSource(utf8(xml));
  assert.deepEqual(source.elements.map(element => [element.elementIndex, element.parentElementIndex ?? null, element.local]), [
    [1, null, 'root'], [2, 1, 'p'], [3, 2, 'r'], [4, 3, 't'], [5, 2, 'r'], [6, 5, 't'],
    [7, 2, 'AlternateContent'], [8, 7, 'Choice'], [9, 8, 'r'], [10, 9, 't'], [11, 7, 'Fallback'], [12, 11, 'r'], [13, 12, 't'],
  ]);
  assert.ok(source.elements.every(Object.isFrozen));
  const output = await source.replaceText([{ elementIndex: 6, namespaceUri: 'urn:a', local: 't', expectedText: 'same', text: 'changed' }]);
  assert.equal(string(output.bytes), xml.replace('<a:r><a:t>same</a:t></a:r><mc:AlternateContent>', '<a:r><a:t>changed</a:t></a:r><mc:AlternateContent>'));
  source.dispose();
});

test("raw parent identities are stable across encodings and deep element nesting", async () => {
  const nesting = 80, xml = `<root xmlns="urn:a">${'<n>'.repeat(nesting)}<t>e\u0301🙂</t>${'</n>'.repeat(nesting)}</root>`;
  const cases = [Buffer.from(xml), Buffer.concat([Buffer.from([255, 254]), Buffer.from(xml, 'utf16le')]),
    Buffer.concat([Buffer.from([254, 255]), Buffer.from(xml, 'utf16le').swap16()])];
  for (const bytes of cases) {
    const source = await captureXmlTextSource(bytes);
    assert.equal(source.elements.length, nesting + 2);
    assert.equal(source.elements[0].parentElementIndex, undefined);
    for (let i = 1; i < source.elements.length; i++) assert.equal(source.elements[i].parentElementIndex, i);
    assert.equal(source.elements.at(-1).text, 'e\u0301🙂');
    const result = await source.replaceText([{ elementIndex: nesting + 2, namespaceUri: 'urn:a', local: 't', expectedText: 'e\u0301🙂', text: 'changed' }]);
    const reopened = await captureXmlTextSource(result.bytes);
    assert.deepEqual(reopened.elements.map(item => item.parentElementIndex), source.elements.map(item => item.parentElementIndex));
    reopened.dispose(); source.dispose();
  }
});

test("self-closing siblings do not change the raw parent stack", async () => {
  const source = await captureXmlTextSource(utf8('<r xmlns="urn:a"><a/><b><c/><d>text</d></b><e/></r>'));
  assert.deepEqual(source.elements.map(item => [item.local, item.parentElementIndex ?? null]), [
    ['r', null], ['a', 1], ['b', 1], ['c', 3], ['d', 3], ['e', 1],
  ]);
  source.dispose();
});

test("complete raw element ranges exclude BOM and preserve nested lexical markup", async () => {
  const xml = '<?xml version="1.0"?><r xmlns:a="urn:a">🙂<!-- <a:r/> --><a:r note="a > b"><a:t><![CDATA[x<y]]></a:t><a:br /></a:r><a:r><a:t>same</a:t></a:r></r>';
  const inputs = [Buffer.concat([Buffer.from([239, 187, 191]), Buffer.from(xml)]),
    Buffer.concat([Buffer.from([255, 254]), Buffer.from(xml, 'utf16le')]),
    Buffer.concat([Buffer.from([254, 255]), Buffer.from(xml, 'utf16le').swap16()])];
  const expected = [
    '<r xmlns:a="urn:a">🙂<!-- <a:r/> --><a:r note="a > b"><a:t><![CDATA[x<y]]></a:t><a:br /></a:r><a:r><a:t>same</a:t></a:r></r>',
    '<a:r note="a > b"><a:t><![CDATA[x<y]]></a:t><a:br /></a:r>', '<a:t><![CDATA[x<y]]></a:t>', '<a:br />',
    '<a:r><a:t>same</a:t></a:r>', '<a:t>same</a:t>',
  ];
  for (const input of inputs) {
    const source = await captureXmlTextSource(input);
    assert.deepEqual(source.elements.map(item => xml.slice(item.outerStart, item.outerEnd)), expected);
    for (const element of source.elements) if (element.parentElementIndex !== undefined) {
      const parent = source.elements[element.parentElementIndex - 1];
      assert.ok(parent.outerStart < element.outerStart && parent.outerEnd > element.outerEnd);
    }
    source.dispose();
  }
});

test("bounded raw element reads use the captured source and enforce identity, limits and disposal", async () => {
  const xml = '<root xmlns:a="urn:a"><a:r x="&gt;"><a:t>🙂</a:t><!--keep--></a:r></root>';
  const expected = '<a:r x="&gt;"><a:t>🙂</a:t><!--keep--></a:r>';
  const inputs = [Buffer.from(xml), Buffer.concat([Buffer.from([255, 254]), Buffer.from(xml, 'utf16le')]),
    Buffer.concat([Buffer.from([254, 255]), Buffer.from(xml, 'utf16le').swap16()])];
  for (const bytes of inputs) {
    const source = await captureXmlTextSource(bytes);
    bytes.fill(0);
    const request = { elementIndex: 2, namespaceUri: 'urn:a', local: 'r' };
    assert.equal(source.readElement(request), expected);
    assert.equal(source.readElement(request, { maxCharacters: expected.length }), expected);
    assert.throws(() => source.readElement(request, { maxCharacters: expected.length - 1 }), code('xml-edit-limit-exceeded'));
    assert.throws(() => source.readElement({ ...request, namespaceUri: 'wrong' }), code('xml-edit-conflict'));
    assert.throws(() => source.readElement({ ...request, elementIndex: 0 }), TypeError);
    assert.throws(() => source.readElement(request, { maxCharacters: 0 }), RangeError);
    const controller = new AbortController(); controller.abort();
    assert.throws(() => source.readElement(request, { signal: controller.signal }), code('operation-aborted'));
    source.dispose();
    assert.throws(() => source.readElement(request), code('xml-source-disposed'));
  }
  const source = await captureXmlTextSource(utf8(xml));
  const hostile = { get elementIndex() { source.dispose(); return 2; }, namespaceUri: 'urn:a', local: 'r' };
  assert.throws(() => source.readElement(hostile), code('xml-source-disposed'));
  const other = await captureXmlTextSource(utf8(xml));
  assert.throws(() => other.readElement({ elementIndex: 2, namespaceUri: 'urn:a', local: 'r' }, { signal: { get aborted() { other.dispose(); return false; } } }), code('xml-source-disposed'));
});

test("run split preflight binds a grapheme boundary to direct raw ancestry without writing", async () => {
  const xml = '<a:p xmlns:a="urn:a"><a:r><a:t>Ae\u0301🙂</a:t></a:r></a:p>';
  const source = await captureXmlTextSource(utf8(xml));
  const request = { elementIndex: 3, namespaceUri: 'urn:a', expectedText: 'Ae\u0301🙂', offset: 3 };
  const plan = await source.preflightRunSplit(request);
  assert.deepEqual(plan, { left: 'Ae\u0301', right: '🙂', runElementIndex: 2, paragraphElementIndex: 1 });
  assert.ok(Object.isFrozen(plan));
  assert.equal(source.readElement({ elementIndex: 1, namespaceUri: 'urn:a', local: 'p' }), xml);
  await assert.rejects(source.preflightRunSplit({ ...request, offset: 2 }), { code: 'split-invalid' });
  await assert.rejects(source.preflightRunSplit({ ...request, expectedText: 'wrong' }), { code: 'xml-edit-conflict' });
  await assert.rejects(source.preflightRunSplit({ ...request, namespaceUri: 'urn:wrong' }), { code: 'xml-edit-conflict' });
  await assert.rejects(source.preflightRunSplit(request, { maxCharacters: 2 }), { code: 'split-limit-exceeded' });
  const blocked = new AbortController(); blocked.abort();
  await assert.rejects(source.preflightRunSplit(request, { signal: blocked.signal }), { code: 'operation-aborted' });
  source.dispose();
  await assert.rejects(source.preflightRunSplit(request), { code: 'xml-source-disposed' });
  for (const invalid of [
    '<a:p xmlns:a="urn:a"><a:fld><a:t>ab</a:t></a:fld></a:p>',
    '<a:p xmlns:a="urn:a"><wrapper><a:r><a:t>ab</a:t></a:r></wrapper></a:p>',
    '<a:p xmlns:a="urn:a"><a:r><a:t>a<!--opaque-->b</a:t></a:r></a:p>',
  ]) {
    const document = await captureXmlTextSource(utf8(invalid));
    const atom = document.elements.find(element => element.local === 't');
    await assert.rejects(document.preflightRunSplit({ elementIndex: atom.elementIndex, namespaceUri: 'urn:a', expectedText: 'ab', offset: 1 }), { code: 'xml-edit-unsupported' });
    document.dispose();
  }
  const longText = 'a'.repeat(20000), longSource = await captureXmlTextSource(utf8(`<a:p xmlns:a="urn:a"><a:r><a:t>${longText}</a:t></a:r></a:p>`));
  const pending = longSource.preflightRunSplit({ elementIndex: 3, namespaceUri: 'urn:a', expectedText: longText, offset: 19999 });
  longSource.dispose();
  await assert.rejects(pending, { code: 'xml-source-disposed' });
});

test("all inactive MCE choices, unknown siblings and local namespace bindings remain exact", async () => {
  const prefix = '<r xmlns="urn:a" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future" mc:Ignorable="f"><mc:AlternateContent><mc:Choice Requires="f"><t f:opaque="yes">same</t></mc:Choice><mc:Fallback><t>same</t></mc:Fallback></mc:AlternateContent>';
  const suffix = '<f:ext xmlns:q="urn:local" ref="q:t">opaque<![CDATA[<t>same</t>]]></f:ext></r>';
  const source = await captureXmlTextSource(utf8(prefix + '<t>chosen</t>' + suffix));
  const result = await source.replaceText([edit(source, "chosen", "updated")]);
  assert.equal(string(result.bytes), prefix + '<t>updated</t>' + suffix);
});

test("CDATA and numeric/entity references are understood without preserving selected character syntax", async () => {
  const source = await captureXmlTextSource(utf8('<r xmlns="urn:a"><t>A&#xD;&#10;&#9;&amp;<![CDATA[<]]>🙂</t></r>'));
  const baseline = "A\r\n\t&<🙂";
  const result = await source.replaceText([edit(source, baseline, ']]> &lt; "quotes" \r\t\n')]);
  assert.equal(string(result.bytes), '<r xmlns="urn:a"><t>]]&gt; &amp;lt; "quotes" &#xD;\t\n</t></r>');
  const reopened = await captureXmlTextSource(result.bytes);
  assert.equal(select(reopened, ']]> &lt; "quotes" \r\t\n').expectedText, ']]> &lt; "quotes" \r\t\n');
});

test("self-closing/empty elements expand locally and equal-text edits retain original syntax", async () => {
  for (const token of ["<a:t/>", "<a:t />", "<a:t x='>' />", "<a:t></a:t>", "<a:t><![CDATA[]]></a:t>"]) {
    const before = `<a:r xmlns:a="urn:a">${token}<a:s/></a:r>`, source = await captureXmlTextSource(utf8(before));
    const noChange = await source.replaceText([edit(source, "", "")]);
    assert.equal(string(noChange.bytes), before); assert.deepEqual(noChange.changes, []);
    const result = await source.replaceText([edit(source, "", "new")]);
    const expectedToken = token.endsWith("/>") ? token.slice(0, -2) + ">new</a:t>" : token.slice(0, token.indexOf(">") + 1) + "new</a:t>";
    assert.equal(string(result.bytes), before.replace(token, expectedToken));
  }
  const source = await captureXmlTextSource(utf8('<r xmlns="urn:a"><t>remove</t></r>'));
  assert.equal(string((await source.replaceText([edit(source, "remove", "")])).bytes), '<r xmlns="urn:a"><t></t></r>');
});

test("UTF-8 BOM and both UTF-16 byte orders remain exact outside selected content", async () => {
  for (const encoding of ["utf-8", "utf-16le", "utf-16be"]) {
    const before = `<?xml version="1.0" encoding="${encoding === "utf-8" ? "UTF-8" : "UTF-16"}"?>\r\n<r xmlns="urn:a" x="中文🙂"><t>before</t><!--保留--></r>`;
    const encode = (text) => encoding === "utf-8" ? Buffer.concat([Buffer.from([239, 187, 191]), Buffer.from(text)]) : Buffer.concat([Buffer.from(encoding === "utf-16le" ? [255, 254] : [254, 255]), encoding === "utf-16le" ? Buffer.from(text, "utf16le") : Buffer.from(text, "utf16le").swap16()]);
    const bytes = encode(before), source = await captureXmlTextSource(bytes);
    assert.equal(source.inspection.encoding, encoding);
    assert.deepEqual(Buffer.from((await source.replaceText([])).bytes), bytes);
    const result = await source.replaceText([edit(source, "before", "新🙂 &\r")]);
    assert.deepEqual(Buffer.from(result.bytes), encode(before.replace(">before</t>", ">新🙂 &amp;&#xD;</t>")));
  }
});

test("replacement escaping blocks element, entity, comment, CDATA and PI injection", async () => {
  const source = await captureXmlTextSource(utf8('<r xmlns="urn:a"><t>old</t><other/></r>'));
  const text = '</t><evil xmlns="urn:evil"/><t>&unknown;<!--comment--><?pi x?><![CDATA[]]>';
  const result = await source.replaceText([edit(source, "old", text)]);
  const reopened = await captureXmlTextSource(result.bytes);
  assert.deepEqual(reopened.elements.map((e) => e.name), source.elements.map((e) => e.name));
  assert.equal(select(reopened, text).expectedText, text);
  assert.ok(!string(result.bytes).includes("<evil"));
});

test("mixed children/comments/PIs are rejected atomically instead of silently dropped", async () => {
  for (const content of ["a<x/>b", "a<!--keep-->b", "a<?keep yes?>b", "<![CDATA[a]]><!--keep-->b"]) {
    const source = await captureXmlTextSource(utf8(`<r xmlns="urn:a"><t>${content}</t><t>valid</t></r>`));
    assert.equal(source.elements[1].text, undefined);
    await assert.rejects(source.replaceText([edit(source, "valid", "new"), { elementIndex: 2, namespaceUri: "urn:a", local: "t", expectedText: "ab", text: "lost" }]), code("xml-edit-unsupported"));
    assert.equal(string((await source.replaceText([])).bytes), `<r xmlns="urn:a"><t>${content}</t><t>valid</t></r>`);
  }
});

test("stale text, missing identities and duplicate targets fail with no source changes", async () => {
  const before = '<r xmlns="urn:a"><t>old</t></r>', source = await captureXmlTextSource(utf8(before)), valid = edit(source, "old", "new");
  for (const request of [
    [{ ...valid, expectedText: "stale" }], [{ ...valid, local: "other" }], [{ ...valid, namespaceUri: "urn:other" }],
    [{ ...valid, elementIndex: 100 }], [valid, valid],
  ]) await assert.rejects(source.replaceText(request), code("xml-edit-conflict"));
  assert.equal(string((await source.replaceText([])).bytes), before);
  assert.equal(string((await source.replaceText([valid])).bytes), before.replace(">old<", ">new<"));
});

test("disjoint edits are order-independent; each candidate starts from the immutable original", async () => {
  const before = '<r xmlns="urn:a"><t>a</t><t>b</t><t>c</t></r>', source = await captureXmlTextSource(utf8(before));
  const edits = [edit(source, "a", "first"), edit(source, "c", "last")];
  const [a, b] = await Promise.all([source.replaceText(edits), source.replaceText([...edits].reverse())]);
  assert.deepEqual(a.bytes, b.bytes);
  assert.deepEqual(a.changes, b.changes);
  assert.equal(string(a.bytes), '<r xmlns="urn:a"><t>first</t><t>b</t><t>last</t></r>');
  a.bytes.fill(0);
  assert.equal(string((await source.replaceText([edit(source, "b", "middle")])).bytes), before.replace(">b<", ">middle<"));
});

test("XML 1.0 invalid characters and lone surrogates are rejected, valid astral/whitespace survive", async () => {
  const source = await captureXmlTextSource(utf8('<r xmlns="urn:a"><t>old</t></r>'));
  for (const text of ["\0", "\u0001", "\u000b", "\ufffe", "\uffff", "\ud800", "\udc00", "\ud800x"]) {
    await assert.rejects(source.replaceText([edit(source, "old", text)]), code("xml-edit-unsupported"));
  }
  for (const text of ["\t\n\r", "\u0085\u2028", "🙂\u{10ffff}", "\ue000"]) {
    const result = await source.replaceText([edit(source, "old", text)]);
    assert.equal(select(await captureXmlTextSource(result.bytes), text).expectedText, text);
  }
});

test("edit count, raw replacement characters, exact encoded bytes and work have independent budgets", async () => {
  const before = '<r xmlns="urn:a"><t>old</t></r>', source = await captureXmlTextSource(utf8(before)), request = [edit(source, "old", "🙂&")];
  const expected = utf8(before.replace(">old<", ">🙂&amp;<"));
  assert.deepEqual(Buffer.from((await source.replaceText(request, { limits: { maxEdits: 1, maxReplacementCharacters: 3, maxOutputBytes: expected.length } })).bytes), expected);
  await assert.rejects(source.replaceText(request, { limits: { maxReplacementCharacters: 2 } }), code("xml-edit-limit-exceeded"));
  await assert.rejects(source.replaceText(request, { limits: { maxOutputBytes: expected.length - 1 } }), code("xml-edit-limit-exceeded"));
  const other = await captureXmlTextSource(utf8('<r xmlns="urn:a"><t>a</t><t>b</t></r>'));
  await assert.rejects(other.replaceText([edit(other, "a", "x"), edit(other, "b", "y")], { limits: { maxEdits: 1 } }), code("xml-edit-limit-exceeded"));
  const noop = await captureXmlTextSource(utf8("<r/>"));
  await noop.replaceText([], { limits: { maxWorkUnits: 4, maxOutputBytes: 4 } });
  await assert.rejects(noop.replaceText([], { limits: { maxWorkUnits: 3 } }), code("xml-edit-limit-exceeded"));
  await assert.rejects(noop.replaceText([], { limits: { maxOutputBytes: 3 } }), code("xml-edit-limit-exceeded"));
  assert.ok(Object.isFrozen(DEFAULT_XML_TEXT_EDIT_LIMITS));
});

test("UTF-16 and entity expansion are charged before candidate allocation", async () => {
  const before = '<r xmlns="urn:a"><t>a</t></r>', bytes = Buffer.concat([Buffer.from([255, 254]), Buffer.from(before, "utf16le")]);
  const source = await captureXmlTextSource(bytes), request = [edit(source, "a", "&\r🙂")];
  const expectedLength = 2 + Buffer.byteLength(before.replace(">a<", ">&amp;&#xD;🙂<"), "utf16le");
  assert.equal((await source.replaceText(request, { limits: { maxOutputBytes: expectedLength } })).bytes.length, expectedLength);
  await assert.rejects(source.replaceText(request, { limits: { maxOutputBytes: expectedLength - 1 } }), code("xml-edit-limit-exceeded"));
  const small = await captureXmlTextSource(utf8("<r><t>a</t></r>"), { limits: { maxInputBytes: 15 } });
  await assert.rejects(small.replaceText([{ elementIndex: 2, namespaceUri: "", local: "t", expectedText: "a", text: "longer" }]), code("xml-limit-exceeded"));
});

test("malformed XML/encodings and invalid capture/edit contracts retain controlled failures", async () => {
  for (const xml of ["<r>", "<r>&entity;</r>", "<!DOCTYPE r><r/>", "<r xmlns:a='urn:a'><a:t></r>", "<?xml version='1.1'?><r/>"]) {
    await assert.rejects(captureXmlTextSource(utf8(xml)), (error) => error.code === "xml-invalid" || error.code === "xml-unsupported");
  }
  await assert.rejects(captureXmlTextSource(Buffer.from([0xc3, 0x28])), code("xml-invalid"));
  await assert.rejects(captureXmlTextSource(utf8("<r/>"), { limits: [] }), TypeError);
  await assert.rejects(captureXmlTextSource(utf8("<r/>"), { limits: { maxInputBytes: 3 } }), code("xml-limit-exceeded"));
  const source = await captureXmlTextSource(utf8("<r/>"));
  for (const value of [0, -1, NaN, Infinity, 1.2, "1", undefined]) {
    await assert.rejects(source.replaceText([], { limits: { maxEdits: value } }), RangeError);
  }
  await assert.rejects(source.replaceText([], { limits: [] }), TypeError);
  await assert.rejects(source.replaceText(null), TypeError);
  await assert.rejects(source.replaceText([{}]), TypeError);
  await assert.rejects(source.replaceText([{ elementIndex: 0, local: "r", namespaceUri: "", expectedText: "", text: "" }]), TypeError);
});

test("input/options/edit scalars are captured before asynchronous work and metadata is frozen", async () => {
  const original = utf8('<r xmlns="urn:a"><t>old</t></r>'), limits = { maxInputBytes: 1000 };
  const pending = captureXmlTextSource(original, { limits });
  original.fill(0); limits.maxInputBytes = 1;
  const source = await pending;
  assert.ok(Object.isFrozen(source) && Object.isFrozen(source.elements) && Object.isFrozen(source.elements[0]));
  assert.ok(Object.isFrozen(source.inspection) && Object.isFrozen(source.inspection.root) && Object.isFrozen(source.inspection.namespaceUris));
  const edits = [edit(source, "old", "new")], editLimits = { maxOutputBytes: 1000 }, options = { limits: editLimits };
  const editing = source.replaceText(edits, options);
  edits[0].text = "mutated"; edits.push(edit(source, "old", "duplicate")); editLimits.maxOutputBytes = 1;
  assert.equal(string((await editing).bytes), '<r xmlns="urn:a"><t>new</t></r>');
  assert.throws(() => { source.elements[0].name = "changed"; }, TypeError);
  if (typeof SharedArrayBuffer !== "undefined") await assert.rejects(captureXmlTextSource(new Uint8Array(new SharedArrayBuffer(4))), code("xml-edit-unsupported"));
});

test("capture and candidate cancellation return no partial result and allow independent retry", async () => {
  const before = '<r xmlns="urn:a"><t>old</t></r>', cancelled = new AbortController(); cancelled.abort();
  await assert.rejects(captureXmlTextSource(utf8(before), { signal: cancelled.signal }), code("operation-aborted"));
  const source = await captureXmlTextSource(utf8(before));
  await assert.rejects(source.replaceText([edit(source, "old", "new")], { signal: cancelled.signal }), code("operation-aborted"));
  const controller = new AbortController(), pending = source.replaceText([edit(source, "old", "x".repeat(100000))], { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, code("operation-aborted"));
  assert.equal(string((await source.replaceText([edit(source, "old", "retry")])).bytes), before.replace(">old<", ">retry<"));
});

test("dispose invalidates future/in-flight/getter-reentrant calls without corrupting another source", async () => {
  const bytes = utf8('<r xmlns="urn:a"><t>old</t></r>'), source = await captureXmlTextSource(bytes), other = await captureXmlTextSource(bytes);
  const pending = source.replaceText([edit(source, "old", "x".repeat(100000))]);
  source.dispose(); source.dispose();
  await assert.rejects(pending, code("xml-source-disposed"));
  await assert.rejects(source.replaceText([]), code("xml-source-disposed"));
  const reentrant = await captureXmlTextSource(bytes);
  await assert.rejects(reentrant.replaceText([], { get limits() { reentrant.dispose(); return {}; } }), code("xml-source-disposed"));
  const duringCopy = await captureXmlTextSource(bytes);
  await assert.rejects(duringCopy.replaceText([{ get text() { duringCopy.dispose(); return ""; } }]), code("xml-source-disposed"));
  assert.equal(string((await other.replaceText([])).bytes), string(bytes));
});

test("many small elements and deep nesting use bounded iterative source binding", async () => {
  const before = '<r xmlns="urn:a">' + Array.from({ length: 2000 }, (_, i) => `<t>${i}</t>`).join("") + '</r>';
  const source = await captureXmlTextSource(utf8(before));
  assert.equal(source.elements.length, 2001);
  assert.equal(string((await source.replaceText([edit(source, "1999", "last")])).bytes), before.replace(">1999<", ">last<"));
  const deep = '<r xmlns="urn:a">' + "<x>".repeat(240) + "<t>old</t>" + "</x>".repeat(240) + "</r>";
  const nested = await captureXmlTextSource(utf8(deep));
  assert.equal(string((await nested.replaceText([edit(nested, "old", "deep")])).bytes), deep.replace(">old<", ">deep<"));
});

test("source binding honors Unicode qualified names and XML line-end normalization without rewriting neighbors", async () => {
  const before = '<根 xmlns:字="urn:a" xmlns:q="urn:q">\r\n<字:t xml:space="preserve">a\r\nb\rc</字:t><q:尾 q:ref="字:t">unchanged\r\n</q:尾></根>';
  const source = await captureXmlTextSource(utf8(before));
  const result = await source.replaceText([edit(source, "a\nb\nc", "x\r\ny\rc")]);
  assert.equal(string(result.bytes), before.replace(">a\r\nb\rc</字:t>", ">x&#xD;\ny&#xD;c</字:t>"));
  const reopened = await captureXmlTextSource(result.bytes);
  assert.equal(select(reopened, "x\r\ny\rc").expectedText, "x\r\ny\rc");
});

test("capture cancellation remains live during lexical binding; a completed capture has a separate edit signal", async () => {
  const before = utf8("<r>" + "<t/>".repeat(2000) + "</r>"), controller = new AbortController();
  const pending = captureXmlTextSource(before, { signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(pending, code("operation-aborted"));
  const captureController = new AbortController();
  const source = await captureXmlTextSource(utf8("<r><t>old</t></r>"), { signal: captureController.signal });
  captureController.abort();
  assert.equal(string((await source.replaceText([{ elementIndex: 2, namespaceUri: "", local: "t", expectedText: "old", text: "new" }])).bytes), "<r><t>new</t></r>");
});

test("64 generated cases match independently assembled lexical output across duplicate text and encodings", async () => {
  let random = 19;
  const next = () => (random = (Math.imul(random, 1664525) + 1013904223) >>> 0);
  for (let sample = 0; sample < 64; sample++) {
    const count = 3 + next() % 17, chosen = next() % count, newText = `case${sample}🙂`, encoding = sample % 2 ? "utf-8" : "utf-16le";
    const tokens = Array.from({ length: count }, (_, i) => i % 3 === 0 ? '<q:t><![CDATA[same]]></q:t>' : '<q:t xml:space="preserve">same</q:t>');
    const prefix = '<q:r xmlns:q="urn:a" x="&gt;"><!--same-->', suffix = '<?keep same?></q:r>';
    const before = prefix + tokens.join("") + suffix;
    const encode = (text) => encoding === "utf-8" ? utf8(text) : Buffer.concat([Buffer.from([255, 254]), Buffer.from(text, "utf16le")]);
    const source = await captureXmlTextSource(encode(before));
    const element = source.elements[chosen + 1];
    const result = await source.replaceText([{ elementIndex: element.elementIndex, namespaceUri: "urn:a", local: "t", expectedText: "same", text: newText }]);
    tokens[chosen] = chosen % 3 === 0 ? `<q:t>${newText}</q:t>` : `<q:t xml:space="preserve">${newText}</q:t>`;
    assert.deepEqual(Buffer.from(result.bytes), encode(prefix + tokens.join("") + suffix), `generated case ${sample}`);
  }
});

test("three PPTX text-patch rounds preserve every other part, unknown effects, source IDs and relationship closure", async () => {
  for (const strict of [false, true]) {
    const f = contentFixture(shape({ txBox: true, props: '<a:effectLst><a:glow rad="12700"><a:srgbClr val="112233"/></a:glow></a:effectLst>', text: body() }), strict);
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", '<p:extLst><p:ext uri="urn:opaque"><q:payload xmlns:q="urn:future" value="q:ref"> A &amp; B </q:payload></p:ext></p:extLst></p:sld>');
    let bytes = encodeStructure(f);
    const initialParts = readGeneratedZip(bytes), initialClosure = await inspectOpcClosure(bytes, { roots: ["/deck/z.xml"] });
    for (let round = 1; round <= 3; round++) {
      const parts = readGeneratedZip(bytes), before = parts.get("deck/z.xml").toString(), source = await captureXmlTextSource(parts.get("deck/z.xml"));
      const oldText = round === 1 ? " A & B " : `round ${round - 1}`, newText = `round ${round}`;
      const result = await source.replaceText([edit(source, oldText, newText, "t", f.A)]);
      parts.set("deck/z.xml", Buffer.from(result.bytes));
      const candidate = repack(parts), reread = readGeneratedZip(candidate);
      assert.deepEqual([...reread.keys()].sort(), [...initialParts.keys()].sort());
      for (const [name, payload] of reread) if (name !== "deck/z.xml") assert.equal(digest(payload), digest(initialParts.get(name)), name);
      assert.equal(string(result.bytes), before.replace(round === 1 ? "> A &amp; B </a:t>" : `>round ${round - 1}</a:t>`, `>${newText}</a:t>`));
      const compared = await compareOpcPackages(bytes, candidate);
      assert.equal(compared.equal, false);
      assert.ok(compared.differences.every((d) => d.part === "/deck/z.xml"));
      const closure = await inspectOpcClosure(candidate, { roots: ["/deck/z.xml"] });
      assert.deepEqual(closure.relationships, initialClosure.relationships);
      assert.deepEqual(closure.incomingRelationships, initialClosure.incomingRelationships);
      assert.deepEqual(closure.unreachedParts, initialClosure.unreachedParts);
      const session = await importPptx(candidate, { projection: "content" });
      try {
        const ready = await session.prepareSave({ mode: "no-op" }); assert.equal(ready.status, "ready");
        assert.equal(digest((await session.save(ready.plan)).bytes), digest(candidate));
        assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
      } finally { session.dispose(); }
      source.dispose(); bytes = candidate;
    }
  }
});
