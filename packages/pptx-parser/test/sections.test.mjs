import test from "node:test";
import assert from "node:assert/strict";
import { inspectPptx, inspectMarkupCompatibility } from "../dist/index.js";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
import { SECTION_NS, SECTION_URI, SECTION_IDS, section, sectionExtension, withSections } from "./sections-fixture.mjs";
import { MC } from "./metadata-fixture.mjs";
const read = (f, options) => inspectPptx(encodeStructure(f), options);
async function invalid(body, change = () => {}, code = "presentation-invalid") {
  const f = withSections(structureFixture(), body); change(f);
  await assert.rejects(read(f), (e) => { assert.equal(e.code, code); assert.ok(e.entry); return true; });
}

test("native section source order, empty sections, names and stable slide identities work in both dialects", async () => {
  for (const strict of [false, true]) {
    const result = await read(withSections(structureFixture(strict)));
    assert.deepEqual(result.sectionList, { part: "/deck/main.data", extensionUri: SECTION_URI, sections: [
      { id: SECTION_IDS[0], name: "Opening", slides: [{ id: 300, part: "/deck/z.xml", relationshipId: "z" }], extensions: [] },
      { id: SECTION_IDS[1], name: "Empty", slides: [], extensions: [] },
      { id: SECTION_IDS[2], name: "Closing", slides: [{ id: 256, part: "/deck/a.xml", relationshipId: "a" }], extensions: [] },
    ] });
    assert.deepEqual(result.findings, []);
    assert.deepEqual(result.slides.map((s) => s.id), [300, 256]);
    assert.equal(result.presentationExtensions[0].status, "interpreted");
  }
});

test("absent native sections, an empty section and an invalid empty section list remain distinct", async () => {
  assert.equal((await read(structureFixture())).sectionList, undefined);
  const f = withSections(structureFixture(), section());
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/, "<p:sldIdLst/>");
  const result = await read(f);
  assert.equal(result.sectionList.sections.length, 1); assert.deepEqual(result.sectionList.sections[0].slides, []);
  assert.ok(!("id" in result.sectionList.sections[0])); assert.ok(!("name" in result.sectionList.sections[0]));
  await invalid("");
});

test("section names are decoded text and never identities; duplicate and empty names do not merge", async () => {
  const f = withSections(structureFixture(), section([300], 'name="中文 &amp; &#xA;"') + section([256], 'name="中文 &amp; &#xA;"') + section([], 'name=""'));
  const result = await read(f);
  assert.deepEqual(result.sectionList.sections.map((s) => s.name), ["中文 & \n", "中文 & \n", ""]);
  assert.equal(result.findings.length, 0);
});

test("native section GUID identities validate schema spelling and normalized duplicates", async () => {
  for (const id of ["", "invalid", SECTION_IDS[0].toLowerCase(), SECTION_IDS[0].slice(1, -1)]) await invalid(section([300, 256], `id="${id}"`));
  await invalid(section([300], `id="${SECTION_IDS[0]}"`) + section([256], `id=" ${SECTION_IDS[0]} "`));
  assert.equal((await read(withSections(structureFixture(), section([300, 256], `id=" ${SECTION_IDS[0]} "`)))).sectionList.sections[0].id, SECTION_IDS[0]);
});

test("section slide references validate numeric identity, uniqueness, source membership and bounds", async () => {
  for (const ids of [[300, 300], ["300", "+0300"], [255], [2147483648], ["1e3"], [""], [257]]) await invalid(section(ids));
  await invalid(section([300]) + section([300, 256]));
  const f = withSections(structureFixture(), section(["+0300", "0256"]));
  assert.deepEqual((await read(f)).sectionList.sections[0].slides.map((s) => s.id), [300, 256]);
  await invalid(section([300, 256]), (f) => { f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('<p:sldId id="256" r:id="a"/>', ""); });
});

test("partial and out-of-order section membership yields located findings without rewriting page order", async () => {
  const partial = await read(withSections(structureFixture(), section([256])));
  assert.deepEqual(partial.findings.map((f) => [f.code, f.sectionIndex, f.slideId]), [["section-order", 0, 256], ["unsectioned-slide", undefined, 300]]);
  const reversed = await read(withSections(structureFixture(), section([256]) + section([300])));
  assert.deepEqual(reversed.sectionList.sections.map((s) => s.slides[0].id), [256, 300]);
  assert.deepEqual(reversed.slides.map((s) => s.id), [300, 256]);
  assert.ok(reversed.findings.every((f) => f.code === "section-order" && f.part === "/deck/main.data"));
});

test("known extension envelopes and section child grammar reject malformed structure", async () => {
  for (const body of ['<s:wrong/>', '<s:section/>', '<s:section><s:sldIdLst/><s:sldIdLst/></s:section>', '<s:section><s:extLst/><s:sldIdLst/></s:section>', '<s:section extra="x"><s:sldIdLst/></s:section>', '<s:section>x<s:sldIdLst/></s:section>', '<s:section><s:sldIdLst><s:sldId/></s:sldIdLst></s:section>', '<s:section><s:sldIdLst><s:sldId id="300"><s:child/></s:sldId></s:sldIdLst></s:section>', '<s:section><s:sldIdLst><s:sldId id="300" extra="x"/></s:sldIdLst></s:section>', '<s:section><s:sldIdLst>x</s:sldIdLst></s:section>', section([], "", '<s:extLst><s:ext/></s:extLst>')]) await invalid(body);
  for (const change of [
    (s) => s.replace('<s:sectionLst>', '<s:sectionLst extra="x">'),
    (s) => s.replace('</p:extLst>', `${sectionExtension(section())}</p:extLst>`),
    (s) => s.replace('</s:sectionLst>', '</s:sectionLst><s:sectionLst/>'),
    (s) => s.replace('xmlns:s=', 'extra="x" xmlns:s='),
    (s) => s.replaceAll('s:sectionLst', 's:other'),
  ]) await invalid(section([300, 256]), (f) => { f.parts["deck/main.data"] = change(f.parts["deck/main.data"]); });
});

test("unknown presentation and section extensions expose descriptors without reading payloads", async () => {
  for (const strict of [false, true]) {
    const f = structureFixture(strict), unknown = `<p:ext uri="urn:future"><u:thing xmlns:u="urn:future" xmlns:mc="${MC}" mc:MustUnderstand="missing"><u:bad/></u:thing></p:ext>`;
    // Section extLst has a Transitional extension-list type even in a Strict document.
    const nested = unknown.replaceAll('p:ext', 'old:ext').replace('uri="urn:future"', 'uri="urn:future" xmlns:old="http://schemas.openxmlformats.org/presentationml/2006/main"');
    withSections(f, section([300, 256], "", `<s:extLst>${nested}</s:extLst>`), unknown);
    const result = await read(f);
    assert.equal(result.presentationExtensions[1].status, "uninterpreted");
    assert.equal(result.sectionList.sections[0].extensions[0].status, "uninterpreted");
    assert.equal(result.sectionList.sections[0].extensions[0].children[0].namespaceUri, "urn:future");
    assert.ok(result.markupCompatibility[0].report.decisions.some((d) => d.action === "opaque-extension"));
  }
});

test("a section-shaped payload under an unknown URI or slide extension is not promoted to native sections", async () => {
  const f = withSections(structureFixture()); f.parts["deck/main.data"] = f.parts["deck/main.data"].replace(SECTION_URI, "urn:other");
  assert.equal((await read(f)).sectionList, undefined);
  const slide = structureFixture(); slide.parts["deck/a.xml"] = slide.parts["deck/a.xml"].replace('</p:sld>', `<p:extLst>${sectionExtension(section([300]))}</p:extLst></p:sld>`);
  assert.equal((await read(slide)).sectionList, undefined);
});

test("extension URI token whitespace is normalized before recognition and duplicate detection", async () => {
  const f = withSections(structureFixture());
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace(SECTION_URI, ` &#x9;${SECTION_URI}&#xA; `);
  const result = await read(f);
  assert.equal(result.sectionList.extensionUri, SECTION_URI);
  assert.equal(result.presentationExtensions[0].uri, SECTION_URI);
  await invalid(section([300, 256]), (f) => { f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('</p:extLst>', `${sectionExtension(section()).replace(SECTION_URI, ` ${SECTION_URI} `)}</p:extLst>`); });
});

test("recognized section MCE honors inherited namespace scopes, ignorable attributes and branch selection", async () => {
  const f = structureFixture();
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('<p:presentation ', `<p:presentation xmlns:mc="${MC}" xmlns:u="urn:future" mc:Ignorable="u" `);
  withSections(f, `<mc:AlternateContent><mc:Choice Requires="u">${section([999])}</mc:Choice><mc:Fallback>${section([300, 256], 'u:extra="ignored"')}</mc:Fallback></mc:AlternateContent>`);
  const result = await read(f);
  assert.deepEqual(result.sectionList.sections[0].slides.map((s) => s.id), [300, 256]);
  const decisions = result.markupCompatibility[0].report.decisions;
  assert.ok(decisions.some((d) => d.action === "ignored-attribute"));
  assert.ok(decisions.some((d) => d.action === "alternate-content"));
  assert.ok(!decisions.some((d) => d.action === "opaque-extension"));
});

test("MCE around the extension and root keeps section capabilities without opting nested islands in", async () => {
  for (const wrapper of ["alternate", "process", "root"]) {
    const f = withSections(structureFixture(), section([300, 256], "", `<s:extLst><p:ext uri="${SECTION_URI}"><u:unknown xmlns:u="urn:unknown"/></p:ext></s:extLst>`));
    let xml = f.parts["deck/main.data"].replace('<p:presentation ', `<p:presentation xmlns:mc="${MC}" xmlns:u="urn:future" xmlns:s="${SECTION_NS}" mc:Ignorable="u" mc:ProcessContent="u:wrapper" `);
    if (wrapper === "alternate") xml = xml.replace('<p:extLst>', '<p:extLst><mc:AlternateContent><mc:Choice Requires="s">').replace('</p:extLst>', '</mc:Choice><mc:Fallback/></mc:AlternateContent></p:extLst>');
    if (wrapper === "process") xml = xml.replace('<p:extLst>', '<p:extLst><u:wrapper>').replace('</p:extLst>', '</u:wrapper></p:extLst>');
    if (wrapper === "root") xml = `<mc:AlternateContent xmlns:mc="${MC}" xmlns:s="${SECTION_NS}"><mc:Choice Requires="s">${xml}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
    f.parts["deck/main.data"] = xml;
    const result = await read(f);
    assert.equal(result.sectionList.sections[0].slides.length, 2, wrapper);
    assert.equal(result.sectionList.sections[0].extensions[0].status, "uninterpreted");
  }
});

test("unknown active section markup rejects while the generic MCE API retains opaque ext behavior", async () => {
  await invalid(`<u:unknown xmlns:u="urn:unknown"/>`, (f) => { f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('<p:presentation ', `<p:presentation xmlns:mc="${MC}" `); }, "mce-unsupported");
  const f = structureFixture(), xml = `<p:presentation xmlns:p="${f.P}"><p:extLst>${sectionExtension('<u:unknown xmlns:u="urn:unknown"/>')}</p:extLst></p:presentation>`;
  const result = await inspectMarkupCompatibility(new TextEncoder().encode(xml), { understoodNamespaces: [f.P], extensionElements: [{ namespaceUri: f.P, local: "ext" }] });
  assert.ok(result.decisions.some((d) => d.action === "opaque-extension"));
});

test("a reused section URI on a different main-part owner remains opaque", async () => {
  const f = structureFixture();
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace('</p:presentation>', `<p:custShowLst><p:custShow name="Other owner" id="1"><p:sldLst/><p:extLst><p:ext uri="${SECTION_URI}"><u:unknown xmlns:u="urn:unknown" xmlns:mc="${MC}" mc:MustUnderstand="missing"/></p:ext></p:extLst></p:custShow></p:custShowLst></p:presentation>`);
  const result = await read(f);
  assert.equal(result.sectionList, undefined);
  assert.deepEqual(result.presentationExtensions, []);
  assert.ok(result.markupCompatibility[0].report.decisions.some((d) => d.action === "opaque-extension"));
});

test("sections share XML budgets and cancellation, and result mutation cannot affect source or later reads", async () => {
  const f = withSections(structureFixture(), section([300, 256], 'name="stable"') + section([], 'name="unused"').repeat(300));
  await assert.rejects(read(f, { limits: { maxXmlNodes: 100 } }), (e) => e.code === "package-limit-exceeded");
  const first = await read(f); first.sectionList.sections[0].slides[0].id = 999; first.sectionList.sections[0].name = "changed";
  const later = await read(f); assert.equal(later.sectionList.sections[0].slides[0].id, 300); assert.equal(later.sectionList.sections[0].name, "stable");
  const { readXml } = await import("../dist/xml/read-xml.js"), { readSections } = await import("../dist/presentation/sections.js");
  const { root } = await readXml(new TextEncoder().encode(f.parts["deck/main.data"]), {}, true);
  let count = 0; const stopped = new Error("stop section traversal");
  await assert.rejects(readSections(root, f.P, "/deck/main.data", later.slides, async () => { if (++count === 64) throw stopped; }), (e) => e === stopped);
  assert.equal(count, 64);
});
