import test from "node:test";
import assert from "node:assert/strict";
import { inspectPptx, inspectOpcPackage } from "../dist/index.js";
import { structureFixture, encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
const MAIN = "deck/main.data";
async function invalid(change, code = "presentation-invalid") {
  const f = structureFixture(); change(f);
  await assert.rejects(inspectPptx(encodeStructure(f)), (error) => { assert.equal(error.code, code); assert.ok(error.entry); return true; });
}
function withNotes(f) {
  f.parts["deck/notes.xml"] = f.p("notes", '<p:cSld/>'); f.types["deck/notes.xml"] = "presentationml.notesSlide";
  f.parts["deck/notesMaster.xml"] = f.p("notesMaster", '<p:cSld/>'); f.types["deck/notesMaster.xml"] = "presentationml.notesMaster";
  f.parts[MAIN] = f.parts[MAIN].replace('<p:sldIdLst>', '<p:notesMasterIdLst><p:notesMasterId r:id="notesMaster"/></p:notesMasterIdLst><p:sldIdLst>');
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("a", "slide", "a.xml"), f.rel("notesMaster", "notesMaster", "notesMaster.xml")]);
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes.xml")]);
  f.parts["deck/_rels/notes.xml.rels"] = relationships([f.rel("slide", "slide", "z.xml"), f.rel("master", "notesMaster", "notesMaster.xml")]);
  f.parts["deck/_rels/notesMaster.xml.rels"] = relationships([f.rel("theme", "theme", "theme.xml")]);
}

test("reads authoritative slide order, identities, exact EMUs, hidden state and shared topology", async () => {
  for (const strict of [false, true]) {
    const f = structureFixture(strict), result = await inspectPptx(encodeStructure(f));
    assert.deepEqual(result.slides.map((s) => [s.id, s.part, s.name, s.hidden]), [[300, "/deck/z.xml", "First", true], [256, "/deck/a.xml", "Second", false]]);
    assert.deepEqual(result.slideSize, { widthEmu: 9144000, heightEmu: 6858000 });
    assert.deepEqual(result.notesSize, { widthEmu: 6858000, heightEmu: 9144000 });
    assert.equal(result.firstSlideNumber, 1); assert.equal(result.firstSlideNumberDeclared, false); assert.equal(result.dialect, strict ? "strict" : "transitional");
    assert.deepEqual(result.layouts, [{ part: "/deck/layout.xml", name: "Shared layout", masterPart: "/deck/master.xml" }]);
    assert.equal(result.slideMasters[0].themePart, "/deck/theme.xml"); assert.equal(result.slideMasters[0].layouts[0].id, 2147483649);
    assert.deepEqual(result.findings, []);
    assert.deepEqual(result.package.xmlTotals, (await inspectOpcPackage(encodeStructure(f))).xmlTotals);
  }
});

test("missing slide list represents an empty deck and absent slide size stays absent", async () => {
  const f = structureFixture(); f.parts[MAIN] = f.p("presentation", '<p:notesSz cx="0" cy="0"/>');
  f.parts["deck/_rels/main.data.rels"] = relationships([]);
  const result = await inspectPptx(encodeStructure(f));
  assert.deepEqual(result.slides, []); assert.deepEqual(result.slideMasters, []); assert.equal(result.slideSize, undefined);
  assert.deepEqual(result.notesSize, { widthEmu: 0, heightEmu: 0 });
  await invalid((f) => { f.parts[MAIN] = f.p("presentation"); });
});

test("dimension and integer fields enforce schema ranges and XML numeric lexical rules", async () => {
  const f = structureFixture(); f.parts[MAIN] = f.p("presentation", '<p:sldSz cx="+00914400" cy="51206400"/><p:notesSz cx="0" cy="27273042316900"/>', 'firstSlideNum="-2147483648"');
  f.parts["deck/_rels/main.data.rels"] = relationships([]);
  const result = await inspectPptx(encodeStructure(f));
  assert.equal(result.slideSize.widthEmu, 914400); assert.equal(result.firstSlideNumber, -2147483648); assert.equal(result.firstSlideNumberDeclared, true);
  for (const value of ["914399", "51206401", "NaN", "1e7", "1.5", "", " 9 14400 ", "&#160;914400"]) await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('cx="9144000"', `cx="${value}"`); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('cx="6858000"', 'cx="27273042316901"'); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('<p:presentation ', '<p:presentation firstSlideNum="2147483648" '); });
});

test('explicit first slide number remains distinguishable from the default in both dialects', async () => {
  for (const strict of [false, true]) for (const [lexical, expected] of [['1', 1], ['+0001', 1], ['0', 0], ['9999', 9999], ['10000', 10000], ['-1', -1], ['2147483647', 2147483647]]) {
    const f = structureFixture(strict); f.parts[MAIN] = f.parts[MAIN].replace('<p:presentation ', `<p:presentation firstSlideNum="${lexical}" `);
    const result = await inspectPptx(encodeStructure(f));
    assert.equal(result.firstSlideNumber, expected); assert.equal(result.firstSlideNumberDeclared, true);
  }
  for (const lexical of ['', '1.5', '1e2', '-2147483649', '2147483648']) await invalid(f => { f.parts[MAIN] = f.parts[MAIN].replace('<p:presentation ', `<p:presentation firstSlideNum="${lexical}" `); });
  const qualified = structureFixture(); qualified.parts[MAIN] = qualified.parts[MAIN].replace('<p:presentation ', '<p:presentation p:firstSlideNum="17" ');
  const result = await inspectPptx(encodeStructure(qualified)); assert.equal(result.firstSlideNumber, 1); assert.equal(result.firstSlideNumberDeclared, false);
  await invalid(f => { f.parts[MAIN] = f.parts[MAIN].replace('<p:presentation ', '<p:presentation firstSlideNum="1" firstSlideNum="2" '); }, 'xml-invalid');
});

test("numeric identity bounds and duplicate identities reject without string-spelling loopholes", async () => {
  for (const value of ["255", "2147483648", "-1", "300.0"]) await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('id="300"', `id="${value}"`); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('id="256"', 'id="+0300"'); });
  for (const value of ["2147483647", "4294967296"]) await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('id="2147483648"', `id="${value}"`); });
  await invalid((f) => { f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace('2147483649', '2147483647'); });
});

test("explicit r:id references must exist, have correct type/dialect, and target whole internal parts", async () => {
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('r:id="z"', 'r:id="missing"'); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('r:id="z"', 'bad="z"'); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('r:id="a"', 'r:id="z"'); });
  for (const entry of [{ Type: "urn:wrong", Target: "z.xml" }, { Target: "https://example.invalid/s.xml", TargetMode: "External" }, { Target: "z.xml#anchor" }]) await invalid((f) => { f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("a", "slide", "a.xml"), f.rel("m", "slideMaster", "master.xml"), { ...f.rel("z", "slide", "z.xml"), ...entry }]); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace(`xmlns:r="${f.R}"`, 'xmlns:r="urn:wrong"'); });
});

test("typed structural targets require matching content types and root names/namespaces", async () => {
  await invalid((f) => { f.types["deck/z.xml"] = "presentationml.slideLayout"; });
  await invalid((f) => { f.parts["deck/z.xml"] = f.p("sldLayout"); });
  await invalid((f) => { f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace(f.P, "http://purl.oclc.org/ooxml/presentationml/main"); });
  await invalid((f) => { f.parts["deck/theme.xml"] = '<theme/>'; });
  await invalid((f) => { f.types["deck/layout.xml"] = "presentationml.slide"; });
});

test("show boolean uses schema lexical values and cSld names are not inferred titles", async () => {
  for (const [show, hidden] of [["false", true], ["0", true], ["true", false], ["1", false], [" true ", false]]) {
    const f = structureFixture(); f.parts["deck/z.xml"] = f.p("sld", '<p:cSld name=""/>', `show="${show}"`);
    const result = await inspectPptx(encodeStructure(f)); assert.equal(result.slides[0].hidden, hidden); assert.equal(result.slides[0].name, "");
  }
  for (const show of ["False", "2", "yes", ""]) await invalid((f) => { f.parts["deck/z.xml"] = f.p("sld", "", `show="${show}"`); });
});

test("missing implicit inheritance is reported while contradictory or ambiguous links reject", async () => {
  const f = structureFixture(); delete f.parts["deck/_rels/z.xml.rels"];
  delete f.parts["deck/_rels/layout.xml.rels"];
  f.parts["deck/_rels/master.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml")]);
  const result = await inspectPptx(encodeStructure(f));
  assert.equal(result.findings.filter((item) => item.code === "missing-inheritance").length, 3);
  await invalid((f) => { f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("l1", "slideLayout", "layout.xml"), f.rel("l2", "slideLayout", "layout.xml")]); });
  await invalid((f) => {
    f.parts["deck/master2.xml"] = f.p("sldMaster"); f.types["deck/master2.xml"] = "presentationml.slideMaster";
    f.parts["deck/_rels/layout.xml.rels"] = relationships([f.rel("master", "slideMaster", "master2.xml")]);
  });
});

test("unlisted layouts and slide references are findings, never silently inserted into page order", async () => {
  const f = structureFixture(); f.parts["deck/master.xml"] = f.p("sldMaster");
  f.parts[MAIN] = f.parts[MAIN].replace('<p:sldId id="256" r:id="a"/>', '');
  const result = await inspectPptx(encodeStructure(f));
  assert.equal(result.slides.length, 1);
  assert.deepEqual(result.findings.map((f) => f.code).sort(), ["unlisted-layout", "unlisted-slide"]);
});

test("notes/master references retain provenance and validate back-links and ownership", async () => {
  const f = structureFixture(); withNotes(f);
  const result = await inspectPptx(encodeStructure(f));
  assert.equal(result.slides[0].notesPart, "/deck/notes.xml"); assert.equal(result.notesMasters[0].part, "/deck/notesMaster.xml"); assert.deepEqual(result.findings, []);
  await invalid((f) => { withNotes(f); f.parts["deck/_rels/notes.xml.rels"] = relationships([f.rel("slide", "slide", "a.xml")]); });
  await invalid((f) => { withNotes(f); f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("notes", "notesSlide", "notes.xml")]); });
  await invalid((f) => { withNotes(f); f.parts[MAIN] = f.parts[MAIN].replace('<p:notesMasterId r:id="notesMaster"/>', '<p:notesMasterId r:id="notesMaster"/><p:notesMasterId r:id="notesMaster"/>'); });
});

test("MCE selection determines page order and reports source decisions without interpreting extension islands", async () => {
  const f = structureFixture();
  f.parts[MAIN] = f.parts[MAIN].replace('<p:sldIdLst><p:sldId id="300" r:id="z"/><p:sldId id="256" r:id="a"/></p:sldIdLst>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><p:sldIdLst><p:sldId id="301" r:id="absent"/></p:sldIdLst></mc:Choice><mc:Fallback><p:sldIdLst><p:sldId id="256" r:id="a"/><p:sldId id="300" r:id="z"/></p:sldIdLst></mc:Fallback></mc:AlternateContent>');
  f.parts[MAIN] = f.parts[MAIN].replace('</p:presentation>', '<p:extLst><p:ext uri="urn:custom"><x:unknown xmlns:x="urn:future"/></p:ext></p:extLst></p:presentation>');
  const result = await inspectPptx(encodeStructure(f));
  assert.deepEqual(result.slides.map((s) => s.id), [256, 300]);
  assert.deepEqual(result.markupCompatibility[0].report.decisions.map((d) => d.action), ["alternate-content", "opaque-extension"]);
});

test("root AlternateContent is interpreted before checking the effective presentation root", async () => {
  const f = structureFixture(); f.parts[MAIN] = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x"><mc:Choice Requires="x"><x:unknown/></mc:Choice><mc:Fallback>${f.parts[MAIN]}</mc:Fallback></mc:AlternateContent>`;
  assert.equal((await inspectPptx(encodeStructure(f))).slides.length, 2);
  f.parts[MAIN] = f.parts[MAIN].replace('</mc:Fallback>', '<extra/></mc:Fallback>');
  await assert.rejects(inspectPptx(encodeStructure(f)), { code: "presentation-invalid" });
});

test("duplicate singleton fields and illegal identity list children reject", async () => {
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('</p:presentation>', '<p:sldSz cx="914400" cy="914400"/></p:presentation>'); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('<p:sldIdLst>', '<p:sldIdLst>unexpected'); });
  await invalid((f) => { f.parts[MAIN] = f.parts[MAIN].replace('<p:sldIdLst>', '<p:sldIdLst><unknown/>'); });
});

test("input/options snapshots, XML budgets, cancellation and detached repeated results remain intact", async () => {
  const f = structureFixture(), original = encodeStructure(f), mutable = Uint8Array.from(original), controller = new AbortController();
  const limits = { maxXmlNodes: 1000 }, options = { limits, signal: controller.signal };
  const pending = inspectPptx(mutable, options); mutable.fill(0); limits.maxXmlNodes = 1; options.signal = AbortSignal.abort();
  const first = await pending; first.slides.length = 0;
  assert.equal((await inspectPptx(original)).slides.length, 2);
  await assert.rejects(inspectPptx(original, { limits: { maxXmlNodes: 1 } }), { code: "package-limit-exceeded" });
  await assert.rejects(inspectPptx(original, { signal: AbortSignal.abort() }), { code: "operation-aborted" });
  await assert.rejects(inspectPptx(original.slice(0, 20)), { code: "archive-invalid" });
});

test("handout master identities use their own typed list and permit at most one", async () => {
  const f = structureFixture();
  f.parts[MAIN] = f.parts[MAIN].replace('<p:sldIdLst>', '<p:handoutMasterIdLst><p:handoutMasterId r:id="handout"/></p:handoutMasterIdLst><p:sldIdLst>');
  f.parts["deck/handout.xml"] = f.p("handoutMaster"); f.types["deck/handout.xml"] = "presentationml.handoutMaster";
  f.parts["deck/_rels/handout.xml.rels"] = relationships([f.rel("theme", "theme", "theme.xml")]);
  f.parts["deck/_rels/main.data.rels"] = relationships([f.rel("m", "slideMaster", "master.xml"), f.rel("z", "slide", "z.xml"), f.rel("a", "slide", "a.xml"), f.rel("handout", "handoutMaster", "handout.xml")]);
  const result = await inspectPptx(encodeStructure(f));
  assert.deepEqual(result.handoutMasters, [{ part: "/deck/handout.xml", relationshipId: "handout" }]);
  assert.deepEqual(result.findings, []);
  f.parts[MAIN] = f.parts[MAIN].replace('<p:handoutMasterId r:id="handout"/>', '<p:handoutMasterId r:id="handout"/><p:handoutMasterId r:id="handout"/>');
  await assert.rejects(inspectPptx(encodeStructure(f)), { code: "presentation-invalid" });
});

test("hundreds of slides share cyclic layout/master topology without duplicating layouts or changing order", async () => {
  const f = structureFixture(), ids = [], relations = [f.rel("m", "slideMaster", "master.xml")];
  for (const name of ["a", "z"]) { delete f.parts[`deck/${name}.xml`]; delete f.parts[`deck/_rels/${name}.xml.rels`]; delete f.types[`deck/${name}.xml`]; }
  for (let i = 249; i >= 0; i--) {
    const name = `s${i}.xml`;
    f.parts[`deck/${name}`] = f.p("sld", '<p:cSld/>', i % 2 ? 'show="false"' : 'show="true"');
    f.types[`deck/${name}`] = "presentationml.slide";
    f.parts[`deck/_rels/${name}.rels`] = relationships([f.rel("layout", "slideLayout", "layout.xml")]);
    relations.push(f.rel(`s${i}`, "slide", name));
    ids.unshift(`<p:sldId id="${256 + i}" r:id="s${i}"/>`);
  }
  f.parts["deck/_rels/main.data.rels"] = relationships(relations);
  f.parts[MAIN] = f.parts[MAIN].replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids.join("")}</p:sldIdLst>`);
  const result = await inspectPptx(encodeStructure(f));
  assert.deepEqual(result.slides.map((slide) => slide.id), Array.from({ length: 250 }, (_, i) => 256 + i));
  assert.equal(result.slides.filter((slide) => slide.hidden).length, 125);
  assert.equal(result.layouts.length, 1); assert.equal(result.slideMasters.length, 1); assert.deepEqual(result.findings, []);
});
