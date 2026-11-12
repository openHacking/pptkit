import assert from "node:assert/strict";
import test from "node:test";
import { inspectPptx, inspectPptxNotes, importPptx, DEFAULT_NOTES_INSPECTION_LIMITS } from "../dist/index.js";
import { encodeStructure, structureFixture } from "./structure-fixture.mjs";
import { notesFixture, notesShape, notesTree } from "./notes-fixture.mjs";
import { body } from "./content-fixture.mjs";
import { notesInventoryFixture } from "./notes-inventory-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";

const notesPart = r => r.parts.find(p => p.kind === "notes");
const masterPart = r => r.parts.find(p => p.kind === "notesMaster");
const notesPage = r => r.pages.find(p => p.notesPart);

for (const strict of [false, true]) test(`notes source inventory roles/fields/local geometry/type association and body isolation (${strict ? "Strict" : "Transitional"})`, async () => {
  const f = notesInventoryFixture(strict), bytes = encodeStructure(f), before = await inspectPptx(bytes), r = await inspectPptxNotes(bytes), n = notesPart(r), m = masterPart(r), page = notesPage(r);
  assert.equal(r.dialect, strict ? "strict" : "transitional"); assert.deepEqual(r.notesSize, { widthEmu: 6858000, heightEmu: 9144000 });
  assert.deepEqual(r.pages.map(p => p.slideId), [300, 256]); assert.equal(r.pages[1].body.status, "absent"); assert.equal(r.parts.length, 2);
  assert.deepEqual(n.drawings.map(d => d.role).sort(), ["group-root", "body", "slide-image", "header", "footer", "date", "slide-number", "artwork", "artwork", "body"].sort());
  assert.equal(page.body.status, "unique-direct"); assert.equal(page.body.candidates.length, 1);
  const localBody = n.drawings.find(d => d.rawId === "2"), nestedBody = n.drawings.find(d => d.rawId === "10");
  assert.equal(localBody.direct, true); assert.equal(nestedBody.direct, false); assert.equal(localBody.placeholder.index, "1");
  const match = page.placeholderMatches.find(p => p.source.elementIndex === localBody.placeholder.source.elementIndex);
  assert.equal(match.status, "matched-direct-type"); assert.equal(match.candidates.length, 1);
  assert.equal(m.drawings.find(d => d.role === "body").placeholder.index, "99");
  assert.equal(page.placeholderMatches.find(p => p.source.elementIndex === nestedBody.placeholder.source.elementIndex).status, "grouped");
  assert.equal(n.fields.length, 3); assert.deepEqual(n.fields.map(f => f.type), ["unrecognized-custom-type", "datetimeCustom", "slidenum"]);
  assert.ok(n.fields.every(f => f.cachedTexts[0].value === "2042-01-02 & stored")); assert.equal(r.fieldEvaluation, "not-performed");
  assert.ok(!n.texts.some(t => t.value.includes("NOT A FIELD"))); assert.equal(n.extensions.length, 1); assert.equal(n.extensions[0].children[0].namespaceUri, "urn:future");
  assert.equal(n.attributes.find(a => a.local === "showMasterSp").value, "false"); assert.equal(m.headerFooter[0].attributes.find(a => a.local === "dt").value, "false");
  const transform = n.geometries.find(g => g.kind === "transform" && g.ownerDrawingElementIndex === localBody.source.elementIndex);
  assert.equal(transform.attributes.find(a => a.local === "rot").value, "60000");
  assert.equal(transform.components.find(c => c.name.local === "off").attributes.find(a => a.local === "x").value, "-10");
  assert.equal(transform.components.find(c => c.name.local === "off").attributes.find(a => a.local === "y").value, "00020");
  assert.equal(r.geometryPolicy, "unresolved-source-only"); assert.equal(r.authoredMapping, "not-performed");
  assert.equal(Object.isFrozen(r), true); assert.equal(Object.isFrozen(n.fields[0].cachedTexts), true); assert.throws(() => { n.fields[0].cachedTexts[0].value = "changed"; }, TypeError);
  const session = await importPptx(bytes, { projection: "content" });
  try { assert.equal(session.document.slides[0].notes[0].runs.map(r => r.text).join(""), "LOCAL BODY"); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes)); } finally { session.dispose(); }
  assert.deepEqual(await inspectPptx(bytes), before);
});

test("body ambiguity, absent local text and unspecified placeholder types do not synthesize master content", async () => {
  for (const [content, expected] of [["", "absent"], [notesShape({ text: "" }), "missing-local-text"], [notesShape() + notesShape({ id: "3" }), "ambiguous-direct"], [notesShape({ text: body() + body() }), "ambiguous-local-text"]]) {
    const r = await inspectPptxNotes(encodeStructure(notesFixture(content))); assert.equal(notesPage(r).body.status, expected);
  }
  const f = notesFixture(notesShape().replace('type="body" ', "")), r = await inspectPptxNotes(encodeStructure(f));
  assert.equal(notesPage(r).body.status, "absent"); assert.equal(notesPage(r).placeholderMatches[0].status, "missing-type"); assert.ok(r.findings.some(f => f.code === "unknown-placeholder"));
});

test("MCE inactive drawings keep raw identities without false duplicate bodies or opaque extension text", async () => {
  const f = notesFixture();
  const alternate = '<mc:AlternateContent><mc:Choice Requires="u">' + notesShape({ text: body('<a:p><a:r><a:t>RAW</a:t></a:r></a:p>') }) + '</mc:Choice><mc:Fallback>' + notesShape({ text: body('<a:p><a:r><a:t>EFFECTIVE</a:t></a:r></a:p>') }) + '</mc:Fallback></mc:AlternateContent>';
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(alternate), `xmlns:a="${f.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"`);
  const r = await inspectPptxNotes(encodeStructure(f)), n = notesPart(r);
  assert.equal(notesPage(r).body.status, "unique-direct"); assert.equal(n.drawings.filter(d => d.role === "body").length, 2);
  assert.deepEqual(n.texts.map(t => [t.value, t.view]), [["RAW", "raw-only"], ["EFFECTIVE", "effective"]]);
  assert.equal(r.findings.some(f => f.code === "duplicate-drawing-id"), false);
  assert.equal(n.drawings.find(d => d.role === "body" && d.view === "effective").direct, true);
});

test("master matching reports ambiguity/grouping and never substitutes a listed master for a missing actual edge", async () => {
  const missing = notesFixture(); missing.parts["deck/notes/_rels/z.xml.rels"] = relationships([missing.rel("back", "slide", "../z.xml")]);
  const noMaster = await inspectPptxNotes(encodeStructure(missing)); assert.equal(notesPage(noMaster).masterPart, undefined); assert.equal(notesPage(noMaster).placeholderMatches[0].status, "missing-master"); assert.equal(masterPart(noMaster).kind, "notesMaster");
  const ambiguous = notesFixture(); ambiguous.parts["deck/notes-master.xml"] = ambiguous.parts["deck/notes-master.xml"].replace('</p:spTree>', notesShape({ id: "6" }) + '</p:spTree>');
  const r = await inspectPptxNotes(encodeStructure(ambiguous)); assert.equal(notesPage(r).placeholderMatches[0].status, "ambiguous"); assert.equal(notesPage(r).placeholderMatches[0].candidates.length, 2);
});

test("drawing identity aliases/findings are scoped to effective parts and unknown drawing islands stay uninterpreted", async () => {
  const f = notesFixture(notesShape() + notesShape({ id: "+0002", role: "hdr" }) + notesShape({ id: "weird", role: "ftr" }) + '<u:art xmlns:u="urn:art">' + notesShape({ id: "2" }) + '</u:art>');
  const r = await inspectPptxNotes(encodeStructure(f));
  assert.deepEqual(r.findings.map(f => f.code), ["duplicate-drawing-id", "uninterpreted-drawing-id", "unknown-drawing"]);
  assert.equal(notesPart(r).drawings.filter(d => d.rawId === "2").length, 1);
  const duplicatePh = notesFixture(notesShape().replace('</p:nvPr>', '<p:ph type="hdr"/></p:nvPr>'));
  await assert.rejects(inspectPptxNotes(encodeStructure(duplicatePh)), e => e.code === "presentation-invalid");
  for (const role of ["__proto__", "constructor", "toString"]) {
    const report = await inspectPptxNotes(encodeStructure(notesFixture(notesShape({ role }))));
    assert.equal(notesPart(report).drawings.find(d => d.rawId === "2").role, "other-placeholder"); assert.deepEqual(JSON.parse(JSON.stringify(report)), report);
  }
});

test("notes limits snapshot own data properties and bound work/records/text/findings at exact observed thresholds", async () => {
  const bytes = encodeStructure(notesInventoryFixture()), original = Uint8Array.from(bytes), r = await inspectPptxNotes(bytes);
  for (const [key, value] of [["maxWorkUnits", r.workUnits], ["maxRecords", r.recordCount], ["maxTextCharacters", r.textCharacters]]) {
    assert.deepEqual(await inspectPptxNotes(bytes, { notesLimits: { [key]: value } }), r);
    await assert.rejects(inspectPptxNotes(bytes, { notesLimits: { [key]: value - 1 } }), e => e.code === "package-limit-exceeded");
  }
  const limits = { ...DEFAULT_NOTES_INSPECTION_LIMITS }, pending = inspectPptxNotes(bytes, { notesLimits: limits }); bytes.fill(0); limits.maxRecords = 1;
  assert.deepEqual(await pending, r); assert.deepEqual(await inspectPptxNotes(original), r);
  for (const value of [null, [], new Date(), { maxRecords: 0 }, { maxWorkUnits: NaN }, { unknown: 2 }]) await assert.rejects(inspectPptxNotes(original, { notesLimits: value }), e => e instanceof TypeError || e instanceof RangeError);
  let reads = 0; const accessor = Object.defineProperty({}, "maxRecords", { enumerable: true, get() { reads++; return 10; } });
  await assert.rejects(inspectPptxNotes(original, { notesLimits: accessor }), TypeError); assert.equal(reads, 0);
  const bad = notesFixture(notesShape({ id: "bad" }) + notesShape({ id: "wrong", role: "undefined-kind" }));
  await assert.rejects(inspectPptxNotes(encodeStructure(bad), { notesLimits: { maxFindings: 1 } }), e => e.code === "package-limit-exceeded");
});

test("abort and structural backlink failures remain explicit, with no implicit network use", async () => {
  const f = notesInventoryFixture(), bytes = encodeStructure(f), controller = new AbortController(); controller.abort();
  await assert.rejects(inspectPptxNotes(bytes, { signal: controller.signal }), e => e.code === "operation-aborted");
  const live = new AbortController(), pending = inspectPptxNotes(bytes, { signal: live.signal }); queueMicrotask(() => live.abort());
  await assert.rejects(pending, e => e.code === "operation-aborted");
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../a.xml"), f.rel("master", "notesMaster", "../notes-master.xml")]);
  await assert.rejects(inspectPptxNotes(encodeStructure(f)), e => e.code === "presentation-invalid");
});

test("no notes produce no invented parts, and shared masters are scanned once across pages", async () => {
  const absent = await inspectPptxNotes(encodeStructure(structureFixture())); assert.deepEqual(absent.parts, []); assert.ok(absent.pages.every(p => !p.notesPart && p.body.status === "absent"));
  const f = notesFixture();
  f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('</p:spTree>', Array.from({ length: 100 }, (_, i) => notesShape({ id: String(i + 20), placeholder: false, text: body(`<a:p><a:r><a:t>Master artwork ${i}</a:t></a:r></a:p>`) })).join("") + '</p:spTree>');
  const one = await inspectPptxNotes(encodeStructure(f));
  f.parts["deck/notes/a.xml"] = f.parts["deck/notes/z.xml"]; f.types["deck/notes/a.xml"] = "presentationml.notesSlide";
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes/a.xml")]);
  f.parts["deck/notes/_rels/a.xml.rels"] = relationships([f.rel("back", "slide", "../a.xml"), f.rel("master", "notesMaster", "../notes-master.xml")]);
  const two = await inspectPptxNotes(encodeStructure(f));
  assert.equal(two.parts.length, 3); assert.equal(two.parts.filter(p => p.kind === "notesMaster").length, 1);
  assert.equal(masterPart(two).drawings.length, masterPart(one).drawings.length); assert.ok(two.workUnits - one.workUnits < masterPart(two).drawings.length);
  assert.deepEqual(two.pages[0].placeholderMatches[0].candidates, two.pages[1].placeholderMatches[0].candidates);
});

test("a grouped master candidate stays unresolved, and actual unlisted edges never borrow listed properties", async () => {
  const grouped = notesFixture();
  const tree = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' + notesShape({ id: "5" }) + '</p:grpSp>';
  grouped.parts["deck/notes-master.xml"] = grouped.p("notesMaster", notesTree(tree), `xmlns:a="${grouped.A}"`);
  assert.equal(notesPage(await inspectPptxNotes(encodeStructure(grouped))).placeholderMatches[0].status, "grouped");
  const f = notesFixture(); f.parts["deck/other-master.xml"] = f.parts["deck/notes-master.xml"]; f.types["deck/other-master.xml"] = "presentationml.notesMaster";
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("master", "notesMaster", "../other-master.xml")]);
  const r = await inspectPptxNotes(encodeStructure(f)); assert.equal(notesPage(r).masterPart, "/deck/other-master.xml"); assert.equal(notesPage(r).masterThemePart, undefined); assert.equal(r.parts.length, 3);
  assert.ok(r.structureFindings.some(f => f.code === "unlisted-master"));
});

test("MCE root alternatives retain raw-only source roots without treating nested fake trees as direct bodies", async () => {
  const f = notesFixture(); const original = f.parts["deck/notes/z.xml"];
  f.parts["deck/notes/z.xml"] = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unknown"><mc:Choice Requires="u">${original.replace('Any name', 'Raw root')}</mc:Choice><mc:Fallback>${original}</mc:Fallback></mc:AlternateContent>`;
  const r = await inspectPptxNotes(encodeStructure(f)); assert.equal(notesPage(r).body.status, "unique-direct"); assert.equal(notesPart(r).drawings.filter(d => d.role === "body").length, 2); assert.ok(notesPart(r).source.elementIndex > 1);
  const malformed = notesFixture(notesShape({ props: '<p:spTree>' + notesShape({ id: "50" }) + '</p:spTree>' }));
  const noFake = await inspectPptxNotes(encodeStructure(malformed)); assert.equal(notesPage(noFake).body.candidates.length, 1); assert.equal(notesPart(noFake).drawings.some(d => d.rawId === "50"), false);
});

test("external theme overrides are rejected without fetching, and inherited structure findings consume notes budgets", async () => {
  const f = notesFixture();
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("master", "notesMaster", "../notes-master.xml"), { ...f.rel("override", "themeOverride", "https://invalid.example/never-fetch.xml"), TargetMode: "External" }]);
  await assert.rejects(inspectPptxNotes(encodeStructure(f)), e => e.code === "presentation-invalid");
  const missing = structureFixture(); missing.parts["deck/_rels/a.xml.rels"] = relationships([]); missing.parts["deck/_rels/z.xml.rels"] = relationships([]);
  await assert.rejects(inspectPptxNotes(encodeStructure(missing), { notesLimits: { maxFindings: 1 } }), e => e.code === "package-limit-exceeded");
});

test("ProcessContent promotes actual drawing/text/field owners through raw-only unknown wrappers", async () => {
  const f = notesFixture();
  const text = body('<a:p><a:r><a:t>FLATTENED</a:t></a:r><a:fld type="unknown"><a:t>cached</a:t></a:fld></a:p>');
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree('<u:wrapper>' + notesShape({ text }) + '</u:wrapper>'), `xmlns:a="${f.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper"`);
  const r = await inspectPptxNotes(encodeStructure(f)), n = notesPart(r), b = n.drawings.find(d => d.role === "body");
  assert.equal(notesPage(r).body.status, "unique-direct"); assert.equal(b.view, "effective"); assert.equal(b.direct, true);
  assert.equal(n.drawings.find(d => d.kind === "uninterpreted").view, "raw-only");
  assert.deepEqual(n.texts.map(t => [t.value, t.ownerDrawingElementIndex]), [["FLATTENED", b.source.elementIndex]]);
  assert.equal(n.fields[0].ownerDrawingElementIndex, b.source.elementIndex); assert.equal(n.fields[0].cachedTexts[0].value, "cached");
  assert.ok(n.attributes.some(a => a.local === "Ignorable" && a.value === "u"));
});

test("UTF-16 source and empty field caches retain the same literal inventory without evaluation", async () => {
  const f = notesInventoryFixture(), expected = await inspectPptxNotes(encodeStructure(f));
  for (const part of ["deck/notes/z.xml", "deck/notes-master.xml"]) f.parts[part] = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(f.parts[part], "utf16le")]);
  assert.deepEqual(await inspectPptxNotes(encodeStructure(f)), expected);
  const empty = notesFixture(notesShape({ text: body('<a:p><a:fld type="date"><a:t></a:t></a:fld><a:fld type="unknown"/></a:p>') }));
  const fields = notesPart(await inspectPptxNotes(encodeStructure(empty))).fields;
  assert.deepEqual(fields.map(f => f.cachedTexts.map(t => t.value)), [[""], []]);
});

test("MCE inside nonvisual and text ownership uses the selected projection with raw source identities", async () => {
  const f = notesFixture();
  const ns = `xmlns:a="${f.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper"`;
  let shape = notesShape();
  shape = shape.replace('<p:nvSpPr>', '<u:wrapper><p:nvSpPr>').replace('</p:nvSpPr>', '</p:nvSpPr></u:wrapper>');
  shape = shape.replace('<p:nvPr>', '<p:nvPr><u:wrapper>').replace('</p:nvPr>', '</u:wrapper></p:nvPr>');
  shape = shape.replace('<p:txBody>', '<u:wrapper><p:txBody>').replace('</p:txBody>', '</p:txBody></u:wrapper>');
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(shape), ns);
  const r = await inspectPptxNotes(encodeStructure(f)), n = notesPart(r), b = n.drawings.find(d => d.role === "body");
  assert.ok(b); assert.equal(b.placeholder.view, "effective"); assert.equal(b.nonVisual.view, "effective");
  assert.equal(b.textBodyCount, 1); assert.equal(notesPage(r).body.status, "unique-direct");
});

test("picture, connector, frame and custom geometry remain source-only artwork with exact nonvisual attributes", async () => {
  for (const strict of [false, true]) {
    const nv = (tag, id) => `<p:${tag}><p:cNvPr id="${id}" name="Art ${id}" hidden="1" descr="Stored description"/><p:nvPr/></p:${tag}>`;
    const pic = '<p:pic>' + nv('nvPicPr', '20') + '<p:spPr><a:xfrm flipV="1"><a:off x="-2" y="3"/><a:ext cx="4" cy="5"/></a:xfrm><a:custGeom><a:avLst/><a:pathLst><a:path w="10" h="10"><a:moveTo><a:pt x="1" y="2"/></a:moveTo></a:path></a:pathLst></a:custGeom></p:spPr></p:pic>';
    const connector = '<p:cxnSp>' + nv('nvCxnSpPr', '21') + '<p:spPr><a:prstGeom prst="line"/></p:spPr></p:cxnSp>';
    const frame = '<p:graphicFrame>' + nv('nvGraphicFramePr', '22') + '<p:xfrm><a:off x="9" y="8"/><a:ext cx="7" cy="6"/></p:xfrm><a:graphic><a:graphicData uri="urn:uninterpreted"/></a:graphic></p:graphicFrame>';
    const r = await inspectPptxNotes(encodeStructure(notesFixture(notesShape() + pic + connector + frame, strict))), n = notesPart(r);
    const art = n.drawings.filter(d => ['pic', 'cxnSp', 'graphicFrame'].includes(d.kind));
    assert.deepEqual(art.map(d => d.kind), ['pic', 'cxnSp', 'graphicFrame']);
    assert.ok(art.every(d => d.role === 'artwork' && d.direct && d.nonVisual.attributes.some(a => a.local === 'hidden' && a.value === '1')));
    const custom = n.geometries.find(g => g.kind === 'custom'); assert.equal(custom.ownerDrawingElementIndex, art[0].source.elementIndex);
    assert.deepEqual(custom.components.map(c => c.name.local), ['avLst', 'pathLst']);
    assert.equal(n.geometries.find(g => g.ownerDrawingElementIndex === art[2].source.elementIndex).name.local, 'xfrm');
    assert.equal(r.authoredMapping, 'not-performed'); assert.equal(r.geometryPolicy, 'unresolved-source-only');
  }
});

test("field caches through selected MCE wrappers retain active text and do not publish inactive alternatives as values", async () => {
  const f = notesFixture(notesShape({ text: body('<a:p><a:fld type="slidenum"><u:wrapper><a:t>selected cache</a:t></u:wrapper></a:fld><a:fld type="unknown"><mc:AlternateContent><mc:Choice Requires="u"><a:t>inactive</a:t></mc:Choice><mc:Fallback><a:t>fallback cache</a:t></mc:Fallback></mc:AlternateContent></a:fld></a:p>') }));
  f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace('<p:notes ', '<p:notes xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper" ');
  const r = await inspectPptxNotes(encodeStructure(f)), n = notesPart(r);
  assert.deepEqual(n.fields[0].cachedTexts.map(t => [t.value, t.view]), [['selected cache', 'effective']]);
  assert.ok(n.fields[0].cachedTexts[0].source.elementIndex > n.fields[0].source.elementIndex + 1);
  assert.deepEqual(n.fields[1].cachedTexts.map(t => [t.value, t.view]), [["fallback cache", "effective"]]);
  assert.equal(n.texts.some(t => t.value === "inactive"), false);
});

test("extension payloads inside notes drawings are not interpreted as field or body text", async () => {
  for (const strict of [false, true]) {
    const extension = '<a:extLst><a:ext uri="urn:future"><a:fld type="slidenum"><a:t>FALSE FIELD</a:t></a:fld><a:t>FALSE TEXT</a:t></a:ext></a:extLst>';
    const shape = notesShape({ text: body('<a:p><a:r><a:t>REAL</a:t></a:r></a:p>' + extension) });
    const r = await inspectPptxNotes(encodeStructure(notesFixture(shape, strict))), n = notesPart(r);
    assert.deepEqual(n.texts.map(t => t.value), ['REAL']);
    assert.equal(n.fields.length, 0);
    assert.ok(n.extensions.some(e => e.children.some(child => child.local === 'fld')));
    assert.equal(notesPage(r).body.status, 'unique-direct');
    const promoted = notesFixture(shape, strict);
    promoted.parts['deck/notes/z.xml'] = promoted.parts['deck/notes/z.xml'].replace('<p:notes ', '<p:notes xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper" ').replace('<a:fld type="slidenum">', '<u:wrapper><a:fld type="slidenum">').replace('</a:fld><a:t>FALSE TEXT</a:t>', '</a:fld><a:t>FALSE TEXT</a:t></u:wrapper>');
    const projected = notesPart(await inspectPptxNotes(encodeStructure(promoted)));
    assert.deepEqual(projected.texts.map(t => t.value), ['REAL']); assert.equal(projected.fields.length, 0);
  }
});

test("field caches and text retain their enclosing source paragraph identity", async () => {
  for (const strict of [false, true]) {
    const text = body('<a:p><a:r><a:t>first</a:t></a:r><a:fld type="slidenum"><a:t>1</a:t></a:fld></a:p><a:p><a:r><a:t>second</a:t></a:r><a:fld type="datetime"><a:t>date</a:t></a:fld></a:p>');
    const n = notesPart(await inspectPptxNotes(encodeStructure(notesFixture(notesShape({ text }), strict))));
    assert.equal(n.fields.length, 2); assert.equal(n.texts.length, 2);
    for (let i = 0; i < 2; i++) {
      assert.ok(n.fields[i].paragraphElementIndex > 0);
      assert.equal(n.fields[i].paragraphElementIndex, n.texts[i].paragraphElementIndex);
      assert.equal(n.fields[i].cachedTexts[0].paragraphElementIndex, n.fields[i].paragraphElementIndex);
      assert.ok(n.fields[i].paragraphElementIndex < n.fields[i].source.elementIndex);
    }
    assert.notEqual(n.fields[0].paragraphElementIndex, n.fields[1].paragraphElementIndex);
  }
});

test("ambiguous nonvisual children in inactive MCE are source-only without rejecting or selecting an identity", async () => {
  const f = notesFixture(), inactive = notesShape().replace('</p:nvSpPr>', '</p:nvSpPr><p:nvSpPr><p:cNvPr id="99" name="inactive second"/><p:nvPr><p:ph type="hdr"/></p:nvPr></p:nvSpPr>');
  const alternate = '<mc:AlternateContent><mc:Choice Requires="u">' + inactive + '</mc:Choice><mc:Fallback>' + notesShape() + '</mc:Fallback></mc:AlternateContent>';
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(alternate), `xmlns:a="${f.A}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported"`);
  const r = await inspectPptxNotes(encodeStructure(f)), n = notesPart(r), raw = n.drawings.find(d => d.kind === 'sp' && d.view === 'raw-only');
  assert.equal(raw.rawId, undefined); assert.equal(raw.placeholder, undefined); assert.equal(raw.direct, false);
  assert.equal(notesPage(r).body.status, 'unique-direct'); assert.equal(r.findings.length, 0);
});
