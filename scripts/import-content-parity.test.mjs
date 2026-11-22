import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx, inspectOpcPackage, inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { createContentImportFixture, createImageImportFixture, createReferenceImportFixture, createNotesImportFixture, createTableImportFixture } from "../tests/fixtures/content-import.mjs";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { png, gif } from "../packages/pptx-parser/test/image-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { structureFixture, encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { withSections } from "../packages/pptx-parser/test/sections-fixture.mjs";
import { createFixture } from "../tests/fixtures/presentations.mjs";

test("PowerPoint saved authored sections match independent XML and inspection", async () => {
  const paths = ["tests/fixtures/authored-native-sections.pptx", "tests/fixtures/authored-native-sections-powerpoint.pptx"];
  const hashes = ["194f12063b331f75cc262e1421ca16448b7bddca9527c51259b8d34a2f311344", "c9917f74ae175e53bb383442b56536b36d46400d58dd557bba305881cc819189"];
  const expected = [
    { id: "{FAB7F38F-4347-8A71-BA1E-92B419C6FFB6}", name: "Introduction", slides: [256] },
    { id: "{904FEFB3-D01C-82AE-B51E-1AEA66183C25}", name: "Empty", slides: [] },
    { id: "{C2F00318-F053-800A-9AE1-DFF72F376939}", name: "Conclusion", slides: [257] },
  ];
  const oracle = spawnSync("python3", ["scripts/research/inspect-authored-sections.py", ...paths], { encoding: "utf8" });
  assert.equal(oracle.status, 0, oracle.stderr); assert.deepEqual(JSON.parse(oracle.stdout), [expected, expected]);
  for (const [index, path] of paths.entries()) {
    const bytes = await readFile(path);
    assert.equal(createHash("sha256").update(bytes).digest("hex"), hashes[index]);
    const inspection = await inspectPptx(bytes);
    assert.deepEqual(inspection.sectionList.sections.map(s => ({ id: s.id, name: s.name, slides: s.slides.map(slide => slide.id) })), expected);
    assert.deepEqual(inspection.findings, []);
  }
  const document = createFixture("text", createPresentation), [a, b] = document.slides;
  document.setSections([{ id: "intro", name: "Introduction", slideIds: [a.id] }, { id: "empty", name: "Empty", slideIds: [] }, { id: "end", name: "Conclusion", slideIds: [b.id] }]);
  assert.deepEqual(Buffer.from((await generatePptx(document)).bytes), await readFile(paths[0]));
});

test("explicit native section collection projection survives three fresh exports", async () => {
  let bytes = encodeStructure(withSections(structureFixture()));
  let sourceGuid;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content", sectionProjection: "collection-v1" });
    try {
      const order = new Map(session.document.slides.map((slide, index) => [slide.id, index]));
      assert.deepEqual(session.document.sections.map(section => [section.id, section.name, section.slideIds.map(id => order.get(id))]), [
        ["source-section-1", "Opening", [0]], ["source-section-2", "Empty", []], ["source-section-3", "Closing", [1]],
      ]);
      const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready");
      assert.equal(Buffer.compare(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(bytes)), 0);
      const guids = session.report.inspection.sectionList.sections.map(s => s.id);
      if (round === 0) sourceGuid = guids[0]; else assert.notEqual(guids[0], sourceGuid);
      bytes = (await generatePptx(session.document)).bytes;
      const native = await inspectPptx(bytes);
      assert.deepEqual(native.sectionList.sections.map(s => [s.name, s.slides.map(slide => slide.id)]), [["Opening", [256]], ["Empty", []], ["Closing", [257]]]);
    } finally { session.dispose(); }
  }
});

test("authored native sections preserve names, empty sections and numeric slide membership", async () => {
  const deck = createPresentation();
  deck.addSlide({ id: "a", section: "legacy" }); deck.addSlide({ id: "b" }); deck.addSlide({ id: "c" });
  deck.setSections([{ id: "first", name: "same &\n", slideIds: ["a", "b"] }, { id: "empty", slideIds: [] }, { id: "blank", name: "", slideIds: [] }, { id: "last", name: "same &\n", slideIds: ["c"] }]);
  const output = await generatePptx(deck);
  const parsed = await inspectPptx(output.bytes);
  assert.deepEqual(parsed.sectionList.sections.map(s => [s.name, s.slides.map(slide => slide.id)]), [["same &\n", [256, 257]], [undefined, []], ["", []], ["same &\n", [258]]]);
  const guids = parsed.sectionList.sections.map(s => s.id);
  assert.equal(new Set(guids).size, 4);
  const xml = readGeneratedZip(output.bytes).get("ppt/presentation.xml").toString();
  assert.match(xml, /same &amp;&#10;/);
  const session = await importPptx(output.bytes, { projection: "content" });
  const noOp = await session.prepareSave({ mode: "no-op" });
  assert.equal(noOp.status, "ready"); assert.deepEqual((await session.save(noOp.plan)).bytes, output.bytes);
  session.document.setSections([{ id: "local", slideIds: session.document.slides.map(s => s.id) }]);
  assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
  session.dispose();
  deck.moveSlide("b", 0);
  assert.deepEqual((await inspectPptx((await generatePptx(deck)).bytes)).sectionList.sections.map(s => s.id), guids);
  deck.setSectionStructure(["c", "b", "a"], [{ id: "last", name: "same &\n", slideIds: ["c", "b"] }, { id: "empty", slideIds: [] }, { id: "first", name: "same &\n", slideIds: ["a"] }]);
  const moved = await inspectPptx((await generatePptx(deck)).bytes);
  assert.deepEqual(moved.sectionList.sections.map(s => [s.id, s.slides.map(slide => slide.id)]), [[guids[3], [256, 257]], [guids[1], []], [guids[0], [258]]]);
  deck.addSlide({ id: "unassigned" });
  await assert.rejects(generatePptx(deck), /cover every slide/);
  deck.removeSlide("unassigned"); deck.moveSlide("a", 1);
  await assert.rejects(generatePptx(deck), /cover every slide/);
});

// Fresh export compares effective colors/fonts; exact no-op byte checks below retain
// the separate guarantee for source theme references and their relationships.
function effectiveStyleField(key, value, theme) {
  if (!["color", "fontFamily"].includes(key) || value === null || typeof value !== "object"
    || Object.keys(value).length !== 1 || !Object.hasOwn(value, "theme")) return value;
  const roles = key === "color" ? theme.colors : theme.fonts;
  assert.ok(Object.hasOwn(roles, value.theme), `Unknown theme ${key} role: ${value.theme}`);
  return roles[value.theme];
}

function tableSemantics(document) {
  const normalized = normalizePresentation(document), order = new Map(normalized.slides.map((s, i) => [s.id, String(i)]));
  const visit = value => {
    if (value === null || typeof value !== "object") return value;
    if (Array.isArray(value)) return value.map(visit);
    // Compare logical table semantics across the explicit native transform wrapper.
    if (value.type === "group" && value.children.length === 1 && value.children[0].type === "table" && value.name === value.children[0].name && value.children[0].box.x === 0 && value.children[0].box.y === 0) {
      const table = { ...value.children[0], box: value.box, transform: value.transform, accessibility: value.accessibility, opacity: value.opacity };
      if (value.action) table.action = value.action;
      return visit(table);
    }
    if (value.type === "table") {
      value = { ...value, nativeFirstRowStyle: value.nativeFirstRowStyle ?? true,
        nativeTableStylePreset: value.nativeTableStylePreset ?? "medium-2-accent-1",
        nativeTableStyleFlags: { firstColumn: false, lastColumn: false, lastRow: false, bandedRows: true, bandedColumns: false, ...value.nativeTableStyleFlags } };
    }
    return Object.fromEntries(Object.entries(value).filter(([key]) => key !== "id").map(([key, v]) => [key, key === "slideId" ? order.get(v) : visit(effectiveStyleField(key, v, normalized.theme))]));
  };
  return normalized.slides.map(s => s.elements.map(visit));
}

test("merged table text, styles, actions and transforms survive three fresh projections with independent physical-grid checks", async () => {
  const original = createTableImportFixture(createPresentation), expected = tableSemantics(original), sizes = [];
  let generated = await generatePptx(original);
  assert.deepEqual(generated.warnings.map(w => w.code), ["table-transform-grouped", "table-transform-grouped"]);
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(generated.bytes, { projection: "content" });
    assert.deepEqual(tableSemantics(s.document), expected);
    const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.deepEqual((await s.save(prepared.plan)).bytes, generated.bytes);
    const fresh = await generatePptx(s.document), parts = readGeneratedZip(fresh.bytes);
    assert.deepEqual(fresh.warnings.map(w => w.code), ["native-table-preset-direct-formatting", "native-table-preset-direct-formatting"]);
    const xml = parts.get("ppt/slides/slide1.xml").toString();
    assert.equal((xml.match(/<a:tbl>/g) ?? []).length, 2);
    assert.equal((xml.match(/<a:tc(?:\s|>)/g) ?? []).length, 13);
    assert.equal((xml.match(/<a:tc gridSpan="2" rowSpan="2">/g) ?? []).length, 2);
    assert.equal((xml.match(/ hMerge="1"/g) ?? []).length, 4);
    assert.equal((xml.match(/ vMerge="1"/g) ?? []).length, 4);
    assert.match(xml, /<a:xfrm rot="900000" flipH="1"><a:off x="381000" y="508000"\/><a:ext cx="3810000" cy="1270000"/);
    assert.match(xml, /<a:xfrm rot="-1200000" flipV="1">/);
    assert.match(xml, /<a:tblGrid><a:gridCol w="1143000"\/><a:gridCol w="1270000"\/><a:gridCol w="1397000"\/>/);
    assert.match(xml, /Merged α😀 &amp; &lt;literal&gt;/);
    assert.match(xml, /<a:prstDash val="sysDot"/);
    assert.match(xml, /<a:alpha val="75000"/);
    const rels = parts.get("ppt/slides/_rels/slide1.xml.rels").toString();
    assert.match(rels, /Target="slide1.xml"/); assert.match(rels, /Target="slide2.xml"/);
    assert.match(rels, /Target="https:\/\/example.invalid\/table\?a=1&amp;b=%E4%B8%AD"/);
    const graph = await inspectOpcPackage(fresh.bytes);
    assert.ok(graph.relationships.filter(r => r.type.endsWith("/slide")).every(r => r.targetPart));
    const preview = await renderPresentationToSvg(s.document);
    assert.match(preview.slides[0].svg, /Merged α😀 &amp; &lt;literal&gt;/);
    assert.match(preview.slides[0].svg, /aria-rowspan="2"/); assert.match(preview.slides[0].svg, /aria-colspan="2"/);
    sizes.push([...parts.values()].reduce((n, part) => n + part.length, 0));
    generated = fresh; s.dispose();
  }
  assert.equal(new Set(sizes).size, 1, "table source identities/resources do not grow across fresh projections");
});

test("imported table cell edits and copied self-links reach fresh output while preservation edits remain blocked", async () => {
  const source = await generatePptx(createTableImportFixture(createPresentation)), s = await importPptx(source.bytes, { projection: "content" });
  const p = s.document, first = p.slides[0], oldTable = first.elements[0], prepared = await s.prepareSave({ mode: "no-op" });
  const copy = p.duplicateSlide(first.id), copiedTable = copy.elements[0];
  assert.equal(copiedTable.action.slideId, copy.id);
  assert.equal(copiedTable.children[0].rows[0].cells[0].content[0].runs[0].action.slideId, copy.id);
  p.removeSlide(first.id);
  const changed = structuredClone(copiedTable);
  changed.children[0].rows[2].cells[0].content = "Edited cell";
  copy.removeElement(copiedTable.id); copy.insertElement(0, changed);
  const fresh = await generatePptx(p), again = await importPptx(fresh.bytes, { projection: "content" });
  assert.deepEqual(tableSemantics(again.document), tableSemantics(p));
  assert.equal(again.document.slides[0].elements[0].children[0].rows[2].cells[0].content[0].runs[0].text, "Edited cell");
  assert.equal(oldTable.children[0].rows[2].cells[0].content[0].runs[0].text, "A");
  await assert.rejects(s.save(prepared.plan), e => e.code === "save-snapshot-stale");
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked");
  again.dispose(); s.dispose();
});

test("native section labels reach fresh slide metadata while native collection loss remains explicit", async () => {
  const bytes = encodeStructure(withSections(structureFixture())), session = await importPptx(bytes, { projection: "content" });
  const saved = await session.prepareSave({ mode: "no-op" }); assert.equal(saved.status, "ready");
  assert.equal(Buffer.compare(Buffer.from((await session.save(saved.plan)).bytes), bytes), 0);
  const fresh = await generatePptx(session.document), entries = readGeneratedZip(fresh.bytes);
  assert.deepEqual(fresh.warnings, []);
  assert.match(entries.get("ppt/slides/slide1.xml").toString(), /&quot;section&quot;:&quot;Opening&quot;/);
  assert.match(entries.get("ppt/slides/slide2.xml").toString(), /&quot;section&quot;:&quot;Closing&quot;/);
  assert.equal((await inspectPptx(fresh.bytes)).sectionList, undefined);
  assert.ok(session.report.findings.some(f => f.feature === "section.native-collection" && f.operation === "export" && f.tier === "unsupported"));
  assert.ok(session.report.findings.some(f => f.feature === "section.label" && f.operation === "export" && f.tier === "fallback"));
  const again = await importPptx(fresh.bytes, { projection: "content" });
  assert.deepEqual(again.document.slides.map(slide => slide.section), [undefined, undefined]);
  assert.equal(session.report.inspection.sectionList.sections[1].name, "Empty");
  assert.equal(session.report.inspection.sectionList.sections[1].slides.length, 0);
  again.dispose(); session.dispose();
});

function noteSemantics(document) {
  const p = normalizePresentation(document);
  for (const slide of p.slides) for (const paragraph of slide.notes) for (const run of paragraph.runs) {
    if (run.action?.type === "slide") run.action.slideId = String(p.slides.findIndex(s => s.id === run.action.slideId));
  }
  return p.slides.map(s => s.notes);
}

test("speaker notes text, styles and actions survive three fresh projection rounds with notes-part relationship checks", async () => {
  const original = createNotesImportFixture(createPresentation), expected = noteSemantics(original);
  let generated = await generatePptx(original);
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(generated.bytes, { projection: "content" });
    assert.deepEqual(noteSemantics(s.document), expected);
    const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.deepEqual((await s.save(prepared.plan)).bytes, generated.bytes);
    const fresh = await generatePptx(s.document); assert.deepEqual(fresh.warnings, []);
    const parts = readGeneratedZip(fresh.bytes), noteXml = parts.get("ppt/notesSlides/notesSlide1.xml").toString(), rels = parts.get("ppt/notesSlides/_rels/notesSlide1.xml.rels").toString();
    assert.equal((noteXml.match(/<a:p>/g) ?? []).length, 3); assert.match(noteXml, /Speaker α😀 &amp; &lt;literal&gt;/);
    assert.match(noteXml, /<a:buAutoNum type="arabicPeriod" startAt="3"/);
    assert.equal((noteXml.match(/ppaction:\/\/hlinksldjump/g) ?? []).length, 2);
    assert.match(rels, /Target="\.\.\/slides\/slide2\.xml"/); assert.match(rels, /Target="\.\.\/slides\/slide1\.xml"/);
    assert.match(rels, /Target="https:\/\/example.invalid\/notes\?a=1&amp;b=%E4%B8%AD"/);
    const emptyXml = parts.get("ppt/notesSlides/notesSlide2.xml").toString(); assert.match(emptyXml, /sz="1400"/); assert.match(emptyXml, /b="1"/);
    const preview = await renderPresentationToSvg(s.document);
    assert.doesNotMatch(preview.slides[0].svg, /Speaker α/);
    assert.ok(s.report.findings.some(f => f.feature === "notes.body.mapped-fields" && f.operation === "preview" && f.tier === "unsupported"));
    s.dispose(); generated = fresh;
  }
});

test("edited and copied notes update fresh output while source-preservation edits remain blocked", async () => {
  const source = await generatePptx(createNotesImportFixture(createPresentation)), s = await importPptx(source.bytes, { projection: "content" });
  const p = s.document, first = p.slides[0], originalNotes = first.notes;
  const plan = await s.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
  const copy = p.duplicateSlide(first.id);
  assert.equal(copy.notes[1].runs[1].action.slideId, copy.id);
  p.removeSlide(first.id);
  copy.setNotes([{ runs: [{ text: "Edited notes", action: { type: "slide", slideId: copy.id } }] }]);
  p.slides.find(slide => slide.id !== copy.id).setNotes();
  const fresh = await generatePptx(p), again = await importPptx(fresh.bytes, { projection: "content" });
  assert.deepEqual(noteSemantics(again.document), noteSemantics(p));
  assert.equal(again.document.slides[0].notes[0].runs[0].text, "Edited notes");
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  await assert.rejects(s.save(plan.plan), e => e.code === "save-snapshot-stale");
  assert.equal(originalNotes[0].runs[0].text, "Speaker α😀 & <literal>");
  again.dispose(); s.dispose();
});


function withoutIds(element) {
  const { id, ...value } = element;
  if (value.type === "group") value.children = value.children.map(withoutIds);
  return value;
}
test("generated direct text and shape/group fields import through public APIs and render with native geometry", async () => {
  const original = createContentImportFixture(createPresentation), generated = await generatePptx(original), session = await importPptx(generated.bytes, { projection: "content" });
  const a = normalizePresentation(original), b = normalizePresentation(session.document);
  assert.deepEqual(b.slides[0].elements.map(withoutIds), a.slides[0].elements.map(withoutIds));
  const svg = await renderPresentationToSvg(session.document);
  assert.match(svg.slides[0].svg, /Editable &amp; exact &lt;text&gt;/);
  assert.match(svg.slides[0].svg, /Inside group/);
  const fresh = await generatePptx(session.document), parts = readGeneratedZip(fresh.bytes), slide = parts.get("ppt/slides/slide1.xml").toString();
  assert.match(slide, /<a:chExt cx="1905000" cy="1143000"/);
  assert.match(slide, /<a:normAutofit fontScale="80000"/);
  assert.match(slide, /<a:buAutoNum type="arabicPeriod" startAt="2"/);
  const again = await importPptx(fresh.bytes, { projection: "content" });
  assert.deepEqual(normalizePresentation(again.document).slides[0].elements.map(withoutIds), b.slides[0].elements.map(withoutIds));
  again.dispose(); session.dispose();
});

function imageSemantics(document) {
  const normalized = normalizePresentation(document);
  const inspect = element => {
    const { id, assetId, ...value } = element;
    if (assetId) {
      const asset = normalized.assets.find(a => a.id === assetId);
      value.asset = { mimeType: asset.mimeType, source: asset.source };
    }
    if (value.type === "group") value.children = value.children.map(inspect);
    return value;
  };
  return normalized.slides.map(s => s.elements.map(inspect));
}

test("embedded image bytes, geometry and shared references survive three fresh projection exports with independent package checks", async () => {
  const original = createImageImportFixture(createPresentation, png), expected = imageSemantics(original);
  let generated = await generatePptx(original);
  const sizes = [];
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(generated.bytes, { projection: "content" });
    assert.deepEqual(imageSemantics(s.document), expected);
    assert.equal(s.document.assets.length, 1);
    const plan = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(plan.plan)).bytes, generated.bytes);
    const fresh = await generatePptx(s.document);
    assert.deepEqual(fresh.warnings, []);
    const parts = readGeneratedZip(fresh.bytes), media = [...parts].filter(([name]) => name.startsWith("ppt/media/"));
    assert.equal(media.length, 1); assert.deepEqual(media[0][1], png);
    const slide = parts.get("ppt/slides/slide1.xml").toString();
    assert.match(slide, /<a:srcRect l="12500" t="25000" r="25000" b="0"/);
    assert.match(slide, /<a:alphaModFix amt="60000"/);
    assert.match(slide, /<a:xfrm rot="1800000" flipH="1"><a:off x="254000" y="254000"\/><a:ext cx="1524000" cy="1016000"/);
    const graph = await inspectOpcPackage(fresh.bytes), edges = graph.relationships.filter(r => r.type.endsWith("/image"));
    assert.equal(new Set(edges.map(r => r.source)).size, 2);
    assert.ok(edges.every(r => r.targetPart === `/${media[0][0]}`));
    const rendered = await renderPresentationToSvg(s.document);
    assert.deepEqual(rendered.warnings, []); assert.match(rendered.slides[0].svg, /data:image\/png;base64,/);
    sizes.push([...parts.values()].reduce((n, bytes) => n + bytes.length, 0));
    generated = fresh; s.dispose();
  }
  assert.equal(new Set(sizes).size, 1, "no uncompressed resource/markup growth on repeated fresh projections");
});

test("shared imported image replacement changes all fresh-output relationships and bytes but refuses preservation edits", async () => {
  const generated = await generatePptx(createImageImportFixture(createPresentation, png)), s = await importPptx(generated.bytes, { projection: "content" });
  const before = imageSemantics(s.document), assetId = s.document.assets[0].id;
  s.document.replaceAsset(assetId, { kind: "image", mimeType: "image/gif", source: { type: "bytes", value: gif } });
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  const fresh = await generatePptx(s.document), parts = readGeneratedZip(fresh.bytes);
  const media = [...parts].filter(([name]) => name.startsWith("ppt/media/"));
  assert.equal(media.length, 1); assert.match(media[0][0], /\.gif$/); assert.deepEqual(media[0][1], gif);
  const again = await importPptx(fresh.bytes, { projection: "content" }), after = imageSemantics(again.document);
  const replace = rows => rows.map(row => row.map(function update(e) {
    if (e.asset) return { ...e, asset: { mimeType: "image/gif", source: { type: "bytes", encoding: "base64", value: gif.toString("base64") } } };
    if (e.children) return { ...e, children: e.children.map(update) };
    return e;
  }));
  assert.deepEqual(after, replace(before));
  assert.deepEqual([...readGeneratedZip(generated.bytes)].find(([name]) => name.startsWith("ppt/media/"))[1], png);
  again.dispose(); s.dispose();
});

test("edited Core content can be freshly exported while the source session refuses edit-preservation claims", async () => {
  const source = await generatePptx(createContentImportFixture(createPresentation)), session = await importPptx(source.bytes, { projection: "content" });
  const plan = await session.prepareSave({ mode: "no-op" }), slide = session.document.slides[0], first = slide.removeElement(slide.elements[0].id);
  slide.insertElement(0, { ...first, content: "Edited <&> content" });
  assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
  assert.equal((await session.prepareSave({ mode: "edit" })).report.assurance, "none");
  await assert.rejects(session.save(plan.plan), { code: "save-snapshot-stale" });
  const fresh = await generatePptx(session.document), imported = await importPptx(fresh.bytes, { projection: "content" });
  assert.equal(normalizePresentation(imported.document).slides[0].elements[0].plainText, "Edited <&> content");
  assert.match(readGeneratedZip(source.bytes).get("ppt/slides/slide1.xml").toString(), /Editable &amp; exact/);
  imported.dispose(); session.dispose();
});

function referenceSemantics(document) {
  const p = normalizePresentation(document), names = new Map(), slideIndices = new Map(p.slides.map((s, i) => [s.id, i]));
  const collect = elements => { for (const e of elements) { names.set(e.id, e.name); if (e.type === "group") collect(e.children); } };
  p.slides.forEach(s => collect(s.elements));
  const visit = value => {
    if (!value || typeof value !== "object") return value;
    if (Array.isArray(value)) return value.map(visit);
    const result = {};
    for (const [key, field] of Object.entries(value)) {
      if (key === "id") continue;
      result[key] = key === "elementId" ? names.get(field) : key === "slideId" ? slideIndices.get(field) : visit(effectiveStyleField(key, field, p.theme));
    }
    // Attached geometry is resolved by Layout, not by the source cache box.
    if (value.type === "connector") delete result.box;
    return result;
  };
  return p.slides.map(s => visit(s.elements));
}

test("fresh color parity resolves source theme roles and still detects incorrect colors and unrelated style changes", async () => {
  const colors = { text1: "2A4C6E", accent1: "C05090", background1: "F0E0D0", hyperlink: "123ABC" };
  const make = (direct = false, changed = {}) => {
    const p = createPresentation({ theme: { colors } });
    p.addSlide().addElement({
      type: "text", name: "Theme color parity", accessibility: { description: "Theme color parity" },
      box: { x: 20, y: 30, width: 400, height: 100 },
      content: [{ runs: Object.entries(colors).map(([theme, value]) => ({
        text: theme, style: { fontFamily: "Arial", color: direct ? value : { theme }, ...changed },
      })) }],
    });
    return p;
  };
  const original = make(), direct = make(true), wrong = make(true, { color: "000000" });
  for (const semantics of [tableSemantics, referenceSemantics]) {
    assert.deepEqual(semantics(original), semantics(direct));
    assert.notDeepEqual(semantics(original), semantics(wrong), "color assertions must detect a wrong palette or fallback");
    assert.notDeepEqual(semantics(original), semantics(make(true, { bold: true })), "other style fields remain compared");
    assert.deepEqual(normalizePresentation(original).slides[0].elements[0].content[0].runs[0].style.color, { theme: "text1" });
  }
  let output = await generatePptx(original);
  const sourceParts = readGeneratedZip(output.bytes);
  assert.match(sourceParts.get("ppt/theme/theme1.xml").toString(), /<a:dk1><a:srgbClr val="2A4C6E"/);
  assert.match(sourceParts.get("ppt/slideMasters/slideMaster1.xml").toString(), /tx1="dk1"/);
  assert.match(sourceParts.get("ppt/slides/slide1.xml").toString(), /<a:schemeClr val="tx1"/);
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(output.bytes, { projection: "content" });
    try {
      const runs = normalizePresentation(session.document).slides[0].elements[0].content[0].runs;
      assert.deepEqual(runs.map(run => run.style.color), Object.values(colors));
      assert.deepEqual(tableSemantics(session.document), tableSemantics(original));
      assert.deepEqual(referenceSemantics(session.document), referenceSemantics(original));
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, output.bytes);
      output = await generatePptx(session.document);
      const xml = readGeneratedZip(output.bytes).get("ppt/slides/slide1.xml").toString();
      for (const color of Object.values(colors)) assert.ok(xml.includes(`<a:srgbClr val="${color}"`));
    } finally {
      session.dispose();
    }
  }
});

test("fresh font parity compares actual families and detects wrong role resolution with independent source font XML", async () => {
  const fonts = { heading: "Heading Typeface", body: "Body Typeface" };
  const make = (direct = false, wrong = false) => {
    const p = createPresentation({ theme: { fonts } });
    p.addSlide().addElement({ type: "text", name: "Font controls", accessibility: { description: "Font controls" }, box: { x: 10, y: 10, width: 200, height: 80 }, content: [{ runs: Object.entries(fonts).map(([role, family]) => ({ text: role, style: { fontFamily: direct ? wrong ? "Wrong Typeface" : family : { theme: role } } })) }] });
    return p;
  };
  const original = make();
  for (const semantics of [tableSemantics, referenceSemantics]) {
    assert.deepEqual(semantics(original), semantics(make(true)));
    assert.notDeepEqual(semantics(original), semantics(make(true, true)), "font comparisons must detect wrong families");
  }
  assert.deepEqual(normalizePresentation(original).slides[0].elements[0].content[0].runs.map(r => r.style.fontFamily), [{ theme: "heading" }, { theme: "body" }]);
  let bytes = (await generatePptx(original)).bytes;
  const parts = readGeneratedZip(bytes), theme = parts.get("ppt/theme/theme1.xml").toString(), slide = parts.get("ppt/slides/slide1.xml").toString();
  assert.match(theme, /<a:majorFont><a:latin typeface="Heading Typeface"/); assert.match(theme, /<a:minorFont><a:latin typeface="Body Typeface"/);
  for (const ref of ["+mj-lt", "+mn-lt"]) for (const script of ["latin", "ea", "cs"]) assert.ok(slide.includes(`<a:${script} typeface="${ref}"/>`));
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content" });
    try {
      assert.deepEqual(normalizePresentation(s.document).slides[0].elements[0].content[0].runs.map(r => r.style.fontFamily), Object.values(fonts));
      for (const semantics of [tableSemantics, referenceSemantics]) assert.deepEqual(semantics(s.document), semantics(original));
      const plan = await s.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await s.save(plan.plan)).bytes, bytes);
      bytes = (await generatePptx(s.document)).bytes;
    } finally { s.dispose(); }
  }
});

test("native connector/action imports preserve semantic references, resolved geometry and fresh native output over three rounds", async () => {
  const original = createReferenceImportFixture(createPresentation), expected = referenceSemantics(original);
  let output = await generatePptx(original, { connectorMode: "native" });
  const expectedGeometry = resolveLayout(original).slides[0].elements;
  assert.deepEqual(expectedGeometry[0].start, { x: 60, y: 60 });
  assert.ok(Math.abs(expectedGeometry[0].end.x - 230) < 1e-9); assert.ok(Math.abs(expectedGeometry[0].end.y - 90) < 1e-9);
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(output.bytes, { projection: "content" });
    assert.deepEqual(referenceSemantics(s.document), expected);
    const layout = resolveLayout(s.document), line = layout.slides[0].elements[0];
    assert.deepEqual(line.start, expectedGeometry[0].start); assert.deepEqual(line.end, expectedGeometry[0].end);
    assert.deepEqual(layout.slides[0].elements[4].children[0].start, { x: 40, y: 10 });
    assert.deepEqual(layout.slides[0].elements[4].children[0].end, { x: 130, y: 80 });
    const fresh = await generatePptx(s.document, { connectorMode: "native" }), parts = readGeneratedZip(fresh.bytes);
    assert.deepEqual(fresh.warnings, []);
    const xml = parts.get("ppt/slides/slide1.xml").toString(), graph = await inspectOpcPackage(fresh.bytes);
    assert.equal((xml.match(/<p:cxnSp>/g) ?? []).length, 2);
    assert.match(xml, /<a:stCxn id="3" idx="3"\/><a:endCxn id="4" idx="2"/);
    assert.match(xml, /<a:stCxn id="8" idx="0"/);
    const slideLinks = graph.relationships.filter(r => r.source === "/ppt/slides/slide1.xml" && r.type.endsWith("/slide"));
    assert.ok(slideLinks.some(r => r.targetPart === "/ppt/slides/slide1.xml"));
    assert.ok(slideLinks.some(r => r.targetPart === "/ppt/slides/slide2.xml"));
    assert.ok(graph.relationships.some(r => r.type.endsWith("/hyperlink") && r.target === "https://example.invalid/a?x=1&y=%E4%B8%AD"));
    assert.match(parts.get("ppt/slides/_rels/slide1.xml.rels").toString(), /a\?x=1&amp;y=%E4%B8%AD/);
    const svg = await renderPresentationToSvg(s.document);
    assert.match(svg.slides[0].svg, /60,60 230,90/);
    assert.match(svg.slides[0].svg, /href="https:\/\/example.invalid\/a\?x=1&amp;y=%E4%B8%AD"/);
    assert.ok(svg.warnings.some(w => w.code === "slide-action-not-linked"));
    const unchanged = await s.prepareSave({ mode: "no-op" });
    assert.deepEqual((await s.save(unchanged.plan)).bytes, output.bytes);
    output = fresh; s.dispose();
  }
});

test("imported reference edits, copied self-links and moved targets reach fresh native output without source mutation", async () => {
  const source = await generatePptx(createReferenceImportFixture(createPresentation), { connectorMode: "native" });
  const s = await importPptx(source.bytes, { projection: "content" }), first = s.document.slides[0], copy = s.document.duplicateSlide(first.id, 0);
  s.document.removeSlide(first.id);
  const oldTarget = copy.removeElement(copy.elements[1].id);
  copy.insertElement(1, { ...oldTarget, box: { x: 100, y: 100, width: 40, height: 40 } });
  assert.equal(copy.elements[0].action.slideId, copy.id);
  assert.equal(copy.elements[0].start.elementId, oldTarget.id);
  assert.deepEqual(resolveLayout(s.document).slides[0].elements[0].start, { x: 140, y: 120 });
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  const fresh = await generatePptx(s.document, { connectorMode: "native" });
  assert.deepEqual(fresh.warnings, []);
  const again = await importPptx(fresh.bytes, { projection: "content" });
  assert.deepEqual(resolveLayout(again.document).slides[0].elements[0].start, { x: 140, y: 120 });
  assert.equal(again.document.slides[0].elements[0].action.slideId, again.document.slides[0].id);
  const original = await importPptx(source.bytes, { projection: "content" });
  assert.deepEqual(resolveLayout(original.document).slides[0].elements[0].start, { x: 60, y: 60 });
  original.dispose(); again.dispose(); s.dispose();
});
