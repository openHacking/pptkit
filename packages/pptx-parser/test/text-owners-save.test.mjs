import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { normalizePresentation } from "../../core/dist/index.js";
import { importPptx, compareOpcPackages, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { analyzeTextEdits } from "../dist/import/analyze-text-edits.js";
import { body, shape } from "./content-fixture.mjs";
import { notesFixture, notesShape, notesTree } from "./notes-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { zip } from "./zip-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v2" };
const code = expected => error => { assert.equal(error.code, expected); return true; };
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const paragraphs = text => `<a:p><a:r><a:rPr sz="2400" b="1"/><a:t>${text}</a:t></a:r></a:p>`;
const textBody = text => body(paragraphs(text));
const repack = parts => zip([...parts].map(([name, bytes]) => ({ name, body: bytes }))).bytes;
const group = content => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000"><a:off x="254000" y="127000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="127000" y="254000"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;

function fixture(strict = false) {
  const f = notesFixture(notesShape({ text: textBody("Notes old"), props: '<a:effectLst><a:glow rad="12700"/></a:effectLst>' }), strict);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace(shape({ txBox: true, text: body() }),
    shape({ text: textBody("Shape old"), preset: "roundRect", fill: '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>' }) +
    shape({ id: "3", txBox: true, text: textBody("Textbox old") }) +
    group(shape({ id: "5", preset: "ellipse", text: textBody("Nested old") })));
  f.parts["orphan.xml"] = '<unknown xmlns="urn:future">preserve me</unknown>';
  return f;
}
const open = (f = fixture(), options = {}) => importPptx(encodeStructure(f), { ...profile, ...options });
const all = elements => elements.flatMap(element => [element, ...(element.type === "group" ? all(element.children) : [])]);
function replaceDrawing(session, sourceId, modify) {
  const slide = session.document.slides[0], id = `${slide.id}-object-${sourceId}`;
  const top = slide.elements.find(element => all([element]).some(item => item.id === id));
  assert.ok(top, id);
  const index = slide.elements.indexOf(top), cloned = structuredClone(top);
  modify(all([cloned]).find(element => element.id === id));
  slide.removeElement(top.id); slide.insertElement(index, cloned);
}
const editShape = (session, text, id = 2) => replaceDrawing(session, id, element => {
  (element.type === "text" ? element.content : element.text.content)[0].runs[0].text = text;
});
function editNotes(session, text, slideIndex = 0) {
  const slide = session.document.slides[slideIndex], content = structuredClone(slide.notes);
  content[0].runs[0].text = text; slide.setNotes(content);
}
async function save(session) {
  const prepared = await session.prepareSave({ mode: "edit" });
  assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
  return { ...await session.save(prepared.plan), plan: prepared.plan };
}
function checkUntouched(before, after, changed) {
  const source = readGeneratedZip(before), candidate = readGeneratedZip(after);
  assert.deepEqual([...candidate.keys()].sort(), [...source.keys()].sort());
  for (const [name, bytes] of source) if (!changed.includes(name)) assert.equal(digest(candidate.get(name)), digest(bytes), name);
}

test("v2 saves shape, text-box, grouped shape and notes characters with distinct source ownership in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict), input = encodeStructure(f), session = await open(f);
    try {
      assert.deepEqual(session.report.textRuns.map(run => run.owner).sort(), ["shape-text", "shape-text", "text-box"]);
      assert.equal(session.report.notesTextRuns.length, 1);
      const notes = session.report.notesTextRuns[0];
      assert.equal(notes.slideId, session.document.slides[0].id);
      assert.equal("elementId" in notes, false); assert.equal("owner" in notes, false);
      assert.equal(notes.source.sourceObjectId, "2"); assert.equal(notes.source.part, "/deck/notes/z.xml");
      assert.equal(notes.sourceElementIndex, notes.source.elementIndex);
      assert.ok(Object.isFrozen(notes) && Object.isFrozen(notes.source));
      editShape(session, "Shape <&>🙂\r\n"); editShape(session, "Box new", 3); editShape(session, "Nested new", 5); editNotes(session, "Notes <&>🙂\r\n");
      const result = await save(session), parts = readGeneratedZip(result.bytes);
      assert.equal(result.report.assurance, "scoped-preservation"); assert.deepEqual(result.report.losses, []);
      assert.equal(result.report.changes.length, 4);
      const notesChange = result.report.changes.find(change => change.source.part === "/deck/notes/z.xml");
      assert.equal(notesChange.documentPath, "/slides/0/notes/0/runs/0/text");
      assert.equal("elementId" in notesChange, false);
      assert.equal(notesChange.slideId, session.document.slides[0].id);
      assert.ok(result.report.changes.some(change => change.documentPath.endsWith("/text/content/0/runs/0/text")));
      assert.ok(result.report.changes.every(change => change.kind === "text"));
      assert.equal(parts.get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace(">Shape old</a:t>", ">Shape &lt;&amp;&gt;🙂&#xD;\n</a:t>").replace(">Textbox old</a:t>", ">Box new</a:t>").replace(">Nested old</a:t>", ">Nested new</a:t>"));
      assert.equal(parts.get("deck/notes/z.xml").toString(), f.parts["deck/notes/z.xml"].replace(">Notes old</a:t>", ">Notes &lt;&amp;&gt;🙂&#xD;\n</a:t>"));
      checkUntouched(input, result.bytes, ["deck/z.xml", "deck/notes/z.xml"]);
      assert.deepEqual((await compareOpcPackages(input, result.bytes)).differences.map(d => [d.code, d.part]),
        [["payload-changed", "/deck/notes/z.xml"], ["payload-changed", "/deck/z.xml"]]);
      const reopened = await importPptx(result.bytes, profile);
      assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document));
      reopened.dispose();
    } finally { session.dispose(); }
  }
});

test("default/v1 scopes and report shapes stay compatible while v2 requires content projection", async () => {
  for (const editSaveProfile of [undefined, "text-runs-v1"]) {
    const session = await open(fixture(), { editSaveProfile });
    assert.equal(session.report.notesTextRuns, undefined);
    if (editSaveProfile) {
      assert.equal(session.report.textRuns.length, 1);
      assert.equal("owner" in session.report.textRuns[0], false);
    } else assert.equal(session.report.textRuns, undefined);
    editShape(session, "blocked shape"); editNotes(session, "blocked notes");
    const prepared = await session.prepareSave({ mode: "edit" });
    assert.equal(prepared.status, "blocked");
    assert.equal(prepared.report.diagnostics[0].code, editSaveProfile ? "save-edit-unsupported" : "save-mode-unsupported");
    session.dispose();
  }
  await assert.rejects(importPptx(encodeStructure(fixture()), { projection: "structure", editSaveProfile: "text-runs-v2" }), TypeError);
  const v1 = await open(fixture(), { editSaveProfile: "text-runs-v1" });
  editShape(v1, "old profile still works", 3); await save(v1); v1.dispose();
});

test("all seven mapped shape presets and painted text boxes keep geometry/styles unchanged", async () => {
  for (const preset of ["rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow", "chevron"]) {
    const f = fixture();
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('prst="roundRect"', `prst="${preset}"`).replace('<p:cNvSpPr txBox="0"/>', '<p:cNvSpPr txBox="1"/>');
    const input = encodeStructure(f), session = await open(f);
    const binding = session.report.textRuns.find(run => run.source.sourceObjectId === "2");
    assert.equal(binding.owner, "shape-text"); editShape(session, "painted text");
    const result = await save(session); checkUntouched(input, result.bytes, ["deck/z.xml"]);
    assert.equal(readGeneratedZip(result.bytes).get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace(">Shape old</a:t>", ">painted text</a:t>"));
    session.dispose();
  }
});

test("unmapped shape families do not acquire bindings or permission through v2", async () => {
  const f = fixture(); f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('prst="roundRect"', 'prst="cloud"');
  const session = await open(f);
  assert.equal(session.report.textRuns.some(run => run.source.sourceObjectId === "2"), false);
  session.document.slides[0].addElement({ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 50, height: 50 }, text: { content: "new" } });
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); session.dispose();
});

test("repeated text in different owners changes only selected raw atoms and leaves effects/master text intact", async () => {
  const f = fixture();
  for (const part of ["deck/z.xml", "deck/notes/z.xml"]) f.parts[part] = f.parts[part].replaceAll("Shape old", "Same").replaceAll("Textbox old", "Same").replaceAll("Nested old", "Same").replaceAll("Notes old", "Same");
  const session = await open(f); editShape(session, "shape selected"); editNotes(session, "notes selected");
  const result = await save(session), parts = readGeneratedZip(result.bytes);
  assert.equal(parts.get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace(">Same</a:t>", ">shape selected</a:t>"));
  assert.equal(parts.get("deck/notes/z.xml").toString(), f.parts["deck/notes/z.xml"].replace(">Same</a:t>", ">notes selected</a:t>"));
  assert.ok(parts.get("deck/notes/z.xml").toString().includes('<a:glow rad="12700"/>'));
  assert.equal(parts.get("deck/notes-master.xml").toString(), f.parts["deck/notes-master.xml"]);
  session.dispose();
});

test("notes aliases and inactive MCE duplicates bind the selected raw body without rewriting alternatives", async () => {
  const f = fixture(true), selected = notesShape({ text: textBody("Notes old") });
  const alternatives = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x">${selected.replace("Notes old", "Inactive")}</mc:Choice><mc:Fallback>${selected}</mc:Fallback></mc:AlternateContent>`;
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(alternatives), `xmlns:a="${f.A}"`).replace(/(<\/?)p:/g, "$1n:").replaceAll("xmlns:p=", "xmlns:n=").replace(/(<\/?)a:/g, "$1d:").replaceAll("xmlns:a=", "xmlns:d=");
  const session = await open(f); editNotes(session, "Selected new");
  const result = await save(session), xml = readGeneratedZip(result.bytes).get("deck/notes/z.xml").toString();
  assert.equal(xml, f.parts["deck/notes/z.xml"].replace(">Notes old</d:t>", ">Selected new</d:t>"));
  assert.ok(xml.includes(">Inactive</d:t>")); assert.ok(session.report.notesTextRuns[0].sourceElementIndex > 20);
  session.dispose();
});

test("mixed notes atoms and valid shape changes block together and recover after restoring the unmapped run", async () => {
  const f = fixture(); f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace("Notes old", "Notes <!--retain-->old");
  const session = await open(f);
  assert.equal(session.report.notesTextRuns[0].status, "unmapped");
  editShape(session, "safe"); editNotes(session, "unsafe");
  const blocked = await session.prepareSave({ mode: "edit" });
  assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.changes.length, 2);
  assert.equal(blocked.report.diagnostics[0].code, "save-edit-unsupported");
  editNotes(session, "Notes old");
  const result = await save(session); checkUntouched(encodeStructure(f), result.bytes, ["deck/z.xml"]); session.dispose();
});

test("break/synthetic notes runs remain unmapped and adjacent fields are retained while real text changes", async () => {
  for (const noteText of [body("<a:p><a:br/></a:p>"), body("<a:p><a:endParaRPr b=\"1\"/></a:p>")]) {
    const f = notesFixture(notesShape({ text: noteText })), session = await open(f);
    assert.equal(session.report.notesTextRuns[0].status, "unmapped"); editNotes(session, "new");
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); session.dispose();
  }
  const f = fixture();
  f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace("<a:r><a:rPr", '<a:fld id="{source-field}" type="slidenum"><a:t>99</a:t></a:fld><a:r><a:rPr');
  const session = await open(f); editNotes(session, "real run");
  const result = await save(session);
  assert.equal(readGeneratedZip(result.bytes).get("deck/notes/z.xml").toString(), f.parts["deck/notes/z.xml"].replace(">Notes old</a:t>", ">real run</a:t>"));
  session.dispose();
});

test("notes range guards are scoped by source part and drawing identity, not the equal slide drawing ID", async () => {
  for (const guardedPart of ["deck/z.xml", "deck/notes/z.xml"]) {
    const f = fixture(), close = guardedPart.includes("/notes/") ? "</p:notes>" : "</p:sld>";
    f.parts[guardedPart] = f.parts[guardedPart].replace(close, `<p:timing><p:spTgt spid=" +0002 "><p:txEl><p:charRg st="0" end="3"/></p:txEl></p:spTgt></p:timing>${close}`);
    const session = await open(f);
    if (guardedPart.includes("/notes/")) editNotes(session, "blocked"); else editShape(session, "blocked");
    const blocked = await session.prepareSave({ mode: "edit" });
    assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.diagnostics[0].code, "save-text-range-reference");
    assert.equal(blocked.report.diagnostics[0].source.part, `/${guardedPart}`);
    if (guardedPart.includes("/notes/")) { editNotes(session, "Notes old"); editShape(session, "safe"); }
    else { editShape(session, "Shape old"); editNotes(session, "safe"); }
    const result = await save(session); checkUntouched(encodeStructure(f), result.bytes, [guardedPart.includes("/notes/") ? "deck/z.xml" : "deck/notes/z.xml"]);
    session.dispose();
  }
});

test("inactive notes range targets block while whole-shape and other-owner targets retain exact bytes", async () => {
  for (const [target, expected] of [['<p:spTgt spid="2"><p:txEl/></p:spTgt>', "blocked"], ['<p:spTgt spid="2"/>', "ready"], ['<p:spTgt spid="3"><p:txEl/></p:spTgt>', "ready"]]) {
    const f = fixture();
    f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace("</p:notes>", `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><p:timing>${target}</p:timing></mc:Choice><mc:Fallback/></mc:AlternateContent></p:notes>`);
    const session = await open(f); editNotes(session, "new");
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, expected);
    if (expected === "ready") assert.ok(readGeneratedZip((await session.save(prepared.plan)).bytes).get("deck/notes/z.xml").toString().includes(target));
    session.dispose();
  }
});

test("shared notes parts retain import rejection and independent notes analysis still reconciles aliased source atoms", async () => {
  const f = fixture();
  f.parts["deck/_rels/a.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes/z.xml")]);
  await assert.rejects(open(f), code("import-invalid-package"));
  const session = await open(), baseline = normalizePresentation(session.document);
  baseline.slides[1].notes = structuredClone(baseline.slides[0].notes);
  const candidate = structuredClone(baseline), source = session.report.notesTextRuns[0];
  const bindings = baseline.slides.map(slide => ({ ...source, slideId: slide.id, owner: "notes", baselineText: "Notes old", namespaceUri: fixture().A }));
  candidate.slides[0].notes[0].runs[0].text = "shared new";
  const conflict = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.equal(conflict.diagnostics[0].code, "save-source-shared-conflict");
  assert.equal("elementId" in conflict.diagnostics[0], false);
  candidate.slides[1].notes[0].runs[0].text = "shared new";
  const agreed = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.deepEqual(agreed.diagnostics, []); assert.equal(agreed.changes.length, 2);
  assert.equal(agreed.parts.get("/deck/notes/z.xml").length, 1); session.dispose();
});

test("ambiguous notes bodies and master-only content stay unbound without blocking unrelated drawing edits", async () => {
  for (const noteMarkup of [notesShape() + notesShape({ id: "3" }), notesShape({ text: "" }), notesShape({ role: "sldImg" })]) {
    const f = fixture(); f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(noteMarkup), `xmlns:a="${f.A}"`);
    const session = await open(f); assert.deepEqual(session.report.notesTextRuns, []); editShape(session, "unrelated");
    const result = await save(session); checkUntouched(encodeStructure(f), result.bytes, ["deck/z.xml"]);
    session.document.slides[0].setNotes("invented source");
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); session.dispose();
  }
});

test("notes and shape run hyperlinks retain their owning relationships and actions", async () => {
  const f = fixture();
  for (const part of ["deck/z.xml", "deck/notes/z.xml"]) f.parts[part] = f.parts[part].replace('<a:rPr sz="2400" b="1"/>', '<a:rPr sz="2400" b="1"><a:hlinkClick r:id="url" tooltip="Keep &amp;"/></a:rPr>');
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("notes", "notesSlide", "notes/z.xml"), { ...f.rel("url", "hyperlink", "https://example.invalid/shape?x=1&y=2"), TargetMode: "External" }]);
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("master", "notesMaster", "../notes-master.xml"), { ...f.rel("url", "hyperlink", "https://example.invalid/notes?x=1&y=2"), TargetMode: "External" }]);
  const session = await open(f), before = normalizePresentation(session.document);
  editShape(session, "shape link"); editNotes(session, "notes link");
  const result = await save(session); checkUntouched(encodeStructure(f), result.bytes, ["deck/z.xml", "deck/notes/z.xml"]);
  const reopened = await importPptx(result.bytes, profile), after = normalizePresentation(reopened.document);
  assert.deepEqual(after.slides[0].elements[0].text.content[0].runs[0].action, before.slides[0].elements[0].text.content[0].runs[0].action);
  assert.deepEqual(after.slides[0].notes[0].runs[0].action, before.slides[0].notes[0].runs[0].action);
  assert.notEqual(after.slides[0].notes[0].runs[0].action.url, after.slides[0].elements[0].text.content[0].runs[0].action.url);
  reopened.dispose(); session.dispose();
});

test("style/geometry/topology and fake derived custom-data mutations cannot hide in a mixed v2 plan", async () => {
  const mutations = [
    ["style", s => replaceDrawing(s, 2, e => { e.text.content[0].runs[0].style.bold = false; })],
    ["geometry", s => replaceDrawing(s, 2, e => { e.box.width += 5; })],
    ["structure", s => { const n = structuredClone(s.document.slides[0].notes); n[0].runs.push({ text: "added" }); s.document.slides[0].setNotes(n); }],
    ["other", s => { s.document.slides[0].customData = { text: { plainText: "pretend" } }; }],
  ];
  for (const [kind, mutate] of mutations) {
    const session = await open(), before = structuredClone(session.document.slides[0].elements);
    editNotes(session, "supported"); mutate(session);
    const blocked = await session.prepareSave({ mode: "edit" });
    assert.equal(blocked.status, "blocked", kind); assert.ok(blocked.report.changes.some(change => change.kind === kind), JSON.stringify(blocked.report.changes));
    assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported"));
    assert.deepEqual(before[0].text.content[0].runs[0].text, "Shape old"); session.dispose();
  }
  const session = await open();
  try {
    const baseline = normalizePresentation(session.document);
    baseline.slides[0].customData = { text: { plainText: "original" } };
    const candidate = structuredClone(baseline);
    candidate.slides[0].customData.text.plainText = "changed";
    const analysis = await analyzeTextEdits(baseline, candidate, [], [], DEFAULT_SESSION_LIMITS, async () => {});
    assert.equal(analysis.changes[0].documentPath, "/slides/0/customData/text/plainText");
    assert.equal(analysis.changes[0].kind, "other");
    assert.equal(analysis.diagnostics[0].code, "save-edit-unsupported");
  } finally { session.dispose(); }
});

test("restoring shape/notes state restores original-byte no-op and saved plans retain full-state freshness", async () => {
  const f = fixture(), input = encodeStructure(f), session = await open(f), slide = session.document.slides[0];
  const originalShape = structuredClone(slide.elements[0]), originalNotes = structuredClone(slide.notes);
  const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready");
  editShape(session, "edited"); editNotes(session, "edited");
  const edited = await save(session); editNotes(session, "stale notes");
  await assert.rejects(session.save(edited.plan), code("save-snapshot-stale"));
  editNotes(session, "edited"); assert.equal(digest((await session.save(edited.plan)).bytes), digest(edited.bytes));
  replaceDrawing(session, 2, element => Object.assign(element, originalShape)); slide.setNotes(originalNotes);
  assert.equal(digest((await session.save(noop.plan)).bytes), digest(input));
  assert.equal((await session.prepareSave({ mode: "edit" })).report.diagnostics[0].code, "save-edit-no-changes");
  session.dispose();
});

test("new notes provenance is bounded across the combined drawing/notes run collection", async () => {
  const f = fixture();
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(notesShape({ text: body(`<a:p>${"<a:r><a:t>bounded</a:t></a:r>".repeat(1000)}</a:p>`) })), `xmlns:a="${f.A}"`);
  await assert.rejects(open(f, { sessionLimits: { maxReportFindings: 500 } }), error => {
    assert.equal(error.code, "import-limit-exceeded"); assert.match(error.message, /Text run source bindings/); return true;
  });
});

test("whole-plan change limits and invalid notes characters release capacity for valid independent retries", async () => {
  const session = await open(fixture(), { sessionLimits: { maxSaveChanges: 1, maxLivePlans: 1 } });
  editShape(session, "new shape"); editNotes(session, "new notes");
  await assert.rejects(session.prepareSave({ mode: "edit" }), code("save-limit-exceeded"));
  editShape(session, "Shape old"); editNotes(session, "\0");
  const blocked = await session.prepareSave({ mode: "edit" });
  assert.equal(blocked.status, "blocked"); assert.equal(blocked.report.diagnostics[0].code, "save-text-candidate-unsupported");
  editNotes(session, "valid retry");
  const result = await save(session); assert.equal(result.report.changes.length, 1);
  result.plan.dispose(); const again = await save(session); again.plan.dispose(); session.dispose();
});

test("UTF-16LE drawing and UTF-16BE notes retain encodings and every unselected source byte", async () => {
  const f = fixture(), parts = readGeneratedZip(encodeStructure(f));
  for (const [part, be] of [["deck/z.xml", false], ["deck/notes/z.xml", true]]) {
    const encoded = Buffer.from(`<?xml version="1.0" encoding="UTF-16"?>\r\n${parts.get(part).toString()}`, "utf16le");
    parts.set(part, Buffer.concat([Buffer.from(be ? [254, 255] : [255, 254]), be ? encoded.swap16() : encoded]));
  }
  const input = repack(parts), session = await importPptx(input, profile);
  editShape(session, "编码🙂<&"); editNotes(session, "备注🙂<&");
  const result = await save(session), output = readGeneratedZip(result.bytes);
  for (const [part, be, before, after] of [["deck/z.xml", false, "Shape old", "编码🙂&lt;&amp;"], ["deck/notes/z.xml", true, "Notes old", "备注🙂&lt;&amp;"]]) {
    const decoded = (be ? Buffer.from(output.get(part).subarray(2)).swap16() : output.get(part).subarray(2)).toString("utf16le");
    assert.equal(decoded, `<?xml version="1.0" encoding="UTF-16"?>\r\n${f.parts[part].replace(`>${before}</a:t>`, `>${after}</a:t>`)}`);
    assert.deepEqual([...output.get(part).subarray(0, 2)], be ? [254, 255] : [255, 254]);
  }
  checkUntouched(input, result.bytes, ["deck/z.xml", "deck/notes/z.xml"]); session.dispose();
});

test("three v2 save/reimport rounds in both dialects retain unknown content, owners and stable output size", async () => {
  for (const strict of [false, true]) {
    const f = fixture(strict); let input = encodeStructure(f), previousSize;
    const original = input;
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(input, profile); editShape(session, `Shape round ${round}`); editNotes(session, `Notes round ${round}`);
      const saved = await save(session); checkUntouched(original, saved.bytes, ["deck/z.xml", "deck/notes/z.xml"]);
      assert.equal(saved.report.changes.length, 2);
      assert.equal(session.report.textRuns.find(run => run.source.sourceObjectId === "2").owner, "shape-text");
      assert.equal(session.report.notesTextRuns[0].status, "bound");
      if (previousSize !== undefined) assert.equal(saved.bytes.length, previousSize);
      previousSize = saved.bytes.length; input = saved.bytes; session.dispose();
    }
  }
});

test("notes edit operations use bound v2 source saves and cannot grant action changes", async () => {
  for (const strict of [false, true]) {
    let input = encodeStructure(fixture(strict));
    for (const value of ["First & value", "中文 notes", "Third value"]) {
      const session = await importPptx(input, profile);
      try {
        const doc = session.document, slide = doc.slides[0], old = slide.notes[0].runs[0].text;
        const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
          { type: 'notes-run', slideId: slide.id, paragraphIndex: 0, runIndex: 0, expected: old, value },
        ] }).inverse;
        const ready = await session.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        const redo = doc.applyEditOperation(undo).inverse; await assert.rejects(session.save(ready.plan), code('save-snapshot-stale'));
        doc.applyEditOperation(redo);
        const result = await save(session); checkUntouched(input, result.bytes, ['deck/notes/z.xml']);
        assert.equal(result.report.changes[0].documentPath, '/slides/0/notes/0/runs/0/text');
        const reopened = await importPptx(result.bytes, profile);
        try { assert.equal(reopened.document.slides[0].notes[0].runs[0].text, value); } finally { reopened.dispose(); }
        input = result.bytes;
      } finally { session.dispose(); }
    }
    const session = await importPptx(input, profile);
    try {
      const doc = session.document, slide = doc.slides[0];
      const inverse = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
        { type: 'notes-run-action', slideId: slide.id, paragraphIndex: 0, runIndex: 0, expected: null, value: { type: 'url', url: 'https://example.org', tooltip: null } },
      ] }).inverse;
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      doc.applyEditOperation(inverse); const noop = await session.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready');
      assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(input));
    } finally { session.dispose(); }
  }
});
