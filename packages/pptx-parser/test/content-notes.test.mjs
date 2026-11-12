import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation } from "@pptkit/core";
import { importPptx, PptxImportError } from "../dist/index.js";
import { body } from "./content-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { notesBytes, notesFixture, notesShape, notesTree } from "./notes-fixture.mjs";

const open = (input = notesBytes(), extra = {}) => importPptx(input, { projection: "content", ...extra });
const notes = s => normalizePresentation(s.document).slides[0].notes;
const code = expected => e => e instanceof PptxImportError && e.code === expected;
const unmapped = s => s.report.findings.filter(f => f.operation === "import" && f.feature === "notes.unmapped-property");

test("direct body notes import rich text without requiring notes-page geometry in both dialects", async () => {
  for (const strict of [false, true]) {
    const s = await open(notesBytes(undefined, strict)), n = notes(s);
    assert.equal(n[0].runs[0].text, " A & B "); assert.equal(n[0].runs[0].style.fontSize, 24);
    assert.equal(n[0].runs[0].style.bold, true); assert.equal(n[0].runs[0].style.fontFamily, "Test Font");
    assert.equal(s.document.slides[1].notes, undefined);
    const id = s.report.identities.find(i => i.kind === "notes");
    assert.equal(id.id, s.document.slides[0].id); assert.equal(id.source.part, "/deck/notes/z.xml");
    assert.equal(id.source.sourceObjectId, "2"); assert.match(id.source.path, /spTree\[1\]\/p:sp\[1\]$/);
    assert.ok(Object.isFrozen(id.source));
    const findings = s.report.findings.filter(f => f.feature === "notes.body.mapped-fields");
    assert.equal(findings.length, 7);
    assert.equal(findings.find(f => f.operation === "preview").tier, "unsupported");
    assert.equal(findings.find(f => f.operation === "export").tier, "editable");
    assert.equal(findings.find(f => f.operation === "edit-save").tier, "unsupported");
    assert.equal(findings.find(f => f.operation === "no-op-save").tier, "preserved");
    s.dispose();
  }
});

test("notes paragraphs preserve blanks, explicit breaks, numbering, direct defaults and Unicode", async () => {
  const text = body('<a:p><a:pPr marL="25400" indent="-12700"><a:buAutoNum type="arabicPeriod" startAt="3"/><a:defRPr b="1" sz="1800"/></a:pPr><a:r><a:t>α😀é</a:t></a:r><a:br/><a:r><a:rPr b="0"/><a:t>  end </a:t></a:r></a:p><a:p><a:endParaRPr sz="1200"/></a:p>');
  const s = await open(notesBytes(notesShape({ text }))), n = notes(s);
  assert.equal(n.length, 2); assert.deepEqual(n[0].runs.map(r => r.text), ["α😀é", "\n", "  end "]);
  assert.equal(n[0].runs[0].style.bold, true); assert.equal(n[0].runs[2].style.bold, false);
  assert.equal(n[1].runs[0].text, ""); assert.equal(n[1].runs[0].style.fontSize, 12);
  assert.deepEqual(n[0].style.bullet, { type: "number", style: "arabicPeriod", startAt: 3 });
  s.dispose();
});

test("slide images, page numbers, headers, footers, named and nested arbitrary text never become speaker notes", async () => {
  const shapes = ["sldImg", "sldNum", "hdr", "ftr", "dt"].map((role, i) => notesShape({ id: String(i + 3), role })).join("");
  const arbitrary = notesShape({ id: "9", placeholder: false, name: "Speaker Notes" });
  const nested = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${notesShape({ id: "11" })}</p:grpSp>`;
  const s = await open(notesBytes(shapes + arbitrary + nested));
  assert.equal(s.document.slides[0].notes, undefined);
  assert.ok(unmapped(s).length >= 7); assert.equal(s.report.identities.some(i => i.kind === "notes"), false);
  s.dispose();
});

test("absent notes, absent body and explicit empty body remain distinguishable", async () => {
  for (const markup of ["", notesShape({ text: body("<a:p/>") })]) {
    const s = await open(notesBytes(markup));
    assert.equal(s.document.slides[0].notes === undefined, markup === "");
    assert.deepEqual(notes(s).map(p => p.runs.map(r => r.text)), [[""]]);
    assert.equal(s.document.slides[1].notes, undefined); s.dispose();
  }
});

test("ambiguous body placeholders and missing local text stay source-only with a reason", async () => {
  for (const markup of [notesShape() + notesShape({ id: "3" }), notesShape({ text: "" })]) {
    const s = await open(notesBytes(markup));
    assert.equal(s.document.slides[0].notes, undefined);
    assert.ok(s.report.diagnostics.some(d => d.code === "notes-body-unmapped"));
    assert.ok(s.report.findings.some(f => f.feature === "notes.body" && f.operation === "import" && f.tier === "unsupported"));
    assert.equal(s.report.identities.some(i => i.kind === "notes"), false); s.dispose();
  }
});

test("notes layout, fields, unknown effects and master inheritance remain located unmapped findings", async () => {
  const text = body('<a:p><a:fld id="{source}" type="slidenum"><a:t>77</a:t></a:fld><a:r><a:t>Own</a:t></a:r></a:p>');
  const s = await open(notesBytes(notesShape({ text, props: '<a:effectLst><a:outerShdw blurRad="900"/></a:effectLst>' })));
  assert.equal(notes(s)[0].runs[0].text, "Own");
  for (const suffix of ["a:bodyPr[1]", "p:spPr[1]", "a:fld[1]"]) {
    const f = unmapped(s).find(f => f.source.path.endsWith(suffix));
    assert.ok(f, suffix); assert.equal(f.source.sourceObjectId, "2"); assert.equal(f.source.part, "/deck/notes/z.xml");
  }
  assert.ok(s.report.findings.some(f => f.feature === "notes.inheritance-and-defaults" && f.tier === "fallback"));
  assert.notEqual(notes(s)[0].runs[0].style.fontSize, 99);
  assert.ok(s.report.findings.some(f => f.source.part === "/deck/notes-master.xml" && f.operation === "no-op-save" && f.tier === "preserved"));
  s.dispose();
});

test("drawing identities remain unique within each notes part including omitted drawing families", async () => {
  for (const markup of [notesShape({ id: "1" }), notesShape() + notesShape({ id: "2", role: "sldImg" }), notesShape({ id: "bad" }), notesShape({ id: "4294967296" })]) {
    await assert.rejects(open(notesBytes(markup)), code("import-invalid-package"));
  }
  const s = await open(); // Source slide and notes can each have drawing ID 2.
  assert.equal(s.report.identities.filter(i => i.source.sourceObjectId === "2").length, 2); s.dispose();
});

test("notes run clicks use notes-part relationships for exact URL and internal slide targets", async () => {
  const text = body('<a:p><a:r><a:rPr><a:hlinkClick r:id="url" tooltip="Open &amp;"/></a:rPr><a:t>URL</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="slide" action="ppaction://hlinksldjump"/></a:rPr><a:t>Next</a:t></a:r></a:p>');
  const f = notesFixture(notesShape({ text }));
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("slide", "slide", "../a.xml"), { ...f.rel("url", "hyperlink", "https://example.invalid/a?x=1&y=%E4%B8%AD"), TargetMode: "External" }]);
  const s = await open(encodeStructure(f)), n = notes(s);
  assert.deepEqual(n[0].runs[0].action, { type: "url", url: "https://example.invalid/a?x=1&y=%E4%B8%AD", tooltip: "Open &" });
  assert.deepEqual(n[0].runs[1].action, { type: "slide", slideId: s.document.slides[1].id });
  assert.equal(s.report.findings.filter(f => f.feature === "notes.click-action" && f.operation === "import").length, 2);
  assert.ok(s.report.findings.some(f => f.feature === "notes.click-action" && f.source.relationshipId === "url" && f.source.part === "/deck/notes/z.xml" && f.slideId === s.document.slides[0].id && !f.elementId));
  s.dispose();
});

test("notes owning-slide backlinks remain distinct from slide-jump relationships and may share a self-link", async () => {
  const text = body('<a:p><a:r><a:rPr><a:hlinkClick r:id="back" action="ppaction://hlinksldjump"/></a:rPr><a:t>Self</a:t></a:r></a:p>');
  const f = notesFixture(notesShape({ text })), s = await open(encodeStructure(f));
  assert.deepEqual(notes(s)[0].runs[0].action, { type: "slide", slideId: s.document.slides[0].id });
  assert.equal(s.report.inspection.findings.some(f => f.message.includes("slide backlink")), false);
  s.dispose();
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../a.xml")]);
  const missing = await open(encodeStructure(f));
  assert.equal(notes(missing)[0].runs[0].action.slideId, missing.document.slides[1].id);
  assert.ok(missing.report.inspection.findings.some(f => f.code === "missing-inheritance" && f.part === "/deck/notes/z.xml" && f.message.includes("slide backlink")));
  missing.dispose();
});

test("invalid implicit notes owners cannot be hidden by explicit slide actions", async () => {
  const text = body('<a:p><a:r><a:rPr><a:hlinkClick r:id="jump"/></a:rPr><a:t>Jump</a:t></a:r></a:p>');
  for (const extra of [[["back", "../a.xml"], ["jump", "../z.xml"]], [["back", "../z.xml"], ["duplicate", "../z.xml"], ["jump", "../a.xml"]]]) {
    const f = notesFixture(notesShape({ text }));
    f.parts["deck/notes/_rels/z.xml.rels"] = relationships(extra.map(([id, target]) => f.rel(id, "slide", target)));
    await assert.rejects(open(encodeStructure(f)), code("import-invalid-package"));
  }
});

test("unsafe or unresolved notes actions remain source-only and insertion defaults never create clicks", async () => {
  const f = notesFixture(notesShape({ text: body('<a:p><a:pPr><a:defRPr><a:hlinkClick r:id="unsafe"/></a:defRPr></a:pPr><a:r><a:rPr><a:hlinkClick r:id="unsafe"/></a:rPr><a:t>Unsafe</a:t></a:r><a:r><a:rPr><a:hlinkClick r:id="missing"/></a:rPr><a:t>Missing</a:t></a:r></a:p>') }));
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), { ...f.rel("unsafe", "hyperlink", "javascript:alert(1)"), TargetMode: "External" }]);
  const s = await open(encodeStructure(f));
  assert.ok(notes(s)[0].runs.every(r => r.action === undefined));
  assert.equal(s.report.diagnostics.filter(d => d.code === "notes-click-action-unmapped").length, 2);
  assert.equal(s.report.findings.filter(f => f.feature === "notes.click-action" && f.operation === "import").length, 2);
  assert.ok(unmapped(s).some(f => f.source.path.includes("a:defRPr[1]/a:hlinkClick[1]"))); s.dispose();
});

test("effective MCE selection and namespace aliases determine the notes body", async () => {
  const f = notesFixture();
  const choice = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:future"><mc:Choice Requires="future">${notesShape({ text: body("<a:p><a:r><a:t>Wrong</a:t></a:r></a:p>") })}</mc:Choice><mc:Fallback>${notesShape()}</mc:Fallback></mc:AlternateContent>`;
  f.parts["deck/notes/z.xml"] = f.p("notes", notesTree(choice), `xmlns:a="${f.A}"`).replace(/(<\/?)p:/g, "$1n:").replaceAll("xmlns:p=", "xmlns:n=").replace(/(<\/?)a:/g, "$1d:").replaceAll("xmlns:a=", "xmlns:d=");
  const s = await open(encodeStructure(f));
  assert.equal(notes(s)[0].runs[0].text, " A & B ");
  assert.match(s.report.identities.find(i => i.kind === "notes").source.path, /^\/n:notes\[1\].*n:sp\[1\]$/);
  s.dispose();
});

test("malformed mapped note styles reject, while report and snapshot budgets bound notes work", async () => {
  await assert.rejects(open(notesBytes(notesShape({ text: body('<a:p><a:r><a:rPr sz="oops"/><a:t>Bad</a:t></a:r></a:p>') }))), code("import-invalid-package"));
  const baseline = await open(), length = baseline.report.findings.length;
  baseline.dispose();
  await assert.rejects(open(notesBytes(), { sessionLimits: { maxReportFindings: length - 1 } }), code("import-limit-exceeded"));
  const large = notesBytes(notesShape({ text: body(Array(500).fill("<a:p><a:r><a:t>bounded</a:t></a:r></a:p>").join("")) }));
  await assert.rejects(open(large, { sessionLimits: { maxSnapshotValues: 2000 } }), code("import-limit-exceeded"));
  const abort = new AbortController(), pending = open(large, { signal: abort.signal }); setTimeout(() => abort.abort(), 0);
  await assert.rejects(pending, code("operation-aborted"));
  const recovered = await open(); assert.equal(notes(recovered)[0].runs[0].text, " A & B "); recovered.dispose();
});

test("notes edits invalidate no-op plans, exact undo restores them, and disposed documents keep their notes", async () => {
  const input = notesBytes(), s = await open(input), slide = s.document.slides[0], baseline = slide.notes;
  const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  slide.setNotes("Edited");
  await assert.rejects(s.save(prepared.plan), code("save-snapshot-stale"));
  assert.equal((await s.prepareSave({ mode: "no-op" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  slide.setNotes(baseline);
  assert.equal(Buffer.compare(Buffer.from((await s.save(prepared.plan)).bytes), input), 0);
  s.dispose(); assert.equal(notes(s)[0].runs[0].text, " A & B ");
  slide.setNotes(); assert.equal(slide.notes, undefined);
});

test("structural projection keeps notes source-only", async () => {
  const s = await importPptx(notesBytes(), { projection: "structure" });
  assert.equal(s.document.slides[0].notes, undefined);
  assert.equal(s.report.identities.some(i => i.kind === "notes"), false);
  s.dispose();
});
