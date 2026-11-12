import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { importPptx, inspectPptxTiming, compareOpcPackages } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { structureContext } from "../dist/presentation/structure-context.js";
import { paragraphStableOwners } from "../dist/import/paragraph-targets.js";
import { paragraphProfile, paragraphTarget, paragraphFixture, paragraphBytes, editParagraph } from "./paragraph-target-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";
import { notesFixture, notesShape } from "./notes-fixture.mjs";
import { tableFixture } from "./table-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const open = (f, options = {}) => importPptx(encodeStructure(f), { ...paragraphProfile, ...options });
async function ready(session) {
  const prepared = await session.prepareSave({ mode: "edit" });
  assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
  return session.save(prepared.plan);
}
async function blocked(session, code = "save-text-range-reference") {
  const p = await session.prepareSave({ mode: "edit" }); assert.equal(p.status, "blocked");
  assert.ok(p.report.diagnostics.some(d => d.code === code), JSON.stringify(p.report));
  assert.ok(!("plan" in p)); return p;
}

test("explicit paragraph profile preserves complete timing and exact surrounding payloads through three Unicode scalar save rounds in both dialects and all text profiles", async () => {
  for (const strict of [false, true]) for (const editSaveProfile of ["text-runs-v1", "text-runs-v2", "text-runs-v3"]) {
    const f = paragraphFixture({ strict, txBox: editSaveProfile === "text-runs-v1" });
    let bytes = encodeStructure(f), xml = f.parts["deck/z.xml"];
    const inventory = await inspectPptxTiming(bytes);
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(bytes, { ...paragraphProfile, editSaveProfile });
      try {
        const a = `Round ${round}<&>👩‍💻\r\né`, b = `尾${round}😀`;
        const element = session.document.slides[0].elements[0], content = (element.type === "text" ? element : element.text).content;
        const oldA = content[0].runs[0].text, oldB = content[2].runs[0].text;
        editParagraph(session, a); editParagraph(session, b, 2);
        const saved = await ready(session), parts = readGeneratedZip(saved.bytes);
        const escaped = text => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\r", "&#xD;");
        xml = xml.replace(`>${escaped(oldA)}</a:t>`, `>${escaped(a)}</a:t>`).replace(`>${escaped(oldB)}</a:t>`, `>${escaped(b)}</a:t>`);
        assert.equal(parts.get("deck/z.xml").toString(), xml);
        const before = readGeneratedZip(bytes); for (const [name, payload] of before) if (name !== "deck/z.xml") assert.deepEqual(parts.get(name), payload, name);
        assert.deepEqual(await inspectPptxTiming(saved.bytes), inventory);
        assert.equal(saved.report.assurance, "scoped-preservation"); assert.equal(saved.report.changes.length, 2);
        assert.deepEqual((await compareOpcPackages(bytes, saved.bytes)).differences.map(d => [d.code, d.part]), [["payload-changed", "/deck/z.xml"]]);
        bytes = saved.bytes;
      } finally { session.dispose(); }
    }
  }
});

test("default guards stay unchanged, option combinations validate, and async options/bytes capture preserves explicit authority", async () => {
  const bytes = paragraphBytes(), session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v1" });
  editParagraph(session, "blocked by default"); await blocked(session); session.dispose();
  for (const options of [
    { projection: "structure", editSaveProfile: "text-runs-v1", textTargetSaveProfile: "paragraph-stable-v1" },
    { projection: "content", textTargetSaveProfile: "paragraph-stable-v1" },
    { ...paragraphProfile, textTargetSaveProfile: "all-ranges" },
    { projection: "content", chartEditSaveProfile: "existing-values-v1", textTargetSaveProfile: "paragraph-stable-v1" },
  ]) await assert.rejects(importPptx(bytes, options), TypeError);
  const source = Uint8Array.from(bytes), options = { ...paragraphProfile }, pending = importPptx(source, options);
  source.fill(0); options.textTargetSaveProfile = undefined;
  const captured = await pending; editParagraph(captured, "captured"); await ready(captured); captured.dispose();
});

test("paragraph ranges admit exact single/multiple boundary indices and normalized numeric owner/index spellings", async () => {
  for (const range of ['<p:pRg st="0" end="0"/>', '<p:pRg st="2" end="2"/>', '<p:pRg st=" +0001 " end=" +0002 "/>']) {
    const f = paragraphFixture({ id: " +0002 ", target: paragraphTarget(range, " +0002 ") }), session = await open(f);
    editParagraph(session, "new"); const saved = await ready(session);
    assert.ok(readGeneratedZip(saved.bytes).get("deck/z.xml").toString().includes(range)); session.dispose();
  }
});

test("malformed, opaque, character, mixed and out-of-bounds targets retain located blocking diagnostics", async () => {
  const targets = [
    paragraphTarget('<p:pRg st="0" end="3"/>'), paragraphTarget('<p:pRg st="2" end="1"/>'),
    paragraphTarget('<p:pRg st="-1" end="1"/>'), paragraphTarget('<p:pRg st="0.0" end="1"/>'),
    paragraphTarget('<p:pRg st="0"/>'), paragraphTarget('<p:pRg st="0" end="4294967296"/>'),
    paragraphTarget('<p:charRg st="0" end="1"/>'), paragraphTarget(''),
    paragraphTarget('<p:pRg st="0" end="1"/><p:charRg st="0" end="1"/>'),
    paragraphTarget('<p:pRg st="0" end="1" custom="future"/>'),
    paragraphTarget('<p:pRg st="0" end="1"><p:extLst/></p:pRg>'),
    paragraphTarget('<p:pRg st="0" end="1">opaque</p:pRg>'),
    paragraphTarget('<p:pRg st="0" end="1"/><!--opaque-->'),
    '<p:spTgt spid="2" x="future"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt>',
    '<p:spTgt spid="2"><p:txEl x="future"><p:pRg st="0" end="1"/></p:txEl></p:spTgt>',
    '<p:spTgt spid="2"><p:txEl><p:pRg st="0" end="1"/></p:txEl><p:bg/></p:spTgt>',
    '<p:spTgt spid="bad"><p:txEl><p:pRg st="0" end="1"/></p:txEl></p:spTgt>',
    '<p:spTgt spid="2"><p:txEl><x:pRg xmlns:x="urn:future" st="0" end="1"/></p:txEl></p:spTgt>',
  ];
  for (const target of targets) {
    const session = await open(paragraphFixture({ target }));
    try {
      editParagraph(session, "new"); const result = await blocked(session);
      assert.equal(result.report.diagnostics[0].source.part, "/deck/z.xml"); assert.ok(result.report.diagnostics[0].source.elementIndex > 0);
      editParagraph(session, "First"); const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready");
    } finally { session.dispose(); }
  }
});

test("all references participate: a supported paragraph target cannot hide an additional character or opaque target", async () => {
  for (const other of [paragraphTarget('<p:charRg st="0" end="1"/>'), '<p:spTgt spid="2" unknown="x"/>']) {
    const session = await open(paragraphFixture({ target: paragraphTarget() + other }));
    editParagraph(session, "new"); await blocked(session); session.dispose();
  }
});

test("inactive references use the same proof while inactive duplicate owners and nontransparent paragraph bodies never acquire it", async () => {
  const f = paragraphFixture({ target: '<p:spTgt spid="2"/>' });
  const raw = f.parts["deck/z.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ');
  for (const duplicate of [false, true]) {
    const next = { ...f, parts: { ...f.parts } };
    next.parts["deck/z.xml"] = raw.replace('</p:sld>', `<mc:AlternateContent><mc:Choice Requires="x"><p:timing>${paragraphTarget()}</p:timing>${duplicate ? shape({ id: "02", txBox: true, text: body() }) : ""}</mc:Choice><mc:Fallback/></mc:AlternateContent></p:sld>`);
    const session = await open(next); editParagraph(session, "new");
    if (duplicate) await blocked(session); else await ready(session);
    session.dispose();
  }
  const opaque = paragraphFixture();
  opaque.parts["deck/z.xml"] = opaque.parts["deck/z.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
    .replace('<a:p><a:pPr lvl="1"/><a:endParaRPr/></a:p>', '<mc:AlternateContent><mc:Choice Requires="x"><a:p><a:r><a:t>hidden</a:t></a:r></a:p></mc:Choice><mc:Fallback><a:p><a:endParaRPr/></a:p></mc:Fallback></mc:AlternateContent>');
  const session = await open(opaque); editParagraph(session, "new"); await blocked(session); session.dispose();
});

test("mixed safe/blocked owners stay atomic, unrelated owner saves work after retry, and paragraph/run topology edits still block", async () => {
  const f = paragraphFixture({ target: paragraphTarget() + paragraphTarget('<p:charRg st="0" end="1"/>', "3"), extra: shape({ id: "3", txBox: true, text: body() }) });
  const session = await open(f);
  try {
    editParagraph(session, "safe"); editParagraph(session, "unsafe", 0, 0, 1); const result = await blocked(session);
    assert.equal(result.report.changes.length, 2);
    editParagraph(session, " A & B ", 0, 0, 1); await ready(session);
    const slide = session.document.slides[0], element = structuredClone(slide.elements[0]);
    element.content.reverse(); slide.replaceElement(element.id, element); await blocked(session, "save-edit-unsupported");
    element.content.reverse(); element.content[0].runs.push({ text: "inserted" }); slide.replaceElement(element.id, element); await blocked(session, "save-edit-unsupported");
  } finally { session.dispose(); }
});

test("notes and table paragraph targets retain their separate conservative guards", async () => {
  const notes = notesFixture(notesShape({ text: body() }));
  notes.parts["deck/notes/z.xml"] = notes.parts["deck/notes/z.xml"].replace('</p:notes>', `<p:timing>${paragraphTarget('<p:pRg st="0" end="0"/>')}</p:timing></p:notes>`);
  const n = await open(notes, { editSaveProfile: "text-runs-v2" });
  const content = structuredClone(n.document.slides[0].notes); content[0].runs[0].text = "new"; n.document.slides[0].setNotes(content); await blocked(n); n.dispose();
  const table = tableFixture(); table.parts["deck/z.xml"] = table.parts["deck/z.xml"].replace('</p:sld>', `<p:timing>${paragraphTarget('<p:pRg st="0" end="0"/>')}</p:timing></p:sld>`);
  const t = await open(table, { editSaveProfile: "text-runs-v3" }), slide = t.document.slides[0], element = structuredClone(slide.elements[0]);
  element.rows[0].cells[0].content[0].runs[0].text = "new"; slide.replaceElement(element.id, element); await blocked(t); t.dispose();
});

test("proof records are bounded and cancellation/failure permits independent retry", async () => {
  const bytes = paragraphBytes(), session = await importPptx(bytes, paragraphProfile);
  const bindings = session.report.textRuns, context = await structureContext(bytes);
  const count = bindings.length + 1 + 3;
  assert.equal((await paragraphStableOwners(context, bindings, count, async () => {})).get("/deck/z.xml").get("2"), 3);
  const wrongParagraph = bindings.map(b => b.paragraphIndex === 2 ? { ...b, sourceElementIndex: bindings[0].sourceElementIndex } : b);
  assert.equal((await paragraphStableOwners(context, wrongParagraph, count, async () => {})).get("/deck/z.xml").has("2"), false);
  const missingAtom = bindings.map((b, i) => i === 0 ? { ...b, sourceElementIndex: 999999 } : b);
  assert.equal((await paragraphStableOwners(context, missingAtom, count, async () => {})).get("/deck/z.xml").has("2"), false);
  await assert.rejects(paragraphStableOwners(context, bindings, count - 1, async () => {}), { code: "import-limit-exceeded" });
  let calls = 0; await assert.rejects(paragraphStableOwners(context, bindings, count, async () => { if (++calls === 10) throw new Error("injected cancellation"); }), /injected cancellation/);
  assert.equal((await paragraphStableOwners(context, bindings, count, async () => {})).get("/deck/z.xml").get("2"), 3);
  const abort = new AbortController(); abort.abort(); await assert.rejects(importPptx(bytes, { ...paragraphProfile, signal: abort.signal }), { code: "operation-aborted" });
  editParagraph(session, "retry"); await ready(session); session.dispose();
});

test("Node file imports retain the explicit paragraph profile across asynchronous file capture", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-paragraph-"));
  try {
    const source = path.join(dir, "source.pptx"); await writeFile(source, paragraphBytes());
    const options = { ...paragraphProfile }, pending = importPptxFile(source, options); options.textTargetSaveProfile = undefined;
    const { session } = await pending; editParagraph(session, "file profile"); await ready(session); session.dispose();
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test("nested group owners and identical identities in separate slide parts retain independent paragraph proofs", async () => {
  const f = paragraphFixture();
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"].replace('<p:pRg st="0" end="2"/>', '<p:charRg st="0" end="2"/>');
  const session = await open(f);
  try {
    const context = await structureContext(encodeStructure(f));
    // A one-part proof budget cannot be reset for the second slide.
    await assert.rejects(paragraphStableOwners(context, session.report.textRuns, 8, async () => {}), { code: "import-limit-exceeded" });
    editParagraph(session, "safe slide"); await ready(session);
    const slide = session.document.slides[1], element = structuredClone(slide.elements[0]); element.content[0].runs[0].text = "blocked slide"; slide.replaceElement(element.id, element);
    const result = await blocked(session); assert.equal(result.report.diagnostics[0].source.part, "/deck/a.xml");
  } finally { session.dispose(); }
  const group = paragraphFixture(), xml = group.parts["deck/z.xml"], start = xml.indexOf('<p:sp>'), end = xml.indexOf('</p:sp>', start) + '</p:sp>'.length;
  const owned = xml.slice(start, end);
  group.parts["deck/z.xml"] = xml.replace(owned, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="2540000"/><a:chOff x="0" y="0"/><a:chExt cx="2540000" cy="2540000"/></a:xfrm></p:grpSpPr>${owned}</p:grpSp>`);
  const nested = await open(group);
  try {
    const slide = nested.document.slides[0], element = structuredClone(slide.elements[0]); element.children[0].content[2].runs[0].text = "nested"; slide.replaceElement(element.id, element);
    await ready(nested);
  } finally { nested.dispose(); }
});

test("inactive lexical negative zero cannot evade raw drawing identity collision detection", async () => {
  const f = paragraphFixture({ id: "0", target: paragraphTarget(undefined, "0") });
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x" ')
    .replace('</p:sld>', `<mc:AlternateContent><mc:Choice Requires="x">${shape({ id: "-000", txBox: true, text: body() })}</mc:Choice><mc:Fallback/></mc:AlternateContent></p:sld>`);
  const session = await open(f); editParagraph(session, "new"); await blocked(session); session.dispose();
});


test("number changes combine with proven paragraph targets without weakening range guards", async () => {
  for (const strict of [false, true]) {
    let bytes = encodeStructure(paragraphFixture({ strict })), previous;
    for (const [round, value] of [17, 0, undefined].entries()) {
      const session = await importPptx(bytes, { ...paragraphProfile, firstSlideNumberSaveProfile: "attribute-v2", textFontProfile: "script-families-v3" });
      try {
        assert.equal(session.document.firstSlideNumber, previous); session.document.setFirstSlideNumber(value); editParagraph(session, `combined ${round}`);
        const saved = await ready(session), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
        assert.equal(saved.report.changes.length, 2);
        const oldSlide = before.get("deck/z.xml").toString(), newSlide = after.get("deck/z.xml").toString();
        assert.equal(newSlide, oldSlide.replace(round ? `combined ${round - 1}` : "First", `combined ${round}`));
        for (const [name, payload] of before) if (!["deck/z.xml", "deck/main.data"].includes(name)) assert.deepEqual(after.get(name), payload);
        assert.deepEqual(await inspectPptxTiming(saved.bytes), await inspectPptxTiming(bytes));
        bytes = saved.bytes; previous = value;
      } finally { session.dispose(); }
    }
    const bad = await open(paragraphFixture({ strict, target: paragraphTarget('<p:charRg st="0" end="1"/>') }), { firstSlideNumberSaveProfile: "attribute-v2" });
    try { bad.document.setFirstSlideNumber(17); editParagraph(bad, "blocked"); await blocked(bad); } finally { bad.dispose(); }
  }
});


test("combined options retain legacy restrictions and independent profile prerequisites", async () => {
  const bytes = paragraphBytes();
  for (const options of [
    { ...paragraphProfile, firstSlideNumberSaveProfile: "existing-v1" },
    { projection: "content", firstSlideNumberSaveProfile: "existing-v1", textFontProfile: "script-families-v3" },
    { projection: "content", firstSlideNumberSaveProfile: "attribute-v2", textTargetSaveProfile: "paragraph-stable-v1" },
    { ...paragraphProfile, firstSlideNumberSaveProfile: "attribute-v2", textFontProfile: "unknown" },
    { ...paragraphProfile, firstSlideNumberSaveProfile: "attribute-v2", notesTextProfile: "unknown" },
    { ...paragraphProfile, firstSlideNumberSaveProfile: "attribute-v2", pictureEditSaveProfile: "replace-v1" },
  ]) await assert.rejects(importPptx(bytes, options), TypeError);
});
