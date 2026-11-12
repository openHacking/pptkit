import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createPresentation, normalizePresentation } from "../../core/dist/index.js";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
import { importPptx } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { scriptFontImportFixture, sourceFonts, secondFontMaster } from "./script-font-import-fixture.mjs";
import { textAppearanceFixture, fontCollection, textReference } from "./table-text-appearance-fixture.mjs";
import { addTextThemeOverride } from "./text-style-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { zip } from "./zip-fixture.mjs";
import { createScriptFontFixture } from "../../../tests/fixtures/script-fonts.mjs";

const profile = { projection: "content", textFontProfile: "script-families-v1", editSaveProfile: "text-runs-v3" };
const open = f => importPptx(encodeStructure(f), profile);
const run = (s, page = 0) => s.document.slides[page].elements[0].text.content[0].runs[0];
const fonts = r => [r.style.fontFamily, r.style.eastAsianFontFamily, r.style.complexScriptFontFamily];
const props = (s, part = "/deck/z.xml") => s.report.textStyleInheritance.properties.filter(p => p.target.part === part && /FontFamily$|fontFamily$/.test(p.property));

test("explicit source font profile maps independent direct slots in both dialects without changing default projection", async () => {
  for (const strict of [false, true]) {
    const f = scriptFontImportFixture({ strict }), bytes = encodeStructure(f), old = await importPptx(bytes, { projection: "content" }), s = await importPptx(bytes, profile);
    try {
      assert.equal(Object.hasOwn(old.report, "textFontProfile"), false); assert.equal(props(old)[0].status, "unsupported");
      assert.equal(s.report.textFontProfile, "script-families-v1"); assert.deepEqual(fonts(run(s)), ["Latin Source", "Eastern Source", "Complex Source"]);
      assert.deepEqual(props(s).map(p => p.property).sort(), ["run.complexScriptFontFamily", "run.eastAsianFontFamily", "run.fontFamily"]);
      assert.ok(props(s).every(p => p.status === "mapped" && p.source.path.endsWith("/@typeface")));
      assert.ok(Object.isFrozen(s.report.textStyleInheritance.properties));
    } finally { old.dispose(); s.dispose(); }
  }
  for (const options of [{ projection: "structure", textFontProfile: "script-families-v1" }, { projection: "content", textFontProfile: "unknown" }, { projection: "content", textFontProfile: null }]) await assert.rejects(importPptx(encodeStructure(scriptFontImportFixture()), options), TypeError);
});

test("Latin-only local override retains independently inherited script families with selected source levels", async () => {
  for (const strict of [false, true]) {
    const f = scriptFontImportFixture({ strict, inherited: true, list: '<a:lvl1pPr><a:defRPr><a:ea typeface="Slide East"/></a:defRPr></a:lvl1pPr>', paragraph: '<a:defRPr><a:cs typeface="Paragraph Complex"/></a:defRPr>', run: '<a:latin typeface="Run Latin"/>' });
    const s = await open(f);
    try {
      assert.deepEqual(fonts(run(s)), ["Run Latin", "Slide East", "Paragraph Complex"]);
      const byName = Object.fromEntries(props(s).map(p => [p.property, p]));
      assert.equal(byName["run.fontFamily"].level, "run"); assert.equal(byName["run.eastAsianFontFamily"].level, "slide-list-style"); assert.equal(byName["run.complexScriptFontFamily"].level, "paragraph");
    } finally { s.dispose(); }
    const inherited = await open(scriptFontImportFixture({ strict, inherited: true, run: '<a:latin typeface="Local Latin"/>' }));
    try { assert.deepEqual(fonts(run(inherited)), ["Local Latin", "Master East", "Master Complex"]); assert.equal(props(inherited).find(p => p.property === "run.eastAsianFontFamily").source.part, "/deck/master.xml"); }
    finally { inherited.dispose(); }
  }
  const f = scriptFontImportFixture({ inherited: true, run: "" });
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace(/<a:p>[\s\S]*?<\/a:p>/, '<a:p><a:endParaRPr><a:ea typeface="Insertion East"/></a:endParaRPr></a:p>');
  const insertion = await open(f);
  try {
    assert.equal(run(insertion).text, ""); assert.deepEqual(fonts(run(insertion)), ["Layout Latin", "Insertion East", "Master Complex"]);
    assert.equal(props(insertion).find(p => p.property === "run.eastAsianFontFamily").level, "end-paragraph");
  } finally { insertion.dispose(); }
});

test("unsupported higher script declaration blocks only that field and never reveals an older script value", async () => {
  for (const declaration of ['<a:ea typeface=""/>', '<a:ea/>', '<a:ea typeface="+unknown"/>']) {
    const s = await open(scriptFontImportFixture({ inherited: true, run: '<a:latin typeface="Local Latin"/>' + declaration }));
    try {
      assert.equal(run(s).style.fontFamily, "Local Latin"); assert.equal(run(s).style.eastAsianFontFamily, undefined); assert.equal(run(s).style.complexScriptFontFamily, "Master Complex");
      const value = props(s).find(p => p.property === "run.eastAsianFontFamily"); assert.equal(value.status, "unsupported"); assert.equal(value.level, "run"); assert.notEqual(run(s).style.eastAsianFontFamily, "Master East");
    } finally { s.dispose(); }
  }
});

test("explicit theme slots select the target component independently and partial nearer components block missing slots", async () => {
  for (const strict of [false, true]) {
    const f = scriptFontImportFixture({ strict, run: sourceFonts("+mn-lt", "+mn-ea", "+mn-cs") });
    const s = await open(f);
    try { assert.deepEqual(fonts(run(s)), ["Body Source", "Body East", "Body Complex"]); assert.ok(props(s).every(p => p.resolvedFrom.some(d => d.part === "/deck/theme.xml"))); } finally { s.dispose(); }
    addTextThemeOverride(f, "z", "partial-font", '<a:fontScheme name="Partial"><a:minorFont><a:latin typeface="Near Latin"/><a:ea typeface="Near East"/></a:minorFont></a:fontScheme>');
    const partial = await open(f);
    try { assert.deepEqual(fonts(run(partial)), ["Near Latin", "Near East", undefined]); const blocked = props(partial).find(p => p.property === "run.complexScriptFontFamily"); assert.equal(blocked.status, "unsupported"); assert.ok(blocked.resolvedFrom.some(d => d.part === "/deck/partial-font.xml" && d.path.includes("fontScheme"))); assert.deepEqual(fonts(run(partial, 1)), ["Body Source", "Body East", "Body Complex"]); } finally { partial.dispose(); }
  }
  const f = scriptFontImportFixture({ run: sourceFonts("+mn-lt", "+mn-ea", "+mn-cs") });
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace("</a:minorFont>", '<a:font script="Hans" typeface="Supplemental East"/></a:minorFont>');
  const s = await open(f);
  try {
    assert.deepEqual(fonts(run(s)), ["Body Source", "Body East", "Body Complex"], "only the declared named-slot projection is evaluated");
    const findings = s.report.findings.filter(p => /Theme supplemental/.test(p.reason)); assert.equal(findings.length, 1); assert.equal(findings[0].source.part, "/deck/theme.xml"); assert.ok(findings[0].source.path.endsWith("a:font[1]"));
  } finally { s.dispose(); }
});

test("table collections, reference slots and local run declarations have independent per-field source records", async () => {
  for (const strict of [false, true]) for (const reference of [false, true]) {
    const f = textAppearanceFixture({ strict, text: reference ? textReference() : fontCollection("Table Latin", "Table East", "Table Complex") });
    if (reference) f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:ea typeface="Arial"/>', '<a:ea typeface="Table East"/>').replaceAll('<a:cs typeface="Arial"/>', '<a:cs typeface="Table Complex"/>');
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll("<a:r>", '<a:r><a:rPr><a:ea typeface="Local East"/></a:rPr>');
    const s = await open(f);
    try {
      const r = s.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0]; assert.deepEqual(fonts(r), [reference ? "Arial" : "Table Latin", "Local East", "Table Complex"]);
      const values = s.report.tableStyleInheritance.properties.filter(p => p.rowIndex === 0 && p.columnIndex === 0 && p.property.endsWith("FontFamily"));
      assert.equal(values.length, 2); assert.ok(values.some(p => p.property === "run.eastAsianFontFamily" && p.level === "run"));
    } finally { s.dispose(); }
  }
});

test("partial/unsupported table collections block missing fields and supplemental selection cannot be silently discarded", async () => {
  for (const text of ['<a:font><a:latin typeface="Latin"/></a:font>', fontCollection("Latin", "Eastern", "Complex", '<a:font script="Hans" typeface="Other"/>')]) {
    const s = await open(textAppearanceFixture({ text }));
    try {
      const p = s.report.tableStyleInheritance.properties; assert.ok(p.some(p => p.property === "run.eastAsianFontFamily" && p.status === "unsupported"));
      assert.ok(p.some(p => p.property === "run.complexScriptFontFamily" && p.status === "unsupported"));
    } finally { s.dispose(); }
  }
  const s = await open(textAppearanceFixture({ text: fontCollection("Same", "Same", "Same", '<a:font script="Hans" typeface="Same"/>') }));
  try { assert.deepEqual(fonts(s.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0]), ["Same", "Same", "Same"]); } finally { s.dispose(); }
});

test("new profile rejects duplicate interpreted font slots and uses effective MCE choices", async () => {
  for (const strict of [false, true]) {
    await assert.rejects(open(scriptFontImportFixture({ strict, run: sourceFonts() + '<a:ea typeface="Duplicate"/>' })), e => e.code === "import-invalid-package" && e.cause?.code === "presentation-invalid");
    const f = scriptFontImportFixture({ strict, run: '<a:latin typeface="Latin"/><mc:AlternateContent><mc:Choice Requires="future"><future:font/></mc:Choice><mc:Fallback><a:ea typeface="MCE East"/></mc:Fallback></mc:AlternateContent><a:cs typeface="Complex"/>' });
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:font:future" mc:Ignorable="future" ');
    f.parts["deck/a.xml"] = f.parts["deck/a.xml"].replace('<p:sld ', '<p:sld xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:font:future" mc:Ignorable="future" ');
    const s = await open(f); try { assert.deepEqual(fonts(run(s)), ["Latin", "MCE East", "Complex"]); } finally { s.dispose(); }
  }
});

test("notes direct script slots have bounded provenance while notes theme references never borrow slide themes", async () => {
  const bytes = (await generatePptx(createScriptFontFixture(createPresentation))).bytes, s = await importPptx(bytes, profile);
  try {
    assert.deepEqual(fonts(s.document.slides[0].notes[0].runs[0]), ["Latin Face", "Eastern Face", "Complex Face"]); assert.equal(props(s, "/ppt/notesSlides/notesSlide1.xml").length, 3);
    const count = s.report.textStyleInheritance.properties.length;
    await assert.rejects(importPptx(bytes, { ...profile, sessionLimits: { maxTextStyleProperties: count - 1 } }), { code: "import-limit-exceeded" });
    const exact = await importPptx(bytes, { ...profile, sessionLimits: { maxTextStyleProperties: count } }); exact.dispose();
  } finally { s.dispose(); }
  const parts = readGeneratedZip(bytes); parts.set("ppt/notesSlides/notesSlide1.xml", Buffer.from(parts.get("ppt/notesSlides/notesSlide1.xml").toString().replace('typeface="Eastern Face"', 'typeface="+mn-ea"')));
  const changed = await importPptx(zip([...parts].map(([name, body]) => ({ name, body }))).bytes, profile);
  try { assert.equal(changed.document.slides[0].notes[0].runs[0].style.eastAsianFontFamily, undefined); assert.equal(props(changed, "/ppt/notesSlides/notesSlide1.xml").find(p => p.property === "run.eastAsianFontFamily").status, "unsupported"); } finally { changed.dispose(); }
});

test("three mixed-owner scalar saves reproject the captured font profile and retain every unedited source payload", async () => {
  let bytes = (await generatePptx(createScriptFontFixture(createPresentation))).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, profile);
    try {
      const slide = s.document.slides[0];
      for (const e of [...slide.elements]) {
        const next = structuredClone(e);
        if (next.type === "text") next.content[0].runs[0].text = `Text ${round} 中文العربية`;
        else if (next.type === "shape") next.text.content[0].runs[0].text = `Shape ${round} 中文العربية`;
        else if (next.type === "table") next.rows[0].cells[0].content[0].runs[0].text = `Cell ${round} 中文العربية`;
        slide.replaceElement(e.id, next);
      }
      const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Note ${round} 中文العربية`; slide.setNotes(notes);
      const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
      for (const [name, payload] of before) if (!name.includes("slides/slide1.xml") && !name.includes("notesSlides/notesSlide1.xml")) assert.deepEqual(after.get(name), payload);
      for (const part of ["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"]) {
        const allSlots = xml => [...xml.toString().matchAll(/<a:(?:latin|ea|cs)\b[^>]*\/>/g)].map(m => m[0]); assert.deepEqual(allSlots(after.get(part)), allSlots(before.get(part)));
      }
      bytes = saved.bytes;
    } finally { s.dispose(); }
  }
});

test("font-style mutations mixed with text are blocked atomically and restoration permits retry", async () => {
  const bytes = (await generatePptx(createScriptFontFixture(createPresentation))).bytes, s = await importPptx(bytes, profile);
  try {
    const slide = s.document.slides[0], old = slide.elements[0], text = structuredClone(old); text.content[0].runs[0].text = "Allowed text";
    for (const field of ["fontFamily", "eastAsianFontFamily", "complexScriptFontFamily"]) {
      const bad = structuredClone(text); bad.content[0].runs[0].style[field] = "Changed Font"; slide.replaceElement(old.id, bad); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
    }
    slide.replaceElement(old.id, text); assert.equal((await s.prepareSave({ mode: "edit" })).status, "ready");
  } finally { s.dispose(); }
  const noPermission = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v1" });
  try { const slide = noPermission.document.slides[0], next = structuredClone(slide.elements[0]); next.content[0].runs[0].text = "Not admitted"; slide.replaceElement(next.id, next); assert.equal((await noPermission.prepareSave({ mode: "edit" })).status, "blocked"); } finally { noPermission.dispose(); }
});

test("shared inherited declarations resolve under two masters without unsupported audit contamination", async () => {
  const f = secondFontMaster(scriptFontImportFixture({ inherited: true, run: "" }));
  f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(sourceFonts("Master Latin", "Master East", "Master Complex"), sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"));
  f.parts["deck/master2.xml"] = f.parts["deck/master2.xml"].replace(sourceFonts("Master Latin", "Master East", "Master Complex"), sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"));
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('typeface="Body East"', 'typeface=""');
  f.parts["deck/theme2.xml"] = f.parts["deck/theme2.xml"].replaceAll("Body East", "Second East");
  const s = await open(f);
  try {
    assert.equal(run(s).style.eastAsianFontFamily, undefined); assert.equal(run(s, 1).style.eastAsianFontFamily, "Second East");
    assert.equal(props(s).find(p => p.property === "run.eastAsianFontFamily").status, "unsupported"); assert.equal(props(s, "/deck/a.xml").find(p => p.property === "run.eastAsianFontFamily").status, "mapped");
  } finally { s.dispose(); }
});

test("shared source declarations under per-slide overrides keep diagnostic state scoped", async () => {
  const f = scriptFontImportFixture({ inherited: true, run: "" });
  f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(sourceFonts("Master Latin", "Master East", "Master Complex"), sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"));
  addTextThemeOverride(f, "z", "missing-east", '<a:fontScheme name="Partial"><a:minorFont><a:latin typeface="Near Latin"/><a:cs typeface="Near Complex"/></a:minorFont></a:fontScheme>');
  const s = await open(f);
  try {
    assert.equal(run(s).style.eastAsianFontFamily, undefined); assert.equal(run(s, 1).style.eastAsianFontFamily, "Body East");
    const audits = s.report.findings.filter(f => f.feature === "drawing.inherited-text.unmapped-property" && /explicit typeface/.test(f.reason)); assert.equal(audits.length, 1, "the valid scope must not inherit the unsupported scope's audit notes");
  } finally { s.dispose(); }
});

test("script provenance/dependencies and table work use exact cumulative budgets with retry", async () => {
  const f = scriptFontImportFixture({ run: sourceFonts("+mn-lt", "+mn-ea", "+mn-cs") }), bytes = encodeStructure(f), baseline = await importPptx(bytes, profile), expected = normalizePresentation(baseline.document); baseline.dispose();
  for (const key of ["maxTextStyleProperties", "maxTextStyleDependencies"]) {
    let lo = 1, hi = 1_000_000;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { const s = await importPptx(bytes, { ...profile, sessionLimits: { [key]: mid } }); s.dispose(); hi = mid; } catch (e) { assert.equal(e.code, "import-limit-exceeded"); lo = mid + 1; } }
    await assert.rejects(importPptx(bytes, { ...profile, sessionLimits: { [key]: lo - 1 } }), { code: "import-limit-exceeded" });
    const retry = await importPptx(bytes, { ...profile, sessionLimits: { [key]: lo } }); try { assert.deepEqual(normalizePresentation(retry.document), expected); } finally { retry.dispose(); }
  }
  const table = encodeStructure(textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") }));
  for (const key of ["maxTableStyleProperties", "maxTableStyleDependencies", "maxTableStyleWork"]) {
    let lo = 1, hi = 2_000_000;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { const s = await importPptx(table, { ...profile, sessionLimits: { [key]: mid } }); s.dispose(); hi = mid; } catch (e) { assert.equal(e.code, "import-limit-exceeded"); lo = mid + 1; } }
    await assert.rejects(importPptx(table, { ...profile, sessionLimits: { [key]: lo - 1 } }), { code: "import-limit-exceeded" });
    const retry = await importPptx(table, { ...profile, sessionLimits: { [key]: lo } }); retry.dispose();
  }
});

test("byte/file option capture and cancellation retain the selected profile across asynchronous work", async t => {
  const bytes = encodeStructure(scriptFontImportFixture()), source = Uint8Array.from(bytes), options = { ...profile }, pending = importPptx(source, options); source.fill(0); options.textFontProfile = undefined;
  const captured = await pending; try { assert.equal(captured.report.textFontProfile, "script-families-v1"); assert.equal(run(captured).style.eastAsianFontFamily, "Eastern Source"); } finally { captured.dispose(); }
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-script-font-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "source.pptx"); await writeFile(file, bytes);
  const config = { ...profile }, filePending = importPptxFile(file, config); config.textFontProfile = "invalid"; const fromFile = await filePending;
  try { assert.equal(fromFile.session.report.textFontProfile, "script-families-v1"); } finally { fromFile.session.dispose(); }
  const c = new AbortController(), operation = importPptx(bytes, { ...profile, signal: c.signal }); queueMicrotask(() => c.abort()); await assert.rejects(operation, { code: "operation-aborted" });
  const retry = await importPptx(bytes, profile); retry.dispose();
});
