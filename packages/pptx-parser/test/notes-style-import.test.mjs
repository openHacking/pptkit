import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { normalizePresentation, resolveTextFontSegments } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { notesStyleFixture, notesFonts, addSecondNotes } from "./notes-style-fixture.mjs";
import { notesShape, notesTree } from "./notes-fixture.mjs";
import { body } from "./content-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const options = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" };
const open = f => importPptx(encodeStructure(f), options);
const note = (s, index = 0) => s.document.slides[index].notes[0].runs[0];
const props = (s, part = "/deck/notes/z.xml") => s.report.textStyleInheritance.properties.filter(p => p.target.part === part);
const expectedMap = { Hant: "Notes Traditional", Arab: "Notes Arabic", Hebr: "Notes Hebrew" };

test("notes profile independently resolves notesStyle/type-matched master list/fonts/colors in both dialects", async () => {
  for (const strict of [false, true]) {
    const f = notesStyleFixture({ strict }), s = await open(f), old = await importPptx(encodeStructure(f), { ...options, notesTextProfile: undefined });
    try {
      const r = note(s); assert.equal(s.report.notesTextProfile, "master-style-v1"); assert.equal(Object.hasOwn(old.report, "notesTextProfile"), false);
      assert.equal(r.text, "A中عא"); assert.equal(r.style.fontSize, 22); assert.equal(r.style.italic, true); assert.equal(r.style.language, "zh-TW"); assert.equal(r.style.color, "204060");
      assert.deepEqual([r.style.fontFamily, r.style.eastAsianFontFamily, r.style.complexScriptFontFamily], ["Notes Latin", "Notes East", "Notes Complex"]); assert.deepEqual(r.style.scriptFontFamilies, expectedMap);
      assert.equal(note(old).style.fontFamily, undefined); assert.equal(note(old).style.scriptFontFamilies, undefined);
      assert.ok(props(s).some(p => p.property === "run.fontFamily" && p.level === "notes-master-text-style" && p.source.part === "/deck/notes-master.xml" && p.resolvedFrom.some(d => d.part === "/deck/notes-theme.xml")));
      const size = props(s).find(p => p.property === "run.fontSize"); assert.equal(size.level, "notes-master-placeholder"); assert.equal(size.source.sourceObjectId, "5"); assert.ok(size.resolvedFrom.some(d => d.part === "/deck/notes/z.xml" && d.path.endsWith("/@type")));
      assert.ok(props(s).every(p => !p.source.part.includes("layout"))); assert.ok(Object.isFrozen(s.report.textStyleInheritance.properties));
      assert.equal(JSON.stringify(normalizePresentation(s.document)).includes("MASTER PROMPT"), false);
    } finally { s.dispose(); old.dispose(); }
  }
});

test("notes local list/paragraph/run override selected master fields independently, preserving unsupported blockers", async () => {
  const f = notesStyleFixture({ local: '<a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr>', paragraph: '<a:defRPr i="0"><a:cs typeface="Paragraph Complex"/></a:defRPr>', run: '<a:ea typeface=""/><a:latin typeface="Local Latin"/>' });
  const s = await open(f);
  try {
    const r = note(s); assert.equal(r.style.fontSize, 24); assert.equal(r.style.italic, false); assert.equal(r.style.fontFamily, "Local Latin"); assert.equal(r.style.complexScriptFontFamily, "Paragraph Complex"); assert.equal(r.style.eastAsianFontFamily, undefined); assert.equal(r.style.scriptFontFamilies, undefined);
    assert.equal(props(s).find(p => p.property === "run.eastAsianFontFamily").status, "unsupported"); assert.equal(props(s).find(p => p.property === "run.fontSize").level, "notes-direct");
  } finally { s.dispose(); }
});

test("notes list levels and empty insertion styles use master defaults without synthesizing prompt content", async () => {
  const f = notesStyleFixture({ global: `<a:defPPr><a:defRPr lang="ja">${notesFonts()}</a:defRPr></a:defPPr><a:lvl2pPr><a:defRPr sz="3300" b="1"/></a:lvl2pPr>`, placeholder: "" });
  f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace('<a:p><a:r>', '<a:p><a:pPr lvl="1"/><a:r>');
  const level = await open(f); try { assert.equal(note(level).style.fontSize, 33); assert.equal(note(level).style.bold, true); } finally { level.dispose(); }
  f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace('<a:r><a:t>A中عא</a:t></a:r>', '<a:endParaRPr sz="2700"/>');
  const empty = await open(f); try { assert.equal(note(empty).text, ""); assert.equal(note(empty).style.fontSize, 27); assert.equal(note(empty).style.fontFamily, "Notes Latin"); assert.equal(props(empty).find(p => p.property === "run.fontSize").level, "end-paragraph"); } finally { empty.dispose(); }
});

test("missing/ambiguous/grouped master body matches keep global notesStyle and report the boundary", async () => {
  for (const variant of ["missing", "ambiguous", "grouped"]) {
    const f = notesStyleFixture();
    if (variant === "missing") f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('type="body"', 'type="ftr"');
    if (variant === "ambiguous") f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('</p:spTree>', `${notesShape({ id: "6", text: body() }).replace('idx="1"', 'idx="100"')}</p:spTree>`);
    if (variant === "grouped") {
      const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${notesShape({ id: "5", text: body() })}</p:grpSp>`;
      f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace(/<p:cSld>[\s\S]*?<\/p:cSld>/, notesTree(group));
    }
    const s = await open(f); try { assert.equal(note(s).style.fontSize, 16, variant); assert.equal(note(s).style.fontFamily, "Notes Latin"); assert.ok(s.report.findings.some(f => f.feature === "notes.inherited-text.unmapped-property" && /no body|Ambiguous|Grouped/.test(f.reason))); } finally { s.dispose(); }
  }
});

test("missing notes master/theme never substitutes listed or slide master dependencies", async () => {
  for (const variant of ["master", "theme"]) {
    const f = notesStyleFixture({ run: notesFonts() });
    if (variant === "master") f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml")]);
    else f.parts["deck/_rels/notes-master.xml.rels"] = relationships([]);
    const s = await open(f);
    try {
      assert.equal(note(s).style.fontFamily, undefined); assert.equal(note(s).style.scriptFontFamilies, undefined);
      const font = props(s).find(p => p.property === "run.fontFamily"); assert.equal(font.status, "unsupported"); assert.equal((font.resolvedFrom ?? []).some(d => d.part === "/deck/theme.xml"), false);
    } finally { s.dispose(); }
  }
});

test("notes-master identities/placeholder indices are checked without interpreting extension islands", async () => {
  for (const [before, after] of [['id="5"', 'id="1"'], ['idx="99"', 'idx="4294967295"']]) {
    const f = notesStyleFixture(); f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace(before, after); await assert.rejects(open(f), e => e.code === "import-invalid-package");
  }
  const f = notesStyleFixture(); f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('</p:spTree>', `<p:extLst><p:ext uri="fixture">${notesShape({ id: "5", text: body() })}</p:ext></p:extLst></p:spTree>`);
  const s = await open(f); try { assert.equal(note(s).style.fontSize, 22); } finally { s.dispose(); }
  const duplicate = notesStyleFixture({ extraMaster: notesShape({ id: "6", role: "ftr", text: body() }).replace('idx="1"', 'idx="99"') });
  await assert.rejects(open(duplicate), e => e.code === "import-invalid-package");
});

test("notes-master styles use effective MCE body choices and report unsupported list attributes", async () => {
  for (const strict of [false, true]) {
    const f = notesStyleFixture({ strict, placeholder: '<a:lvl1pPr unsupported="retain"><a:defRPr sz="2200"/></a:lvl1pPr>' });
    f.parts["deck/notes-master.xml"] = f.parts["deck/notes-master.xml"].replace('<p:notesMaster ', '<p:notesMaster xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:notes:future" mc:Ignorable="future" ')
      .replace(/(<p:sp>.*?<\/p:sp>)/, `<mc:AlternateContent><mc:Choice Requires="future">${notesShape({ id: "8", text: body().replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr><a:defRPr sz="8800"/></a:lvl1pPr></a:lstStyle>') }).replace('idx="1"', 'idx="100"')}</mc:Choice><mc:Fallback>$1</mc:Fallback></mc:AlternateContent>`);
    const s = await open(f);
    try { assert.equal(note(s).style.fontSize, 22); assert.ok(s.report.findings.some(f => f.feature === "notes.inherited-text.unmapped-property" && f.source.path?.endsWith("/@unsupported"))); }
    finally { s.dispose(); }
  }
});

test("two notes parts sharing or changing actual masters keep independent theme scopes and colors", async () => {
  for (const different of [false, true]) {
    const s = await open(addSecondNotes(notesStyleFixture(), different));
    try { assert.equal(note(s).style.fontFamily, "Notes Latin"); assert.equal(note(s, 1).style.fontFamily, different ? "Second Notes Latin" : "Notes Latin"); assert.equal(note(s, 1).style.scriptFontFamilies.Hant, different ? "Second Notes Traditional" : "Notes Traditional"); }
    finally { s.dispose(); }
  }
  const f = notesStyleFixture(); f.parts["deck/notes/z.xml"] = f.parts["deck/notes/z.xml"].replace('</p:notes>', '<p:clrMapOvr><a:overrideClrMapping tx1="accent1"/></p:clrMapOvr></p:notes>');
  const s = await open(f); try { assert.equal(note(s).style.color, "112233"); assert.ok(props(s).find(p => p.property === "run.color").resolvedFrom.some(d => d.part === "/deck/notes/z.xml" && d.path.endsWith("/@tx1"))); } finally { s.dispose(); }
});

test("near notes theme override replaces fonts only and unsupported audit state does not pollute another scope", async () => {
  const f = addSecondNotes(notesStyleFixture());
  f.parts["deck/notes/_rels/z.xml.rels"] = relationships([f.rel("back", "slide", "../z.xml"), f.rel("master", "notesMaster", "../notes-master.xml"), f.rel("override", "themeOverride", "../notes-override.xml")]);
  f.parts["deck/notes-override.xml"] = `<a:themeOverride xmlns:a="${f.A}"><a:fontScheme name="Near"><a:minorFont><a:latin typeface="Near Notes"/><a:ea typeface=""/><a:cs typeface="Near Complex"/></a:minorFont></a:fontScheme></a:themeOverride>`; f.types["deck/notes-override.xml"] = "themeOverride";
  const s = await open(f);
  try {
    assert.equal(note(s).style.fontFamily, "Near Notes"); assert.equal(note(s).style.eastAsianFontFamily, undefined); assert.equal(note(s).style.scriptFontFamilies, undefined); assert.equal(note(s).style.color, "204060");
    assert.equal(note(s, 1).style.fontFamily, "Notes Latin"); assert.deepEqual(note(s, 1).style.scriptFontFamilies, expectedMap);
    const fonts = props(s, "/deck/notes/a.xml").filter(p => p.property.endsWith("FontFamily") || p.property === "run.fontFamily"); assert.ok(fonts.every(p => p.status === "mapped"));
  } finally { s.dispose(); }
  const colors = addSecondNotes(notesStyleFixture());
  colors.parts["deck/notes/z.xml"] = colors.parts["deck/notes/z.xml"].replace('</p:notes>', '<p:clrMapOvr><a:overrideClrMapping bg1="lt1"/></p:clrMapOvr></p:notes>');
  const legacyFonts = await importPptx(encodeStructure(colors), { ...options, textFontProfile: undefined });
  try {
    assert.equal(note(legacyFonts).style.color, undefined); assert.equal(note(legacyFonts, 1).style.color, "204060");
    assert.equal(legacyFonts.report.findings.filter(f => f.feature === "notes.inherited-text.unmapped-property" && /Color mapping override omits/.test(f.reason)).length, 1, "notes reader isolation must also hold without an explicit font profile");
  } finally { legacyFonts.dispose(); }
});

test("notes inherited scalar saves retain original identities/rules over three rounds and block style mutations", async () => {
  let bytes = encodeStructure(addSecondNotes(notesStyleFixture()));
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, options);
    try {
      for (const slide of s.document.slides) {
        const original = structuredClone(slide.notes), notes = structuredClone(original); notes[0].runs[0].text = `Round${round} ${slide.id} 中عא`;
        for (const field of ["fontFamily", "language", "scriptFontFamilies"]) { const bad = structuredClone(notes); bad[0].runs[0].style[field] = field === "scriptFontFamilies" ? { Arab: "Changed" } : "Changed"; slide.setNotes(bad); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); }
        slide.setNotes(notes);
      }
      const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
      for (const [name, payload] of before) if (!/^deck\/notes\/[az]\.xml$/.test(name)) assert.deepEqual(after.get(name), payload); else assert.equal(after.get(name).toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"), payload.toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"));
      bytes = saved.bytes;
    } finally { s.dispose(); }
    const next = await importPptx(bytes, options); try { assert.deepEqual(note(next).style.scriptFontFamilies, expectedMap); assert.equal(note(next).style.fontSize, 22); assert.match(note(next).text, new RegExp(`^Round${round}`)); } finally { next.dispose(); }
  }
});

test("notes profile works with legacy/v1 font projections and has independent public admission/save permissions", async () => {
  for (const font of [undefined, "script-families-v1", "script-families-v2"]) {
    const s = await importPptx(encodeStructure(notesStyleFixture()), { ...options, textFontProfile: font, editSaveProfile: undefined });
    try {
      assert.equal(note(s).style.fontSize, 22); assert.equal(Boolean(note(s).style.scriptFontFamilies), font === "script-families-v2");
      const notes = structuredClone(s.document.slides[0].notes); notes[0].runs[0].text = "No permission"; s.document.slides[0].setNotes(notes); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { s.dispose(); }
  }
  for (const config of [{ projection: "structure", notesTextProfile: "master-style-v1" }, { projection: "content", notesTextProfile: "unknown" }, { projection: "content", notesTextProfile: null }]) await assert.rejects(importPptx(encodeStructure(notesStyleFixture()), config), TypeError);
});

test("notes option is captured for bytes/file and abort is retryable", async t => {
  const bytes = encodeStructure(notesStyleFixture()), data = Uint8Array.from(bytes), config = { ...options }, pending = importPptx(data, config); data.fill(0); config.notesTextProfile = undefined;
  const s = await pending; try { assert.equal(s.report.notesTextProfile, "master-style-v1"); assert.equal(note(s).style.fontFamily, "Notes Latin"); } finally { s.dispose(); }
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-notes-style-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "source.pptx"); await writeFile(file, bytes);
  const opts = { ...options }, fromFile = importPptxFile(file, opts); opts.notesTextProfile = "invalid"; const imported = await fromFile; try { assert.equal(note(imported.session).style.fontSize, 22); } finally { imported.session.dispose(); }
  const c = new AbortController(), op = importPptx(bytes, { ...options, signal: c.signal }); queueMicrotask(() => c.abort()); await assert.rejects(op, { code: "operation-aborted" }); const retry = await open(notesStyleFixture()); retry.dispose();
});

test("notes and drawings share exact text property/dependency budgets across master projections", async () => {
  const bytes = encodeStructure(addSecondNotes(notesStyleFixture())), baseline = await importPptx(bytes, options), expected = normalizePresentation(baseline.document); baseline.dispose();
  for (const key of ["maxTextStyleProperties", "maxTextStyleDependencies", "maxNotesStyleWork"]) {
    let lo = 1, hi = 1_000_000;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { const s = await importPptx(bytes, { ...options, sessionLimits: { [key]: mid } }); s.dispose(); hi = mid; } catch (e) { assert.equal(e.code, "import-limit-exceeded"); lo = mid + 1; } }
    await assert.rejects(importPptx(bytes, { ...options, sessionLimits: { [key]: lo - 1 } }), { code: "import-limit-exceeded" }); const exact = await importPptx(bytes, { ...options, sessionLimits: { [key]: lo } }); try { assert.deepEqual(normalizePresentation(exact.document), expected); } finally { exact.dispose(); }
  }
  const direct = await importPptx(bytes, { ...options, notesTextProfile: undefined, sessionLimits: { maxNotesStyleWork: 1 } }); direct.dispose();
});
