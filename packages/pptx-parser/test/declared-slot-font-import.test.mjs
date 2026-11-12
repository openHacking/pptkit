import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { normalizePresentation } from "@pptkit/core";
import { importPptx } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { supplementalFontFixture, supplementalFonts } from "./supplemental-font-fixture.mjs";
import { sourceFonts, secondFontMaster } from "./script-font-import-fixture.mjs";
import { textAppearanceFixture, textReference, fontCollection } from "./table-text-appearance-fixture.mjs";
import { notesStyleFixture, addSecondNotes } from "./notes-style-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const latin = '<a:latin typeface="+mn-lt"/>';
const open = (fixture, extra = {}) => importPptx(encodeStructure(fixture), { ...options, ...extra });
const run = (session, index = 0) => session.document.slides[index].elements[0].text.content[0].runs[0];
const rules = session => session.report.textStyleInheritance.properties.filter(p => p.property.startsWith("run.scriptFontFamilies."));
const allMap = { Hans: "Simplified Source", Hant: "Traditional Source", Jpan: "Japanese Source", Hang: "Korean Source", Arab: "Arabic Source", Hebr: "Hebrew Source", Grek: "Greek Source" };

test("v3 admits only declared native buckets; v2/default/v1 remain compatible in both dialects", async () => {
  for (const strict of [false, true]) {
    const fixture = supplementalFontFixture({ strict, run: latin });
    for (const profile of [undefined, "script-families-v1", "script-families-v2", "script-families-v3"]) {
      const session = await open(fixture, { textFontProfile: profile });
      try {
        assert.deepEqual(run(session).style.scriptFontFamilies, profile === "script-families-v3" ? { Grek: allMap.Grek } : profile === "script-families-v2" ? allMap : undefined);
        assert.equal(run(session).style.eastAsianFontFamily, undefined);
        if (profile !== "script-families-v3") continue;
        assert.equal(session.report.textFontProfile, profile);
        const blocked = rules(session).filter(p => p.status === "unsupported");
        assert.equal(blocked.length, 12); // Six declarations on each of two source slides.
        assert.ok(blocked.every(p => p.value === undefined && p.source.part === "/deck/theme.xml" && p.source.path.endsWith("/@typeface") && p.resolvedFrom.some(d => d.part === p.target.part)));
        assert.ok(blocked.every(p => /no effective run\.(eastAsian|complexScript)FontFamily/.test(p.reason)));
        assert.ok(session.report.findings.some(f => f.source.part === "/deck/theme.xml" && /does not supply an absent native slot/.test(f.reason)));
        assert.ok(session.report.textStyleInheritance.properties.every(p => !Object.hasOwn(p, "fontMap")));
        assert.ok(Object.isFrozen(blocked[0]));
      } finally { session.dispose(); }
    }
  }
});

test("v3 inherits actual slots, keeps empty generic-theme maps and honors literal/unsupported/cross-slot overrides", async () => {
  for (const strict of [false, true]) {
    const fixture = supplementalFontFixture({ strict, inherited: true, run: latin });
    fixture.parts["deck/master.xml"] = fixture.parts["deck/master.xml"].replace(sourceFonts("Master Latin", "Master East", "Master Complex"), sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"));
    const inherited = await open(fixture);
    try { assert.deepEqual(run(inherited).style.scriptFontFamilies, allMap); assert.equal(rules(inherited).some(p => p.status === "unsupported"), false); }
    finally { inherited.dispose(); }
    for (const [slots, expected] of [
      [sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"), allMap],
      [sourceFonts("+mn-lt", "Local East", "+mn-cs"), { Grek: allMap.Grek, Arab: allMap.Arab, Hebr: allMap.Hebr }],
      [sourceFonts("+mn-lt", "", "+mn-cs"), { Grek: allMap.Grek, Arab: allMap.Arab, Hebr: allMap.Hebr }],
      [sourceFonts("+mn-lt", "+mn-lt", "+mn-lt"), { Grek: allMap.Grek }],
    ]) {
      const session = await open(supplementalFontFixture({ strict, run: slots, emptySlots: true }));
      try { assert.deepEqual(run(session).style.scriptFontFamilies, expected); assert.equal(rules(session).some(p => /absent native slot/.test(p.reason ?? "")), false); }
      finally { session.dispose(); }
    }
  }
});

test("v3 table fontRefs/explicit collections retain their own rules; missing run buckets cannot borrow Latin", async () => {
  for (const strict of [false, true]) for (const choice of ["reference", "collection", "absent", "partial-collection"]) {
    const text = choice === "reference" ? textReference() : choice === "collection" ? fontCollection("Table Latin", "Table East", "Table Complex", supplementalFonts) : choice === "partial-collection" ? '<a:font><a:font script="Arab" typeface="Only Arabic"/></a:font>' : "";
    const fixture = textAppearanceFixture({ strict, text, ...(choice === "absent" ? { paragraphs: `<a:p><a:r><a:rPr>${latin}</a:rPr><a:t>中عΩ</a:t></a:r></a:p>` } : {}) });
    fixture.parts["deck/theme.xml"] = fixture.parts["deck/theme.xml"].replace('</a:minorFont>', `${supplementalFonts}</a:minorFont>`);
    const session = await open(fixture);
    try {
      const first = session.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0];
      assert.deepEqual(first.style.scriptFontFamilies, choice === "absent" ? { Grek: allMap.Grek } : choice === "partial-collection" ? { Arab: "Only Arabic" } : allMap);
      assert.ok(session.report.tableStyleInheritance.properties.every(p => !Object.hasOwn(p, "fontMap")));
      if (choice === "absent") {
        const blocked = session.report.tableStyleInheritance.properties.filter(p => p.property.startsWith("run.scriptFontFamilies.") && p.status === "unsupported");
        assert.equal(blocked.length, 24); assert.ok(blocked.every(p => p.source.part === "/deck/theme.xml" && Number.isInteger(p.rowIndex) && Number.isInteger(p.columnIndex) && p.value === undefined));
        assert.ok(session.report.findings.some(f => /absent native slot/.test(f.reason)));
      }
    } finally { session.dispose(); }
  }
  const conditional = await open(textAppearanceFixture({ text: fontCollection("Base", "Base", "Base", supplementalFonts), flags: 'firstRow="1"', extraRegions: `<a:firstRow><a:tcTxStyle>${fontCollection("Near", "Near", "Near", '<a:font script="Hant" typeface="First Traditional"/>')}</a:tcTxStyle></a:firstRow>` }));
  try {
    const rows = conditional.document.slides[0].elements[0].rows;
    assert.deepEqual(rows[0].cells[0].content[0].runs[0].style.scriptFontFamilies, { Hant: "First Traditional" });
    assert.deepEqual(rows[1].cells[0].content[0].runs[0].style.scriptFontFamilies, allMap);
  } finally { conditional.dispose(); }
});

test("v3 empty insertion/end properties use the same slot-bound projection without inventing a text run", async () => {
  for (const strict of [false, true]) {
    const fixture = supplementalFontFixture({ strict, run: latin });
    for (const part of ["deck/z.xml", "deck/a.xml"]) fixture.parts[part] = fixture.parts[part].replace(/<a:r>[\s\S]*?<\/a:r>/, `<a:endParaRPr>${latin}</a:endParaRPr>`);
    const session = await open(fixture);
    try {
      assert.equal(run(session).text, ""); assert.deepEqual(run(session).style.scriptFontFamilies, { Grek: allMap.Grek });
      const blocked = rules(session).filter(p => p.status === "unsupported"); assert.equal(blocked.length, 12); assert.ok(blocked.every(p => p.target.path.includes("endParaRPr")));
    } finally { session.dispose(); }
    const note = notesStyleFixture({ strict, global: `<a:defPPr><a:defRPr>${latin}</a:defRPr></a:defPPr>` });
    note.parts["deck/notes/z.xml"] = note.parts["deck/notes/z.xml"].replace('<a:r><a:t>A中عא</a:t></a:r>', '<a:endParaRPr/>');
    const notes = await open(note, { notesTextProfile: "master-style-v1" });
    try { assert.equal(notes.document.slides[0].notes[0].runs[0].text, ""); assert.equal(notes.document.slides[0].notes[0].runs[0].style.scriptFontFamilies, undefined); }
    finally { notes.dispose(); }
  }
});

test("v3 notes use their actual master/theme buckets and scope the missing-selection findings independently", async () => {
  for (const strict of [false, true]) {
    const fixture = addSecondNotes(notesStyleFixture({ strict, global: `<a:defPPr><a:defRPr lang="zh-TW">${latin}</a:defRPr></a:defPPr>` }), true);
    const session = await open(fixture, { notesTextProfile: "master-style-v1" });
    try {
      for (const slide of session.document.slides) assert.equal(slide.notes[0].runs[0].style.scriptFontFamilies, undefined);
      const blocked = rules(session).filter(p => p.target.part.includes("/notes/") && p.status === "unsupported");
      assert.equal(blocked.length, 6); assert.ok(blocked.every(p => p.value === undefined && p.source.part.includes("notes-theme") && !p.resolvedFrom.some(d => d.part === "/deck/theme.xml")));
      assert.ok(session.report.findings.some(f => f.feature === "notes.inherited-text.unmapped-property" && /absent native slot/.test(f.reason)));
    } finally { session.dispose(); }
    const explicit = await open(notesStyleFixture({ strict }), { notesTextProfile: "master-style-v1" });
    try { assert.deepEqual(explicit.document.slides[0].notes[0].runs[0].style.scriptFontFamilies, { Hant: "Notes Traditional", Arab: "Notes Arabic", Hebr: "Notes Hebrew" }); }
    finally { explicit.dispose(); }
  }
});

test("v3 scope/MCE/unknown/bad-rule behavior never reuses an older or another master's map", async () => {
  const fixture = secondFontMaster(supplementalFontFixture({ run: latin }));
  fixture.parts["deck/theme2.xml"] = fixture.parts["deck/theme2.xml"].replace("Greek Source", "Second Greek");
  const scoped = await open(fixture);
  try { assert.deepEqual(run(scoped).style.scriptFontFamilies, { Grek: "Greek Source" }); assert.deepEqual(run(scoped, 1).style.scriptFontFamilies, { Grek: "Second Greek" }); }
  finally { scoped.dispose(); }
  for (const strict of [false, true]) {
    const majorFixture = supplementalFontFixture({ strict, run: sourceFonts("+mj-lt", "+mj-ea", "+mj-cs") });
    majorFixture.parts["deck/theme.xml"] = majorFixture.parts["deck/theme.xml"].replace('</a:majorFont>', `${supplementalFonts.replaceAll('typeface="', 'typeface="Major ')}</a:majorFont>`);
    const major = await open(majorFixture);
    try { assert.deepEqual(run(major).style.scriptFontFamilies, Object.fromEntries(Object.entries(allMap).map(([key, value]) => [key, `Major ${value}`]))); }
    finally { major.dispose(); }
    const f = supplementalFontFixture({ strict, run: latin, entries: '<mc:AlternateContent><mc:Choice Requires="future"><a:font script="Grek" typeface="Wrong"/></mc:Choice><mc:Fallback><a:font script="Grek" typeface="Effective"/><a:font script="Hant" typeface="Traditional"/></mc:Fallback></mc:AlternateContent><a:font script="Ethi" typeface="Source Only"/>' });
    f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('<a:theme ', '<a:theme xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:font:v3" mc:Ignorable="future" ');
    const session = await open(f);
    try { assert.deepEqual(run(session).style.scriptFontFamilies, { Grek: "Effective" }); assert.ok(session.report.findings.some(f => /Ethi/.test(f.reason))); }
    finally { session.dispose(); }
    const malformed = await open(supplementalFontFixture({ strict, run: latin, entries: '<a:font typeface="No Identity"/>' }));
    try { assert.equal(run(malformed).style.scriptFontFamilies, undefined); assert.ok(rules(malformed).every(p => p.status === "unsupported" && p.value === undefined)); assert.ok(malformed.report.findings.some(f => /without a script identity/.test(f.reason))); }
    finally { malformed.dispose(); }
    await assert.rejects(open(supplementalFontFixture({ strict, run: latin, entries: '<a:font script="Ethi" typeface="One"/><a:font script="Ethi" typeface="Two"/>' })), e => e.code === "import-invalid-package");
  }
});

test("v3 actual candidates retain their profile across three scalar script changes and atomically block font edits", async () => {
  let bytes = encodeStructure(supplementalFontFixture({ run: latin }));
  for (const text of ["日本語あΩ", "العربية עברית", "繁體中文 Ω plain"]) {
    const session = await importPptx(bytes, options);
    try {
      const slide = session.document.slides[0], before = slide.elements[0], next = structuredClone(before); next.text.content[0].runs[0].text = text;
      for (const mutate of [style => { style.scriptFontFamilies.Hant = "Unproved"; }, style => { delete style.scriptFontFamilies; }, style => { style.language = "ko-KR"; }, style => { style.eastAsianFontFamily = "New East"; }]) {
        const bad = structuredClone(next); mutate(bad.text.content[0].runs[0].style); slide.replaceElement(before.id, bad); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
      }
      slide.replaceElement(before.id, next); const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready");
      const saved = await session.save(plan.plan), initial = readGeneratedZip(bytes), actual = readGeneratedZip(saved.bytes);
      assert.deepEqual([...actual.keys()].sort(), [...initial.keys()].sort());
      for (const [name, payload] of initial) assert.deepEqual(name === "deck/z.xml" ? actual.get(name).toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>") : actual.get(name), name === "deck/z.xml" ? payload.toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>") : payload);
      const again = await importPptx(saved.bytes, options); try { assert.deepEqual(run(again).style.scriptFontFamilies, { Grek: allMap.Grek }); assert.equal(run(again).text, text); } finally { again.dispose(); }
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
});

test("v3 notes candidates and byte/file options capture cannot silently revert to v2 inference", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "pptkit-font-v3-"));
  try {
    const bytes = encodeStructure(notesStyleFixture({ global: `<a:defPPr><a:defRPr>${latin}</a:defRPr></a:defPPr>` }));
    const file = path.join(directory, "source.pptx"); await writeFile(file, bytes);
    for (const fromFile of [false, true]) {
      const config = { ...options, notesTextProfile: "master-style-v1" }, input = bytes.slice(), pending = fromFile ? importPptxFile(file, config) : importPptx(input, config); config.textFontProfile = "script-families-v2"; input.fill(0);
      const result = await pending, session = fromFile ? result.session : result;
      try {
        assert.equal(session.report.textFontProfile, "script-families-v3"); assert.equal(session.document.slides[0].notes[0].runs[0].style.scriptFontFamilies, undefined);
        const notes = structuredClone(session.document.slides[0].notes); notes[0].runs[0].text = "Notes 中عא"; session.document.slides[0].setNotes(notes);
        assert.equal((await session.prepareSave({ mode: "edit" })).status, "ready");
      } finally { session.dispose(); }
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test("v3 new blocked provenance/finding work stays bounded, cancelable and retryable", async () => {
  const fixture = supplementalFontFixture({ run: latin }), bytes = encodeStructure(fixture), baseline = await open(fixture);
  const count = baseline.report.textStyleInheritance.properties.length, dependencies = baseline.report.textStyleInheritance.properties.reduce((sum, p) => sum + (p.resolvedFrom?.length ?? 0), 0);
  baseline.dispose();
  for (const key of ["maxTextStyleProperties", "maxTextStyleDependencies"]) {
    const value = key === "maxTextStyleProperties" ? count : dependencies;
    const exact = await open(fixture, { sessionLimits: { [key]: value } }); exact.dispose();
    await assert.rejects(open(fixture, { sessionLimits: { [key]: value - 1 } }), e => e.code === "import-limit-exceeded");
  }
  await assert.rejects(open(fixture, { sessionLimits: { maxReportFindings: 1 } }), e => e.code === "import-limit-exceeded");
  const controller = new AbortController(), pending = importPptx(bytes, { ...options, signal: controller.signal }); queueMicrotask(() => controller.abort()); await assert.rejects(pending, e => e.code === "operation-aborted");
  await assert.rejects(open(fixture, { projection: "structure" }), TypeError);
  await assert.rejects(open(fixture, { textFontProfile: "script-families-v4" }), TypeError);
  const noWriter = await open(fixture, { editSaveProfile: undefined });
  try { const slide = noWriter.document.slides[0], next = structuredClone(slide.elements[0]); next.text.content[0].runs[0].text = "No permission"; slide.replaceElement(next.id, next); assert.equal((await noWriter.prepareSave({ mode: "edit" })).status, "blocked"); }
  finally { noWriter.dispose(); }
  const retry = await open(fixture); try { assert.equal(normalizePresentation(retry.document).slides.length, 2); } finally { retry.dispose(); }
});
