import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createPresentation, normalizePresentation, resolveTextFontSegments } from "@pptkit/core";
import { generatePptx } from "../../pptx-exporter/dist/index.js";
import { importPptx } from "../dist/index.js";
import { importPptxFile } from "../dist/node.js";
import { supplementalFontFixture, supplementalFonts } from "./supplemental-font-fixture.mjs";
import { sourceFonts, secondFontMaster } from "./script-font-import-fixture.mjs";
import { textAppearanceFixture, textReference, fontCollection } from "./table-text-appearance-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { addTextThemeOverride } from "./text-style-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const profile = { projection: "content", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" };
const open = f => importPptx(encodeStructure(f), profile);
const run = (s, index = 0) => s.document.slides[index].elements[0].text.content[0].runs[0];
const props = (s, part = "/deck/z.xml") => s.report.textStyleInheritance.properties.filter(p => p.target.part === part && p.property.startsWith("run.scriptFontFamilies."));
const map = { Hans: "Simplified Source", Hant: "Traditional Source", Jpan: "Japanese Source", Hang: "Korean Source", Arab: "Arabic Source", Hebr: "Hebrew Source", Grek: "Greek Source" };
const families = s => { const n = normalizePresentation(s.document), r = run(s); return resolveTextFontSegments(r.text, r.style, n.theme).map(v => v.family); };

test("v2 transports all admitted theme rules independently of current text in both dialects and keeps old profiles", async () => {
  for (const strict of [false, true]) {
    const f = supplementalFontFixture({ strict }), bytes = encodeStructure(f), s = await open(f), v1 = await importPptx(bytes, { ...profile, textFontProfile: "script-families-v1" }), old = await importPptx(bytes, { projection: "content" });
    try {
      assert.equal(s.report.textFontProfile, "script-families-v2"); assert.deepEqual(run(s).style.scriptFontFamilies, map);
      assert.deepEqual(families(s), ["Body Source", "Traditional Source", "Arabic Source", "Hebrew Source", "Greek Source"]);
      assert.equal(run(v1).style.scriptFontFamilies, undefined); assert.equal(run(old).style.scriptFontFamilies, undefined);
      assert.ok(v1.report.findings.some(f => /Theme supplemental/.test(f.reason))); assert.equal(s.report.findings.some(f => /Theme supplemental/.test(f.reason)), false);
      assert.equal(props(s).length, 7); assert.ok(props(s).every(p => p.source.part === "/deck/theme.xml" && p.source.path.endsWith("/@typeface") && p.resolvedFrom.some(d => d.part === "/deck/z.xml")));
      assert.ok(s.report.textStyleInheritance.properties.every(p => !Object.hasOwn(p, "fontMap"))); assert.ok(Object.isFrozen(s.report.textStyleInheritance.properties));
      assert.equal(run(s).text, "A中عאΩ"); assert.equal(s.document.slides[0].elements[0].text.content[0].runs.length, 1);
    } finally { s.dispose(); v1.dispose(); old.dispose(); }
  }
});

test("empty theme ea/cs retain independent generic blockers and supported supplemental rules without a Latin guess", async () => {
  const s = await open(supplementalFontFixture({ emptySlots: true }));
  try {
    assert.equal(run(s).style.eastAsianFontFamily, undefined); assert.equal(run(s).style.complexScriptFontFamily, undefined);
    assert.deepEqual(run(s).style.scriptFontFamilies, map); assert.equal(families(s)[1], "Traditional Source");
    for (const property of ["run.eastAsianFontFamily", "run.complexScriptFontFamily"]) assert.ok(s.report.textStyleInheritance.properties.some(p => p.property === property && p.status === "unsupported"));
    assert.ok(props(s).every(p => p.status === "mapped"));
  } finally { s.dispose(); }
});

test("effective slot cascade applies local literal/unsupported/other-collection overrides without resurrecting older maps", async () => {
  for (const strict of [false, true]) for (const [override, expected] of [
    ['<a:latin typeface="Literal"/>', { ...map, Grek: undefined }],
    ['<a:ea typeface="Local East"/>', { Arab: map.Arab, Hebr: map.Hebr, Grek: map.Grek }],
    ['<a:cs typeface=""/>', { Hans: map.Hans, Hant: map.Hant, Jpan: map.Jpan, Hang: map.Hang, Grek: map.Grek }],
  ]) {
    const f = supplementalFontFixture({ strict, inherited: true, run: override });
    f.parts["deck/master.xml"] = f.parts["deck/master.xml"].replace(sourceFonts("Master Latin", "Master East", "Master Complex"), sourceFonts("+mn-lt", "+mn-ea", "+mn-cs"));
    f.parts["deck/layout.xml"] = f.parts["deck/layout.xml"].replace('<a:latin typeface="Layout Latin"/>', '<a:latin typeface="+mn-lt"/>');
    const s = await open(f);
    try { assert.deepEqual(run(s).style.scriptFontFamilies, Object.fromEntries(Object.entries(expected).filter(([, v]) => v !== undefined))); }
    finally { s.dispose(); }
  }
  const f = supplementalFontFixture({ run: sourceFonts("+mn-lt", "+mj-ea", "+mn-cs") });
  f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('</a:majorFont>', '<a:font script="Hant" typeface="Heading Traditional"/></a:majorFont>');
  const s = await open(f); try { assert.deepEqual(run(s).style.scriptFontFamilies, { Grek: map.Grek, Hant: "Heading Traditional", Arab: map.Arab, Hebr: map.Hebr }); } finally { s.dispose(); }
});

test("Latin theme collection supplies omitted native slots, while explicit cross-slot tokens stay concrete", async () => {
  const inherited = await open(supplementalFontFixture({ run: '<a:latin typeface="+mn-lt"/>' }));
  try { assert.deepEqual(run(inherited).style.scriptFontFamilies, map); assert.equal(run(inherited).style.eastAsianFontFamily, undefined); }
  finally { inherited.dispose(); }
  const cross = await open(supplementalFontFixture({ run: sourceFonts("+mn-lt", "+mn-lt", "+mn-lt") }));
  try { assert.deepEqual(run(cross).style.scriptFontFamilies, { Grek: map.Grek }); assert.equal(run(cross).style.eastAsianFontFamily, "Body Source"); }
  finally { cross.dispose(); }
});

test("unknown/invalid supplemental rules have located findings, block their own key and reject duplicates", async () => {
  const s = await open(supplementalFontFixture({ entries: '<a:font script="Hant" typeface="Traditional"/><a:font script="Arab" typeface="+mn-lt"/><a:font script="Hebr" typeface=""/><a:font script="Deva" typeface="Bad&#10;Font"/><a:font script="Ethi" typeface="Unknown"/>' }));
  try {
    assert.deepEqual(run(s).style.scriptFontFamilies, { Hant: "Traditional" });
    assert.equal(props(s).filter(p => p.status === "unsupported").length, 3);
    assert.ok(s.report.findings.some(f => f.source.part === "/deck/theme.xml" && /Ethi/.test(f.reason)));
  } finally { s.dispose(); }
  const bad = await open(supplementalFontFixture({ entries: supplementalFonts + '<a:font typeface="No Identity"/>' }));
  try { assert.equal(run(bad).style.scriptFontFamilies, undefined); assert.equal(props(bad).length, 37); assert.ok(props(bad).every(p => p.status === "unsupported")); }
  finally { bad.dispose(); }
  for (const entries of ['<a:font script="Hant" typeface="One"/><a:font script="Hant" typeface="Two"/>', '<a:font script="Ethi" typeface="One"/><a:font script="Ethi" typeface="Two"/>']) await assert.rejects(open(supplementalFontFixture({ entries })), e => e.code === "import-invalid-package" && e.cause?.code === "presentation-invalid");
});

test("nearer font schemes replace maps as a component; multiple masters and scopes keep maps/audits separate", async () => {
  const f = secondFontMaster(supplementalFontFixture());
  f.parts["deck/theme2.xml"] = f.parts["deck/theme2.xml"].replaceAll("Traditional Source", "Second Traditional");
  addTextThemeOverride(f, "z", "partial-map", '<a:fontScheme name="Near"><a:minorFont><a:latin typeface="Near Latin"/><a:ea typeface="Near East"/><a:cs typeface="Near Complex"/><a:font script="Hant" typeface="Near Traditional"/></a:minorFont></a:fontScheme>');
  const s = await open(f);
  try { assert.deepEqual(run(s).style.scriptFontFamilies, { Hant: "Near Traditional" }); assert.equal(run(s, 1).style.scriptFontFamilies.Hant, "Second Traditional"); assert.ok(props(s).every(p => p.source.part === "/deck/partial-map.xml")); }
  finally { s.dispose(); }
  const same = supplementalFontFixture();
  addTextThemeOverride(same, "z", "bad-map", '<a:fontScheme name="Near"><a:minorFont><a:latin typeface="Near"/><a:ea typeface="Near"/><a:cs typeface="Near"/><a:font script="Arab" typeface=""/></a:minorFont></a:fontScheme>');
  const scopes = await open(same);
  try { assert.equal(run(scopes).style.scriptFontFamilies, undefined); assert.deepEqual(run(scopes, 1).style.scriptFontFamilies, map); assert.equal(scopes.report.findings.filter(f => /Supplemental typeface/.test(f.reason)).length, 1); }
  finally { scopes.dispose(); }
});

test("table collections and fontRefs transport maps, conditional collection replaces maps, local native slot wins", async () => {
  for (const strict of [false, true]) for (const reference of [false, true]) {
    const f = textAppearanceFixture({ strict, text: reference ? textReference() : fontCollection("Table Latin", "Table East", "Table Complex", supplementalFonts) });
    if (reference) f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('</a:minorFont>', `${supplementalFonts}</a:minorFont>`);
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll('<a:r>', '<a:r><a:rPr><a:cs typeface="Local Complex"/></a:rPr>');
    const s = await open(f);
    try {
      const r = s.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0];
      assert.deepEqual(r.style.scriptFontFamilies, { Hans: map.Hans, Hant: map.Hant, Jpan: map.Jpan, Hang: map.Hang, Grek: map.Grek }); assert.equal(r.style.complexScriptFontFamily, "Local Complex");
      assert.ok(s.report.tableStyleInheritance.properties.some(p => p.property === "run.scriptFontFamilies.Hant" && p.source.part === (reference ? "/deck/theme.xml" : "/deck/tableStyles.xml")));
    } finally { s.dispose(); }
  }
  const conditional = await open(textAppearanceFixture({ text: fontCollection("Base", "Base", "Base", supplementalFonts), flags: 'firstRow="1"', extraRegions: `<a:firstRow><a:tcTxStyle>${fontCollection("Near", "Near", "Near", '<a:font script="Hant" typeface="First Traditional"/>')}</a:tcTxStyle></a:firstRow>` }));
  try {
    const rows = conditional.document.slides[0].elements[0].rows;
    assert.deepEqual(rows[0].cells[0].content[0].runs[0].style.scriptFontFamilies, { Hant: "First Traditional" });
    assert.deepEqual(rows[1].cells[0].content[0].runs[0].style.scriptFontFamilies, map);
  } finally { conditional.dispose(); }
  const partial = await open(textAppearanceFixture({ text: '<a:font><a:font script="Arab" typeface="Only Arabic"/></a:font>' }));
  try {
    assert.deepEqual(partial.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].style.scriptFontFamilies, { Arab: "Only Arabic" });
    const generic = partial.report.tableStyleInheritance.properties.filter(p => p.rowIndex === 0 && p.columnIndex === 0 && ["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"].includes(p.property));
    assert.equal(generic.length, 3); assert.ok(generic.every(p => p.status === "unsupported"));
  } finally { partial.dispose(); }
  const own = textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs", '<a:font script="Hant" typeface="Own Traditional"/>') });
  own.parts["deck/theme.xml"] = own.parts["deck/theme.xml"].replace('</a:minorFont>', `${supplementalFonts}</a:minorFont>`);
  const replaced = await open(own);
  try { assert.deepEqual(replaced.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].style.scriptFontFamilies, { Hant: "Own Traditional" }); }
  finally { replaced.dispose(); }
});

test("supplemental mapping uses effective MCE branches and retains unknown attributes as source-only", async () => {
  for (const strict of [false, true]) {
    const f = supplementalFontFixture({ strict, entries: '<mc:AlternateContent><mc:Choice Requires="future"><a:font script="Hant" typeface="Wrong Choice"/></mc:Choice><mc:Fallback><a:font script="Hant" typeface="Effective Traditional" future:hint="retain"/></mc:Fallback></mc:AlternateContent>' });
    f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replace('<a:theme ', '<a:theme xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:pptkit:supplemental:future" mc:Ignorable="future" ');
    const s = await open(f);
    try { assert.deepEqual(run(s).style.scriptFontFamilies, { Hant: "Effective Traditional" }); assert.ok(props(s).every(p => p.value !== "Wrong Choice")); }
    finally { s.dispose(); }
  }
  const f = supplementalFontFixture({ entries: '<a:font script="Hant" typeface="Traditional" charset="128"/>' }), s = await open(f);
  try { assert.equal(run(s).style.scriptFontFamilies.Hant, "Traditional"); assert.ok(s.report.findings.some(f => f.source.part === "/deck/theme.xml" && f.source.path?.endsWith("/@charset"))); }
  finally { s.dispose(); }
});

test("scalar character saves retain source rules over three text/script changes; map/style edits stay atomic", async () => {
  let bytes = encodeStructure(supplementalFontFixture({ emptySlots: true }));
  for (const text of ["Japanese あ中", "العربية עברית 中", "中文 Ω plain"]) {
    const s = await importPptx(bytes, profile);
    try {
      const slide = s.document.slides[0], old = slide.elements[0], next = structuredClone(old); next.text.content[0].runs[0].text = text;
      for (const mutate of [r => { r.style.scriptFontFamilies.Hant = "Changed"; }, r => { delete r.style.scriptFontFamilies; }, r => { r.style.language = "ja"; }]) {
        const bad = structuredClone(next); mutate(bad.text.content[0].runs[0]); slide.replaceElement(old.id, bad); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
      }
      slide.replaceElement(old.id, next); const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan);
      const before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes); for (const [name, payload] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), payload);
      assert.deepEqual(after.get("deck/z.xml").toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"), before.get("deck/z.xml").toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"));
      bytes = saved.bytes;
    } finally { s.dispose(); }
    const next = await importPptx(bytes, profile); try { assert.equal(run(next).text, text); assert.deepEqual(run(next).style.scriptFontFamilies, map); } finally { next.dispose(); }
  }
});

test("v2 direct notes keep existing scope boundary; supplemental options capture through bytes/file/cancel", async t => {
  const p = createPresentation(); p.addSlide({ notes: [{ runs: [{ text: "Notes", style: { fontFamily: "Literal", eastAsianFontFamily: "Eastern" } }] }] });
  const notes = await importPptx((await generatePptx(p)).bytes, profile);
  try { assert.equal(notes.document.slides[0].notes[0].runs[0].style.eastAsianFontFamily, "Eastern"); assert.equal(notes.document.slides[0].notes[0].runs[0].style.scriptFontFamilies, undefined); }
  finally { notes.dispose(); }
  const bytes = encodeStructure(supplementalFontFixture()), source = Uint8Array.from(bytes), config = { ...profile }, pending = importPptx(source, config); config.textFontProfile = "invalid"; source.fill(0);
  const captured = await pending; try { assert.deepEqual(run(captured).style.scriptFontFamilies, map); } finally { captured.dispose(); }
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-supplemental-")); t.after(() => rm(dir, { recursive: true, force: true })); const file = path.join(dir, "source.pptx"); await writeFile(file, bytes);
  const opts = { ...profile }, filePending = importPptxFile(file, opts); opts.textFontProfile = undefined; const fromFile = await filePending; try { assert.equal(fromFile.session.report.textFontProfile, "script-families-v2"); } finally { fromFile.session.dispose(); }
  const controller = new AbortController(), operation = importPptx(bytes, { ...profile, signal: controller.signal }); queueMicrotask(() => controller.abort()); await assert.rejects(operation, { code: "operation-aborted" });
  for (const options of [{ projection: "structure", textFontProfile: "script-families-v2" }, { projection: "content", textFontProfile: "script-families-v99" }]) await assert.rejects(importPptx(bytes, options), TypeError);
  const retry = await importPptx(bytes, profile); retry.dispose();
});

test("map properties/dependencies and table work charge existing cumulative limits at exact boundaries", async () => {
  const bytes = encodeStructure(supplementalFontFixture()), base = await importPptx(bytes, profile), n = normalizePresentation(base.document); base.dispose();
  const table = textAppearanceFixture({ text: fontCollection("Latin", "East", "Complex", supplementalFonts) });
  for (const [source, keys] of [[bytes, ["maxTextStyleProperties", "maxTextStyleDependencies"]], [encodeStructure(table), ["maxTableStyleProperties", "maxTableStyleDependencies", "maxTableStyleWork"]]]) for (const key of keys) {
    let lo = 1, hi = 1_000_000;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { const s = await importPptx(source, { ...profile, sessionLimits: { [key]: mid } }); s.dispose(); hi = mid; } catch (e) { assert.equal(e.code, "import-limit-exceeded"); lo = mid + 1; } }
    await assert.rejects(importPptx(source, { ...profile, sessionLimits: { [key]: lo - 1 } }), { code: "import-limit-exceeded" });
    const exact = await importPptx(source, { ...profile, sessionLimits: { [key]: lo } }); try { if (source === bytes) assert.deepEqual(normalizePresentation(exact.document), n); } finally { exact.dispose(); }
  }
});
