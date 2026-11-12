import test from "node:test";
import assert from "node:assert/strict";
import { importPptx } from "../dist/index.js";
import { normalizePresentation } from "../../core/dist/index.js";
import { encodeStructure } from "./structure-fixture.mjs";
import { textAppearanceFixture, textColor, textReference, fontCollection } from "./table-text-appearance-fixture.mjs";
import { matrixOverride } from "./theme-matrix-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { textStyleFixture, setTextTheme, textBody } from "./text-style-fixture.mjs";
import { inheritedShape } from "./placeholder-fixture.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const open = f => importPptx(encodeStructure(f), profile);
const run = (s, row = 0) => s.document.slides[0].elements[0].rows[row].cells[0].content[0].runs[0];
const props = (s, key) => s.report.tableStyleInheritance.properties.filter(p => p.property === key);

test("fontRef supplies source-backed default color/alpha and font in both dialects", async () => {
  for (const strict of [false, true]) {
    const s = await open(textAppearanceFixture({ strict, text: textReference(textColor("accent2", '<a:alpha val="25000"/>')) }));
    try {
      assert.equal(run(s).style.color, "CA6420"); assert.equal(run(s).style.colorOpacity, .25); assert.equal(run(s).style.fontFamily, "Arial");
      assert.ok(props(s, "run.color").every(p => p.source.path.includes("fontRef") && p.resolvedFrom.some(d => d.part === "/deck/theme.xml")));
      assert.ok(props(s, "run.fontFamily").every(p => p.source.path.endsWith("fontRef[1]/@idx")));
      assert.equal(s.report.findings.filter(p => /Unsupported source element.*fontRef/.test(p.reason)).length, 0);
    } finally { s.dispose(); }
  }
});

test("explicit region then local paragraph/run colors replace fontRef including opacity", async () => {
  for (const strict of [false, true]) {
    const f = textAppearanceFixture({ strict, text: textReference(textColor("accent2", '<a:alpha val="25000"/>')) + textColor("112233", "", "srgbClr") });
    let s = await open(f);
    try { assert.equal(run(s).style.color, "112233"); assert.equal(run(s).style.colorOpacity, 1); assert.ok(props(s, "run.color")[0].source.path.endsWith("tcTxStyle[1]/a:srgbClr[1]")); } finally { s.dispose(); }
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll("<a:p>", '<a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:defRPr></a:pPr>')
      .replaceAll("<a:r>", '<a:r><a:rPr><a:solidFill><a:srgbClr val="778899"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:rPr>');
    s = await open(f);
    try { assert.equal(run(s).style.color, "778899"); assert.equal(run(s).style.colorOpacity, 0); assert.equal(props(s, "run.color")[0].level, "run"); } finally { s.dispose(); }
  }
});

test("higher reference color blocks older color when unsupported, independently of font eligibility", async () => {
  for (const color of ['<a:prstClr val="red"/>', textColor("missing"), textColor("accent1", '<a:unknown val="50000"/>')]) {
    const s = await open(textAppearanceFixture({ text: textReference(textColor("112233", "", "srgbClr")), flags: 'firstRow="1"', extraRegions: `<a:firstRow><a:tcTxStyle>${textReference(color, "none")}</a:tcTxStyle></a:firstRow>` }));
    try {
      assert.ok(props(s, "run.color").some(p => p.rowIndex === 0 && p.status === "unsupported"));
      assert.notEqual(run(s).style.color, "112233"); assert.equal(run(s, 1).style.color, "112233");
      assert.ok(props(s, "run.fontFamily").some(p => p.rowIndex === 0 && p.status === "unsupported"));
    } finally { s.dispose(); }
  }
  const s = await open(textAppearanceFixture({ text: textReference(textColor("112233", "", "srgbClr"), "none") }));
  try { assert.equal(run(s).style.color, "112233"); assert.equal(props(s, "run.fontFamily")[0].status, "unsupported"); } finally { s.dispose(); }
});

test("explicit table font collections map only equal families and retain all alternative dependencies", async () => {
  for (const strict of [false, true]) for (const themed of [false, true]) {
    const collection = themed ? fontCollection("+mn-lt", "+mn-ea", "+mn-cs") : fontCollection();
    const s = await open(textAppearanceFixture({ strict, text: collection }));
    try {
      assert.equal(run(s).style.fontFamily, "Arial"); const p = props(s, "run.fontFamily")[0]; assert.equal(p.status, "mapped");
      assert.ok(p.resolvedFrom.some(d => d.path.includes("ea[1]/@typeface")));
      assert.ok(p.resolvedFrom.some(d => d.path.includes("cs[1]/@typeface")));
      if (themed) assert.equal(p.resolvedFrom.filter(d => d.part === "/deck/theme.xml").length, 3);
    } finally { s.dispose(); }
  }
});

test("different/empty/missing script slots block old Latin family without reviving after a later equal alternative", async () => {
  for (const text of [fontCollection("Arial", "Other", "Arial"), fontCollection("Arial", "", "Arial"), fontCollection("Arial", "+mn-ea", "Arial"), '<a:font><a:ea typeface="Arial"/></a:font>', '<a:font><a:latin typeface="Arial"/><a:ea/><a:cs typeface="Arial"/></a:font>']) {
    const f = textAppearanceFixture({ text });
    if (text.includes("+mn-ea")) f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:ea typeface="Arial"/>', '<a:ea typeface=""/>');
    const s = await open(f);
    try { assert.equal(props(s, "run.fontFamily")[0].status, "unsupported"); assert.notEqual(run(s).style.fontFamily, "Arial"); } finally { s.dispose(); }
  }
});

test("equal theme script slots work for local shape runs and different slots block inherited font", async () => {
  for (const strict of [false, true]) for (const different of [false, true]) {
    const f = setTextTheme(textStyleFixture(strict), { major: "Arial", minor: "Arial" });
    f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:latin typeface="Arial"/>', `<a:latin typeface="Arial"/><a:ea typeface="${different ? "Other" : "Arial"}"/><a:cs typeface="Arial"/>`);
    f.set("deck/z.xml", "sld", inheritedShape({ text: textBody({ paragraphs: '<a:p><a:r><a:rPr><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>Latin 中文 العربية</a:t></a:r></a:p>' }) }));
    const s = await open(f);
    try {
      const p = s.report.textStyleInheritance.properties.find(p => p.property === "run.fontFamily" && p.target.part === "/deck/z.xml");
      assert.equal(p.status, different ? "unsupported" : "mapped");
      if (!different) assert.equal(s.document.slides[0].elements[0].text.content[0].runs[0].style.fontFamily, "Arial");
    } finally { s.dispose(); }
  }
});

test("shared styles evaluate alternate font dependencies independently for two masters", async () => {
  const f = textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + textColor() });
  f.parts["deck/a.xml"] = f.parts["deck/z.xml"];
  f.parts["deck/_rels/a.xml.rels"] = f.parts["deck/_rels/a.xml.rels"].replace("layout.xml", "layout2.xml");
  for (const kind of ["layout", "master", "theme"]) { f.parts[`deck/${kind}2.xml`] = f.parts[`deck/${kind}.xml`]; f.types[`deck/${kind}2.xml`] = f.types[`deck/${kind}.xml`]; }
  f.parts["deck/_rels/layout2.xml.rels"] = f.parts["deck/_rels/layout.xml.rels"].replace("master.xml", "master2.xml");
  f.parts["deck/_rels/master2.xml.rels"] = f.parts["deck/_rels/master.xml.rels"].replace("layout.xml", "layout2.xml").replace("theme.xml", "theme2.xml");
  f.parts["deck/theme2.xml"] = f.parts["deck/theme2.xml"].replaceAll('typeface="Arial"', 'typeface="Second Font"');
  f.parts["deck/main.data"] = f.parts["deck/main.data"].replace("</p:sldMasterIdLst>", '<p:sldMasterId id="2147483650" r:id="m2"/></p:sldMasterIdLst>');
  f.parts["deck/_rels/main.data.rels"] = f.parts["deck/_rels/main.data.rels"].replace("</Relationships>", `<Relationship Id="m2" Type="${f.R}/slideMaster" Target="master2.xml"/></Relationships>`);
  for (const missing of [false, true]) {
    if (missing) f.parts["deck/theme.xml"] = f.parts["deck/theme.xml"].replaceAll('<a:ea typeface="Arial"/>', '<a:ea typeface=""/>');
    const s = await open(f);
    try {
      const second = s.document.slides[1].elements[0].rows[0].cells[0].content[0].runs[0]; assert.equal(second.style.fontFamily, "Second Font");
      assert.equal(props(s, "run.fontFamily").find(p => p.target.part === "/deck/z.xml").status, missing ? "unsupported" : "mapped");
      assert.ok(props(s, "run.fontFamily").filter(p => p.target.part === "/deck/a.xml").every(p => p.status === "mapped" && p.resolvedFrom.some(d => d.part === "/deck/theme2.xml")));
    } finally { s.dispose(); }
  }
});

test("supplemental font collections agree explicitly or block, duplicate scripts reject", async () => {
  for (const [extra, mapped] of [['<a:font script="Hans" typeface="Arial"/>', true], ['<a:font script="Arab" typeface="Other"/>', false], ['<a:font script="" typeface="Arial"/>', false]]) {
    const s = await open(textAppearanceFixture({ text: fontCollection("Arial", "Arial", "Arial", extra) }));
    try { assert.equal(props(s, "run.fontFamily")[0].status, mapped ? "mapped" : "unsupported"); if (mapped) assert.ok(props(s, "run.fontFamily")[0].resolvedFrom.some(d => d.path.includes("font[1]/a:font[1]"))); } finally { s.dispose(); }
  }
  const f = textAppearanceFixture({ text: fontCollection("Arial", "Arial", "Arial", '<a:font script="Hans" typeface="Arial"/><a:font script="Hans" typeface="Arial"/>') });
  await assert.rejects(open(f), e => e.code === "import-invalid-package" && e.cause?.code === "presentation-invalid");
});

test("nearer font components do not mix with old slots and leave color inheritance independent", async () => {
  const f = textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + textColor() });
  matrixOverride(f, '<a:fontScheme name="Partial"><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>');
  const s = await open(f);
  try { assert.equal(props(s, "run.fontFamily")[0].status, "unsupported"); assert.equal(run(s).style.color, "CA6420"); } finally { s.dispose(); }
});

test("duplicate reference colors and script font slots reject while valid independent retry succeeds", async () => {
  for (const text of [textReference(textColor() + textColor("accent1")), '<a:font><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:ea typeface="Arial"/></a:font>']) {
    await assert.rejects(open(textAppearanceFixture({ text })), e => e.code === "import-invalid-package" && e.cause?.code === "presentation-invalid");
  }
  const s = await open(textAppearanceFixture()); try { assert.equal(run(s).style.color, "CA6420"); } finally { s.dispose(); }
});

test("three source scalar saves retain text appearances, source styles and themes; font edits block atomically", async () => {
  let bytes = encodeStructure(textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + textColor() }));
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, profile);
    try {
      assert.equal(run(s).style.fontFamily, "Arial"); assert.equal(run(s).style.color, "CA6420");
      const slide = s.document.slides[0], old = slide.elements[0], next = structuredClone(old); next.rows[0].cells[0].content[0].runs[0].text = `轮次 ${round} α`;
      slide.replaceElement(old.id, next); const ready = await s.prepareSave({ mode: "edit" }); assert.equal(ready.status, "ready");
      const result = await s.save(ready.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(result.bytes);
      for (const [name, value] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), value);
      const bad = structuredClone(next); bad.rows[0].cells[0].content[0].runs[0].style.fontFamily = "Other"; slide.replaceElement(old.id, bad);
      assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(old.id, next); assert.equal((await s.prepareSave({ mode: "edit" })).status, "ready"); bytes = result.bytes;
    } finally { s.dispose(); }
  }
});

test("font/color provenance budgets remain exact and input capture/cancellation leave clean retries", async () => {
  const bytes = encodeStructure(textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs", '<a:font script="Hans" typeface="Arial"/>') + textColor() }));
  const s = await importPptx(bytes, profile), expected = normalizePresentation(s.document); s.dispose();
  for (const key of ["maxTableStyleDependencies", "maxTableStyleWork"]) {
    let lo = 1, hi = 2_000_000;
    while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { const session = await importPptx(bytes, { ...profile, sessionLimits: { [key]: mid } }); session.dispose(); hi = mid; } catch (e) { assert.equal(e.code, "import-limit-exceeded"); lo = mid + 1; } }
    await assert.rejects(importPptx(bytes, { ...profile, sessionLimits: { [key]: lo - 1 } }), { code: "import-limit-exceeded" });
    const session = await importPptx(bytes, { ...profile, sessionLimits: { [key]: lo } }); assert.deepEqual(normalizePresentation(session.document), expected); session.dispose();
  }
  const controller = new AbortController(), pending = importPptx(bytes, { ...profile, signal: controller.signal }); queueMicrotask(() => controller.abort()); await assert.rejects(pending, { code: "operation-aborted" });
  const input = Uint8Array.from(bytes), options = { ...profile }, captured = importPptx(input, options); input.fill(0); options.projection = "structure";
  const retry = await captured; try { assert.deepEqual(normalizePresentation(retry.document), expected); } finally { retry.dispose(); }
});
