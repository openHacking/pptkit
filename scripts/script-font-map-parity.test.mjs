import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createScriptFontMapFixture, mapText, mapStyle } from "../tests/fixtures/script-font-map.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const expectedFonts = ["Latin Map", "Traditional Map", "Arabic Map", "Hebrew Map"];
const moduleUrl = s => `data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;

test("script-map Node/ESM/global canonical, selection, layout and outputs are identical without network", async t => {
  const p = createScriptFontMapFixture(createPresentation), canonical = normalizePresentation(p), output = await generatePptx(p);
  const expected = { canonical, layout: resolveLayout(p), bytes: [...output.bytes], warnings: output.warnings, preview: await renderPresentationToSvg(p), segments: resolveTextFontSegments(mapText, mapStyle, canonical.theme) };
  const browser = await chromium.launch(); t.after(() => browser.close()); const urls = {};
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) {
    let source = await readFile(`packages/${name}/dist/browser/index.js`, "utf8");
    for (const [dependency, url] of Object.entries(urls)) source = source.replaceAll(`"@pptkit/${dependency}"`, JSON.stringify(url));
    urls[name] = moduleUrl(source);
  }
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
    const actual = await page.evaluate(async ({ urls, factory, mapText, mapStyle }) => {
      const c = urls ? await import(urls.core) : PPTKit, l = urls ? await import(urls.layout) : PPTKit, e = urls ? await import(urls["pptx-exporter"]) : PPTKit, s = urls ? await import(urls["svg-renderer"]) : PPTKit;
      const p = Function("createPresentation", "mapText", "mapStyle", `return (${factory})(createPresentation)`)(c.createPresentation, mapText, mapStyle);
      const canonical = c.normalizePresentation(p), output = await e.generatePptx(p);
      return { canonical, layout: l.resolveLayout(p), bytes: [...output.bytes], warnings: output.warnings, preview: await s.renderPresentationToSvg(p), segments: c.resolveTextFontSegments(mapText, mapStyle, canonical.theme) };
    }, { urls: format === "esm" ? urls : null, factory: createScriptFontMapFixture.toString(), mapText, mapStyle });
    assert.deepEqual(actual, expected); assert.deepEqual(requests, []); await page.close();
  }
});

test("independent DOM verifies materialized slots, alpha/actions, notes, warnings and both preview paths", async t => {
  const p = createScriptFontMapFixture(createPresentation), before = normalizePresentation(p), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const output = await generatePptx(p), parts = readGeneratedZip(output.bytes);
  const warnings = output.warnings.filter(w => w.code === "script-font-map-materialized");
  assert.equal(warnings.length, 6);
  assert.deepEqual(warnings.filter(w => w.elementId).map(w => w.elementId).sort(), ["map-linked", "map-native", "map-shape", "map-table"]);
  assert.ok(warnings.some(w => w.slideId === "map-slide" && w.path === "notes"));
  assert.ok(warnings.some(w => w.layoutId === "map-layout" && w.path === "placeholders.0.textStyle.run.scriptFontFamilies"));
  const actual = await page.evaluate(({ slide, notes, svg }) => {
    const parser = new DOMParser(), ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const read = xml => {
      const doc = parser.parseFromString(xml, "application/xml"); if (doc.querySelector("parsererror")) throw new Error("Invalid XML");
      return [...doc.getElementsByTagNameNS(ns, "r")].map(r => ({ text: r.getElementsByTagNameNS(ns, "t")[0]?.textContent, fonts: ["latin", "ea", "cs"].map(n => r.getElementsByTagNameNS(ns, n)[0]?.getAttribute("typeface")), alpha: r.getElementsByTagNameNS(ns, "alpha")[0]?.getAttribute("val"), links: r.getElementsByTagNameNS(ns, "hlinkClick").length }));
    };
    const preview = parser.parseFromString(svg, "image/svg+xml");
    return { slide: read(slide), notes: read(notes), native: [...preview.getElementsByTagName("tspan")].map(n => [n.textContent, n.getAttribute("font-family").slice(1, -1)]), html: [...preview.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "span")].filter(n => n.hasAttribute("lang")).map(n => [n.textContent, n.style.fontFamily.replaceAll('"', "")]), links: preview.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "a").length };
  }, { slide: parts.get("ppt/slides/slide1.xml").toString(), notes: parts.get("ppt/notesSlides/notesSlide1.xml").toString(), svg: (await renderPresentationToSvg(p)).slides[0].svg });
  assert.equal(actual.slide.length, 16); assert.equal(actual.notes.length, 4);
  for (const [i, r] of actual.slide.entries()) { assert.equal(r.text, [...mapText][i % 4]); assert.deepEqual(r.fonts, Array(3).fill(expectedFonts[i % 4])); assert.equal(r.alpha, "50000"); assert.equal(r.links, i >= 4 && i < 8 ? 1 : 0); }
  for (const [i, r] of actual.notes.entries()) assert.deepEqual(r.fonts, Array(3).fill(expectedFonts[i]));
  for (const kind of ["native", "html"]) for (const [i, text] of [...mapText].entries()) assert.ok(actual[kind].some(pair => pair[0] === text && pair[1] === expectedFonts[i]), `${kind}/${text}`);
  assert.equal(actual.links, 1); assert.deepEqual(normalizePresentation(p), before);
});

test("three fresh generations retain effective fonts but lose map topology; source no-op exact, map edits blocked", async () => {
  let bytes = (await generatePptx(createScriptFontMapFixture(createPresentation))).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v1", editSaveProfile: "text-runs-v3" });
    try {
      const n = normalizePresentation(session.document), run = n.slides[0].elements[0].content[0].runs;
      assert.equal(run.map(r => r.text).join(""), mapText); assert.deepEqual(run.map(r => r.style.fontFamily), expectedFonts);
      assert.ok(run.every(r => !Object.hasOwn(r.style, "scriptFontFamilies"))); assert.equal(run.length, 4);
      const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      const slide = session.document.slides[0], original = slide.elements[0], changed = structuredClone(original);
      changed.content[0].runs[0].style.scriptFontFamilies = { Latn: "Other" }; slide.replaceElement(original.id, changed);
      assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(original.id, original);
      assert.equal((await session.prepareSave({ mode: "no-op" })).status, "ready");
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test("empty and currently-unused maps warn, retain insertion/action; absent maps preserve native default output", async () => {
  for (const text of ["", "Only Latin"]) {
    const p = createPresentation(); p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: [{ runs: [{ text, style: { fontFamily: { theme: "body" }, scriptFontFamilies: { Arab: "Unused" } }, action: { type: "url", url: "https://example.invalid" } }] }] });
    const result = await generatePptx(p), xml = readGeneratedZip(result.bytes).get("ppt/slides/slide1.xml").toString();
    assert.equal(result.warnings.filter(w => w.code === "script-font-map-materialized").length, 1); assert.match(xml, /<a:hlinkClick/); assert.doesNotMatch(xml, /Unused|\+mn-lt/);
    const e = structuredClone(p.slides[0].elements[0]); delete e.content[0].runs[0].style.scriptFontFamilies; p.slides[0].replaceElement(e.id, e);
    const legacy = await generatePptx(p); assert.equal(legacy.warnings.length, 0); assert.match(readGeneratedZip(legacy.bytes).get("ppt/slides/slide1.xml").toString(), /\+mn-lt/);
  }
});

test("fresh native-font controls export independent slots versus materialized map without installed fonts", async t => {
  const { createRealFontFreshFixture } = await import("../tests/fixtures/real-font-fresh.mjs");
  const { nativeFontFamilies } = await import("./research/font-native-cases.mjs");
  const fixture = createRealFontFreshFixture(createPresentation, nativeFontFamilies), before = normalizePresentation(fixture.document);
  const result = await generatePptx(fixture.document), parts = readGeneratedZip(result.bytes);
  assert.deepEqual(normalizePresentation(fixture.document), before);
  assert.equal(fixture.controls.length, 16);
  assert.deepEqual(result.warnings.filter(w => w.code === "script-font-map-materialized").map(w => [w.slideId, w.elementId]), fixture.controls.filter(c => c.mode === "map").map(c => [c.slideId, c.elementId]));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const actual = await page.evaluate(xmls => xmls.map(xml => {
    const dom = new DOMParser().parseFromString(xml, "application/xml"), ns = "http://schemas.openxmlformats.org/drawingml/2006/main", p = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const shapes = [...dom.getElementsByTagNameNS(p, "sp")];
    const probe = shapes.find(s => s.getElementsByTagNameNS(p, "cNvPr")[0].getAttribute("name") === "Probe");
    const runs = [...probe.getElementsByTagNameNS(ns, "r")];
    return { label: shapes[0].getElementsByTagNameNS(ns, "t")[0].textContent, runs: runs.map(r => ({ text: r.getElementsByTagNameNS(ns, "t")[0].textContent, language: r.getElementsByTagNameNS(ns, "rPr")[0].getAttribute("lang"), slots: ["latin", "ea", "cs"].map(slot => r.getElementsByTagNameNS(ns, slot)[0].getAttribute("typeface")) })) };
  }), fixture.controls.map((_, i) => parts.get(`ppt/slides/slide${i + 1}.xml`).toString()));
  assert.deepEqual(actual, fixture.controls.map(c => ({ label: c.id, runs: [{ text: c.text, language: c.language, slots: c.expectedSlots }] })));
});

test("original mixed-script author runs materialize three families and reimport without reusable maps", async t => {
  const { createRealFontMixedFixture } = await import("../tests/fixtures/real-font-mixed.mjs");
  const { nativeFontFamilies } = await import("./research/font-native-cases.mjs");
  const fixture = createRealFontMixedFixture(createPresentation, nativeFontFamilies), before = normalizePresentation(fixture.document);
  for (const slide of before.slides) assert.equal(slide.elements.find(e => e.name === "Probe").content[0].runs.length, 1);
  const result = await generatePptx(fixture.document), parts = readGeneratedZip(result.bytes);
  assert.deepEqual(normalizePresentation(fixture.document), before);
  assert.deepEqual(result.warnings.filter(w => w.code === "script-font-map-materialized").map(w => [w.slideId, w.elementId]), fixture.controls.map(c => [c.slideId, c.elementId]));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const actual = await page.evaluate(xmls => xmls.map(xml => {
    const doc = new DOMParser().parseFromString(xml, "application/xml"), p = "http://schemas.openxmlformats.org/presentationml/2006/main", a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const probe = [...doc.getElementsByTagNameNS(p, "sp")].find(s => s.getElementsByTagNameNS(p, "cNvPr")[0].getAttribute("name") === "Probe");
    return [...probe.getElementsByTagNameNS(a, "r")].map(r => ({ text: r.getElementsByTagNameNS(a, "t")[0].textContent, slots: ["latin", "ea", "cs"].map(slot => r.getElementsByTagNameNS(a, slot)[0].getAttribute("typeface")) }));
  }), fixture.controls.map((_, i) => parts.get(`ppt/slides/slide${i + 1}.xml`).toString()));
  assert.deepEqual(actual, fixture.controls.map(c => c.expectedRuns.map(r => ({ text: r.text, slots: [r.family, r.family, r.family] }))));
  const session = await importPptx(result.bytes, { projection: "content", textFontProfile: "script-families-v3" });
  try {
    const document = normalizePresentation(session.document);
    for (const [i, control] of fixture.controls.entries()) {
      const runs = document.slides[i].elements.find(e => e.name === "Probe").content[0].runs;
      assert.deepEqual(runs.map(r => [r.text, r.style.fontFamily, r.style.eastAsianFontFamily, r.style.complexScriptFontFamily]), control.expectedRuns.map(r => [r.text, r.family, r.family, r.family]));
      assert.ok(runs.every(r => !Object.hasOwn(r.style, "scriptFontFamilies")));
    }
    const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.deepEqual((await session.save(prepared.plan)).bytes, result.bytes);
  } finally { session.dispose(); }
});

test("one original paragraph materializes all language-selected Han and RTL families without recovering maps", async t => {
  const { createRealFontCombinedFixture } = await import("../tests/fixtures/real-font-combined.mjs");
  const { nativeFontFamilies } = await import("./research/font-native-cases.mjs");
  const fixture = createRealFontCombinedFixture(createPresentation, nativeFontFamilies), before = normalizePresentation(fixture.document);
  const probe = before.slides[0].elements.find(e => e.name === "Probe"); assert.equal(probe.content.length, 1); assert.equal(probe.content[0].runs.length, 8);
  const generated = await generatePptx(fixture.document), parts = readGeneratedZip(generated.bytes);
  assert.deepEqual(normalizePresentation(fixture.document), before);
  assert.equal(generated.warnings.length, 1); assert.equal(generated.warnings[0].code, "script-font-map-materialized");
  assert.equal(generated.warnings[0].slideId, "combined-scripts"); assert.equal(generated.warnings[0].elementId, "probe-combined");
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const actual = await page.evaluate(xml => {
    const dom = new DOMParser().parseFromString(xml, "application/xml"), p = "http://schemas.openxmlformats.org/presentationml/2006/main", a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const probe = [...dom.getElementsByTagNameNS(p, "sp")].find(s => s.getElementsByTagNameNS(p, "cNvPr")[0].getAttribute("name") === "Probe");
    return { paragraphs: probe.getElementsByTagNameNS(a, "p").length, runs: [...probe.getElementsByTagNameNS(a, "r")].map(r => ({ text: r.getElementsByTagNameNS(a, "t")[0].textContent, language: r.getElementsByTagNameNS(a, "rPr")[0].getAttribute("lang"), slots: ["latin", "ea", "cs"].map(slot => r.getElementsByTagNameNS(a, slot)[0].getAttribute("typeface")) })) };
  }, parts.get("ppt/slides/slide1.xml").toString());
  assert.deepEqual(actual, { paragraphs: 1, runs: fixture.expectedRuns.map(r => ({ text: r.text, language: r.language, slots: [r.family, r.family, r.family] })) });
  const session = await importPptx(generated.bytes, { projection: "content", textFontProfile: "script-families-v3" });
  try {
    const runs = normalizePresentation(session.document).slides[0].elements.find(e => e.name === "Probe").content[0].runs;
    assert.deepEqual(runs.map(r => [r.text, r.style.language, r.style.fontFamily]), fixture.expectedRuns.map(r => [r.text, r.language, r.family]));
    assert.ok(runs.every(r => !Object.hasOwn(r.style, "scriptFontFamilies")));
    const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready"); assert.deepEqual((await session.save(prepared.plan)).bytes, generated.bytes);
  } finally { session.dispose(); }
});
