import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createScriptFontFixture, scriptStyle, scriptText } from "../tests/fixtures/script-fonts.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const moduleUrl = s => `data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;

test("Node/ESM/global script-font authoring, selection, layout and output remain identical", async t => {
  const p = createScriptFontFixture(createPresentation), canonical = normalizePresentation(p), layout = resolveLayout(p), output = await generatePptx(p), preview = await renderPresentationToSvg(p);
  const expected = { canonical, layout, bytes: [...output.bytes], warnings: output.warnings, preview, segments: resolveTextFontSegments(scriptText, scriptStyle, canonical.theme) };
  const browser = await chromium.launch(); t.after(() => browser.close());
  const urls = {};
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) {
    let source = await readFile(`packages/${name}/dist/browser/index.js`, "utf8");
    for (const [dependency, url] of Object.entries(urls)) source = source.replaceAll(`"@pptkit/${dependency}"`, JSON.stringify(url));
    urls[name] = moduleUrl(source);
  }
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
    const actual = await page.evaluate(async ({ urls, factory, text, style }) => {
      const c = urls ? await import(urls.core) : PPTKit, l = urls ? await import(urls.layout) : PPTKit, e = urls ? await import(urls["pptx-exporter"]) : PPTKit, s = urls ? await import(urls["svg-renderer"]) : PPTKit;
      // This factory is our checked-in test fixture, never source document code.
      const p = Function("createPresentation", "scriptStyle", "scriptText", `return (${factory})(createPresentation)`)(c.createPresentation, style, text);
      const canonical = c.normalizePresentation(p), layout = l.resolveLayout(p), output = await e.generatePptx(p), preview = await s.renderPresentationToSvg(p);
      return { canonical, layout, bytes: [...output.bytes], warnings: output.warnings, preview, segments: c.resolveTextFontSegments(text, style, canonical.theme) };
    }, { urls: format === "esm" ? urls : null, factory: createScriptFontFixture.toString(), text: scriptText, style: scriptStyle });
    assert.deepEqual(actual, expected); assert.deepEqual(requests, []); await page.close();
  }
});

test("independent browser DOM checks native run slots and both SVG text paths without splitting authored runs", async t => {
  const p = createScriptFontFixture(createPresentation), before = normalizePresentation(p), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (let round = 0; round < 3; round++) {
    const output = await generatePptx(p), parts = readGeneratedZip(output.bytes), svg = (await renderPresentationToSvg(p)).slides[0].svg;
    const result = await page.evaluate(({ slide, notes, svg }) => {
      const ns = "http://schemas.openxmlformats.org/drawingml/2006/main", parser = new DOMParser(), source = parser.parseFromString(slide, "application/xml"), note = parser.parseFromString(notes, "application/xml"), preview = parser.parseFromString(svg, "image/svg+xml");
      const slots = doc => [...doc.getElementsByTagNameNS(ns, "r")].filter(r => r.getElementsByTagNameNS(ns, "t")[0]?.textContent === "Latin中文العربية").map(r => ["latin", "ea", "cs"].map(n => r.getElementsByTagNameNS(ns, n)[0]?.getAttribute("typeface")));
      const spans = [...preview.getElementsByTagName("tspan")].filter(n => n.textContent).map(n => [n.textContent, n.getAttribute("font-family").slice(1, -1)]);
      const html = [...preview.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "span")].filter(n => n.textContent && n.hasAttribute("lang")).map(n => [n.textContent, n.style.fontFamily]);
      return { slots: slots(source), notes: slots(note), spans, html, links: preview.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "a").length };
    }, { slide: parts.get("ppt/slides/slide1.xml").toString(), notes: parts.get("ppt/notesSlides/notesSlide1.xml").toString(), svg });
    assert.equal(result.slots.length, 5); assert.deepEqual(result.slots.slice(0, 4), Array(4).fill(["Latin Face", "Eastern Face", "Complex Face"]));
    assert.deepEqual(result.slots[4], ["Latin Face", "+mj-lt", "Latin Face"]); assert.deepEqual(result.notes, [["Latin Face", "Eastern Face", "Complex Face"]]);
    for (const pair of [["Latin", "Latin Face"], ["中文", "Eastern Face"], ["العربية", "Complex Face"]]) assert.ok(result.spans.some(v => v[0] === pair[0] && v[1] === pair[1]));
    for (const pair of [["Latin", "Latin Face"], ["中文", "Eastern Face"], ["العربية", "Complex Face"]]) assert.ok(result.html.some(v => v[0] === pair[0] && v[1].replaceAll('"', "") === pair[1]));
    assert.equal(result.links, 1); assert.deepEqual(normalizePresentation(p), before);
  }
});

test("source no-op retains all script slots while current semantic import keeps its explicit admission boundary", async () => {
  let bytes = (await generatePptx(createScriptFontFixture(createPresentation))).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    try {
      assert.ok(s.report.textStyleInheritance.properties.some(p => p.property === "run.fontFamily" && p.status === "unsupported"));
      const plan = await s.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan); assert.deepEqual(saved.bytes, bytes);
      const slide = s.document.slides[0], original = slide.elements[0], next = structuredClone(original); next.content[0].runs[0].style.eastAsianFontFamily = "Different"; slide.replaceElement(original.id, next);
      assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(original.id, original); assert.equal((await s.prepareSave({ mode: "no-op" })).status, "ready"); bytes = saved.bytes;
    } finally { s.dispose(); }
  }
});

test("font CSS strings cannot turn control characters or quotes into new style declarations", async t => {
  const p = createPresentation(), runStyle = { fontFamily: 'Bad";color:red;/*\nName', eastAsianFontFamily: 'Eastern";color:red;/*', complexScriptFontFamily: 'Complex";color:red;/*', color: "123456" };
  p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 300, height: 80 }, content: [{ runs: [{ text: "A中ع", style: runStyle, action: { type: "hyperlink", url: "https://example.invalid" } }] }] });
  const svg = (await renderPresentationToSvg(p)).slides[0].svg, browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  const colors = await page.evaluate(svg => {
    const doc = new DOMParser().parseFromString(svg, "image/svg+xml");
    return [...doc.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "span")].filter(s => s.hasAttribute("lang")).map(s => s.style.color);
  }, svg);
  assert.deepEqual(colors, ["rgb(18, 52, 86)", "rgb(18, 52, 86)", "rgb(18, 52, 86)"]);
});
