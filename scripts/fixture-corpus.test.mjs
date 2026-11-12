import assert from "node:assert/strict";
import test from "node:test";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveNormalizedLayout } from "../packages/layout/dist/index.js";
import { inspectZipArchive, inspectXmlPart, inspectOpcPackage, inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderLayoutToSvg } from "../packages/svg-renderer/dist/index.js";
import { createFixture, fixtureCatalog } from "../tests/fixtures/presentations.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { inspectXmlPackage } from "../tests/support/inspect-xml-package.mjs";

test("named corpus validates, exports deterministically, and has complete XML relationships", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const fixture of fixtureCatalog) await t.test(fixture.id, async () => {
    const document = createFixture(fixture.id, createPresentation);
    assert.deepEqual(validatePresentation(document), []);
    const normalized = normalizePresentation(document);
    assert.equal(normalized.irVersion, 2);
    assert.deepEqual(normalizePresentation(createFixture(fixture.id, createPresentation)), normalized);
    const before = structuredClone(normalized);
    const layout = resolveNormalizedLayout(normalized);
    assert.deepEqual(normalized, before);
    const generated = await generatePptx(document);
    assert.ok(Buffer.from(generated.bytes).equals(Buffer.from((await generatePptx(document)).bytes)), `${fixture.id}: repeated export is byte-identical`);
    assert.deepEqual(generated.warnings.map(({ code }) => code).sort(), [...fixture.warnings.export].sort());
    const svg = await renderLayoutToSvg(layout);
    assert.deepEqual(svg.warnings.map(({ code }) => code).sort(), [...fixture.warnings.preview].sort());
    assert.equal(svg.slides.length, document.slides.length);
    const entries = readGeneratedZip(generated.bytes);
    const inspected = await inspectZipArchive(generated.bytes);
    const presentation = await inspectPptx(generated.bytes);
    const imported = await importPptx(generated.bytes, { projection: "structure" });
    const prepared = await imported.prepareSave({ mode: "no-op" });
    assert.equal(prepared.status, "ready");
    assert.equal(Buffer.compare(Buffer.from((await imported.save(prepared.plan)).bytes), Buffer.from(generated.bytes)), 0);
    assert.deepEqual(imported.document.slides.map((s) => s.hidden), layout.slides.map((s) => s.hidden));
    imported.dispose();
    assert.equal(presentation.slides.length, document.slides.length);
    assert.deepEqual(presentation.slideSize, { widthEmu: Math.round(layout.size.width * 12700), heightEmu: Math.round(layout.size.height * 12700) });
    assert.deepEqual(presentation.slides.map((slide) => slide.hidden), layout.slides.map((slide) => slide.hidden));
    assert.deepEqual(presentation.findings, []);
    const coreProperties = new Map(presentation.metadata.core.properties.map((p) => [p.name, p.value]));
    assert.equal(coreProperties.get("title"), layout.metadata.title);
    assert.equal(coreProperties.get("creator"), layout.metadata.author);
    assert.equal(coreProperties.get("keywords"), layout.metadata.keywords.join(", "));
    assert.equal(coreProperties.get("revision"), String(layout.metadata.revision));
    const appProperties = new Map(presentation.metadata.extended.properties.map((p) => [p.name, p.value.value]));
    assert.equal(appProperties.get("Application"), "PPTKit");
    assert.equal(appProperties.get("Slides"), String(layout.slides.length));
    // Existing authoring section labels are PPTKit slide metadata, not native sections.
    assert.equal(presentation.sectionList, undefined);
    assert.deepEqual(presentation.presentationExtensions, []);
    assert.equal(presentation.dialect, "transitional");
    assert.equal(presentation.suggestedExtension, "pptx");
    assert.equal(presentation.mainPart, "/ppt/presentation.xml");
    const opc = presentation.package;
    assert.equal(opc.parts.length, entries.size - 1);
    assert.ok(opc.relationships.some((relationship) => relationship.source === "/" && relationship.targetPart === "/ppt/presentation.xml"));
    assert.equal(opc.xmlTotals.parts, [...entries.keys()].filter((name) => /\.(xml|rels)$/.test(name)).length);
    assert.deepEqual(inspected.entries.map((entry) => entry.name).sort(), [...entries.keys()].sort());
    assert.equal(inspected.expandedBytes, [...entries.values()].reduce((sum, value) => sum + value.length, 0));
    const xmlParts = [...entries].filter(([name]) => /\.(xml|rels)$/.test(name));
    const xmlInventories = [];
    for (const [partName, bytes] of xmlParts) {
      const result = await inspectXmlPart(bytes, { partName });
      xmlInventories.push({ root: result.root.name, uri: result.root.namespaceUri, elements: result.elements, attributes: result.attributes });
    }
    const browserInventories = await page.evaluate((parts) => parts.map(([, xml]) => {
      const doc = new DOMParser().parseFromString(xml, "application/xml");
      const elements = [...doc.getElementsByTagName("*")];
      return { root: doc.documentElement.nodeName, uri: doc.documentElement.namespaceURI ?? "", elements: elements.length, attributes: elements.reduce((sum, element) => sum + element.attributes.length, 0) };
    }), xmlParts.map(([name, value]) => [name, value.toString()]));
    assert.deepEqual(xmlInventories, browserInventories, `${fixture.id}: XML namespace/count inventory agrees with Chromium`);
    assert.deepEqual(await page.evaluate(inspectXmlPackage, [...entries].map(([name, bytes]) => [name, bytes.toString()])), []);
    for (const [name, bytes] of entries) if (name.endsWith(".xlsx")) {
      const workbook = readGeneratedZip(bytes);
      const workbookOpc = await inspectOpcPackage(bytes);
      assert.equal(workbookOpc.parts.length, workbook.size - 1);
      for (const [partName, data] of workbook) if (/\.(xml|rels)$/.test(partName)) await inspectXmlPart(data, { partName: `${name}!/${partName}` });
      assert.deepEqual(await page.evaluate(inspectXmlPackage, [...workbook].map(([part, value]) => [part, value.toString()])), [], name);
    }
    const slide = entries.get("ppt/slides/slide1.xml").toString();
    if (fixture.id === "text") {
      assert.match(slide, /Rich &amp; &lt;editable&gt;/);
      assert.match(slide, /<a:buAutoNum[^>]*startAt="3"/);
      assert.match(entries.get("ppt/notesSlides/notesSlide2.xml").toString(), /Speaker &amp; reviewer notes/);
    } else if (fixture.id === "theme-layout") {
      assert.doesNotMatch(slide, /<p:bg>/);
      assert.match(slide, /<p:ph type="title"/);
      assert.match(entries.get("ppt/slides/slide2.xml").toString(), /show="0"/);
      assert.equal(normalized.slides[0].backgroundSource, "layout");
    } else if (fixture.id === "geometry") {
      assert.equal((slide.match(/<p:grpSp>/g) ?? []).length, 2);
      assert.deepEqual(layout.slides[0].elements.find(({ id }) => id === "anchored").start, { x: 220, y: 85 });
    } else if (fixture.id === "images") {
      assert.equal([...entries.keys()].filter((name) => name.startsWith("ppt/media/")).length, 1);
      assert.match(slide, /<a:srcRect/);
      const image = entries.get("ppt/media/pixel.png");
      assert.deepEqual([...image.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
      await page.setContent(`<img src="${document.assets[0].source.value}">`);
      assert.deepEqual(await page.locator("img").evaluate(async (image) => { await image.decode(); return [image.naturalWidth, image.naturalHeight]; }), [8, 4]);
    } else if (fixture.id === "table") {
      assert.match(slide, /gridSpan="3"/);
      assert.match(slide, /rowSpan="2"/);
      assert.match(slide, /vMerge="1"/);
    } else if (fixture.id === "charts") {
      assert.equal([...entries.keys()].filter((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length, 3);
      for (const [index, type] of ["bar", "line", "pie"].entries()) assert.match(entries.get(`ppt/charts/chart${index + 1}.xml`).toString(), new RegExp(`<c:${type}Chart>`));
    }
  });
});

test("concurrent exports with async assets are independent and byte-identical", async () => {
  const document = createFixture("charts", createPresentation);
  const expected = await generatePptx(document);
  const results = await Promise.all(Array.from({ length: 8 }, () => generatePptx(document)));
  for (const result of results) assert.ok(Buffer.from(result.bytes).equals(Buffer.from(expected.bytes)), "concurrent output equals isolated output");
});

test("XML package verification catches malformed parts, dangling references, and content type drift", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  const entries = readGeneratedZip((await generatePptx(createFixture("charts", createPresentation))).bytes);
  for (const [part, replacement, expected] of [
    ["ppt/slides/slide1.xml", "<broken>", /invalid XML/],
    ["ppt/charts/chart1.xml", undefined, /missing target.*chart1|absent part.*chart1/],
    ["ppt/slides/_rels/slide1.xml.rels", undefined, /unresolved r:/],
    ["[Content_Types].xml", undefined, /Missing content types/],
  ]) {
    const corrupted = new Map(entries);
    if (replacement === undefined) corrupted.delete(part);
    else corrupted.set(part, Buffer.from(replacement));
    const issues = await page.evaluate(inspectXmlPackage, [...corrupted].map(([name, bytes]) => [name, bytes.toString()]));
    assert.match(issues.join("\n"), expected);
  }
});
