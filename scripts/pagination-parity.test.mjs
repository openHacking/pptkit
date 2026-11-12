import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, evaluatePresentationField, fieldContextForSlide, normalizePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg, renderLayoutToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const area = { x: 20, y: 30, width: 100, height: 55 };
const options = { tablePagination: [{ elementId: "table", area, headerRows: 1 }] };
const row = (index) => ({ height: index === 0 ? 10 : 20, cells: [{ content: [{ runs: [{ text: `Row ${index}`, action: { type: "slide", slideId: "after" } }] }] }] });
function fixture(manual = false) {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const source = document.addSlide({ id: "source", notes: "Original notes", section: "Data", tags: ["data"] });
  source.addElement({ type: "shape", id: "background", shape: "rect", box: { x: 0, y: 0, width: 2, height: 2 } });
  const pages = manual ? [[0, 1, 2], [0, 3, 4], [0, 5, 6]] : [[0, 1, 2, 3, 4, 5, 6]];
  pages.forEach((indices, i) => {
    const slide = i === 0 ? source : document.addSlide({ id: `pptkit:page:${JSON.stringify(["source", "table", i])}`, section: "Data", tags: ["data"] });
    slide.addElement({ type: "table", name: "table", id: i === 0 ? "table" : `pptkit:table:${JSON.stringify(["source", "table", i])}`, box: { x: manual ? 20 : 10, y: manual ? 30 : 20, width: 100, height: manual ? 50 : 130 }, columns: [100], rows: indices.map(row) });
  });
  document.addSlide({ id: "after" }).addElement({ type: "shape", id: "back", shape: "rect", box: { x: 0, y: 0, width: 10, height: 10 }, action: { type: "slide", slideId: "source" } });
  return document;
}

test("field page contexts follow actual Layout continuation order", () => {
  const source = fixture(), layout = resolveLayout(source, options), ids = layout.slides.map(slide => slide.id);
  assert.equal(ids.length, 4); assert.equal(ids[0], 'source'); assert.equal(ids[3], 'after');
  assert.deepEqual(layout.pagination.fragments.map(fragment => fragment.slideId), ids.slice(0, 3));
  assert.deepEqual(ids.map(slideId => evaluatePresentationField({ id: 'number', kind: 'slide-number' }, fieldContextForSlide(layout, { slideId, firstPageNumber: 5, includeHidden: true })).value), ['5', '6', '7', '8']);
  assert.equal(source.slides.length, 2);
});

test("source-boundary pagination policy agrees between Node and the browser global bundle", async t => {
  const input = { size: { width: 300, height: 200 }, slides: [{ id: "s", elements: [{ id: "table", type: "table", box: { x: 0, y: 0, width: 100, height: 100 }, columns: [100], nativeTableStyleFlags: { lastRow: true, bandedRows: false }, rows: Array.from({ length: 5 }, (_, i) => ({ height: 20, cells: [{ content: String(i) }] })) }] }] };
  const options = { tablePagination: [{ elementId: "table", area: { x: 0, y: 0, width: 100, height: 40 }, nativeRowRegions: "source-boundaries" }] };
  const project = result => ({ fragments: result.pagination.fragments, diagnostics: result.analysis.diagnostics, parts: result.slides.map(slide => ({ first: slide.elements[0].nativeFirstRowStyle ?? null, flags: slide.elements[0].nativeTableStyleFlags })) });
  const expected = project(resolveLayout(createPresentation(input), options));
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const actual = await page.evaluate(({ input, options }) => {
    const result = PPTKit.resolveLayout(PPTKit.createPresentation(input), options);
    return { fragments: result.pagination.fragments, diagnostics: result.analysis.diagnostics, parts: result.slides.map(slide => ({ first: slide.elements[0].nativeFirstRowStyle ?? null, flags: slide.elements[0].nativeTableStyleFlags })) };
  }, { input, options });
  assert.deepEqual(actual, expected);
});

test("implicit native style risk reaches PPTX and SVG without changing source row content", async () => {
  const document = fixture(), before = normalizePresentation(document);
  const layout = { tablePagination: [{ ...options.tablePagination[0], area: { ...area, height: 30 } }] };
  const resolved = resolveLayout(document, layout), [native, preview] = await Promise.all([generatePptx(document, { layout }), renderPresentationToSvg(document, { layout })]);
  assert.deepEqual(resolved.pagination.fragments.map(fragment => fragment.sourceRowIndices), [[0, 1], [0, 2], [0, 3], [0, 4], [0, 5], [0, 6]]);
  for (const warnings of [resolved.analysis.diagnostics, native.warnings, preview.warnings]) assert.equal(warnings.filter(w => w.code === "pagination-native-row-style-reapplied").length, 1);
  const parts = readGeneratedZip(native.bytes);
  for (let index = 0; index < 6; index++) {
    const xml = parts.get(`ppt/slides/slide${index + 1}.xml`).toString();
    assert.match(xml, /<a:tblPr firstRow="1" bandRow="1"/);
    assert.match(xml, new RegExp(`Row ${index + 1}`));
    assert.match(preview.slides[index].svg, new RegExp(`Row ${index + 1}`));
  }
  assert.deepEqual(normalizePresentation(document), before);
});

test("pagination keeps materialized row styles but repeats native row-style switches on each fragment", async () => {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  document.addSlide({ id: "styled" }).addElement({
    id: "styled-table", type: "table", box: { x: 0, y: 0, width: 100, height: 90 }, columns: [100],
    nativeFirstRowStyle: true, nativeTableStyleFlags: { lastRow: true, bandedRows: true },
    style: {
      firstRow: { fill: { type: "solid", color: "112233" } },
      lastRow: { fill: { type: "solid", color: "445566" } },
      bandedRows: { first: { fill: { type: "solid", color: "AA0000" } }, second: { fill: { type: "solid", color: "00AA00" } } },
    },
    rows: Array.from({ length: 5 }, (_, index) => ({ height: index ? 20 : 10, cells: [{ content: `Row ${index}` }] })),
  });
  const layout = { tablePagination: [{ elementId: "styled-table", area: { x: 0, y: 0, width: 100, height: 30 }, headerRows: 1 }] };
  const resolved = resolveLayout(document, layout), [native, preview] = await Promise.all([generatePptx(document, { layout }), renderPresentationToSvg(document, { layout })]);
  assert.equal(resolved.analysis.diagnostics.filter(d => d.code === "pagination-native-row-style-reapplied").length, 1);
  assert.equal(native.warnings.filter(w => w.code === "pagination-native-row-style-reapplied").length, 1);
  assert.equal(preview.warnings.filter(w => w.code === "pagination-native-row-style-reapplied").length, 1);
  assert.deepEqual(resolved.pagination.fragments.map(f => f.sourceRowIndices), [[0, 1], [0, 2], [0, 3], [0, 4]]);
  const expected = ["AA0000", "00AA00", "AA0000", "445566"], entries = readGeneratedZip(native.bytes);
  for (let index = 0; index < expected.length; index++) {
    const table = resolved.slides[index].elements[0], xml = entries.get(`ppt/slides/slide${index + 1}.xml`).toString();
    assert.equal(table.rows[0].cells[0].style.fill.color, "112233");
    assert.equal(table.rows[1].cells[0].style.fill.color, expected[index]);
    assert.match(xml, /<a:tblPr firstRow="1" bandRow="1" lastRow="1"/);
    assert.match(xml, new RegExp(`<a:srgbClr val="${expected[index]}"`));
    assert.match(preview.slides[index].svg, new RegExp(`#${expected[index]}`, "i"));
  }
  const boundaryLayout = { tablePagination: [{ ...layout.tablePagination[0], nativeRowRegions: "source-boundaries" }] };
  const guarded = await generatePptx(document, { layout: boundaryLayout }), guardedParts = readGeneratedZip(guarded.bytes);
  assert.equal(guarded.warnings.filter(w => w.code === "pagination-native-row-style-reapplied").length, 1);
  for (let index = 0; index < expected.length; index++) {
    const xml = guardedParts.get(`ppt/slides/slide${index + 1}.xml`).toString();
    assert.match(xml, /<a:tblPr firstRow="1" bandRow="1"/);
    assert.equal(/<a:tblPr[^>]* lastRow="1"/.test(xml), index === expected.length - 1);
    assert.match(xml, new RegExp(`<a:srgbClr val="${expected[index]}"`));
  }
  const noHeaders = { tablePagination: [{ ...layout.tablePagination[0], headerRows: 0, area: { x: 0, y: 0, width: 100, height: 40 }, nativeRowRegions: "source-boundaries" }] };
  const noHeaderOutput = await generatePptx(document, { layout: noHeaders }), noHeaderParts = readGeneratedZip(noHeaderOutput.bytes);
  assert.equal(noHeaderOutput.slideCount, 3);
  for (let index = 0; index < 3; index++) {
    const xml = noHeaderParts.get(`ppt/slides/slide${index + 1}.xml`).toString();
    assert.match(xml, new RegExp(`<a:tblPr firstRow="${index === 0 ? 1 : 0}" bandRow="1"`));
    assert.equal(/<a:tblPr[^>]* lastRow="1"/.test(xml), index === 2);
  }
});

test("pagination matches manually split native PPTX and exact-cell SVG, including reordered slide links", async () => {
  const source = fixture(), before = normalizePresentation(source);
  const manual = fixture(true);
  const expectedLayout = resolveLayout(manual);
  // Independently authored three-row geometry; do not call the pagination or cell-grid helper.
  for (const slide of expectedLayout.slides.slice(0, 3)) {
    const table = slide.elements.find((element) => element.type === "table");
    table.cellLayout = [0, 1, 2].map((rowIndex) => ({ rowIndex, cellIndex: 0, columnIndex: 0, box: { x: 0, y: rowIndex === 0 ? 0 : rowIndex === 1 ? 10 : 30, width: 100, height: rowIndex === 0 ? 10 : 20 } }));
  }
  const [actual, expected, svg, expectedSvg] = await Promise.all([
    generatePptx(source, { layout: options }), generatePptx(manual),
    renderPresentationToSvg(source, { layout: options }), renderLayoutToSvg(expectedLayout),
  ]);
  assert.deepEqual(actual.warnings, []); assert.equal(actual.slideCount, 4);
  const actualParts = readGeneratedZip(actual.bytes), expectedParts = readGeneratedZip(expected.bytes);
  for (const [name, data] of actualParts) {
    const expectedData = expectedParts.get(name);
    if (!data.equals(expectedData)) {
      const a = data.toString(), b = expectedData.toString();
      let i = 0; while (i < Math.min(a.length, b.length) && a[i] === b[i]) i++;
      assert.fail(`${name} differs at ${i}: actual=${a.slice(Math.max(0, i - 80), i + 180)} expected=${b.slice(Math.max(0, i - 80), i + 180)}`);
    }
  }
  assert.deepEqual(actual.bytes, expected.bytes); assert.deepEqual(svg, expectedSvg);
  const entries = readGeneratedZip(actual.bytes);
  for (let i = 1; i <= 3; i++) {
    const xml = entries.get(`ppt/slides/slide${i}.xml`).toString();
    assert.deepEqual([...xml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]) / 12700), [10, 20, 20]);
    assert.match(entries.get(`ppt/slides/_rels/slide${i}.xml.rels`).toString(), /Target="slide4.xml"/);
  }
  assert.match(entries.get("ppt/slides/_rels/slide4.xml.rels").toString(), /Target="slide1.xml"/);
  assert.deepEqual(normalizePresentation(source), before);
});

test("failed pagination keeps native bytes and preview and forwards failure diagnostics", async () => {
  const document = fixture();
  const layout = { tablePagination: [{ ...options.tablePagination[0], maxPages: 1 }] };
  const [actual, plain, svg, plainSvg] = await Promise.all([generatePptx(document, { layout }), generatePptx(document), renderPresentationToSvg(document, { layout }), renderPresentationToSvg(document)]);
  assert.deepEqual(actual.bytes, plain.bytes); assert.deepEqual(svg.slides, plainSvg.slides);
  assert.deepEqual(svg.warnings, [...actual.warnings, ...plainSvg.warnings]); assert.equal(actual.warnings[0].code, "pagination-page-limit");
});

test("merged two-row headers and merged body blocks retain native merges after splitting", async () => {
  function merged(manual) {
    const document = createPresentation({ size: { width: 300, height: 200 } });
    const rows = [
      { height: 10, cells: [{ content: "Header merged", rowSpan: 2 }, { content: "H1" }] },
      { height: 10, cells: [{ content: "H2" }] },
      { height: 20, cells: [{ content: "Body A merged", rowSpan: 2 }, { content: "A1" }] },
      { height: 20, cells: [{ content: "A2" }] },
      { height: 20, cells: [{ content: "Body B merged", rowSpan: 2 }, { content: "B1" }] },
      { height: 20, cells: [{ content: "B2" }] },
    ];
    const pages = manual ? [[0, 1, 2, 3], [0, 1, 4, 5]] : [[0, 1, 2, 3, 4, 5]];
    pages.forEach((indices, i) => {
      const slideId = i === 0 ? "s" : `pptkit:page:${JSON.stringify(["s", "t", i])}`;
      const elementId = i === 0 ? "t" : `pptkit:table:${JSON.stringify(["s", "t", i])}`;
      document.addSlide({ id: slideId }).addElement({ type: "table", id: elementId, name: "t", box: { x: 20, y: 30, width: 100, height: manual ? 60 : 100 }, columns: [50, 50], rows: indices.map((index) => rows[index]) });
    });
    return document;
  }
  const layout = { tablePagination: [{ elementId: "t", area: { x: 20, y: 30, width: 100, height: 65 }, headerRows: 2 }] };
  const [actual, expected] = await Promise.all([generatePptx(merged(false), { layout }), generatePptx(merged(true))]);
  assert.deepEqual(actual.warnings.map(w => w.code), ["pagination-native-row-style-reapplied"]); assert.equal(actual.slideCount, 2);
  assert.deepEqual(actual.bytes, expected.bytes);
  const entries = readGeneratedZip(actual.bytes);
  for (let i = 1; i <= 2; i++) {
    const xml = entries.get(`ppt/slides/slide${i}.xml`).toString();
    assert.equal([...xml.matchAll(/ vMerge="1"/g)].length, 2);
    assert.equal([...xml.matchAll(/ rowSpan="2"/g)].length, 2);
    assert.deepEqual([...xml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]) / 12700), [10, 10, 20, 20]);
  }
});
