import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createZip } from "../packages/pptx-exporter/dist/archive/create-zip.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const specimen = {
  type: "chart", chartType: "bubble", id: "sample", title: "Bubble size probe",
  box: { x: 40, y: 30, width: 640, height: 345 },
  axes: { x: { scale: { min: 0, max: 9, majorUnit: 1 } }, y: { scale: { min: 0, max: 8, majorUnit: 1 } } },
  series: [
    { name: "Blue", color: "#4472C4", xValues: [1, 4], yValues: [2, 4], sizes: [1, 9] },
    { name: "Orange", color: "#ED7D31", xValues: [7], yValues: [6], sizes: [4] },
  ],
};
const author = element => {
  const deck = createPresentation({ size: { width: 720, height: 405 } });
  deck.addSlide({ elements: [element] });
  return deck;
};

test("bubble authoring rejects inconsistent triples before export", async () => {
  for (const sizes of [[0, 9], [-1, 9], [1], [1, NaN], [Infinity, 9], Array(2)]) {
    const element = structuredClone(specimen);
    element.series[0].sizes = sizes;
    const deck = author(element);
    assert.ok(validatePresentation(deck).some(item => item.code === "invalid-bubble-data" && item.path.endsWith(".series")));
    await assert.rejects(generatePptx(deck), error => error.diagnostics?.some(item => item.code === "invalid-bubble-data"));
  }
});

test("bubble chart data, numeric axes, clipped preview and editable workbook agree across Node and Chromium", async t => {
  const deck = author(specimen);
  assert.deepEqual(validatePresentation(deck), []);
  const ir = normalizePresentation(deck);
  assert.notEqual(ir.slides[0].elements[0].series[0].sizes, specimen.series[0].sizes);
  assert.deepEqual(ir.slides[0].elements[0].series[0].sizes, [1, 9]);
  const layout = resolveLayout(deck).slides[0].elements[0].chartLayout;
  const [blue, orange] = layout.bubbles;
  const [small, large] = blue.points, [medium] = orange.points;
  assert.ok(small.x < large.x && large.x < medium.x);
  assert.ok(small.y > large.y && large.y > medium.y);
  assert.ok(Math.abs(large.radius / small.radius - 3) < 1e-10);
  assert.ok(Math.abs(medium.radius / small.radius - 2) < 1e-10);
  const rendered = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.equal([...rendered.matchAll(/data-chart-role="bubble"/g)].length, 3);
  assert.match(rendered, /clip-path="url\(#[^"]+\)" data-chart-layer="series"/);
  const pptx = await generatePptx(deck);
  const parts = readGeneratedZip(pptx.bytes);
  const xml = parts.get("ppt/charts/chart1.xml").toString();
  assert.match(xml, /<c:bubbleChart>.*<c:sizeRepresents val="area"\/><c:axId val="1"\/><c:axId val="2"\/><\/c:bubbleChart>/s);
  assert.equal([...xml.matchAll(/<c:valAx>/g)].length, 2);
  assert.match(xml, /<c:bubbleSize><c:numRef><c:f>Sheet1!\$C\$2:\$C\$3<\/c:f>/);
  const workbook = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
  assert.match(workbook.get("xl/worksheets/sheet1.xml").toString(), /<c r="C3"><v>9<\/v><\/c>/);
  assert.match(parts.get("ppt/charts/_rels/chart1.xml.rels").toString(), /Microsoft_Excel_Worksheet1.xlsx/);

  const imported = await importPptx(pptx.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v2" });
  try {
    const slide = imported.document.slides[0], chart = slide.elements[0];
    assert.equal(chart.chartType, "bubble");
    assert.deepEqual(chart.series.map(s => [s.xValues, s.yValues, s.sizes]), [[[1, 4], [2, 4], [1, 9]], [[7], [6], [4]]]);
    assert.deepEqual(imported.report.chartValues, []);
    const original = await imported.prepareSave({ mode: "no-op" });
    assert.equal(original.status, "ready");
    assert.deepEqual((await imported.save(original.plan)).bytes, pptx.bytes);
    const changed = structuredClone(chart);
    changed.series[0].sizes[1] = 16;
    slide.replaceElement(chart.id, changed);
    assert.equal((await imported.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { imported.dispose(); }

  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer", "pptx-parser"]) {
    await page.addScriptTag({ content: await readFile(new URL(`../packages/${name}/dist/browser/global.js`, import.meta.url), "utf8") });
  }
  const actual = await page.evaluate(async element => {
    const deck = PPTKit.createPresentation({ size: { width: 720, height: 405 } });
    deck.addSlide({ elements: [element] });
    const svg = (await PPTKit.renderPresentationToSvg(deck)).slides[0].svg;
    const result = await PPTKit.generatePptx(deck);
    const imported = await PPTKit.importPptx(result.bytes, { projection: "content" });
    try { return { svg, bytes: [...result.bytes], sizes: imported.document.slides[0].elements[0].series.map(item => item.sizes) }; }
    finally { imported.dispose(); }
  }, specimen);
  assert.equal(actual.svg, rendered);
  assert.deepEqual(actual.bytes, [...pptx.bytes]);
  assert.deepEqual(actual.sizes, [[1, 9], [4]]);
});

test("LibreOffice bubble padding projects without workbook and survives three fresh rounds", async () => {
  let bytes = new Uint8Array(await readFile(new URL("../tests/fixtures/bubble-libreoffice-resave.pptx", import.meta.url)));
  const evidence = JSON.parse(await readFile(new URL("../docs/testing/bubble-public-probe-evidence.json", import.meta.url), "utf8"));
  assert.equal(createHash("sha256").update(bytes).digest("hex"), evidence.saved.sha256);
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content" });
    try {
      const chart = session.document.slides[0].elements[0];
      assert.equal(chart.chartType, "bubble");
      assert.deepEqual(chart.series.map(s => [s.name, s.xValues, s.yValues, s.sizes]),
        [["Blue", [1, 4], [2, 4], [1, 9]], ["Orange", [7], [6], [4]]]);
      assert.equal(session.report.findings.some(f => f.reason.includes("Aligned trailing blank positions")), round === 0);
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      bytes = (await generatePptx(session.document)).bytes;
      assert.ok(readGeneratedZip(bytes).has("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
    } finally { session.dispose(); }
  }
});

test("bubble import keeps unsupported size semantics and partial triples source-only", async () => {
  const source = await generatePptx(author(specimen));
  const parts = readGeneratedZip(source.bytes), name = "ppt/charts/chart1.xml", chart = parts.get(name).toString();
  const patchSize = (xml, fn) => xml.replace(/<c:bubbleSize>[\s\S]*?<\/c:bubbleSize>/, fn);
  const mutations = [
    chart.replace('sizeRepresents val="area"', 'sizeRepresents val="w"'),
    chart.replace('bubbleScale val="100"', 'bubbleScale val="200"'),
    chart.replace('bubble3D val="0"', 'bubble3D val="1"'),
    patchSize(chart, xml => xml.replace("<c:v>9</c:v>", "<c:v>0</c:v>")),
    patchSize(chart, xml => xml.replace(/<c:pt idx="0">[\s\S]*?<\/c:pt>/, "")),
    patchSize(chart, xml => xml.replace(/<c:pt idx="1">[\s\S]*?<\/c:pt>/, "")),
  ];
  for (const altered of mutations) {
    assert.notEqual(altered, chart);
    const bytes = createZip([...parts].map(([part, data]) => ({ name: part, data: part === name ? new TextEncoder().encode(altered) : data })));
    const session = await importPptx(bytes, { projection: "content" });
    try {
      assert.deepEqual(session.document.slides[0].elements, []);
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
    } finally { session.dispose(); }
  }
  await assert.rejects(importPptx(source.bytes, { projection: "content", sessionLimits: { maxChartDataPoints: 2 } }), /maxChartDataPoints/);
});

test("bubble workbook fallback fills an absent cache and never overrides a complete one", async () => {
  const source = await generatePptx(author(specimen));
  const parts = readGeneratedZip(source.bytes), name = "ppt/charts/chart1.xml", chart = parts.get(name).toString();
  for (const [mutate, expected, finding] of [
    [xml => xml.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, ""), [1, 9], "Missing chart cache was mapped"],
    [xml => xml.replace("<c:v>9</c:v>", "<c:v>16</c:v>"), [1, 16], "Chart cache conflicts"],
  ]) {
    const altered = chart.replace(/<c:bubbleSize>[\s\S]*?<\/c:bubbleSize>/, mutate);
    assert.notEqual(altered, chart);
    const bytes = createZip([...parts].map(([part, data]) => ({ name: part, data: part === name ? new TextEncoder().encode(altered) : data })));
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v3" });
    try {
      assert.deepEqual(session.document.slides[0].elements[0].series[0].sizes, expected);
      assert.ok(session.report.findings.some(f => f.reason.includes(finding)));
      const bindings = session.report.chartValues.filter(binding => binding.seriesIndex === 0);
      assert.equal(bindings.length, 6);
      assert.ok(bindings.every(binding => binding.status === "unmapped"));
    } finally { session.dispose(); }
  }
});

test("bubble source-save v3 edits independent X/Y/size atoms over three rounds", async t => {
  let bytes = (await generatePptx(author(specimen))).bytes;
  const options = { projection: "content", chartEditSaveProfile: "existing-values-v3" };
  const chartName = "ppt/charts/chart1.xml", bookName = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const sheetName = "xl/worksheets/sheet1.xml";
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const targets = session.report.chartValues;
      assert.equal(targets.length, 9);
      assert.ok(targets.every(binding => binding.status === "bound"));
      const coordinates = ["x", "y", "size"];
      assert.deepEqual(coordinates.map(coordinate =>
        targets.find(b => b.seriesIndex === 0 && b.pointIndex === 1 && b.coordinate === coordinate).workbook.cellReference),
      ["A3", "B3", "C3"]);
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
      const values = [-2 - round, 5.25 + round, 16 + round];
      const prior = [chart.series[0].xValues[1], chart.series[0].yValues[1], chart.series[0].sizes[1]];
      const doc = session.document, changes = ['xValues', 'yValues', 'sizes'].map((coordinate, index) => ({ type: 'chart-point', slideId: slide.id, elementId: chart.id, seriesIndex: 0, pointIndex: 1, coordinate, expected: prior[index], value: values[index] }));
      const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes }).inverse;
      const redo = doc.applyEditOperation(undo).inverse; doc.applyEditOperation(redo);
      const candidate = await session.prepareSave({ mode: "edit" });
      assert.equal(candidate.status, "ready", JSON.stringify(candidate.report?.diagnostics));
      const saved = await session.save(candidate.plan), output = saved.bytes;
      assert.deepEqual(saved.report.changes.map(change => change.chartValue.coordinate).sort(), ["size", "x", "y"]);
      if (round === 0) {
        const browser = await chromium.launch({ headless: true });
        t.after(() => browser.close());
        const page = await browser.newPage();
        for (const name of ["core", "pptx-parser"]) {
          await page.addScriptTag({ content: await readFile(new URL(`../packages/${name}/dist/browser/global.js`, import.meta.url), "utf8") });
        }
        const browserBytes = await page.evaluate(async ({ bytes, options, values }) => {
          const session = await PPTKit.importPptx(Uint8Array.from(bytes), options);
          try {
            const doc = session.document, slide = doc.slides[0], chart = slide.elements[0];
            const changes = ['xValues', 'yValues', 'sizes'].map((coordinate, index) => ({ type: 'chart-point', slideId: slide.id, elementId: chart.id, seriesIndex: 0, pointIndex: 1, coordinate, expected: chart.series[0][coordinate][1], value: values[index] }));
            const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes }).inverse;
            const redo = doc.applyEditOperation(undo).inverse; doc.applyEditOperation(redo);
            const candidate = await session.prepareSave({ mode: "edit" });
            if (candidate.status !== "ready") throw new Error(JSON.stringify(candidate.report));
            return [...(await session.save(candidate.plan)).bytes];
          } finally { session.dispose(); }
        }, { bytes: [...bytes], options, values });
        assert.deepEqual(browserBytes, [...output]);
      }
      const before = readGeneratedZip(bytes), after = readGeneratedZip(output);
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, data] of before) if (name !== chartName && name !== bookName) assert.deepEqual(after.get(name), data, name);
      let expectedChart = before.get(chartName).toString();
      let expectedSheet = readGeneratedZip(before.get(bookName)).get(sheetName).toString();
      for (const [i, [tag, column]] of [["xVal", "A"], ["yVal", "B"], ["bubbleSize", "C"]].entries()) {
        const start = expectedChart.indexOf(`<c:${tag}>`), end = expectedChart.indexOf(`</c:${tag}>`, start);
        const prefix = expectedChart.slice(start, end);
        const next = prefix.replace(`<c:pt idx="1"><c:v>${prior[i]}</c:v></c:pt>`, `<c:pt idx="1"><c:v>${values[i]}</c:v></c:pt>`);
        assert.notEqual(prefix, next);
        expectedChart = expectedChart.slice(0, start) + next + expectedChart.slice(end);
        expectedSheet = expectedSheet.replace(`<c r="${column}3"><v>${prior[i]}</v></c>`, `<c r="${column}3"><v>${values[i]}</v></c>`);
      }
      assert.equal(after.get(chartName).toString(), expectedChart);
      const oldBook = readGeneratedZip(before.get(bookName)), newBook = readGeneratedZip(after.get(bookName));
      assert.equal(newBook.get(sheetName).toString(), expectedSheet);
      for (const [name, data] of oldBook) if (name !== sheetName) assert.deepEqual(newBook.get(name), data, name);
      const reopened = await importPptx(output, options);
      try { assert.deepEqual(reopened.document.slides[0].elements[0].series, slide.elements[0].series); }
      finally { reopened.dispose(); }
      bytes = output;
    } finally { session.dispose(); }
  }
});

test("bubble source-save v3 keeps cache-only native output unbound", async () => {
  const source = await readFile(new URL("../tests/fixtures/bubble-libreoffice-resave.pptx", import.meta.url));
  const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v3" });
  try {
    assert.equal(session.report.chartValues.length, 9);
    assert.ok(session.report.chartValues.every(b => b.status === "unmapped" && /internal XLSX/.test(b.reason)));
    const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
    chart.series[0].sizes[0] = 16;
    slide.replaceElement(chart.id, chart);
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { session.dispose(); }
});
