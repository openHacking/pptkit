import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { inflateRawSync } from "node:zlib";

import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { generatePptx as generateNodePptx, writePptx } from "../dist/node.js";

function readZipEntries(input) {
  const bytes = Buffer.from(input);
  const entries = new Map();
  let offset = 0;
  while (offset + 30 <= bytes.length && bytes.readUInt32LE(offset) === 0x04034b50) {
    const nameLength = bytes.readUInt16LE(offset + 26);
    const extraLength = bytes.readUInt16LE(offset + 28);
    const compressedSize = bytes.readUInt32LE(offset + 18);
    const method = bytes.readUInt16LE(offset + 8);
    const name = bytes.subarray(offset + 30, offset + 30 + nameLength).toString();
    const data = bytes.subarray(offset + 30 + nameLength + extraLength, offset + 30 + nameLength + extraLength + compressedSize);
    entries.set(name, method === 8 ? inflateRawSync(data) : data);
    offset += 30 + nameLength + extraLength + compressedSize;
  }
  return entries;
}

test("generatePptx exports the IR v2 feature surface", async () => {
  const presentation = createPresentation({
    metadata: { title: "Q1 & <ready>", author: "PPTKit", company: "Example", language: "zh-CN" },
    theme: { colors: { accent1: "123456" }, fonts: { heading: "Arial", body: "Microsoft YaHei" } },
  });
  presentation.defineSlideLayout({
    id: "title-content",
    name: "Title & Content",
    background: { type: "solid", color: { theme: "background2" } },
    elements: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 12, height: 540 }, style: { fill: { type: "solid", color: { theme: "accent1" } } } }],
    placeholders: [{ key: "title", kind: "title", box: { x: 40, y: 30, width: 700, height: 60 }, textStyle: { run: { fontSize: 36, fontFamily: { theme: "heading" } } } }],
  });
  const asset = presentation.registerAsset({ kind: "image", source: { type: "url", value: "https://example.com/hero.png" }, width: 400, height: 200, accessibility: { description: "Hero" } });
  const first = presentation.addSlide({
    id: "intro",
    layoutId: "title-content",
    notes: "Speaker note",
    section: "Opening",
    tags: ["hero"],
    customData: { source: "unit-test" },
  });
  first.addElement({
    type: "text",
    content: [{
      style: { bullet: { type: "bullet", character: "→" }, lineSpacing: 1.15 },
      runs: [
        { text: "Q1 & <ready> ", style: { bold: true, color: { theme: "accent1" } }, action: { type: "url", url: "https://example.com", tooltip: "Visit" } },
        { text: "next", style: { italic: true }, action: { type: "slide", slideId: "details" } },
      ],
    }],
    placeholderKey: "title",
    frame: { verticalAlign: "middle", autoFit: { mode: "shrink", fontScale: 0.95 } },
  });
  first.addElement({ type: "shape", id: "left", shape: "roundRect", box: { x: 40, y: 130, width: 120, height: 60 }, style: { fill: { type: "solid", color: "FF0000", opacity: 0.8 } } });
  first.addElement({ type: "shape", id: "right", shape: "ellipse", box: { x: 300, y: 130, width: 80, height: 80 } });
  first.addElement({ type: "connector", start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" }, route: [{ x: 220, y: 150 }], style: { dash: "dash", endArrow: "triangle" } });
  first.addElement({
    type: "group",
    box: { x: 420, y: 120, width: 240, height: 140 },
    coordinateSize: { width: 240, height: 140 },
    children: [{ type: "image", assetId: asset.id, fit: "cover", box: { x: 0, y: 0, width: 240, height: 140 } }],
  });
  first.addElement({
    type: "table",
    box: { x: 40, y: 300, width: 400, height: 120 },
    columns: [200, 200],
    rows: [
      { height: 50, cells: [{ content: "Header", colSpan: 2, style: { fill: { type: "solid", color: { theme: "accent1" } } } }] },
      { height: 70, cells: [{ content: "A" }, { content: "B" }] },
    ],
  });
  presentation.addSlide({ id: "details", hidden: true, background: { type: "solid", color: "FAFAFA" }, elements: [{ type: "text", content: "Details", box: { x: 40, y: 40, width: 300, height: 50 } }] });

  const originalFetch = globalThis.fetch;
  globalThis.fetch = async () => new Response(Buffer.from("fake-png"), { status: 200, headers: { "content-type": "image/png" } });
  try {
    const result = await generatePptx(presentation);
    const repeated = await generatePptx(presentation);
    assert.deepEqual(result.bytes, repeated.bytes);
    assert.equal(result.status, "generated");
    assert.equal(result.slideCount, 2);
    const entries = readZipEntries(result.bytes);
    const slide1 = entries.get("ppt/slides/slide1.xml").toString();
    const slide2 = entries.get("ppt/slides/slide2.xml").toString();
    const slideRels = entries.get("ppt/slides/_rels/slide1.xml.rels").toString();
    const layout2 = entries.get("ppt/slideLayouts/slideLayout2.xml").toString();
    const theme = entries.get("ppt/theme/theme1.xml").toString();
    const notes = entries.get("ppt/notesSlides/notesSlide1.xml").toString();

    assert.match(entries.get("docProps/core.xml").toString(), /Q1 &amp; &lt;ready&gt;/);
    assert.match(theme, /accent1><a:srgbClr val="123456"/);
    assert.match(theme, /majorFont><a:latin typeface="Arial"/);
    assert.match(theme, /minorFont><a:latin typeface="Microsoft YaHei"/);
    assert.match(layout2, /name="Title &amp; Content"/);
    assert.match(layout2, /<p:ph type="title" idx="1"/);
    assert.match(layout2, /<a:schemeClr val="accent1"/);
    assert.doesNotMatch(slide1, /<p:bg>/);
    assert.match(slide1, /Q1 &amp; &lt;ready&gt;/);
    assert.match(slide1, /<a:buChar char="→"/);
    assert.match(slide1, /<a:pPr algn="l" marL="342900" indent="-342900">/);
    assert.match(slide1, /<a:normAutofit fontScale="95000"/);
    assert.doesNotMatch(slide1, /<p:cxnSp>/);
    assert.match(slide1, /<a:custGeom>/);
    assert.match(slide1, /<p:grpSp>/);
    assert.match(slide1, /<a:srcRect l="7143"[^>]*r="7143"/);
    assert.match(slide1, /<a:tbl>/);
    assert.match(slide1, /gridSpan="2"/);
    assert.match(slide1, /hMerge="1"/);
    assert.match(slide1, /pptkit:slideData/);
    assert.match(slideRels, /TargetMode="External"/);
    assert.match(slideRels, /relationships\/hyperlink/);
    assert.match(slideRels, /Target="\.\.\/slides\/slide2.xml"/);
    assert.match(notes, /Speaker note/);
    assert.match(slide2, /show="0"/);
    assert.match(slide2, /<p:bg>/);
    assert.ok(entries.has("ppt/media/asset-1.png"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("Node entry packages local images and writes nested output", async () => {
  const directory = await mkdtemp(join(tmpdir(), "pptkit-export-"));
  const imagePath = join(directory, "pixel.png");
  await writeFile(imagePath, Buffer.from("fake-png"));
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", id: "hero", source: { type: "path", value: imagePath }, mimeType: "image/png" });
  const missing = presentation.registerAsset({ kind: "image", id: "missing", source: { type: "path", value: join(directory, "missing.png") }, mimeType: "image/png" });
  presentation.addSlide({ elements: [
    { type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } },
    { type: "image", assetId: missing.id, box: { x: 100, y: 0, width: 100, height: 100 } },
  ] });
  const output = join(directory, "deck.pptx");
  const generated = await generateNodePptx(presentation);
  const result = await writePptx(presentation, { output });
  assert.equal(result.status, "written-with-warnings");
  assert.ok(result.warnings.some((warning) => warning.code === "asset-read-failed"));
  assert.ok(readZipEntries(generated.bytes).has("ppt/media/hero.png"));
  assert.equal((await stat(output)).size, result.byteLength);
  assert.ok(readZipEntries(await readFile(output)).has("ppt/media/hero.png"));
});

test("default entry reports path assets as unsupported", async () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", source: { type: "path", value: "./pixel.png" } });
  presentation.addSlide({ elements: [{ type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 100, height: 100 } }] });
  const result = await generatePptx(presentation);
  assert.equal(result.status, "generated-with-warnings");
  assert.match(result.warnings.find((warning) => warning.code === "asset-read-failed")?.message ?? "", /only supported by @pptkit\/pptx-exporter\/node/);
});

test("connector extents stay positive for horizontal and vertical lines", async () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide();
  slide.addElement({
    type: "connector",
    start: { x: 48, y: 128 },
    end: { x: 624, y: 128 },
    style: { width: 3, paint: { type: "solid", color: "0F172A" } },
  });
  slide.addElement({
    type: "connector",
    start: { x: 96, y: 72 },
    end: { x: 96, y: 288 },
    style: { width: 3, paint: { type: "solid", color: "0F172A" } },
  });

  const result = await generatePptx(presentation);
  const slideXml = readZipEntries(result.bytes).get("ppt/slides/slide1.xml").toString();
  assert.equal(slideXml.includes("<p:cxnSp>"), false);
  const extents = slideXml.split("<p:sp>").slice(1).map((connector) => {
    const match = connector.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    assert.ok(match);
    return { cx: Number(match[1]), cy: Number(match[2]) };
  });

  assert.deepEqual(extents, [
    { cx: 7315200, cy: 13 },
    { cx: 13, cy: 2743200 },
  ]);
});

test("shape text exports as editable text inside the native shape", async () => {
  const presentation = createPresentation({
    textStylePresets: {
      title: { frame: { margin: 0, verticalAlign: "middle" }, paragraph: { align: "center" }, run: { fontSize: 24, bold: true } },
    },
  });
  presentation.addSlide({ elements: [{
    type: "shape",
    shape: "roundRect",
    box: { x: 20, y: 20, width: 200, height: 60 },
    text: { content: "Card", textStylePreset: "title" },
  }] });
  const result = await generatePptx(presentation);
  const slideXml = readZipEntries(result.bytes).get("ppt/slides/slide1.xml").toString();
  const shape = slideXml.slice(slideXml.indexOf("<p:sp>"), slideXml.indexOf("</p:sp>") + 7);
  assert.match(shape, /<a:prstGeom prst="roundRect"/);
  assert.match(shape, /<p:txBody>/);
  assert.match(shape, /<a:bodyPr[^>]*anchor="ctr"[^>]*lIns="0"/);
  assert.match(shape, /<a:rPr[^>]*sz="2400"[^>]*b="1"/);
  assert.doesNotMatch(shape, /txBox="1"/);
});

test("generatePptx exports native editable charts", async () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide();
  slide.addElement({
    type: "chart",
    chartType: "bar",
    categories: ["A", "B", "C"],
    series: [
      { name: "Series 1", values: [10, 20, 30], color: "FF0000" },
      { name: "Series 2", values: [15, 25, 35], color: "0000FF" },
    ],
    axes: { category: { majorGridlines: { paint: { type: "solid", color: "ABCDEF" }, width: 0.5 } } },
    box: { x: 10, y: 10, width: 300, height: 200 },
  });
  slide.addElement({
    type: "chart",
    chartType: "line",
    categories: ["X", "Y"],
    series: [{ name: "Line 1", values: [5, 15], color: "00FF00", marker: { shape: "square", size: 7 } }],
    box: { x: 320, y: 10, width: 300, height: 200 },
    legend: { position: "bottom" },
  });
  slide.addElement({
    type: "chart",
    chartType: "pie",
    categories: ["P", "Q", "R", "S"],
    series: [{ name: "Pie 1", values: [25, 35, 30, 10], pointColors: ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"] }],
    box: { x: 10, y: 220, width: 300, height: 200 },
  });
  slide.addElement({
    type: "chart",
    chartType: "bar",
    categories: ["M", "N"],
    series: [{ name: "Hidden Legend", values: [1, 2] }],
    box: { x: 320, y: 220, width: 300, height: 200 },
    orientation: "horizontal",
    legend: { visible: false },
  });

  const result = await generatePptx(presentation);
  const repeated = await generatePptx(presentation);
  assert.deepEqual(result.bytes, repeated.bytes);

  const entries = readZipEntries(result.bytes);
  const slideXml = entries.get("ppt/slides/slide1.xml").toString();
  const slideRels = entries.get("ppt/slides/_rels/slide1.xml.rels").toString();
  const contentTypes = entries.get("[Content_Types].xml").toString();

  assert.match(slideXml, /graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart"/);
  assert.ok(slideXml.includes("<c:chart"));
  assert.match(slideRels, /relationships\/chart/);
  assert.match(slideRels, /Target="\.\.\/charts\/chart1\.xml"/);
  assert.match(slideRels, /Target="\.\.\/charts\/chart2\.xml"/);
  assert.match(slideRels, /Target="\.\.\/charts\/chart3\.xml"/);
  assert.match(slideRels, /Target="\.\.\/charts\/chart4\.xml"/);

  assert.ok(entries.has("ppt/charts/chart1.xml"));
  assert.ok(entries.has("ppt/charts/chart2.xml"));
  assert.ok(entries.has("ppt/charts/chart3.xml"));
  assert.ok(entries.has("ppt/charts/chart4.xml"));
  assert.ok(entries.has("ppt/charts/_rels/chart1.xml.rels"));
  assert.ok(entries.has("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));

  const chart1 = entries.get("ppt/charts/chart1.xml").toString();
  const chart2 = entries.get("ppt/charts/chart2.xml").toString();
  const chart3 = entries.get("ppt/charts/chart3.xml").toString();
  const chart4 = entries.get("ppt/charts/chart4.xml").toString();

  assert.match(chart1, /<c:chartSpace/);
  assert.match(chart1, /<c:barChart/);
  assert.match(chart1, /<c:barDir val="col"/);
  assert.match(chart1, /<c:catAx>[\s\S]*?<c:majorGridlines>[\s\S]*?ABCDEF/);
  assert.match(chart1, /<c:grouping val="clustered"/);
  assert.match(chart1, /<c:tickLblSkip val="1"\/><c:tickMarkSkip val="1"\/>/);
  assert.match(chart1, /<c:crossBetween val="between"\/>/);
  assert.match(chart1, /<c:majorTickMark val="none"\/>/);

  assert.match(chart2, /<c:chartSpace/);
  assert.match(chart2, /<c:lineChart/);
  assert.match(chart2, /<c:grouping val="standard"/);
  assert.match(chart2, /<c:symbol val="square"/);
  assert.match(chart2, /<c:size val="7"/);
  assert.match(chart2, /<c:crossBetween val="midCat"\/>/);
  const lineSeries = chart2.slice(chart2.indexOf("<c:ser>"), chart2.indexOf("</c:ser>") + "</c:ser>".length);
  assert.ok(lineSeries.indexOf("<c:cat>") < lineSeries.indexOf("<c:val>"));
  assert.ok(lineSeries.indexOf("<c:val>") < lineSeries.indexOf('<c:smooth val="0"/>'));

  assert.match(chart3, /<c:chartSpace/);
  assert.match(chart3, /<c:pieChart/);

  assert.match(chart1, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"/);
  assert.match(chart2, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="2"/);
  assert.match(chart3, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="4"/);
  assert.match(chart1, /Sheet1!\$B\$2:\$B\$4/);
  assert.match(chart1, /Sheet1!\$C\$2:\$C\$4/);

  // Legend: right when default / visible, bottom when configured, absent when hidden
  assert.match(chart1, /<c:legend><c:legendPos val="r"\/>[\s\S]*?<c:overlay val="0"\/>/);
  assert.match(chart2, /<c:legend><c:legendPos val="b"\/>[\s\S]*?<c:overlay val="0"\/>/);
  assert.match(chart3, /<c:legend><c:legendPos val="r"\/>[\s\S]*?<c:overlay val="0"\/>/);
  for (const chart of [chart1, chart2, chart3]) {
    const legend = chart.slice(chart.indexOf("<c:legend>"), chart.indexOf("</c:legend>") + "</c:legend>".length);
    assert.doesNotMatch(legend, /<c:layout>/, "native legend position must control row/column flow");
  }
  assert.doesNotMatch(chart4, /<c:legend>/);

  // Pie slice colors: explicit per-data-point fills from the public series model
  const dPtMatches = [...chart3.matchAll(/<c:dPt>/g)];
  assert.equal(dPtMatches.length, 4);
  for (const color of ["FF0000", "00FF00", "0000FF", "FFFF00"]) assert.match(chart3, new RegExp(`<a:srgbClr val="${color}"`));
  assert.match(chart3, /<c:varyColors val="1"/);

  // Native chart data is backed by an embedded workbook.
  for (const chart of [chart1, chart2, chart3, chart4]) assert.match(chart, /<c:externalData r:id="rId1"/);
  const chartRels = entries.get("ppt/charts/_rels/chart1.xml.rels").toString();
  assert.match(chartRels, /relationships\/package/);
  assert.match(chartRels, /Microsoft_Excel_Worksheet1\.xlsx/);
  const workbookEntries = readZipEntries(entries.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"));
  const sheet = workbookEntries.get("xl/worksheets/sheet1.xml").toString();
  assert.match(sheet, /<c r="B1" t="inlineStr"><is><t xml:space="preserve">Series 1<\/t>/);
  assert.match(sheet, /<c r="C1" t="inlineStr"><is><t xml:space="preserve">Series 2<\/t>/);
  assert.match(sheet, /<c r="C4"><v>35<\/v><\/c>/);

  // EMU frame matches pt * 12700
  assert.match(slideXml, /<a:off x="127000" y="127000"\/><a:ext cx="3810000" cy="2540000"\/>/);
  assert.match(slideXml, /<a:off x="4064000" y="127000"\/><a:ext cx="3810000" cy="2540000"\/>/);
  assert.match(slideXml, /<a:off x="127000" y="2794000"\/><a:ext cx="3810000" cy="2540000"\/>/);
  assert.match(slideXml, /<a:off x="4064000" y="2794000"\/><a:ext cx="3810000" cy="2540000"\/>/);

  // Exact cached category counts
  assert.match(chart1, /<c:strCache><c:ptCount val="3"/);
  assert.match(chart2, /<c:strCache><c:ptCount val="2"/);
  assert.match(chart3, /<c:strCache><c:ptCount val="4"/);
  assert.match(chart4, /<c:strCache><c:ptCount val="2"/);

  assert.match(contentTypes, /\/ppt\/charts\/chart1\.xml/);
  assert.match(contentTypes, /application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml/);
  assert.match(contentTypes, /Extension="xlsx" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet"/);
  assert.match(chart4, /<c:barDir val="bar"/);
  assert.match(chart4, /<c:catAx>[\s\S]*?<c:orientation val="maxMin"/);
  assert.match(chart4, /<c:minorTickMark val="none"/);
  assert.match(chart1, /<c:majorUnit val="10"/);
  assert.match(chart1, /<a:srgbClr val="D9D9D9"/);
});
