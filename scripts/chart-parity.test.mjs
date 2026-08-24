import test from "node:test";
import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { resolveNormalizedLayout } from "../packages/layout/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";

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
    const start = offset + 30 + nameLength + extraLength;
    const data = bytes.subarray(start, start + compressedSize);
    entries.set(name, method === 8 ? inflateRawSync(data) : data);
    offset = start + compressedSize;
  }
  return entries;
}

function svgFragment(svg, id) {
  const marker = `data-pptkit-element-id="${id}"`;
  const start = svg.indexOf(marker);
  assert.ok(start >= 0, `${id} present in SVG`);
  const next = svg.indexOf('data-pptkit-element-id="', start + marker.length);
  return next < 0 ? svg.slice(start) : svg.slice(start, next);
}

function xmlValue(xml, tag) {
  const match = xml.match(new RegExp(`<c:${tag} val="([^"]+)"`));
  assert.ok(match, `${tag} is explicit in chart XML`);
  return Number(match[1]);
}

function colorToken(color) {
  return color.replace(/^#/, "").toUpperCase();
}

test("SVG and PPTX consume the same canonical chart layout contract", async () => {
  const presentation = createPresentation({ metadata: { title: "Chart parity" } });
  const slide = presentation.addSlide({ id: "parity" });
  slide.addElement({
    type: "chart", id: "vertical", chartType: "bar", orientation: "vertical",
    categories: ["Q1", "Q2", "Q3"],
    series: [{ name: "Actual", values: [-5, 12, 25], color: "2457D6" }, { name: "Plan", values: [4, 18, 20], color: "E65A3A" }],
    legend: { visible: true, position: "right" },
    box: { x: 20, y: 20, width: 400, height: 260 },
  });
  slide.addElement({
    type: "chart", id: "trend", chartType: "line",
    categories: ["Jan", "Feb", "Mar"],
    series: [{ name: "Visits", values: [2, 8, 15], color: "4472C4", marker: "auto" }, { name: "Orders", values: [1, 5, 11], color: "ED7D31", marker: "auto" }],
    legend: { visible: true, position: "bottom" },
    axes: { value: { scale: { min: 0, max: 20, majorUnit: 5 } } },
    box: { x: 440, y: 20, width: 400, height: 260 },
  });
  slide.addElement({
    type: "chart", id: "horizontal", chartType: "bar", orientation: "horizontal",
    categories: ["A", "B"], series: [{ name: "Count", values: [3, 9], color: "70AD47" }],
    legend: { visible: false }, box: { x: 20, y: 310, width: 400, height: 240 },
  });
  slide.addElement({
    type: "chart", id: "share", chartType: "pie", categories: ["A", "B", "C"],
    series: [{ name: "Share", values: [2, 3, 5], pointColors: ["4472C4", "ED7D31", "A5A5A5"] }],
    legend: { visible: true }, box: { x: 440, y: 310, width: 400, height: 240 },
  });

  const layout = resolveNormalizedLayout(normalizePresentation(presentation));
  const charts = new Map(layout.slides[0].elements.filter((element) => element.type === "chart").map((element) => [element.id, element]));
  const [svgResult, pptxResult] = await Promise.all([renderPresentationToSvg(presentation), generatePptx(presentation)]);
  assert.equal(svgResult.status, "rendered");
  assert.equal(pptxResult.status, "generated");

  const entries = readZipEntries(pptxResult.bytes);
  const xml = [1, 2, 3, 4].map((index) => entries.get(`ppt/charts/chart${index}.xml`).toString());
  const svg = svgResult.slides[0].svg;

  for (const [index, id] of ["vertical", "trend", "horizontal"].entries()) {
    const chart = charts.get(id);
    assert.ok(chart && chart.chartType !== "pie");
    assert.equal(xmlValue(xml[index], "min"), chart.chartLayout.valueScale.min);
    assert.equal(xmlValue(xml[index], "max"), chart.chartLayout.valueScale.max);
    assert.equal(xmlValue(xml[index], "majorUnit"), chart.chartLayout.valueScale.majorUnit);
    const placement = id === "trend" ? "onCategory" : "betweenCategories";
    assert.equal(chart.chartLayout.categoryAxisPlacement, placement);
    assert.match(xml[index], new RegExp(`<c:crossBetween val="${placement === "onCategory" ? "midCat" : "between"}"/>`));
    assert.match(xml[index], /<c:tickLblSkip val="1"\/><c:tickMarkSkip val="1"\/>/);
    const fragment = svgFragment(svg, id);
    assert.ok(fragment.indexOf('data-chart-layer="grid"') < fragment.indexOf('data-chart-layer="series"'));
    assert.equal([...fragment.matchAll(/data-chart-role="gridline"/g)].length, chart.chartLayout.valueScale.ticks.length);
  }

  assert.match(xml[0], /<c:barDir val="col"/);
  assert.match(xml[2], /<c:barDir val="bar"/);
  assert.equal(charts.get("vertical").orientation, "vertical");
  assert.equal(charts.get("horizontal").orientation, "horizontal");

  const line = charts.get("trend");
  assert.equal(line.series[0].marker.shape, "diamond");
  assert.equal(line.series[1].marker.shape, "square");
  assert.match(xml[1], /<c:symbol val="diamond"/);
  assert.match(xml[1], /<c:symbol val="square"/);
  const lineSvg = svgFragment(svg, "trend");
  assert.equal([...lineSvg.matchAll(/data-chart-role="legend-marker"/g)].length, 2);
  assert.equal([...lineSvg.matchAll(/data-chart-role="series-marker"/g)].length, 6);
  assert.match(lineSvg, /data-chart-role="legend-line"/);

  for (const index of [0, 1, 3]) {
    const legend = xml[index].slice(xml[index].indexOf("<c:legend>"), xml[index].indexOf("</c:legend>") + "</c:legend>".length);
    assert.doesNotMatch(legend, /<c:layout>/);
  }

  for (const chart of charts.values()) {
    for (const series of chart.series) {
      const colors = chart.chartType === "pie" ? series.pointColors : [series.color];
      for (const color of colors) {
        assert.match(svgFragment(svg, chart.id).toUpperCase(), new RegExp(color.toUpperCase().replace(/^#/, "#")));
        assert.match(xml[["vertical", "trend", "horizontal", "share"].indexOf(chart.id)], new RegExp(colorToken(color)));
      }
    }
  }

  for (let index = 1; index <= 4; index += 1) {
    assert.match(xml[index - 1], /<c:manualLayout>/);
    assert.match(xml[index - 1], new RegExp(`<c:externalData r:id="rId1"`));
    assert.ok(entries.has(`ppt/charts/_rels/chart${index}.xml.rels`));
    assert.ok(entries.has(`ppt/embeddings/Microsoft_Excel_Worksheet${index}.xlsx`));
  }
});
