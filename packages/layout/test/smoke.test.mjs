import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";

test("resolveLayout returns detached IR with metadata, theme, and layouts", () => {
  const presentation = createPresentation({ metadata: { title: "Layout" } });
  presentation.defineSlideLayout({
    id: "branded",
    name: "Branded",
    elements: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 20, height: 20 } }],
  });
  presentation.addSlide({
    layoutId: "branded",
    elements: [{ type: "text", content: "Hello", box: { x: 10, y: 20, width: 100, height: 30 } }],
  });
  const result = resolveLayout(presentation);
  assert.equal(result.status, "resolved");
  assert.equal(result.metadata.title, "Layout");
  assert.equal(result.theme.fonts.body, "Aptos");
  assert.equal(result.layouts.find((layout) => layout.id === "branded").elements.length, 1);
  assert.equal(result.slides[0].elements.length, 1);
  result.slides[0].elements[0].box.width = 999;
  assert.equal(presentation.slides[0].elements[0].box.width, 100);
});

test("layout resolves connector anchors", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [
    { type: "shape", id: "left", shape: "rect", box: { x: 10, y: 20, width: 100, height: 40 } },
    { type: "shape", id: "right", shape: "ellipse", box: { x: 300, y: 40, width: 80, height: 80 } },
    { type: "connector", start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" }, route: [{ x: 200, y: 40 }] },
  ] });
  const result = resolveLayout(presentation);
  const connector = result.slides[0].elements[2];
  assert.deepEqual(connector.start, { x: 110, y: 40 });
  assert.deepEqual(connector.end, { x: 300, y: 80 });
  assert.deepEqual(connector.box, { x: 110, y: 40, width: 190, height: 40 });
});

test("layout resolves contain and cover using asset dimensions", () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({ kind: "image", source: { type: "url", value: "https://example.com/wide.png" }, width: 400, height: 200 });
  presentation.addSlide({ elements: [
    { type: "image", assetId: asset.id, fit: "contain", box: { x: 0, y: 0, width: 100, height: 100 } },
    { type: "image", assetId: asset.id, fit: "cover", box: { x: 100, y: 0, width: 100, height: 100 } },
  ] });
  const result = resolveNormalizedLayout(normalizePresentation(presentation));
  assert.deepEqual(result.slides[0].elements[0].box, { x: 0, y: 25, width: 100, height: 50 });
  assert.equal(result.slides[0].elements[0].fit, "stretch");
  assert.equal(result.slides[0].elements[1].fit, "crop");
  assert.deepEqual(result.slides[0].elements[1].crop, { left: 0.25, right: 0.25, top: 0, bottom: 0 });
});

test("chart layout resolves shared scales, vertical and horizontal bar geometry", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [
    {
      type: "chart",
      id: "vertical",
      chartType: "bar",
      categories: ["A", "B", "C"],
      series: [{ name: "One", values: [10, 20, 30] }, { name: "Two", values: [15, 25, 35] }],
      legend: { visible: false },
      box: { x: 0, y: 0, width: 300, height: 200 },
    },
    {
      type: "chart",
      id: "horizontal",
      chartType: "bar",
      orientation: "horizontal",
      categories: ["A", "B"],
      series: [{ name: "Mixed", values: [-5, 10] }],
      axes: { value: { scale: { min: -10, max: 10, majorUnit: 5 } } },
      legend: { visible: false },
      box: { x: 320, y: 0, width: 300, height: 200 },
    },
  ] });
  const result = resolveLayout(presentation);
  const vertical = result.slides[0].elements[0];
  const horizontal = result.slides[0].elements[1];
  assert.equal(vertical.chartLayout.valueScale.min, 0);
  assert.equal(vertical.chartLayout.valueScale.max, 40);
  assert.deepEqual(vertical.chartLayout.valueScale.ticks, [0, 10, 20, 30, 40]);
  assert.equal(vertical.chartLayout.bars.length, 6);
  assert.ok(vertical.chartLayout.bars.every((bar) => Number.isFinite(bar.box.x) && bar.box.height >= 0));
  assert.deepEqual(horizontal.chartLayout.valueScale.ticks, [-10, -5, 0, 5, 10]);
  assert.equal(horizontal.chartLayout.bars.length, 2);
  assert.ok(horizontal.chartLayout.bars.every((bar) => bar.box.width > 0));
  assert.notEqual(horizontal.chartLayout.categoryPositions[0].y, horizontal.chartLayout.categoryPositions[1].y);
  assert.equal(vertical.chartLayout.categoryAxisPlacement, "betweenCategories");
  assert.equal(vertical.chartLayout.categoryTickPositions.length, vertical.categories.length + 1);
});

test("chart layout resolves right legends as columns and bottom legends as rows", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [
    {
      type: "chart", chartType: "pie", categories: ["Core", "Layout", "Exporter", "Renderer"],
      series: [{ name: "Packages", values: [1, 1, 1, 1] }],
      legend: { position: "right" }, box: { x: 0, y: 0, width: 624, height: 320 },
    },
    {
      type: "chart", chartType: "bar", categories: ["Q1", "Q2"],
      series: [{ name: "Drafts", values: [1, 2] }, { name: "Exports", values: [2, 3] }],
      legend: { position: "bottom" }, box: { x: 0, y: 340, width: 624, height: 320 },
    },
  ] });
  const [right, bottom] = resolveLayout(presentation).slides[0].elements.map((element) => element.chartLayout);
  assert.equal(right.legendItems.length, 4);
  assert.ok(right.legendItems.every((item) => item.box.x === right.legendBox.x));
  assert.deepEqual(right.legendItems.map((item) => item.box.y), [...right.legendItems.map((item) => item.box.y)].sort((a, b) => a - b));
  assert.equal(bottom.legendItems.length, 2);
  assert.ok(bottom.legendItems.every((item) => item.box.y === bottom.legendBox.y));
  assert.notEqual(bottom.legendItems[0].box.x, bottom.legendItems[1].box.x);
});

test("chart layout handles stacked, percent, constant, and pie data without unsafe geometry", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [
    {
      type: "chart",
      chartType: "bar",
      grouping: "stacked",
      categories: ["A"],
      series: [{ name: "Positive", values: [4] }, { name: "Negative", values: [-2] }],
      box: { x: 0, y: 0, width: 200, height: 160 },
    },
    {
      type: "chart",
      chartType: "bar",
      grouping: "percentStacked",
      categories: ["A"],
      series: [{ name: "One", values: [1] }, { name: "Two", values: [3] }],
      box: { x: 210, y: 0, width: 200, height: 160 },
    },
    {
      type: "chart",
      chartType: "line",
      categories: ["A", "B"],
      series: [{ name: "Constant", values: [5, 5] }],
      box: { x: 420, y: 0, width: 200, height: 160 },
    },
    {
      type: "chart",
      chartType: "pie",
      categories: ["A", "B"],
      series: [{ name: "Share", values: [0, 0] }],
      box: { x: 630, y: 0, width: 200, height: 160 },
    },
  ] });
  const charts = resolveLayout(presentation).slides[0].elements;
  assert.equal(charts[0].chartLayout.bars.length, 2);
  assert.equal(charts[1].chartLayout.valueScale.max, 100);
  assert.ok(charts[2].chartLayout.lines[0].points.every((point) => Number.isFinite(point.y)));
  assert.ok(charts[3].chartLayout.pie.slices.every((slice) => slice.startAngle === slice.endAngle));
});

test("automatic chart scales cover zero, negative, crossing-zero, decimal, and compact inputs", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [
    { type: "chart", chartType: "bar", categories: ["Only"], series: [{ name: "Zero", values: [0] }], legend: { visible: false }, box: { x: 0, y: 0, width: 90, height: 70 } },
    { type: "chart", chartType: "bar", categories: ["A", "B"], series: [{ name: "Negative", values: [-9, -2] }], legend: { visible: false }, box: { x: 100, y: 0, width: 220, height: 160 } },
    { type: "chart", chartType: "line", categories: ["A", "B", "C"], series: [{ name: "Cross", values: [-0.15, 0.05, 0.22] }], legend: { visible: false }, box: { x: 330, y: 0, width: 240, height: 160 } },
  ] });
  const charts = resolveLayout(presentation).slides[0].elements;
  const zero = charts[0].chartLayout;
  const negative = charts[1].chartLayout;
  const decimal = charts[2].chartLayout;
  assert.equal(zero.valueScale.min, 0);
  assert.ok(zero.valueScale.max > 0);
  assert.equal(zero.bars.length, 1);
  assert.ok(zero.bars.every((bar) => Object.values(bar.box).every(Number.isFinite)));
  assert.ok(negative.valueScale.min <= -9 && negative.valueScale.max === 0);
  assert.ok(negative.bars.every((bar) => bar.box.height >= 0));
  assert.ok(decimal.valueScale.min <= -0.15 && decimal.valueScale.max >= 0.22);
  assert.ok(decimal.valueScale.ticks.includes(0));
  assert.ok(decimal.lines[0].points.every((point) => Number.isFinite(point.x) && Number.isFinite(point.y)));
});
