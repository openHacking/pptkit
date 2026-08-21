import test from "node:test";
import assert from "node:assert/strict";

import {
  createPresentation,
  normalizePresentation,
  PresentationValidationError,
  validatePresentation,
} from "../dist/index.js";

test("validation collects independent reference and geometry errors", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    id: "intro",
    elements: [
      { type: "image", assetId: "missing", fit: "crop", crop: { left: 0.8, right: 0.3 }, box: { x: 0, y: 0, width: 100, height: 100 } },
      { type: "shape", id: "node", shape: "rect", box: { x: 0, y: 0, width: -1, height: 20 }, opacity: 2 },
      {
        type: "connector",
        start: { elementId: "node", anchor: "right" },
        end: { elementId: "missing-node" },
        style: { width: -1 },
        action: { type: "slide", slideId: "missing-slide" },
      },
    ],
  });

  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "missing-image-asset"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-image-crop"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-element-box"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-opacity"));
  assert.ok(diagnostics.some((item) => item.code === "missing-connector-target"));
  assert.ok(diagnostics.some((item) => item.code === "missing-action-slide"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-stroke-width"));

  assert.throws(
    () => normalizePresentation(presentation),
    (error) => error instanceof PresentationValidationError && error.diagnostics.length === diagnostics.length,
  );
});

test("validation accepts groups, connectors, tables, and placeholder references", () => {
  const presentation = createPresentation();
  presentation.defineSlideLayout({
    id: "content",
    name: "Content",
    placeholders: [{ key: "body", kind: "body", box: { x: 20, y: 20, width: 400, height: 200 } }],
  });
  presentation.addSlide({
    layoutId: "content",
    elements: [
      { type: "shape", id: "left", shape: "rect", box: { x: 0, y: 0, width: 50, height: 50 } },
      { type: "shape", id: "right", shape: "ellipse", box: { x: 100, y: 0, width: 50, height: 50 } },
      { type: "connector", start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" } },
      { type: "text", content: "Body", placeholderKey: "body" },
    ],
  });
  assert.deepEqual(validatePresentation(presentation), []);
});

test("validation reports missing width for auto-sized text boxes", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "text",
      content: "Needs a width",
      box: { x: 20, y: 20 },
    }],
  });

  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "invalid-text-auto-size-box"));
});

test("mutation methods reject duplicate identities immediately", () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide({ id: "intro" });
  assert.throws(() => presentation.addSlide({ id: "intro" }), /Duplicate slide id/);
  slide.addElement({ type: "shape", id: "same", shape: "rect", box: { x: 0, y: 0, width: 1, height: 1 } });
  assert.throws(() => slide.addElement({ type: "shape", id: "same", shape: "rect", box: { x: 0, y: 0, width: 1, height: 1 } }), /Duplicate element id/);
});

test("validation reports missing and malformed text style presets", () => {
  const presentation = createPresentation({
    textStylePresets: {
      "": { run: { fontSize: 0 } },
    },
  });
  presentation.addSlide({ elements: [{
    type: "shape",
    shape: "rect",
    box: { x: 0, y: 0, width: 100, height: 50 },
    text: { content: "Missing", textStylePreset: "missing" },
  }] });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "invalid-text-style-preset-name"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-font-size"));
  assert.ok(diagnostics.some((item) => item.code === "missing-text-style-preset"));
});

test("validation accepts valid bar, line, and pie charts", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [
      {
        type: "chart",
        chartType: "bar",
        categories: ["Q1", "Q2", "Q3"],
        series: [{ name: "Sales", values: [10, 20, 30] }],
        title: "Revenue",
        showLegend: true,
        xAxis: { show: true, labels: true },
        yAxis: { show: true },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "line",
        categories: ["A", "B"],
        series: [{ name: "X", values: [1, 2] }, { name: "Y", values: [3, 4] }],
        showLegend: false,
        xAxis: { show: true, labels: false },
        yAxis: { show: false },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "pie",
        categories: ["A", "B", "C"],
        series: [{ name: "Share", values: [30, 50, 20], color: { theme: "accent1" } }],
        showLegend: true,
        xAxis: { show: false, labels: false },
        yAxis: { show: false },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
    ],
  });
  assert.deepEqual(validatePresentation(presentation), []);
});

test("validation accepts charts without optional config fields", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [10, 20] }],
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  assert.deepEqual(validatePresentation(presentation), []);
});

test("validation reports pie charts with multiple series", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "pie",
      categories: ["A", "B"],
      series: [{ name: "X", values: [1, 2] }, { name: "Y", values: [3, 4] }],
      showLegend: true,
      xAxis: { show: true, labels: true },
      yAxis: { show: true },
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "pie-single-series"));
});

test("validation reports chart series values length mismatch", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2", "Q3"],
      series: [{ name: "Sales", values: [10, 20] }],
      showLegend: true,
      xAxis: { show: true, labels: true },
      yAxis: { show: true },
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "chart-values-length"));
});

test("validation reports non-finite chart values", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [10, NaN] }],
      showLegend: true,
      xAxis: { show: true, labels: true },
      yAxis: { show: true },
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "chart-values-finite"));
});

test("validation reports empty chart categories", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: [],
      series: [{ name: "Sales", values: [] }],
      showLegend: true,
      xAxis: { show: true, labels: true },
      yAxis: { show: true },
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "empty-chart-categories"));
});
