import test from "node:test";
import assert from "node:assert/strict";

import {
  createPresentation,
  normalizePresentation,
  PresentationValidationError,
  validatePresentation,
} from "../dist/index.js";

test("connector targets resolve only in their immediate coordinate scope", () => {
  const p = createPresentation(), box = { x: 0, y: 0, width: 100, height: 100 };
  p.addSlide({ elements: [
    { id: 'outer-edge', type: 'connector', start: { elementId: 'inner' }, end: { x: 10, y: 10 } },
    { id: 'outer', type: 'group', box, coordinateSize: { width: 100, height: 100 }, children: [
      { id: 'inner', type: 'shape', shape: 'rect', box },
      { id: 'valid-edge', type: 'connector', start: { elementId: 'inner' }, end: { x: 5, y: 5 } },
      { id: 'nested', type: 'group', box, coordinateSize: { width: 100, height: 100 }, children: [{ id: 'deep', type: 'shape', shape: 'rect', box }] },
      { id: 'invalid-edge', type: 'connector', start: { elementId: 'deep' }, end: { x: 5, y: 5 } },
    ] },
  ] });
  const errors = validatePresentation(p).filter(item => item.code === 'missing-connector-target');
  assert.deepEqual(errors.map(item => item.path), ['slides.0.elements.0.start', 'slides.0.elements.1.children.3.start']);
});

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
        legend: { visible: true, position: "right" },
        axes: { category: { visible: true, labels: true }, value: { visible: true } },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "line",
        categories: ["A", "B"],
        series: [{ name: "X", values: [1, 2] }, { name: "Y", values: [3, 4] }],
        legend: { visible: false },
        axes: { category: { visible: true, labels: false }, value: { visible: false } },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "pie",
        categories: ["A", "B", "C"],
        series: [{ name: "Share", values: [30, 50, 20], color: { theme: "accent1" } }],
        legend: { visible: true },
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
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "empty-chart-categories"));
});

test("validation reports invalid chart type", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "surface",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [10, 20] }],
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  const match = diagnostics.find((item) => item.code === "invalid-chart-type");
  assert.ok(match, "expected invalid-chart-type diagnostic");
  assert.equal(match.severity, "error");
});

test("validation reports empty chart series", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [],
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  const match = diagnostics.find((item) => item.code === "empty-chart-series");
  assert.ok(match, "expected empty-chart-series diagnostic");
  assert.equal(match.severity, "error");
});

test("validation reports empty chart series name", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [{ name: "", values: [10, 20] }],
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  const match = diagnostics.find((item) => item.code === "chart-series-name");
  assert.ok(match, "expected chart-series-name diagnostic");
  assert.equal(match.severity, "error");
});

test("validation reports invalid chart series color", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [10, 20], color: "not-a-color" }],
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  const match = diagnostics.find((item) => item.code === "invalid-chart-color");
  assert.ok(match, "expected invalid-chart-color diagnostic");
  assert.equal(match.severity, "error");
});

test("validation reports empty chart title", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "chart",
      chartType: "bar",
      categories: ["Q1", "Q2"],
      series: [{ name: "Sales", values: [10, 20] }],
      title: "",
      box: { x: 0, y: 0, width: 400, height: 300 },
    }],
  });
  const diagnostics = validatePresentation(presentation);
  const match = diagnostics.find((item) => item.code === "invalid-chart-title");
  assert.ok(match, "expected invalid-chart-title diagnostic");
  assert.equal(match.severity, "error");
});

test("validation reports invalid chart axes, legend, scale, and marker config", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [
      {
        type: "chart",
        chartType: "bar",
        categories: ["Q1", "Q2"],
        series: [{ name: "Sales", values: [10, 20], marker: "auto", pointColors: ["FF0000", "00FF00"] }],
        orientation: "diagonal",
        grouping: "overlap",
        legend: { visible: 1, position: "left" },
        axes: { category: { visible: true, labels: true, majorTick: "large" }, value: { scale: { min: 10, max: 5, majorUnit: 0 } } },
        box: { x: 0, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "bar",
        categories: ["Q1", "Q2"],
        series: [{ name: "Sales", values: [10, 20] }],
        legend: { visible: true },
        axes: { category: { visible: "yes", labels: true }, value: { visible: true } },
        seriesGap: 101,
        box: { x: 0, y: 320, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "line",
        categories: ["Q1"],
        series: [{ name: "Line", values: [1], marker: { shape: "circle", size: 100 } }],
        box: { x: 420, y: 0, width: 400, height: 300 },
      },
      {
        type: "chart",
        chartType: "pie",
        categories: ["Q1"],
        series: [{ name: "Pie", values: [-1] }],
        box: { x: 420, y: 320, width: 400, height: 300 },
      },
    ],
  });
  const diagnostics = validatePresentation(presentation);
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-legend"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-axis"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-scale"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-series-field"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-orientation"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-grouping"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-gap"));
  assert.ok(diagnostics.some((item) => item.code === "invalid-chart-marker"));
  assert.ok(diagnostics.some((item) => item.code === "pie-negative-value"));
});

test("fixed chart axes locate out-of-range values without blocking normalization", () => {
  for (const chartType of ["bar", "line"]) {
    const document = createPresentation();
    const slide = document.addSlide({ id: "fixed-axis-slide" });
    slide.addElement({ id: "fixed-axis-chart", type: "chart", chartType, box: { x: 0, y: 0, width: 300, height: 200 },
      categories: ["below", "min", "max", "above", "blank"], series: [{ name: "values", values: [-1, 0, 9, 37.25, null] }],
      axes: { value: { scale: { min: 0, max: 9, majorUnit: 1 } } } });
    const warnings = validatePresentation(document).filter(d => d.code === "chart-value-outside-fixed-axis");
    assert.deepEqual(warnings.map(d => [d.severity, d.slideId, d.elementId, d.path]), [
      ["warning", "fixed-axis-slide", "fixed-axis-chart", "slides.0.elements.0.series.0.values.0"],
      ["warning", "fixed-axis-slide", "fixed-axis-chart", "slides.0.elements.0.series.0.values.3"],
    ]);
    assert.deepEqual(normalizePresentation(document).slides[0].elements[0].series[0].values, [-1, 0, 9, 37.25, null]);
  }
  for (const [grouping, scale] of [["clustered", "auto"], ["clustered", { min: 9, max: 0, majorUnit: 1 }]]) {
    const document = createPresentation();
    document.addSlide().addElement({ type: "chart", chartType: "bar", grouping, box: { x: 0, y: 0, width: 300, height: 200 }, categories: ["A"], series: [{ name: "S", values: [37.25] }], axes: { value: { scale } } });
    assert.equal(validatePresentation(document).filter(d => d.code === "chart-value-outside-fixed-axis").length, 0);
  }
});

test("fixed-axis warnings respect chart family, combo axis ownership and paired coordinates", () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "axis-slide" });
  const box = { x: 0, y: 0, width: 300, height: 200 }, scale = (min, max) => ({ min, max, majorUnit: 1 });
  slide.addElement({ id: "area", type: "chart", chartType: "area", box, categories: ["A"], series: [{ name: "A", values: [12] }], axes: { value: { scale: scale(0, 10) } } });
  slide.addElement({ id: "stacked", type: "chart", chartType: "area", grouping: "stacked", box, categories: ["A"], series: [{ name: "A", values: [8] }], axes: { value: { scale: scale(0, 10) } } });
  slide.addElement({ id: "combo", type: "chart", chartType: "combo", box, categories: ["A", "B"], series: [{ kind: "bar", name: "B", values: [11, 5] }, { kind: "line", name: "L", values: [11, 101] }], axes: { value: { scale: scale(0, 10) }, secondaryValue: { scale: scale(0, 100) } } });
  slide.addElement({ id: "scatter", type: "chart", chartType: "scatter", box, series: [{ name: "S", xValues: [11, 20, 5], yValues: [3, null, -1] }], axes: { x: { scale: scale(0, 10) }, y: { scale: scale(0, 10) } } });
  slide.addElement({ id: "bubble", type: "chart", chartType: "bubble", box, series: [{ name: "B", xValues: [11, 3], yValues: [12, 3], sizes: [1, 2] }], axes: { x: { scale: scale(0, 10) }, y: { scale: scale(0, 10) } } });
  const warnings = validatePresentation(deck).filter(d => d.code === "chart-value-outside-fixed-axis");
  assert.deepEqual(warnings.map(d => [d.elementId, d.path.replace(/^.*\.series\./, "series.")]), [
    ["area", "series.0.values.0"], ["combo", "series.0.values.0"], ["combo", "series.1.values.1"],
    ["scatter", "series.0.xValues.0"], ["scatter", "series.0.yValues.2"], ["bubble", "series.0.xValues.0"], ["bubble", "series.0.yValues.0"],
  ]);
  assert.ok(warnings.every(d => d.severity === "warning" && d.slideId === "axis-slide"));
  assert.doesNotThrow(() => normalizePresentation(deck));
});

test("stacked fixed-axis warnings use cumulative endpoints and family-specific percentage scales", () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "stack-slide" });
  const box = { x: 0, y: 0, width: 300, height: 200 };
  const add = (id, chartType, grouping, values, min, max) => slide.addElement({ id, type: "chart", chartType, grouping, box, categories: ["A"], series: values.map((value, i) => ({ name: `S${i}`, values: [value] })), axes: { value: { scale: { min, max, majorUnit: 0.1 } } } });
  add("bar", "bar", "stacked", [4, 8, -3, -9], -10, 10);
  add("area", "area", "stacked", [8, -3, 7], 0, 10);
  add("bar-percent", "bar", "percentStacked", [4, 4, -3, -1], -80, 80);
  add("area-percent", "area", "percentStacked", [4, -2, 2], -1, 0.4);
  add("zero", "bar", "percentStacked", [0, null, 0], -1, 1);
  add("huge", "bar", "percentStacked", [Number.MAX_VALUE, Number.MAX_VALUE], 0, 75);
  add("rounded-bar", "bar", "percentStacked", [1, 1, 1, 1, 1, 1], 0, 100);
  add("rounded-area", "area", "percentStacked", [1, 1, 1, 1, 1, 1], 0, 1);
  add("outside-rounded", "bar", "percentStacked", [1, 1, 1, 1, 1, 1], 0, 100 - 1e-9);
  add("wide-axis", "bar", "percentStacked", [1], -Number.MAX_VALUE, 90);
  const warnings = validatePresentation(deck).filter(d => d.code === "chart-value-outside-fixed-axis");
  assert.deepEqual(warnings.map(d => [d.elementId, d.path.match(/series\.\d+\.values\.\d+$/)[0]]), [
    ["bar", "series.1.values.0"], ["bar", "series.3.values.0"], ["area", "series.2.values.0"],
    ["bar-percent", "series.1.values.0"], ["bar-percent", "series.3.values.0"],
    ["area-percent", "series.0.values.0"], ["area-percent", "series.2.values.0"], ["huge", "series.1.values.0"], ["outside-rounded", "series.5.values.0"], ["wide-axis", "series.0.values.0"],
  ]);
  assert.ok(warnings.every(d => d.severity === "warning" && d.slideId === "stack-slide" && d.message.includes("cumulative stack endpoint")));
  assert.doesNotThrow(() => normalizePresentation(deck));
  const malformed = createPresentation();
  malformed.addSlide().addElement({ type: "chart", chartType: "bar", grouping: "stacked", box, categories: ["A"], series: [{ name: "A", values: [10] }, { name: "B", values: [undefined] }], axes: { value: { scale: { min: 0, max: 5, majorUnit: 1 } } } });
  assert.equal(validatePresentation(malformed).filter(d => d.code === "chart-value-outside-fixed-axis").length, 0);
});

test("chart validation rejects sparse categories, series, values and pie colors without treating holes as zero", () => {
  const chart = (chartType, categories, series) => {
    const deck = createPresentation();
    deck.addSlide().addElement({ type: "chart", chartType, box: { x: 0, y: 0, width: 300, height: 200 }, categories, series });
    return deck;
  };
  for (const chartType of ["bar", "line", "pie"]) {
    for (const values of [[1, , 3], [1, undefined, 3], [1, Infinity, 3], [1, NaN, 3]]) {
      const sparse = chart(chartType, ["A", "B", "C"], [{ name: "S", values }]);
      const diagnostics = validatePresentation(sparse);
      assert.ok(diagnostics.some(item => item.code === "chart-values-finite" && item.path.endsWith(".series.0.values.1")), chartType);
      assert.throws(() => normalizePresentation(sparse), PresentationValidationError);
    }
    const blank = chart(chartType, ["A", "B", "C"], [{ name: "S", values: [1, null, 3] }]);
    assert.deepEqual(validatePresentation(blank), []); assert.equal(normalizePresentation(blank).slides[0].elements[0].series[0].values[1], null);
    const zero = chart(chartType, ["A", "B", "C"], [{ name: "S", values: [1, 0, 3] }]);
    assert.deepEqual(validatePresentation(zero), []); assert.equal(normalizePresentation(zero).slides[0].elements[0].series[0].values[1], 0);
  }
  const missingCategory = chart("bar", ["A", , "C"], [{ name: "S", values: [1, 2, 3] }]);
  assert.ok(validatePresentation(missingCategory).some(item => item.code === "empty-chart-categories"));
  const missingSeries = chart("bar", ["A"], [, { name: "S", values: [1] }]);
  assert.ok(validatePresentation(missingSeries).some(item => item.code === "invalid-chart-series" && item.path.endsWith(".series.0")));
  const missingColor = chart("pie", ["A", "B", "C"], [{ name: "S", values: [1, 2, 3], pointColors: ["FF0000", , "0000FF"] }]);
  assert.ok(validatePresentation(missingColor).some(item => item.code === "invalid-chart-color" && item.path.endsWith(".pointColors.1")));
});

test('image fit validation locates ignored crops and rejects unknown modes without changing valid output', () => {
  const p = createPresentation();
  p.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/image.png' } });
  const box = { x: 0, y: 0, width: 20, height: 20 };
  p.addSlide({ id: 's', elements: [
    { id: 'default', type: 'image', assetId: 'img', box, crop: { left: .1 } },
    { id: 'stretch', type: 'image', assetId: 'img', box, fit: 'stretch', crop: { right: .2 } },
    { id: 'contain', type: 'image', assetId: 'img', box, fit: 'contain', crop: { top: .1 } },
    { id: 'cover', type: 'image', assetId: 'img', box, fit: 'cover', crop: { bottom: .1 } },
    { id: 'crop', type: 'image', assetId: 'img', box, fit: 'crop', crop: { left: .1 } },
    { id: 'zero', type: 'image', assetId: 'img', box, crop: { left: 0 } },
    { id: 'bad-crop', type: 'image', assetId: 'img', box, crop: { left: -1 } },
    { id: 'bad-fit', type: 'image', assetId: 'img', box, fit: 'invalid', crop: { left: .1 } },
  ] });
  const findings = validatePresentation(p);
  assert.deepEqual(findings.filter(d => d.code === 'image-crop-ignored').map(d => [d.elementId, d.path, d.severity]), [
    ['default', 'slides.0.elements.0.crop', 'warning'], ['stretch', 'slides.0.elements.1.crop', 'warning'],
    ['contain', 'slides.0.elements.2.crop', 'warning'], ['cover', 'slides.0.elements.3.crop', 'warning'],
  ]);
  assert.ok(findings.some(d => d.code === 'invalid-image-crop' && d.elementId === 'bad-crop'));
  assert.ok(findings.some(d => d.code === 'invalid-image-fit' && d.path === 'slides.0.elements.7.fit' && d.severity === 'error'));
  assert.throws(() => normalizePresentation(p), PresentationValidationError);
  const clean = createPresentation(); clean.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/image.png' } });
  clean.addSlide({ elements: [{ id: 'warning', type: 'image', assetId: 'img', box, crop: { left: .1 } }] });
  assert.equal(validatePresentation(clean).filter(d => d.code === 'image-crop-ignored').length, 1);
  assert.equal(normalizePresentation(clean).slides[0].elements[0].fit, 'stretch');
});

test('ignored crop warnings identify nested and shared image owners', () => {
  const p = createPresentation();
  p.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/image.png' } });
  const image = id => ({ id, type: 'image', assetId: 'img', fit: 'cover', crop: { left: .25 }, box: { x: 0, y: 0, width: 10, height: 10 } });
  const group = id => ({ id, type: 'group', box: { x: 0, y: 0, width: 20, height: 20 }, coordinateSize: { width: 20, height: 20 }, children: [image(`${id}-image`)] });
  p.defineSlideMaster({ id: 'm', name: 'Master', elements: [group('master-group')] });
  p.defineSlideLayout({ id: 'l', name: 'Layout', masterId: 'm', elements: [group('layout-group')] });
  p.addSlide({ id: 's', layoutId: 'l', elements: [group('slide-group')] });
  const warnings = validatePresentation(p).filter(d => d.code === 'image-crop-ignored');
  assert.deepEqual(warnings.map(d => [d.elementId, d.path, d.slideId, d.layoutId, d.masterId]), [
    ['layout-group-image', 'layouts.0.elements.0.children.0.crop', undefined, 'l', undefined],
    ['master-group-image', 'masters.0.elements.0.children.0.crop', undefined, undefined, 'm'],
    ['slide-group-image', 'slides.0.elements.0.children.0.crop', 's', undefined, undefined],
  ]);
});
