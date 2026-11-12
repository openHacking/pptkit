import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import { renderLayoutToSvg, renderPresentationToSvg, SvgRenderLimitError, SvgRenderAbortError } from "../dist/index.js";

test('mixed-run list markers ignore leading zero-character styles and retain the chosen full font/paint', async () => {
  const d = createPresentation();
  d.addSlide({ elements: [{ type: 'text', box: { x: 20, y: 20, width: 400, height: 200 }, content: [
    { style: { bullet: { type: 'number' } }, runs: [{ text: '', style: { colorOpacity: 0, fontSize: 12 } }, { text: 'Visible', style: { bold: true, underline: true, strike: true, fontFamily: 'Arial', fontSize: 24, color: '123456' } }] },
    { style: { bullet: { type: 'bullet' } }, runs: [{ text: '', style: { colorOpacity: 0, fontSize: 12 } }, { text: '', style: { bold: true, fontFamily: 'Georgia', fontSize: 18 } }] },
  ] }] });
  const xml = (await renderPresentationToSvg(d)).slides[0].svg;
  const numbered = xml.match(/<span aria-hidden="true" style="([^"]*)">1\. <\/span>/)[1];
  assert.match(numbered, /font-size:24px/); assert.match(numbered, /font-weight:700/); assert.match(numbered, /123456/); assert.match(numbered, /Arial/); assert.doesNotMatch(numbered, /opacity:0/);
  assert.match(numbered, /text-decoration:line-through/); assert.doesNotMatch(numbered, /underline/);
  const empty = xml.match(/<span aria-hidden="true" style="([^"]*)">• <\/span>/)[1];
  assert.match(empty, /font-size:18px/); assert.match(empty, /Georgia/); assert.doesNotMatch(empty, /opacity:0/);
  const single = createPresentation(); single.addSlide({ elements: [{ type: 'text', box: { x: 20, y: 20, width: 300, height: 100 }, content: [{ style: { bullet: { type: 'number' } }, runs: [{ text: 'Visible', style: { underline: true, strike: true } }] }] }] });
  const native = (await renderPresentationToSvg(single)).slides[0].svg;
  assert.match(native, /text-decoration="line-through"[^>]*>1\. <\/tspan>/);
  assert.match(native, /text-decoration="underline line-through"[^>]*>Visible<\/tspan>/);
  const clamped = createPresentation(); clamped.addSlide({ elements: [{ type: 'text', box: { x: 20, y: 20, width: 300, height: 100 }, frame: { margin: 0 }, content: [{ style: { indent: 0, hanging: 27 }, runs: [{ text: 'Plain' }] }] }] });
  assert.match((await renderPresentationToSvg(clamped)).slides[0].svg, /<text x="20"/);
});

function createFixture({ pathAsset = false } = {}) {
  const presentation = createPresentation({
    metadata: { title: "SVG <QA> & review", author: "PPTKit" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A" }, fonts: { heading: "Arial", body: "Arial" } },
  });
  presentation.defineSlideLayout({
    id: "fixture-layout",
    name: "Fixture",
    background: { type: "solid", color: "F8FAFC" },
    elements: [{
      type: "shape",
      id: "layout-band",
      name: "Layout band",
      shape: "rect",
      box: { x: 0, y: 0, width: 960, height: 18 },
      style: { fill: { type: "solid", color: { theme: "accent1" } } },
      accessibility: { decorative: true },
    }],
  });
  const image = presentation.registerAsset({
    id: "hero",
    kind: "image",
    source: pathAsset ? { type: "path", value: "/tmp/hero.png" } : { type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg'/%3E" },
    mimeType: "image/png",
    width: 400,
    height: 200,
    accessibility: { description: "Hero image" },
  });
  presentation.addSlide({
    id: "fixture-slide",
    layoutId: "fixture-layout",
    hidden: true,
    elements: [
      {
        type: "text",
        id: "rich-text",
        name: "Rich text",
        box: { x: 40, y: 34, width: 400, height: 80 },
        frame: { verticalAlign: "middle", autoFit: { mode: "shrink", fontScale: 0.9 } },
        content: [{
          style: { bullet: { type: "bullet", character: "•" }, lineSpacing: 1.2 },
          runs: [
            { text: "Hello <script>alert(1)</script> ", style: { fontSize: 22, bold: true } },
            { text: "PPTKit", style: { color: { theme: "accent2" }, italic: true }, action: { type: "url", url: "https://example.com?q=1&x=2" } },
          ],
        }],
      },
      {
        type: "shape",
        id: "shape-card",
        name: "Shape card",
        shape: "roundRect",
        box: { x: 40, y: 140, width: 240, height: 100 },
        transform: { rotation: 4 },
        style: { fill: { type: "solid", color: "DCE7FF" }, stroke: { paint: { type: "solid", color: "2457D6" }, dash: "dash" } },
        text: { content: "Shape text", frame: { verticalAlign: "middle" } },
      },
      {
        type: "image",
        id: "cropped-image",
        name: "Cropped image",
        assetId: image.id,
        fit: "crop",
        crop: { left: 0.1, right: 0.1, top: 0.2, bottom: 0.2 },
        box: { x: 320, y: 140, width: 240, height: 120 },
      },
      {
        type: "connector",
        id: "routed-connector",
        name: "Routed connector",
        start: { x: 80, y: 300 },
        route: [{ x: 240, y: 330 }],
        end: { x: 420, y: 300 },
        style: { paint: { type: "solid", color: { theme: "accent1" } }, width: 2, beginArrow: "oval", endArrow: "triangle" },
      },
      {
        type: "group",
        id: "nested-group",
        name: "Nested group",
        box: { x: 600, y: 40, width: 260, height: 140 },
        coordinateSize: { width: 260, height: 140 },
        opacity: 0.8,
        children: [
          { type: "shape", id: "group-shape", shape: "diamond", box: { x: 20, y: 20, width: 100, height: 100 }, style: { fill: { type: "solid", color: "FBCFE8" } } },
          { type: "text", id: "group-text", content: "Nested", box: { x: 130, y: 45, width: 100, height: 40 } },
        ],
      },
      {
        type: "table",
        id: "table",
        name: "Table",
        box: { x: 40, y: 360, width: 820, height: 120 },
        columns: [410, 410],
        rows: [
          { height: 50, cells: [{ content: "Merged", colSpan: 2, style: { fill: { type: "solid", color: "E0E7FF" } } }] },
          { height: 50, cells: [{ content: "A" }, { content: "B" }] },
        ],
      },
    ],
  });
  return presentation;
}

test("renders all current element families as stable standalone SVG", async () => {
  const presentation = createFixture();
  const before = normalizePresentation(presentation);
  const result = await renderPresentationToSvg(presentation);
  const again = await renderPresentationToSvg(presentation);

  assert.equal(result.status, "rendered");
  assert.equal(result.width, 960);
  assert.equal(result.height, 540);
  assert.equal(result.slides.length, 1);
  assert.equal(result.slides[0].hidden, true);
  assert.equal(result.slides[0].svg, again.slides[0].svg);
  assert.match(result.slides[0].svg, /^<svg[^>]+xmlns="http:\/\/www\.w3\.org\/2000\/svg"/);
  assert.match(result.slides[0].svg, /<foreignObject/);
  assert.match(result.slides[0].svg, /<image /);
  assert.match(result.slides[0].svg, /<polyline /);
  assert.match(result.slides[0].svg, /<table /);
  assert.match(result.slides[0].svg, /rowspan="1" colspan="2"/);
  assert.match(result.slides[0].svg, /marker-end=/);
  assert.ok(result.slides[0].svg.indexOf("Layout band") < result.slides[0].svg.indexOf("Rich text"));
  assert.ok(!result.slides[0].svg.includes("<script>"));
  assert.match(result.slides[0].svg, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.match(result.slides[0].svg, /q=1&amp;x=2/);
  assert.deepEqual(normalizePresentation(presentation), before);
});

test("keeps XHTML text and table lengths in SVG user units", async () => {
  const result = await renderPresentationToSvg(createFixture());
  const svg = result.slides[0].svg;

  assert.match(svg, /padding:3\.5px 7px 3\.5px 7px/);
  assert.match(svg, /font-size:19\.8px/);
  assert.match(svg, /<foreignObject[^>]+overflow="visible"[^>]+data-pptkit-element-id="rich-text"/);
  assert.match(svg, /overflow:visible;overflow-wrap:anywhere/);
  assert.match(svg, /width:410px/);
  assert.match(svg, /height:50px/);
  assert.doesNotMatch(svg, /<foreignObject[^>]+overflow="visible"[^>]+data-pptkit-element-id="table"/);
  assert.doesNotMatch(svg, /(?:font-size|padding|text-indent|width|height|border):[^;\"]*\bpt\b/);
});

test("uses native SVG baselines and PowerPoint-oriented wrapping for simple text", async () => {
  const presentation = createPresentation();
  const cases = [
    {
      id: "build-traction-title",
      text: "Build traction before launch.",
      fontSize: 43.5,
      fontScale: 0.96105,
      lineSpacing: 0.9,
      box: { x: 55.5, y: 94.5, width: 457.5, height: 97.5 },
      bold: true,
      expected: ["Build traction before", "launch."],
    },
    {
      id: "earlier-visibility-title",
      text: "Earlier visibility\nbecomes the product.",
      fontSize: 24.75,
      fontScale: 0.98555,
      lineSpacing: 0.96,
      box: { x: 598.5, y: 198, width: 213.75, height: 91.5 },
      bold: true,
      expected: ["Earlier visibility", "becomes the", "product."],
    },
    {
      id: "launch-title",
      text: "Launch-day momentum is too narrow.",
      fontSize: 33,
      fontScale: 0.78942,
      lineSpacing: 0.96,
      box: { x: 55.5, y: 84, width: 457.5, height: 66 },
      bold: true,
      expected: ["Launch-day momentum is too", "narrow."],
    },
    {
      id: "founder-visibility",
      text: "A founder needs visibility before the product is fully released, not only on the day it goes live.",
      fontSize: 17.25,
      fontScale: 0.92389,
      lineSpacing: 1.15,
      box: { x: 55.5, y: 165, width: 345, height: 67.5 },
      bold: false,
      expected: ["A founder needs visibility before the product", "is fully released, not only on the day it goes", "live."],
    },
    {
      id: "discovery-title",
      text: "The platform turns discovery into an ongoing loop.",
      fontSize: 31.5,
      fontScale: 0.81018,
      lineSpacing: 0.96,
      box: { x: 55.5, y: 82.5, width: 547.5, height: 66 },
      bold: true,
      expected: ["The platform turns discovery into an", "ongoing loop."],
    },
    {
      id: "coming-soon",
      text: "Coming Soon",
      fontSize: 19.5,
      fontScale: 1,
      lineSpacing: 1,
      box: { x: 72, y: 324, width: 154.5, height: 43.5 },
      bold: true,
      expected: ["Coming Soon"],
    },
    {
      id: "discovery-becomes-easier",
      text: "Discovery becomes easier.",
      fontSize: 21.75,
      fontScale: 0.92595,
      lineSpacing: 1,
      box: { x: 504, y: 243, width: 255, height: 39 },
      bold: true,
      expected: ["Discovery becomes", "easier."],
    },
  ];
  for (const item of cases) presentation.addSlide({
    id: item.id,
    elements: [{
      type: "text",
      id: item.id,
      content: [{
        style: { lineSpacing: item.lineSpacing },
        runs: [{
          text: item.text,
          style: { fontSize: item.fontSize, fontFamily: "Helvetica Neue", bold: item.bold },
        }],
      }],
      box: item.box,
      frame: { autoFit: { mode: "shrink", fontScale: item.fontScale } },
    }],
  });

  const result = await renderPresentationToSvg(presentation);
  for (const [index, item] of cases.entries()) {
    const svg = result.slides[index].svg;
    const lines = Array.from(svg.matchAll(/data-pptkit-text-line="\d+"><tspan [^>]+>([^<]+)<\/tspan>/g), (match) => match[1]);
    assert.deepEqual(lines, item.expected, item.id);
    assert.ok(!svg.includes("<foreignObject"));
  }
  const founderVisibility = result.slides.find((slide) => slide.slideId === "founder-visibility");
  assert.ok(founderVisibility);
  assert.match(founderVisibility.svg, /<text x="62\.5" y="181\.249682" text-anchor="start" xml:space="preserve" style="white-space:pre" data-pptkit-text-line="0">/);
});

test("wraps long single-run bullet paragraphs instead of overflowing the frame", async () => {
  const presentation = createPresentation();
  presentation.addSlide({
    id: "bullet-wrap",
    elements: [{
      type: "text",
      id: "bullet-copy",
      content: [{
        style: { bullet: { type: "bullet" }, indent: 22, hanging: 22, lineSpacing: 1.12 },
        runs: [{ text: "Long-term visibility through listings and Coming Soon pages", style: { fontFamily: "Arial", fontSize: 18 } }],
      }],
      box: { x: 526, y: 258, width: 350, height: 100 },
      frame: { autoFit: { mode: "shrink", fontScale: 0.9 } },
    }],
  });
  const result = await renderPresentationToSvg(presentation);
  const svg = result.slides[0].svg;
  assert.match(svg, /data-pptkit-text-line="1"/);
  assert.match(svg, /<text x="555"[^>]+data-pptkit-text-line="1"/);
  assert.doesNotMatch(svg, />• Long-term visibility through listings and Coming Soon pages</);
});

test("explicit SVG namespaces are deterministic, isolated and validated", async () => {
  const presentation = createFixture();
  const layout = resolveNormalizedLayout(normalizePresentation(presentation));
  const baseline = await renderPresentationToSvg(presentation);
  assert.deepEqual(await renderPresentationToSvg(presentation, { idNamespace: undefined }), baseline);
  const key = `copy / <&"😀`;
  const first = await renderPresentationToSvg(presentation, { idNamespace: key });
  const again = await renderPresentationToSvg(presentation, { idNamespace: key });
  assert.deepEqual(first, again);
  assert.deepEqual(await renderLayoutToSvg(layout, { idNamespace: key }), first);
  const other = await renderPresentationToSvg(presentation, { idNamespace: `${key}2` });
  const ids = result => [...result.slides[0].svg.matchAll(/\sid="([^"]+)"/g)].map(match => match[1]);
  assert.ok(ids(first).length > 2);
  assert.ok(ids(first).every(id => !ids(other).includes(id) && !ids(baseline).includes(id)));
  for (const invalid of ["", null, 3, {}]) {
    await assert.rejects(renderPresentationToSvg(presentation, { idNamespace: invalid }), /idNamespace must be a nonempty string/);
    await assert.rejects(renderLayoutToSvg(layout, { idNamespace: invalid }), /idNamespace must be a nonempty string/);
  }
});

test("renderLayoutToSvg matches the authoring convenience API", async () => {
  const presentation = createFixture();
  const layout = resolveNormalizedLayout(normalizePresentation(presentation));
  const snapshot = structuredClone(layout);
  const [fromDocument, fromLayout] = await Promise.all([
    renderPresentationToSvg(presentation),
    renderLayoutToSvg(layout),
  ]);
  assert.deepEqual(fromLayout, fromDocument);
  assert.deepEqual(layout, snapshot);
});

test("path assets degrade visibly and a custom resolver restores them", async () => {
  const presentation = createFixture({ pathAsset: true });
  const missing = await renderPresentationToSvg(presentation);
  assert.equal(missing.status, "rendered-with-warnings");
  assert.ok(missing.warnings.some((warning) => warning.code === "asset-path-unsupported"));
  assert.match(missing.slides[0].svg, /stroke-dasharray="6 4"/);

  const resolved = await renderPresentationToSvg(presentation, {
    resolveAsset: () => "data:image/png;base64,AA==",
  });
  assert.equal(resolved.status, "rendered");
  assert.match(resolved.slides[0].svg, /data:image\/png;base64,AA==/);
});

test("asset resolver failures and unsafe actions are isolated as warnings", async () => {
  const presentation = createFixture({ pathAsset: true });
  presentation.slides[0].addElement({
    type: "shape",
    id: "unsafe-action",
    shape: "rect",
    box: { x: 880, y: 20, width: 40, height: 40 },
    action: { type: "url", url: "javascript:alert(1)" },
  });
  presentation.slides[0].addElement({
    type: "shape",
    id: "data-action",
    shape: "ellipse",
    box: { x: 880, y: 70, width: 40, height: 40 },
    action: { type: "url", url: "data:text/html,<script>alert(1)</script>" },
  });
  const result = await renderPresentationToSvg(presentation, {
    resolveAsset: async () => { throw new Error("fixture failure"); },
  });
  assert.ok(result.warnings.some((warning) => warning.code === "asset-resolver-failed"));
  assert.ok(result.warnings.some((warning) => warning.code === "unsafe-action-url"));
  assert.ok(!result.slides[0].svg.includes("javascript:"));
  assert.ok(!result.slides[0].svg.includes("data:text/html"));
  assert.equal(result.slides.length, 1);
});

test("non-image data URLs from asset resolvers are rejected", async () => {
  const presentation = createFixture({ pathAsset: true });
  const result = await renderPresentationToSvg(presentation, {
    resolveAsset: () => "data:text/html,<script>alert(1)</script>",
  });
  assert.ok(result.warnings.some((warning) => warning.code === "unsafe-asset-url"));
  assert.ok(!result.slides[0].svg.includes("data:text/html"));
});

test("autoFit resize and slide actions report explicit fidelity warnings", async () => {
  const presentation = createPresentation();
  presentation.addSlide({ id: "target" });
  presentation.addSlide({
    id: "source",
    elements: [{
      type: "text",
      id: "resizing-link",
      content: "Open target",
      box: { x: 20, y: 20, width: 200, height: 60 },
      frame: { autoFit: { mode: "resize" } },
      action: { type: "slide", slideId: "target" },
    }],
  });
  const result = await renderPresentationToSvg(presentation);
  assert.ok(result.warnings.some((warning) => warning.code === "text-autofit-resize-unsupported"));
  assert.ok(result.warnings.some((warning) => warning.code === "slide-action-not-linked"));
  assert.match(result.slides[1].svg, /data-pptkit-slide-target="target"/);
});

test("renders bar, line, and pie charts as deterministic SVG", async () => {
  const presentation = createPresentation({
    metadata: { title: "Chart test" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A" } },
  });
  const slide = presentation.addSlide({ id: "chart-slide" });
  slide.addElement({
    type: "chart",
    id: "bar-chart",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue", values: [120, 200, 150], color: "2457D6" },
      { name: "Cost", values: [80, 140, 110], color: "E65A3A" },
    ],
    title: "Bar Chart",
    legend: { visible: true },
    axes: { category: { majorGridlines: { paint: { type: "solid", color: "ABCDEF" }, width: 0.5 } } },
    box: { x: 40, y: 40, width: 280, height: 200 },
  });
  slide.addElement({
    type: "chart",
    id: "line-chart",
    chartType: "line",
    categories: ["A", "B", "C"],
    series: [
      { name: "Trend", values: [10, 30, 20], color: "22C55E" },
    ],
    title: "Line Chart",
    legend: { visible: false },
    box: { x: 340, y: 40, width: 280, height: 200 },
  });
  slide.addElement({
    type: "chart",
    id: "pie-chart",
    chartType: "pie",
    categories: ["X", "Y", "Z"],
    series: [
      { name: "Share", values: [40, 35, 25], color: "8B5CF6" },
    ],
    title: "Pie Chart",
    legend: { visible: true },
    box: { x: 640, y: 40, width: 280, height: 200 },
  });

  const result = await renderPresentationToSvg(presentation);
  const again = await renderPresentationToSvg(presentation);
  assert.equal(result.status, "rendered");
  assert.equal(result.slides[0].svg, again.slides[0].svg);

  const svg = result.slides[0].svg;
  assert.match(svg, /<rect[^>]+fill="#2457D6"/);
  assert.match(svg, /<polyline[^>]+stroke="#22C55E"/);
  assert.match(svg, /<path[^>]+d="M[^"]+A[^"]+Z"/);
  assert.match(svg, /<text[^>]*>Bar Chart<\/text>/);
  assert.match(svg, /<text[^>]*>Line Chart<\/text>/);
  assert.match(svg, /<text[^>]*>Pie Chart<\/text>/);
  assert.match(svg, /<text[^>]*>Q1<\/text>/);
  assert.match(svg, /<text[^>]*>Q2<\/text>/);
  assert.match(svg, /<text[^>]*>Q3<\/text>/);
  assert.match(svg, /<text[^>]*>Revenue<\/text>/);
  assert.match(svg, /<text[^>]*>X<\/text>/);
  assert.match(svg, /<text[^>]*>Y<\/text>/);
  assert.match(svg, /<text[^>]*>Z<\/text>/);
  assert.doesNotMatch(svg, /<text[^>]*>Share<\/text>/);
  const barFragment = svg.slice(svg.indexOf('data-pptkit-element-id="bar-chart"'), svg.indexOf('data-pptkit-element-id="line-chart"'));
  assert.ok(barFragment.indexOf('data-chart-layer="grid"') < barFragment.indexOf('data-chart-layer="series"'), "grid layer renders before bars");
  assert.equal([...barFragment.matchAll(/data-chart-role="category-gridline"/g)].length, 4);
  const lineFragment = svg.slice(svg.indexOf('data-pptkit-element-id="line-chart"'), svg.indexOf('data-pptkit-element-id="pie-chart"'));
  assert.ok(lineFragment.indexOf('data-chart-layer="grid"') < lineFragment.indexOf('data-chart-layer="series"'), "grid layer renders before lines");
});

test("chart box origin, y-axis ticks, pie slice colors, and safe geometry", async () => {
  const presentation = createPresentation({
    metadata: { title: "Chart geometry" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A" } },
  });
  const slide = presentation.addSlide({ id: "geometry-slide" });
  slide.addElement({
    type: "chart",
    id: "bar-with-yaxis",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [{ name: "Revenue", values: [120, 200, 150], color: "2457D6" }],
    title: "Bar",
    legend: { visible: false },
    box: { x: 100, y: 50, width: 240, height: 180 },
  });
  slide.addElement({
    type: "chart",
    id: "bar-no-yaxis",
    chartType: "bar",
    categories: ["A", "B"],
    series: [{ name: "Count", values: [3, 7], color: "22C55E" }],
    legend: { visible: false },
    axes: { value: { visible: false, labels: false, majorGridlines: false } },
    box: { x: 360, y: 50, width: 200, height: 180 },
  });
  slide.addElement({
    type: "chart",
    id: "pie-slices",
    chartType: "pie",
    categories: ["A", "B", "C", "D"],
    series: [{ name: "Share", values: [10, 20, 30, 40], pointColors: ["FF0000", "00FF00", "0000FF", "FFFF00"] }],
    legend: { visible: false },
    box: { x: 580, y: 50, width: 200, height: 180 },
  });

  const layout = resolveNormalizedLayout(normalizePresentation(presentation));

  const result = await renderLayoutToSvg(layout);
  assert.equal(result.status, "rendered");
  const svg = result.slides[0].svg;

  assert.match(svg, /<g transform="translate\(100 50\)">/);

  const noYaxisStart = svg.indexOf('data-pptkit-element-id="bar-no-yaxis"');
  const noYaxisEnd = svg.indexOf('data-pptkit-element-id="pie-slices"');
  const noYaxisFragment = svg.slice(noYaxisStart, noYaxisEnd);
  assert.doesNotMatch(noYaxisFragment, /text-anchor="end">-?\d/);

  const pieStart = svg.indexOf('data-pptkit-element-id="pie-slices"');
  const pieFragment = svg.slice(pieStart);
  const piePaths = Array.from(pieFragment.matchAll(/<path[^>]*fill="(#[0-9A-Fa-f]+)"/g), (match) => match[1]);
  assert.equal(piePaths.length, 4, "pie renders one path per positive category");
  assert.deepEqual(piePaths, ["#FF0000", "#00FF00", "#0000FF", "#FFFF00"]);

  assert.ok(!svg.includes("NaN"), "no NaN in svg");
  assert.ok(!svg.includes("Infinity"), "no Infinity in svg");

  const barOuter = svg.slice(svg.indexOf('data-pptkit-element-id="bar-with-yaxis"'));
  assert.match(barOuter, /opacity="/);
  assert.match(barOuter, /data-pptkit-element-id="bar-with-yaxis"/);
});

test("chart legend placement follows structured legend config with escaped labels", async () => {
  const presentation = createPresentation({
    metadata: { title: "Legend layout" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A", accent3: "22C55E" } },
  });
  const slide = presentation.addSlide({ id: "legend-slide" });
  slide.addElement({
    type: "chart",
    id: "legend-right",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue & Cost", values: [120, 200, 150], color: "2457D6" },
      { name: "Profit", values: [40, 60, 35], color: "E65A3A" },
    ],
    title: "Right Legend",
    legend: { visible: true, position: "right" },
    box: { x: 40, y: 40, width: 280, height: 200 },
  });
  slide.addElement({
    type: "chart",
    id: "legend-bottom",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue & Cost", values: [120, 200, 150], color: "2457D6" },
      { name: "Profit", values: [40, 60, 35], color: "E65A3A" },
    ],
    title: "Bottom Legend",
    legend: { visible: true, position: "bottom" },
    box: { x: 360, y: 40, width: 280, height: 200 },
  });
  slide.addElement({
    type: "chart",
    id: "legend-none",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [
      { name: "Revenue & Cost", values: [120, 200, 150], color: "2457D6" },
      { name: "Profit", values: [40, 60, 35], color: "E65A3A" },
    ],
    title: "Hidden Legend",
    legend: { visible: false },
    box: { x: 680, y: 40, width: 280, height: 200 },
  });

  const result = await renderPresentationToSvg(presentation);
  assert.equal(result.status, "rendered");
  const svg = result.slides[0].svg;

  const fragmentBetween = (fromId, toId) => {
    const start = svg.indexOf(`data-pptkit-element-id="${fromId}"`);
    const end = svg.indexOf(`data-pptkit-element-id="${toId}"`);
    assert.ok(start >= 0, `${fromId} anchor present`);
    assert.ok(end > start, `${toId} anchor follows ${fromId}`);
    return svg.slice(start, end);
  };
  const fragmentFrom = (id) => {
    const start = svg.indexOf(`data-pptkit-element-id="${id}"`);
    assert.ok(start >= 0, `${id} anchor present`);
    return svg.slice(start);
  };
  const matchXAxisLine = (fragment) => {
    const match = fragment.match(
      /<line[^>]*x1="([\d.]+)"[^>]*y1="([\d.]+)"[^>]*x2="([\d.]+)"[^>]*y2="([\d.]+)"[^>]*data-chart-role="category-axis"/,
    );
    assert.ok(match, "x-axis line found in chart fragment");
    const x1 = Number(match[1]);
    const y1 = Number(match[2]);
    const x2 = Number(match[3]);
    const y2 = Number(match[4]);
    assert.equal(y1, y2, "matched line is horizontal (x-axis)");
    return { x1, y1, x2, y2 };
  };
  const matchLegendText = (fragment) => {
    const match = fragment.match(
      /<text x="([\d.]+)" y="([\d.]+)"[^>]*data-chart-role="legend-label">([^<]+)<\/text>/,
    );
    assert.ok(match, "legend label found in chart fragment");
    return { x: Number(match[1]), y: Number(match[2]), label: match[3] };
  };

  const rightFragment = fragmentBetween("legend-right", "legend-bottom");
  assert.match(rightFragment, /Revenue &amp; Cost/);
  assert.match(rightFragment, /Profit/);
  const rightAxis = matchXAxisLine(rightFragment);
  const rightLegend = matchLegendText(rightFragment);
  assert.equal(rightLegend.label, "Revenue &amp; Cost");
  assert.ok(
    rightLegend.x > rightAxis.x2,
    `right legend text x (${rightLegend.x}) must exceed x-axis x2 (${rightAxis.x2})`,
  );

  const bottomFragment = fragmentBetween("legend-bottom", "legend-none");
  assert.match(bottomFragment, /Revenue &amp; Cost/);
  assert.match(bottomFragment, /Profit/);
  const bottomAxis = matchXAxisLine(bottomFragment);
  const bottomLegend = matchLegendText(bottomFragment);
  assert.ok(
    bottomLegend.y > bottomAxis.y2,
    `bottom legend text y (${bottomLegend.y}) must exceed x-axis y2 (${bottomAxis.y2})`,
  );

  const noneFragment = fragmentFrom("legend-none");
  assert.doesNotMatch(noneFragment, /data-chart-role="legend-label"/);
  assert.doesNotMatch(noneFragment, /Revenue &amp; Cost/);
  assert.doesNotMatch(noneFragment, />Profit</);
});

test("line series and legend independently honor line and marker visibility", async () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide({ id: "line-visibility-slide" });
  for (const [id, marker] of [["both", "auto"], ["line-only", false], ["marker-only", "auto"], ["neither", false]]) {
    slide.addElement({ type: "chart", id, chartType: "line", categories: ["A", "B"],
      series: [{ name: id, values: [2, 5], marker }], legend: { visible: true },
      box: { x: 0, y: 0, width: 250, height: 160 } });
  }
  const layout = resolveNormalizedLayout(normalizePresentation(presentation));
  for (const element of layout.slides[0].elements) {
    if (element.id === "marker-only" || element.id === "neither") element.series[0].line = false;
  }
  const svg = (await renderLayoutToSvg(layout)).slides[0].svg;
  const ids = ["both", "line-only", "marker-only", "neither"];
  for (const [index, id] of ids.entries()) {
    const start = svg.indexOf(`data-pptkit-element-id="${id}"`);
    const end = index + 1 === ids.length ? svg.length : svg.indexOf(`data-pptkit-element-id="${ids[index + 1]}"`);
    assert.ok(start >= 0 && end > start);
    const fragment = svg.slice(start, end);
    const hasLine = id === "both" || id === "line-only";
    const hasMarker = id === "both" || id === "marker-only";
    const count = role => [...fragment.matchAll(new RegExp(`data-chart-role="${role}"`, "g"))].length;
    assert.equal(count("line"), Number(hasLine), `${id}: series line`);
    assert.equal(count("series-marker"), hasMarker ? 2 : 0, `${id}: data markers`);
    assert.equal(count("legend-line"), Number(hasLine), `${id}: legend line`);
    assert.equal(count("legend-marker"), Number(hasMarker), `${id}: legend marker`);
    assert.equal(count("legend-swatch"), 0, `${id}: no bar-style swatch`);
    assert.match(fragment, new RegExp(`>${id}</text>`), `${id}: legend label`);
  }
});

test("chart axis labels and y-axis ticks follow show flags, and category labels are XML-escaped", async () => {
  const presentation = createPresentation({
    metadata: { title: "Axis and escape" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A" } },
  });
  const slide = presentation.addSlide({ id: "axis-escape-slide" });
  slide.addElement({
    type: "chart",
    id: "bar-with-ticks",
    chartType: "bar",
    categories: ["Q1", "Q2", "Q3"],
    series: [{ name: "Revenue", values: [120, 200, 150], color: "2457D6" }],
    legend: { visible: false },
    box: { x: 40, y: 40, width: 240, height: 180 },
  });
  slide.addElement({
    type: "chart",
    id: "bar-no-xlabels",
    chartType: "bar",
    categories: ["A", "B"],
    series: [{ name: "Count", values: [3, 7], color: "22C55E" }],
    legend: { visible: false },
    axes: { category: { labels: false }, value: { visible: false, labels: false, majorGridlines: false } },
    box: { x: 300, y: 40, width: 200, height: 180 },
  });
  slide.addElement({
    type: "chart",
    id: "bar-escaped-cat",
    chartType: "bar",
    categories: ["Q1 <script>alert(1)</script>", "Q2"],
    series: [{ name: "Rev", values: [10, 20], color: "2457D6" }],
    legend: { visible: true, position: "right" },
    box: { x: 520, y: 40, width: 200, height: 180 },
  });

  const result = await renderPresentationToSvg(presentation);
  assert.equal(result.status, "rendered");
  const svg = result.slides[0].svg;

  const withTicksFragment = svg.slice(
    svg.indexOf('data-pptkit-element-id="bar-with-ticks"'),
    svg.indexOf('data-pptkit-element-id="bar-no-xlabels"'),
  );
  assert.match(withTicksFragment, /text-anchor="end"[^>]*data-chart-role="value-label">\d/);
  const tickLabels = Array.from(
    withTicksFragment.matchAll(/<text[^>]*text-anchor="end"[^>]*data-chart-role="value-label">(-?\d+(?:\.\d+)?)<\/text>/g),
    (match) => match[1],
  );
  assert.deepEqual(tickLabels, ["0", "50", "100", "150", "200", "250"], "nice-number value-axis ticks with headroom");
  const gridlineCount = (withTicksFragment.match(/data-chart-role="gridline"/g) ?? []).length;
  const valueTickCount = (withTicksFragment.match(/data-chart-role="value-tick"/g) ?? []).length;
  const categoryTickCount = (withTicksFragment.match(/data-chart-role="category-tick"/g) ?? []).length;
  assert.ok(gridlineCount >= 4, `major gridlines present (got ${gridlineCount})`);
  assert.ok(valueTickCount >= 5, `value-axis tick marks present (got ${valueTickCount})`);
  assert.ok(categoryTickCount >= 3, `category-axis tick marks present (got ${categoryTickCount})`);

  const noXLabelsFragment = svg.slice(
    svg.indexOf('data-pptkit-element-id="bar-no-xlabels"'),
    svg.indexOf('data-pptkit-element-id="bar-escaped-cat"'),
  );
  assert.doesNotMatch(noXLabelsFragment, /<text[^>]*>A<\/text>/);
  assert.doesNotMatch(noXLabelsFragment, /<text[^>]*>B<\/text>/);
  assert.doesNotMatch(noXLabelsFragment, /data-chart-role="gridline"/);
  assert.doesNotMatch(noXLabelsFragment, /data-chart-role="value-label"/);

  const escapedFragment = svg.slice(svg.indexOf('data-pptkit-element-id="bar-escaped-cat"'));
  assert.match(escapedFragment, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
});

test("line chart legend uses a line swatch while bar and pie use filled squares", async () => {
  const presentation = createPresentation({
    metadata: { title: "Legend swatch" },
    theme: { colors: { accent1: "2457D6", accent2: "E65A3A", accent3: "22C55E" } },
  });
  const slide = presentation.addSlide({ id: "legend-swatch-slide" });
  slide.addElement({
    type: "chart",
    id: "line-with-legend",
    chartType: "line",
    categories: ["A", "B", "C"],
    series: [{ name: "Trend", values: [10, 30, 20], color: "22C55E" }],
    legend: { visible: true, position: "right" },
    box: { x: 40, y: 40, width: 280, height: 200 },
  });
  slide.addElement({
    type: "chart",
    id: "pie-with-legend",
    chartType: "pie",
    categories: ["X", "Y", "Z"],
    series: [{ name: "Share", values: [40, 35, 25], color: "8B5CF6" }],
    legend: { visible: true, position: "right" },
    box: { x: 360, y: 40, width: 280, height: 200 },
  });

  const result = await renderPresentationToSvg(presentation);
  assert.equal(result.status, "rendered");
  const svg = result.slides[0].svg;

  const lineFragment = svg.slice(
    svg.indexOf('data-pptkit-element-id="line-with-legend"'),
    svg.indexOf('data-pptkit-element-id="pie-with-legend"'),
  );
  assert.match(lineFragment, /<line[^>]+stroke="#22C55E"[^>]+data-chart-role="legend-line"/);
  assert.match(lineFragment, /data-chart-role="legend-marker"/);
  assert.doesNotMatch(lineFragment, /<rect[^>]+data-chart-role="legend-swatch"/);
  assert.match(lineFragment, /<text[^>]*>Trend<\/text>/);

  const pieFragment = svg.slice(svg.indexOf('data-pptkit-element-id="pie-with-legend"'));
  const pieSwatches = Array.from(
    pieFragment.matchAll(/<rect[^>]+fill="(#[0-9A-Fa-f]+)"[^>]+data-chart-role="legend-swatch"/g),
    (match) => match[1],
  );
  assert.equal(pieSwatches.length, 3, "pie legend shows one swatch per slice");
  assert.match(pieFragment, /<text[^>]*>X<\/text>/);
  assert.match(pieFragment, /<text[^>]*>Y<\/text>/);
  assert.match(pieFragment, /<text[^>]*>Z<\/text>/);
  assert.doesNotMatch(pieFragment, /<text[^>]*>Share<\/text>/);
});

test('slide selection renders only requested pages, preserves source indices and scopes warnings', async () => {
  const d = createPresentation({ id: 'selection', metadata: { title: 'Selected' } });
  for (const id of ['one', 'two', 'three']) {
    d.registerAsset({ id: `asset-${id}`, kind: 'image', source: { type: 'url', value: `https://example.invalid/${id}.png` }, mimeType: 'image/png', width: 10, height: 10 });
    d.addSlide({ id, hidden: id === 'two', elements: [{ id: `image-${id}`, type: 'image', assetId: `asset-${id}`, box: { x: 10, y: 10, width: 10, height: 10 } }] });
  }
  const calls = [], options = { resolveAsset: async asset => { calls.push(asset.id); return `data:image/png;base64,AA==`; } };
  const full = await renderPresentationToSvg(d, options); assert.deepEqual(calls, ['asset-one', 'asset-two', 'asset-three']); calls.length = 0;
  const selected = await renderPresentationToSvg(d, { ...options, slideIndices: [2, 0] });
  assert.deepEqual(calls, ['asset-three', 'asset-one']); assert.deepEqual(selected.slides.map(s => s.index), [2, 0]);
  assert.equal(selected.slides[0].svg, full.slides[2].svg); assert.equal(selected.slides[1].svg, full.slides[0].svg);
  assert.deepEqual(await renderPresentationToSvg(d, { ...options, slideIndices: [] }), { width: full.width, height: full.height, slides: [], warnings: [], status: 'rendered' });
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  layout.analysis = { diagnostics: [
    { severity: 'warning', code: 'one', message: 'one', path: 'a', slideId: 'one' },
    { severity: 'warning', code: 'two', message: 'two', path: 'b', slideId: 'two' },
    { severity: 'warning', code: 'global', message: 'global', path: 'c' },
  ] };
  const scoped = await renderLayoutToSvg(layout, { ...options, slideIndices: [1] });
  assert.deepEqual(scoped.warnings.map(w => w.code), ['two', 'global']); assert.equal(scoped.slides[0].hidden, true);
  for (const invalid of [[-1], [3], [0, 0], [1.2], [NaN], [undefined], Array(1)]) await assert.rejects(renderLayoutToSvg(layout, { ...options, slideIndices: invalid }), RangeError);
  await assert.rejects(renderLayoutToSvg(layout, { slideIndices: '1' }), TypeError);
});

test('slide selection is captured before awaiting an asset resolver', async () => {
  const d = createPresentation();
  d.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a' }, mimeType: 'image/png', width: 1, height: 1 });
  d.addSlide({ id: 'a', elements: [{ type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 10, height: 10 } }] }); d.addSlide({ id: 'b' });
  let release; const gate = new Promise(resolve => { release = resolve; });
  const indices = [0, 1], pending = renderPresentationToSvg(d, { slideIndices: indices, resolveAsset: async () => { await gate; return 'data:image/png;base64,AA=='; } });
  indices.splice(0, 2, 1); release(); const result = await pending; assert.deepEqual(result.slides.map(s => s.slideId), ['a', 'b']);
});

test('SVG byte budget counts UTF-8 exactly across slides and stops before later assets', async () => {
  const d = createPresentation({ id: 'bytes', metadata: { title: '中文 😀' } });
  for (const id of ['one', 'two', 'three']) {
    d.registerAsset({ id, kind: 'image', source: { type: 'url', value: `https://example.invalid/${id}` }, mimeType: 'image/png', width: 2, height: 2 });
    d.addSlide({ id, elements: [{ id: `text-${id}`, type: 'text', box: { x: 0, y: 0, width: 150, height: 30 }, content: 'é 中 😀 \ud800' }, { id: `image-${id}`, type: 'image', assetId: id, box: { x: 200, y: 0, width: 20, height: 20 } }] });
  }
  const calls = [], resolveAsset = async asset => { calls.push(asset.id); return 'data:image/png;base64,AA=='; };
  const all = await renderPresentationToSvg(d, { resolveAsset });
  const sizes = all.slides.map(slide => Buffer.byteLength(slide.svg, 'utf8')), total = sizes.reduce((a, b) => a + b, 0);
  const exact = await renderPresentationToSvg(d, { resolveAsset, maxSvgBytes: total }); assert.deepEqual(exact, all);
  const layout = resolveNormalizedLayout(normalizePresentation(d));
  assert.deepEqual(await renderLayoutToSvg(layout, { resolveAsset, maxSvgBytes: total }), all);
  await assert.rejects(renderLayoutToSvg(layout, { resolveAsset, maxSvgBytes: total - 1 }), SvgRenderLimitError);
  calls.length = 0;
  await assert.rejects(renderPresentationToSvg(d, { resolveAsset, maxSvgBytes: sizes[0] + sizes[1] - 1 }), error => error instanceof SvgRenderLimitError && error.code === 'svg-limit-exceeded' && error.slideIndex === 1 && error.limit === sizes[0] + sizes[1] - 1);
  assert.deepEqual(calls, ['one', 'two']);
  const selected = await renderPresentationToSvg(d, { resolveAsset, slideIndices: [2], maxSvgBytes: sizes[2] }); assert.equal(selected.slides[0].svg, all.slides[2].svg);
  await assert.rejects(renderPresentationToSvg(d, { slideIndices: [2], resolveAsset, maxSvgBytes: sizes[2] - 1 }), SvgRenderLimitError);
  const empty = await renderPresentationToSvg(d, { slideIndices: [], maxSvgBytes: 0 }); assert.deepEqual(empty.slides, []);
  for (const invalid of [-1, 1.1, NaN, Infinity, 9007199254740992, '5']) await assert.rejects(renderPresentationToSvg(d, { maxSvgBytes: invalid }), RangeError);
});

test('abort interrupts a pending resolver, removes listeners and permits a later retry', async () => {
  const d = createPresentation(); d.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a' }, mimeType: 'image/png', width: 1, height: 1 });
  d.addSlide({ id: 'a', elements: [{ type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 10, height: 10 } }] });
  const controller = new AbortController(), signal = controller.signal; let added = 0, removed = 0;
  const add = signal.addEventListener.bind(signal), remove = signal.removeEventListener.bind(signal);
  signal.addEventListener = (...args) => { if (args[0] === 'abort') added++; return add(...args); };
  signal.removeEventListener = (...args) => { if (args[0] === 'abort') removed++; return remove(...args); };
  let started; const ready = new Promise(resolve => { started = resolve; }); let rejectLate;
  const options = { signal, resolveAsset: (_asset, context) => { assert.equal(context.signal, signal); started(); return new Promise((_resolve, reject) => { rejectLate = reject; }); } };
  const pending = renderPresentationToSvg(d, options);
  await ready; options.signal = new AbortController().signal; controller.abort(); await assert.rejects(pending, error => error instanceof SvgRenderAbortError && error.code === 'operation-aborted');
  assert.equal(added, removed); rejectLate(new Error('late resolver rejection')); await new Promise(resolve => setImmediate(resolve));
  const retried = await renderPresentationToSvg(d, { resolveAsset: async () => 'data:image/png;base64,AA==' }); assert.equal(retried.slides.length, 1); assert.ok(!retried.warnings.some(w => w.code === 'asset-resolver-failed'));
  const success = new AbortController(); let successAdded = 0, successRemoved = 0;
  const on = success.signal.addEventListener.bind(success.signal), off = success.signal.removeEventListener.bind(success.signal);
  success.signal.addEventListener = (...args) => { if (args[0] === 'abort') successAdded++; return on(...args); };
  success.signal.removeEventListener = (...args) => { if (args[0] === 'abort') successRemoved++; return off(...args); };
  const complete = await renderPresentationToSvg(d, { signal: success.signal, resolveAsset: async () => 'data:image/png;base64,AA==' });
  assert.equal(successAdded, 1); assert.equal(successRemoved, 1); success.abort(); assert.equal(complete.slides.length, 1);
});

test('abort is distinct from ordinary resolver failures at invocation and completion boundaries', async () => {
  const d = createPresentation(); d.registerAsset({ id: 'img', kind: 'image', source: { type: 'url', value: 'https://example.invalid/a' }, mimeType: 'image/png', width: 1, height: 1 });
  d.addSlide({ elements: [{ type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 10, height: 10 } }] });
  const pre = new AbortController(); pre.abort(); let called = false;
  await assert.rejects(renderPresentationToSvg(d, { signal: pre.signal, resolveAsset: () => { called = true; return undefined; } }), SvgRenderAbortError); assert.equal(called, false);
  const layout = resolveNormalizedLayout(normalizePresentation(d)); await assert.rejects(renderLayoutToSvg(layout, { signal: pre.signal }), SvgRenderAbortError);
  await assert.rejects(renderPresentationToSvg(d, { signal: {} }), TypeError);
  const sync = new AbortController();
  await assert.rejects(renderPresentationToSvg(d, { signal: sync.signal, resolveAsset: () => { sync.abort(); return Promise.reject(new Error('simultaneous')); } }), SvgRenderAbortError);
  await new Promise(resolve => setImmediate(resolve));
  const normal = await renderPresentationToSvg(d, { resolveAsset: async () => { throw new Error('network'); } });
  assert.ok(normal.warnings.some(w => w.code === 'asset-resolver-failed')); assert.equal(normal.status, 'rendered-with-warnings');
});
