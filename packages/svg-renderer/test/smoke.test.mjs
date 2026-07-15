import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import { renderLayoutToSvg, renderPresentationToSvg } from "../dist/index.js";

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
    const lines = Array.from(svg.matchAll(/data-pptkit-text-line="\d+"><tspan [^>]+>([^<]+)<\/tspan><\/text>/g), (match) => match[1]);
    assert.deepEqual(lines, item.expected, item.id);
    assert.ok(!svg.includes("<foreignObject"));
  }
  const founderVisibility = result.slides.find((slide) => slide.slideId === "founder-visibility");
  assert.ok(founderVisibility);
  assert.match(founderVisibility.svg, /<text x="62\.5" y="181\.249682" text-anchor="start" data-pptkit-text-line="0">/);
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
