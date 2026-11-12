import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { pixelPng } from "../tests/fixtures/presentations.mjs";
const ref = (elementId, metric, factor = 1, constant = 0) => ({ terms: [{ source: { elementId }, metric, factor }], constant });
function fixture(sized = false) {
  const document = createPresentation({ size: { width: 500, height: 300 } });
  const image = document.registerAsset({ kind: "image", id: "asset", source: { type: "url", value: pixelPng }, mimeType: "image/png", width: 8, height: 4 });
  const slide = document.addSlide({ id: "slide" });
  slide.addElement({ type: "shape", id: "a", shape: "rect", box: { x: 10, y: 20, width: sized ? 100 : 40, height: 30 } });
  slide.addElement({ type: "image", id: "b", assetId: image.id, fit: "contain", box: { x: sized ? 120 : 50, y: 20, width: sized ? 100 : 40, height: 60 } });
  slide.addElement({ type: "chart", id: "c", chartType: "bar", categories: ["A"], series: [{ name: "Series", values: [2] }], box: { x: sized ? 230 : 90, y: 20, width: sized ? 150 : 80, height: 100 } });
  slide.addElement({ type: "group", id: "g", coordinateSize: { width: 50, height: 50 }, box: { x: 10, y: 150, width: sized ? 200 : 100, height: 100 }, transform: { rotation: 15, flipH: true }, children: [{ type: "shape", id: "child", shape: "rect", box: { x: 5, y: 5, width: 20, height: 20 } }] });
  slide.addElement({ type: "connector", id: "link", start: { elementId: "a", anchor: "right" }, end: { elementId: "g", anchor: "top" } });
  return document;
}
const rule = { type: "constraints", container: { kind: "slide", id: "slide" }, elements: [
  { elementId: "a", width: { value: 100 } },
  { elementId: "b", x: { value: ref("a", "right", 1, 10) }, width: { value: ref("a", "width") } },
  { elementId: "c", x: { value: ref("b", "right", 1, 10) }, width: { value: ref("a", "width", 1.5) } },
  { elementId: "g", width: { value: ref("a", "width", 2) } },
] };

test("constrained images, charts, groups and connectors equal independent PPTX/SVG authoring", async () => {
  const document = fixture(), before = normalizePresentation(document);
  const [actual, expected, svg, expectedSvg] = await Promise.all([
    generatePptx(document, { layout: { composition: [rule] } }), generatePptx(fixture(true)),
    renderPresentationToSvg(document, { layout: { composition: [rule] } }), renderPresentationToSvg(fixture(true)),
  ]);
  assert.deepEqual(actual.warnings, []);
  assert.deepEqual(actual.bytes, expected.bytes);
  assert.deepEqual(svg, expectedSvg);
  assert.deepEqual(normalizePresentation(document), before);
});

test("cyclic constraints preserve output bytes and forward the same actionable warnings", async () => {
  const document = fixture();
  const cyclic = { ...rule, elements: [{ elementId: "a", width: { value: ref("b", "width") } }, { elementId: "b", width: { value: ref("a", "width") } }] };
  const layout = { composition: [cyclic] };
  const [actual, plain, svg, plainSvg] = await Promise.all([generatePptx(document, { layout }), generatePptx(document), renderPresentationToSvg(document, { layout }), renderPresentationToSvg(document)]);
  assert.deepEqual(actual.bytes, plain.bytes);
  assert.deepEqual(svg.slides, plainSvg.slides);
  assert.deepEqual(svg.warnings, actual.warnings);
  assert.equal(actual.warnings[0].code, "composition-constraint-cycle");
});
