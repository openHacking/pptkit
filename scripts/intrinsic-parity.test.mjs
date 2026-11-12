import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { pixelPng } from "../tests/fixtures/presentations.mjs";

function fixture(sized = false) {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const asset = document.registerAsset({ id: "asset", kind: "image", source: { type: "url", value: pixelPng }, mimeType: "image/png", width: 8, height: 4 });
  const slide = document.addSlide({ id: "slide" });
  slide.addElement({ type: "text", id: "text", content: "abcdefghij", box: { x: 10, y: 10, width: 60, height: sized ? 40 : 20 }, frame: { margin: 0 } });
  slide.addElement({ type: "image", id: "image", assetId: asset.id, box: { x: 100, y: 10, width: 80, height: sized ? 40 : 80 }, fit: "contain" });
  slide.addElement({ type: "group", id: "group", box: sized ? { x: 10, y: 90, width: 120, height: 50 } : { x: 50, y: 100, width: 200, height: 100 }, coordinateSize: sized ? { width: 60, height: 50 } : { width: 100, height: 100 }, children: [{ type: "shape", id: "child", shape: "rect", box: { x: sized ? 0 : -20, y: sized ? 0 : -10, width: 60, height: 50 } }] });
  slide.addElement({ type: "connector", id: "link", start: { elementId: "text", anchor: "bottom" }, end: { elementId: "group", anchor: "top" } });
  return document;
}
const options = {
  textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 60, height: 40 } }) },
  intrinsicSizing: [
    { elementId: "text", height: { type: "content" } },
    { elementId: "image", height: { type: "content" } },
    { elementId: "group", width: { type: "content" }, height: { type: "content" } },
  ],
};

test("intrinsic text/image/group geometry matches independently authored PPTX bytes and SVG", async () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const [actualPptx, expectedPptx, actualSvg, expectedSvg] = await Promise.all([
    generatePptx(document, { layout: options }), generatePptx(fixture(true)),
    renderPresentationToSvg(document, { layout: options }), renderPresentationToSvg(fixture(true)),
  ]);
  assert.deepEqual(actualPptx.bytes, expectedPptx.bytes);
  assert.deepEqual(actualPptx.warnings, []);
  assert.deepEqual(actualSvg, expectedSvg);
  assert.deepEqual(normalizePresentation(document), before);
});

test("failed intrinsic batches retain original export and preview and forward identical warnings", async () => {
  const document = fixture();
  const layout = { ...options, intrinsicSizing: [...options.intrinsicSizing, { elementId: "absent", width: { type: "content" } }] };
  const [actual, expected, svg, plainSvg] = await Promise.all([
    generatePptx(document, { layout }), generatePptx(document),
    renderPresentationToSvg(document, { layout }), renderPresentationToSvg(document),
  ]);
  assert.deepEqual(actual.bytes, expected.bytes);
  assert.deepEqual(svg.slides, plainSvg.slides);
  assert.deepEqual(svg.warnings, actual.warnings);
  assert.equal(actual.warnings[0].code, "intrinsic-target-missing");
});
