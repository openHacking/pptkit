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
