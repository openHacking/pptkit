import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "../dist/index.js";

test("createPresentation builds a formal document model with defaults", () => {
  const presentation = createPresentation({ title: "Roadmap" });
  const slide = presentation.addSlide({
    elements: [
      {
        type: "text",
        text: "Launch plan",
        box: {
          x: 48,
          y: 48,
          width: 400,
          height: 32,
        },
      },
    ],
  });

  assert.equal(presentation.id, "presentation-1");
  assert.equal(presentation.title, "Roadmap");
  assert.deepEqual(presentation.size, {
    width: 960,
    height: 540,
    unit: "pt",
  });
  assert.equal(presentation.slides.length, 1);
  assert.equal(slide.id, "slide-1");
});

test("registerAsset deduplicates repeat registrations by dedupe key and source", () => {
  const presentation = createPresentation();
  const first = presentation.registerAsset({
    kind: "image",
    source: {
      type: "path",
      value: "./chart.png",
    },
    dedupeKey: "chart-image",
    mimeType: "image/png",
  });
  const second = presentation.registerAsset({
    kind: "image",
    source: {
      type: "path",
      value: "./chart.png",
    },
    dedupeKey: "chart-image",
    mimeType: "image/png",
  });

  assert.equal(first.id, "asset-1");
  assert.equal(second, first);
  assert.equal(presentation.assets.length, 1);
});

test("registerAsset rejects conflicting duplicate asset metadata", () => {
  const presentation = createPresentation();
  const first = presentation.registerAsset({
    id: "logo",
    kind: "image",
    source: {
      type: "path",
      value: "./logo.png",
    },
    mimeType: "image/png",
  });

  const duplicate = presentation.registerAsset({
    id: "logo",
    kind: "image",
    source: {
      type: "path",
      value: "./logo.png",
    },
    mimeType: "image/png",
  });

  assert.equal(duplicate, first);

  assert.throws(
    () =>
      presentation.registerAsset({
        id: "logo",
        kind: "image",
        source: {
          type: "path",
          value: "./logo.png",
        },
        mimeType: "image/jpeg",
      }),
    /already registered with different metadata/,
  );
});

test("normalizePresentation validates image asset references", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [
      {
        type: "image",
        assetId: "missing-asset",
        box: {
          x: 0,
          y: 0,
          width: 120,
          height: 120,
        },
      },
    ],
  });

  assert.throws(
    () => normalizePresentation(presentation),
    /references missing asset "missing-asset"/,
  );
});

test("normalizePresentation returns detached normalized state", () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({
    kind: "image",
    source: {
      type: "url",
      value: "https://example.com/hero.png",
    },
    width: 800,
    height: 400,
  });

  presentation.addSlide({
    elements: [
      {
        type: "image",
        assetId: asset.id,
        box: {
          x: 32,
          y: 24,
          width: 320,
          height: 160,
        },
        altText: "Hero graphic",
      },
    ],
  });

  const normalized = normalizePresentation(presentation);
  presentation.assets[0].source.value = "https://example.com/changed.png";
  presentation.slides[0].elements[0].box.width = 999;

  assert.equal(normalized.assets[0].source.value, "https://example.com/hero.png");
  assert.equal(normalized.slides[0].elements[0].box.width, 320);
});

test("normalizePresentation preserves slide background and detached text layout styles", () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide({
    background: "#F7F5EF",
    elements: [{
      type: "text",
      text: "First line\nSecond line",
      box: { x: 20, y: 20, width: 240, height: 80 },
      style: {
        fontFamily: "Helvetica Neue",
        lineSpacing: 1.15,
        autoFit: { mode: "shrink", fontScale: 0.96 },
      },
    }],
  });

  const normalized = normalizePresentation(presentation);
  slide.background = "#FFFFFF";
  slide.elements[0].style.autoFit.fontScale = 0.5;

  assert.equal(normalized.slides[0].background, "#F7F5EF");
  assert.deepEqual(normalized.slides[0].elements[0].style.autoFit, {
    mode: "shrink",
    fontScale: 0.96,
  });
});

test("normalizePresentation rejects invalid text layout ratios", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [{
      type: "text",
      text: "Invalid spacing",
      box: { x: 0, y: 0, width: 100, height: 20 },
      style: { lineSpacing: 0 },
    }],
  });
  assert.throws(() => normalizePresentation(presentation), /lineSpacing must be a positive finite number/);

  presentation.slides[0].elements[0].style = {
    autoFit: { mode: "shrink", fontScale: 1.01 },
  };
  assert.throws(() => normalizePresentation(presentation), /fontScale must be greater than zero and at most one/);
});

test("addSlide rejects duplicate slide ids", () => {
  const presentation = createPresentation();
  presentation.addSlide({ id: "intro" });

  assert.throws(() => presentation.addSlide({ id: "intro" }), /Duplicate slide id "intro"/);
});

test("normalizePresentation rejects negative frame sizes", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    elements: [
      {
        type: "shape",
        shape: "rect",
        box: {
          x: 10,
          y: 10,
          width: -1,
          height: 80,
        },
      },
    ],
  });

  assert.throws(() => normalizePresentation(presentation), /negative width or height/);
});

test("normalizePresentation maps text, image, and shape elements", () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({
    kind: "image",
    source: { type: "path", value: "./hero.png" },
  });

  presentation.addSlide({
    elements: [
      { type: "text", text: "Title", box: { x: 0, y: 0, width: 100, height: 20 } },
      { type: "image", assetId: asset.id, box: { x: 0, y: 30, width: 100, height: 80 } },
      { type: "shape", shape: "rect", box: { x: 0, y: 120, width: 100, height: 40 } },
    ],
  });

  const [slide] = normalizePresentation(presentation).slides;
  assert.deepEqual(slide?.elements.map((element) => element.type), ["text", "image", "shape"]);
  assert.deepEqual(slide?.elements[0]?.style, {});
  assert.deepEqual(slide?.elements[2]?.style, {});
});
