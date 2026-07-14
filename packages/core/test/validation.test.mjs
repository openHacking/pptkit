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
