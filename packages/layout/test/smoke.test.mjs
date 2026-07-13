import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";

test("resolveLayout returns an export-ready detached layout", () => {
  const presentation = createPresentation();
  presentation.addSlide({
    id: "intro",
    elements: [{
      type: "text",
      text: "Hello",
      box: { x: 10, y: 20, width: 100, height: 30 },
      style: { fontSize: 18, color: "#123456" },
    }],
  });

  const result = resolveLayout(presentation);

  assert.deepEqual(result, {
    size: { width: 960, height: 540, unit: "pt" },
    slides: [{
      id: "intro",
      elements: [{
        type: "text",
        text: "Hello",
        box: { x: 10, y: 20, width: 100, height: 30 },
        style: { fontSize: 18, color: "#123456" },
      }],
    }],
    slideCount: 1,
    status: "resolved",
  });
});

test("resolveNormalizedLayout consumes detached normalized IR without sharing references", () => {
  const presentation = createPresentation();
  presentation.addSlide({ elements: [{
    type: "text",
    text: "Detached",
    box: { x: 0, y: 0, width: 10, height: 10 },
    style: { autoFit: { mode: "shrink", fontScale: 0.96 } },
  }] });
  const normalized = normalizePresentation(presentation);
  const result = resolveNormalizedLayout(normalized);

  result.slides[0].elements[0].box.width = 99;
  result.slides[0].elements[0].style.autoFit.fontScale = 0.5;
  assert.equal(normalized.slides[0].elements[0].box.width, 10);
  assert.equal(normalized.slides[0].elements[0].style.autoFit.fontScale, 0.96);
});
