import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";

function fixture() {
  const document = createPresentation({ size: { width: 200, height: 100 } });
  document.addSlide({ id: "slide" }).addElement({ type: "text", id: "label", content: "Measured text", box: { x: 190, y: 0, width: 50, height: 20 }, frame: { margin: 0 } });
  return document;
}
const layout = { textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 60, height: 30 } }) }, diagnostics: { textOverflow: true, slideBounds: true } };

import { resolveLayout } from "@pptkit/layout";
import { renderPresentationToSvg, renderLayoutToSvg } from "../dist/index.js";

test("analysis preserves SVG geometry and forwards diagnostics once through both entry points", async () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const plain = await renderPresentationToSvg(document);
  const analyzed = await renderPresentationToSvg(document, { layout });
  assert.deepEqual(analyzed.slides, plain.slides);
  assert.equal(analyzed.status, "rendered-with-warnings");
  assert.deepEqual(analyzed.warnings.map(({ code }) => code), ["slide-overflow", "text-overflow", "slide-overflow"]);
  assert.ok(analyzed.warnings.every((warning) => warning.slideId === "slide" && warning.elementId === "label" && warning.path.startsWith("slides.0.elements.0.")));
  assert.deepEqual(await renderLayoutToSvg(resolveLayout(document, layout)), analyzed);
  assert.deepEqual(normalizePresentation(document), before);
});

test("preview reports unmeasured overflow and rejects invalid configuration", async () => {
  const result = await renderPresentationToSvg(fixture(), { layout: { diagnostics: { textOverflow: true } } });
  assert.equal(result.warnings[0].code, "text-measurement-unavailable");
  await assert.rejects(renderPresentationToSvg(fixture(), { layout: { diagnostics: { tolerance: -1 } } }), /tolerance/);
});
