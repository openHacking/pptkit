import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";

test("tab layout differences are reported for both SVG text paths and table cells", async () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "s" });
  const box = { x: 0, y: 0, width: 300, height: 80 };
  slide.addElement({ id: "native", type: "text", box, content: "a\tb" });
  slide.addElement({ id: "html", type: "text", box, content: [{ runs: [{ text: "a\t" }, { text: "b" }] }] });
  slide.addElement({ id: "table", type: "table", box, columns: [300], rows: [{ height: 80, cells: [{ content: "a\tb" }] }] });
  slide.addElement({ id: "plain", type: "text", box, content: "a b" });
  const original = JSON.stringify(slide.elements), result = await renderPresentationToSvg(deck);
  const warnings = result.warnings.filter(w => w.code === "text-tab-layout-unsupported");
  assert.deepEqual(warnings.map(w => [w.slideId, w.elementId]), [["s", "native"], ["s", "html"], ["s", "table"]]);
  assert.equal(result.status, "rendered-with-warnings");
  assert.equal(JSON.stringify(slide.elements), original);
});
