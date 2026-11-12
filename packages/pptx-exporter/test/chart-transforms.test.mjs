import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

test("native chart graphic frames retain rotation/flips for every supported family and identity frames stay minimal", async () => {
  const p = createPresentation(), slide = p.addSlide();
  for (const [chartType, transform] of [["bar", { rotation: 45, flipH: true }], ["line", { rotation: -30, flipV: true }], ["pie", {}]]) {
    slide.addElement({ type: "chart", chartType, transform, box: { x: 10, y: 20, width: 100, height: 80 }, categories: ["A", "B"], series: [{ name: "Series", values: [1, 2] }] });
  }
  const output = await generatePptx(p); assert.deepEqual(output.warnings, []);
  const xml = readGeneratedZip(output.bytes).get("ppt/slides/slide1.xml").toString();
  assert.match(xml, /<p:xfrm rot="2700000" flipH="1">/);
  assert.match(xml, /<p:xfrm rot="-1800000" flipV="1">/);
  assert.match(xml, /<p:xfrm><a:off x="127000" y="254000"\/><a:ext cx="1270000" cy="1016000"\/><\/p:xfrm>/);
  assert.equal((xml.match(/<p:xfrm/g) ?? []).length, 3);
});
