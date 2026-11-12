import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";

test("SVG replaces invalid XML characters in ordinary text and attributes with a located warning", async () => {
  const deck = createPresentation({ metadata: { title: "Title\u0000\rEnd" } }), slide = deck.addSlide({ id: "s" });
  slide.addElement({ id: "text", name: "Name\u0001", type: "text", box: { x: 0, y: 0, width: 300, height: 100 }, content: "A\u0008 & <B> 😀\rC\ud800" });
  deck.addSlide({ id: "control" });
  const result = await renderPresentationToSvg(deck), svg = result.slides[0].svg;
  assert.equal(result.status, "rendered-with-warnings"); assert.deepEqual(result.warnings.filter(w => w.code === "svg-xml-character-replacement").map(w => [w.slideId, w.path]), [["s", "slides.0"], ["control", "slides.1"]]);
  assert.ok(svg.includes("Title�&#13;End")); assert.ok(svg.includes("Name�")); assert.ok(svg.includes("A� &amp; &lt;B&gt; 😀")); assert.ok(svg.includes("C�"));
  assert.ok(!/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(svg));
  assert.equal(slide.elements[0].content, "A\u0008 & <B> 😀\rC\ud800");
  const limit = new TextEncoder().encode(svg).length;
  assert.equal((await renderPresentationToSvg(deck, { slideIndices: [0], maxSvgBytes: limit })).slides[0].svg, svg);
  await assert.rejects(renderPresentationToSvg(deck, { slideIndices: [0], maxSvgBytes: limit - 1 }), e => e.code === "svg-limit-exceeded");
});

test("SVG warnings follow selected slides and valid emoji remain intact", async () => {
  const deck = createPresentation(); deck.addSlide({ id: "bad", elements: [{ type: "text", box: { x: 0, y: 0, width: 250, height: 80 }, content: "bad\u0000" }] });
  deck.addSlide({ id: "good", elements: [{ type: "text", box: { x: 0, y: 0, width: 250, height: 80 }, content: "safe 😀" }] });
  const result = await renderPresentationToSvg(deck, { slideIndices: [1] }); assert.equal(result.status, "rendered"); assert.deepEqual(result.warnings, []); assert.ok(result.slides[0].svg.includes("safe 😀"));
});
