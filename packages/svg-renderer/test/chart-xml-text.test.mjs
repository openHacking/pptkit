import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";

test("chart previews report and replace XML-invalid displayed text without changing input", async () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "slide" });
  slide.addElement({ type: "chart", chartType: "bar", id: "bar", box: { x: 0, y: 0, width: 400, height: 300 }, title: "title\u0001", categories: ["A\u0000", "😀\rB"], series: [{ name: "name\u000b", values: [1, 2] }], legend: { visible: true } });
  slide.addElement({ type: "chart", chartType: "pie", id: "pie", box: { x: 400, y: 0, width: 400, height: 300 }, categories: ["pie\u001f", "safe"], series: [{ name: "name", values: [1, 2] }], legend: { visible: true } });
  const first = await renderPresentationToSvg(deck), second = await renderPresentationToSvg(deck);
  assert.deepEqual(first, second); assert.equal(first.status, "rendered-with-warnings");
  assert.deepEqual(first.warnings.filter(w => w.code === "chart-text-xml-replacement").map(w => [w.slideId, w.elementId]), [["slide", "bar"], ["slide", "pie"]]);
  const svg = first.slides[0].svg; for (const text of ["title�", "A�", "name�", "pie�", "😀&#13;B"]) assert.ok(svg.includes(text), text);
  assert.equal(slide.elements[0].categories[0], "A\u0000"); assert.equal(slide.elements[1].categories[0], "pie\u001f");
});

test("undisplayed chart labels do not produce replacement warnings", async () => {
  const deck = createPresentation(); deck.addSlide({ elements: [{ type: "chart", chartType: "bar", box: { x: 0, y: 0, width: 400, height: 300 }, categories: ["hidden\u0000"], series: [{ name: "hidden\u0001", values: [1] }], legend: { visible: false }, axes: { category: { labels: false } } }] });
  const result = await renderPresentationToSvg(deck); assert.ok(!result.warnings.some(w => w.code === "chart-text-xml-replacement")); assert.ok(!result.slides[0].svg.includes("hidden"));
});
