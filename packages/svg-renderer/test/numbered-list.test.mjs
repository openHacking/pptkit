import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";

test("alphabetic markers use the browser measurement sequence past Z", async () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "numbered" });
  const cases = [[26, "z"], [27, "aa"], [28, "ab"], [52, "az"], [53, "ba"], [702, "zz"], [703, "aaa"], [32767, "avlg"]];
  for (const [index, [startAt]] of cases.entries()) slide.addElement({ id: `lower-${startAt}`, type: "text", box: { x: 0, y: index * 50, width: 300, height: 45 }, content: [{ style: { bullet: { type: "number", style: "alphaLowerPeriod", startAt } }, runs: [{ text: `L${startAt}` }] }] });
  slide.addElement({ id: "upper", type: "text", box: { x: 400, y: 0, width: 300, height: 80 }, content: [{ style: { bullet: { type: "number", style: "alphaUpperPeriod", startAt: 53 } }, runs: [{ text: "link", action: { type: "url", url: "https://example.invalid" } }] }] });
  slide.addElement({ id: "table", type: "table", box: { x: 400, y: 90, width: 300, height: 80 }, columns: [300], rows: [{ height: 80, cells: [{ content: [{ style: { bullet: { type: "number", style: "alphaLowerPeriod", startAt: 703 } }, runs: [{ text: "cell" }] }] }] }] });
  const result = await renderPresentationToSvg(deck), svg = result.slides[0].svg;
  for (const [startAt, marker] of cases) assert.ok(svg.includes(`${marker}. `), `${startAt}: ${marker}`);
  assert.ok(svg.includes("BA. ")); const warnings = result.warnings.filter(w => w.code === "alphabetic-numbering-application-dependent");
  assert.equal(warnings.length, 8); assert.ok(warnings.every(w => w.slideId === "numbered")); assert.ok(warnings.some(w => w.elementId === "table"));
  assert.ok(!warnings.some(w => w.elementId === "lower-26" || w.elementId === "lower-27"));
});
