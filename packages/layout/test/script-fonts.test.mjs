import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveNormalizedLayout } from "../dist/index.js";

function fixture() {
  const p = createPresentation(); p.addSlide({ id: "slide" }).addElement({ id: "text", type: "text", box: { x: 10, y: 10, width: 200, height: 80 }, content: [{ runs: [{ text: "A中ع", style: { fontFamily: "Latin", eastAsianFontFamily: "Eastern", complexScriptFontFamily: "Complex", fontSize: 20 } }] }] });
  return normalizePresentation(p);
}
test("measurement providers receive concrete script fragments without changing author runs or identities", () => {
  const p = fixture(), before = structuredClone(p), requests = []; let measured;
  const result = resolveNormalizedLayout(p, { textMeasurement: {
    resolveFont(request) { requests.push(request); assert.ok(Object.isFrozen(request)); return { family: request.family }; },
    measure(request) { measured = request; assert.ok(Object.isFrozen(request.paragraphs[0].runs[0].style)); return { bounds: { x: 0, y: 0, width: 60, height: 20 } }; },
  } });
  assert.deepEqual(requests.map(r => [r.text, r.family]), [["A", "Latin"], ["中", "Eastern"], ["ع", "Complex"]]);
  assert.deepEqual(measured.paragraphs[0].runs.map(r => [r.text, r.style.fontFamily]), requests.map(r => [r.text, r.family]));
  assert.ok(measured.paragraphs[0].runs.every(r => !Object.hasOwn(r.style, "eastAsianFontFamily")));
  assert.deepEqual(p, before); assert.equal(result.slides[0].elements[0].content[0].runs.length, 1); assert.equal(result.analysis.measurements.length, 1);
});

test("script-specific missing/fallback fonts retain located diagnostics and prevent partial measurement", () => {
  for (const missing of [true, false]) {
    let calls = 0;
    const result = resolveNormalizedLayout(fixture(), { textMeasurement: {
      resolveFont(request) { return request.family === "Eastern" ? missing ? undefined : { family: "Fallback" } : { family: request.family }; },
      measure() { calls++; return { bounds: { x: 0, y: 0, width: 40, height: 20 } }; },
    } });
    assert.equal(calls, missing ? 0 : 1); assert.equal(result.analysis.diagnostics[0].code, missing ? "font-unavailable" : "font-fallback");
    assert.equal(result.analysis.diagnostics[0].path, "slides.0.elements.0.content.0.runs.0.style.eastAsianFontFamily");
  }
});

test("selection budget failure keeps authored geometry and never calls providers with a partial run", () => {
  const p = fixture(); p.slides[0].elements[0].content[0].runs[0].text = "A中".repeat(50_001);
  let calls = 0;
  const result = resolveNormalizedLayout(p, { textMeasurement: { measure() { calls++; return { bounds: { x: 0, y: 0, width: 1, height: 1 } }; } } });
  assert.equal(calls, 0); assert.equal(result.analysis.measurements.length, 0); assert.equal(result.analysis.diagnostics[0].code, "text-font-selection-failed");
  assert.deepEqual(result.slides[0].elements[0].box, p.slides[0].elements[0].box);
});

test("script maps resolve separate Arabic/Hebrew and language-aware Han fragments with map diagnostics", () => {
  for (const missing of [true, false]) {
    const p = fixture(), run = p.slides[0].elements[0].content[0].runs[0];
    run.text = "عא中"; run.style.language = "zh-TW"; run.style.scriptFontFamilies = { Arab: "Arabic", Hebr: "Hebrew", Hant: "Traditional" };
    const before = structuredClone(p), requests = []; let measured;
    const result = resolveNormalizedLayout(p, { textMeasurement: {
      resolveFont(request) { requests.push([request.text, request.family]); return request.family === "Arabic" ? missing ? undefined : { family: "Fallback" } : { family: request.family }; },
      measure(request) { measured = request; return { bounds: { x: 0, y: 0, width: 50, height: 20 } }; },
    } });
    assert.deepEqual(requests, [["ع", "Arabic"], ["א", "Hebrew"], ["中", "Traditional"]]);
    assert.equal(result.analysis.diagnostics[0].path, "slides.0.elements.0.content.0.runs.0.style.scriptFontFamilies.Arab");
    assert.equal(result.analysis.diagnostics[0].code, missing ? "font-unavailable" : "font-fallback");
    if (missing) assert.equal(measured, undefined);
    else {
      assert.deepEqual(measured.paragraphs[0].runs.map(r => r.style.fontFamily), ["Fallback", "Hebrew", "Traditional"]);
      assert.ok(measured.paragraphs[0].runs.every(r => !Object.hasOwn(r.style, "scriptFontFamilies")));
    }
    assert.deepEqual(p, before); assert.equal(result.slides[0].elements[0].content[0].runs.length, 1);
  }
});

test("RTL font resolution locates the selected slot even with an East Asian font present", () => {
  for (const [complex, expectedFamily, expectedField] of [
    ["Complex", "Complex", "complexScriptFontFamily"],
    [undefined, "Latin", "fontFamily"],
  ]) {
    const p = fixture(), run = p.slides[0].elements[0].content[0].runs[0];
    run.text = "中";
    run.style.rightToLeft = true;
    run.style.complexScriptFontFamily = complex;
    const requested = [];
    const result = resolveNormalizedLayout(p, { textMeasurement: {
      resolveFont(request) { requested.push(request.family); return undefined; },
      measure() { assert.fail("A missing selected font must prevent measurement"); },
    } });
    assert.deepEqual(requested, [expectedFamily]);
    assert.equal(result.analysis.measurements.length, 0);
    assert.equal(result.analysis.diagnostics[0].code, "font-unavailable");
    assert.equal(result.analysis.diagnostics[0].path, `slides.0.elements.0.content.0.runs.0.style.${expectedField}`);
  }
});
