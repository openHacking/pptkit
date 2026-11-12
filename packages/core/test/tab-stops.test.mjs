import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, resolvePresentationInheritance, validatePresentation } from "../dist/index.js";

const stops = [{ position: 36, align: "left" }, { position: 72.5, align: "center" }, { position: 144, align: "right" }, { position: 216, align: "decimal" }];
function fixture(local) {
  const deck = createPresentation({ textStylePresets: { tabs: { paragraph: { tabStops: stops } } } });
  deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 0, y: 0, width: 300, height: 100 }, textStylePreset: "tabs", content: [{ ...(local === undefined ? {} : { style: { tabStops: local } }), runs: [{ text: "A\tB" }] }] });
  return deck;
}
test("tab stops inherit, locally replace or explicitly clear without mutating the document", () => {
  for (const local of [undefined, [], [{ position: 8, align: "right" }]]) {
    const deck = fixture(local), before = JSON.stringify(deck.slides[0].elements), ir = normalizePresentation(deck), expected = local ?? stops;
    assert.deepEqual(ir.slides[0].elements[0].content[0].style.tabStops, expected);
    const origin = resolvePresentationInheritance(deck).origins.find(o => o.target.id === "t" && o.property === "content.0.style.tabStops");
    assert.equal(origin.source.kind, local === undefined ? "preset" : "element");
    assert.equal(JSON.stringify(deck.slides[0].elements), before);
    ir.slides[0].elements[0].content[0].style.tabStops.push({ position: 300, align: "left" });
    assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].content[0].style.tabStops, expected);
  }
});
test("tab stops reject malformed, out-of-range and nonincreasing coordinates", () => {
  for (const invalid of [null, {}, [null], [{ position: -1, align: "left" }], [{ position: Infinity, align: "left" }], [{ position: 2147483648 / 12700, align: "left" }], [{ position: 2, align: "bad" }], [{ position: 2, align: "left", extra: 1 }], [{ position: 2, align: "left" }, { position: 2, align: "right" }], [{ position: 2, align: "left" }, { position: 2.000001, align: "right" }]]) {
    const deck = fixture(invalid); assert.ok(validatePresentation(deck).some(d => d.code === "invalid-tab-stop"), JSON.stringify(invalid));
    assert.throws(() => normalizePresentation(deck));
  }
});
test("table region rules use the same tab-stop validation", () => {
  const deck = createPresentation(), slide = deck.addSlide();
  slide.addElement({ type: "table", box: { x: 0, y: 0, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: "A\tB" }] }], style: { wholeTable: { text: { paragraph: { tabStops: stops } } } } });
  assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].rows[0].cells[0].content[0].style.tabStops, stops);
  const bad = createPresentation(), s = bad.addSlide();
  s.addElement({ type: "table", box: { x: 0, y: 0, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: "x" }] }], style: { wholeTable: { text: { paragraph: { tabStops: [{ position: -1, align: "left" }] } } } } });
  assert.ok(validatePresentation(bad).some(d => d.code === "invalid-table-style" && d.path.endsWith("tabStops.0")));
});

test("tab-stop EMU endpoints are accepted without changing the authored coordinate", () => {
  const edge = [{ position: 0, align: "left" }, { position: 2147483647 / 12700, align: "right" }];
  assert.deepEqual(normalizePresentation(fixture(edge)).slides[0].elements[0].content[0].style.tabStops, edge);
});

test("resetting a preset tab-stop override uses the placeholder declaration", () => {
  const deck = createPresentation({ textStylePresets: { tabs: { paragraph: { tabStops: stops } } } });
  const placeholderStops = [{ position: 40, align: "right" }];
  deck.defineSlideLayout({ id: "l", name: "Layout", placeholders: [{ key: "body", kind: "body", box: { x: 0, y: 0, width: 300, height: 100 }, textStyle: { paragraph: { tabStops: placeholderStops } } }] });
  deck.addSlide({ layoutId: "l" }).addElement({ id: "t", type: "text", placeholderKey: "body", textStylePreset: "tabs", textStylePresetReset: ["paragraph.tabStops"], content: "A\tB" });
  assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].content[0].style.tabStops, placeholderStops);
});
