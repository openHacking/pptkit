import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout } from "../dist/index.js";
import { solveTracks } from "../dist/composition/grid.js";

const fixed = (size) => ({ type: "fixed", size });
const content = (min = 0, max = 100) => ({ type: "content", min, max });
const fraction = (weight = 1, max = 1000) => ({ type: "fraction", weight, max });
function deck(hidden = false) {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const slide = document.addSlide({ id: "slide" });
  for (const [id, width, height] of [["a", 40, 20], ["b", 60, 30], ["c", 80, 40]]) slide.addElement({ type: "shape", id, shape: "rect", hidden: hidden && id === "b", box: { x: 1, y: 2, width, height } });
  return document;
}
const rule = () => ({ type: "grid", container: { kind: "slide", id: "slide" }, columns: [fixed(50), content(), fraction()], rows: [fixed(50), fraction()], columnGap: 10, rowGap: 10, placements: [{ elementId: "a", row: 0, column: 0, horizontal: "center", vertical: "end" }, { elementId: "b", row: 0, column: 1 }, { elementId: "c", row: 1, column: 2, horizontal: "end", vertical: "end" }] });

test("mixed grid tracks, gaps and alignments preserve frame sizes, z-order and source", () => {
  const document = deck(), before = normalizePresentation(document);
  const result = resolveLayout(document, { composition: [rule()] });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.deepEqual(result.slides[0].elements.map((e) => e.box), [
    { x: 5, y: 30, width: 40, height: 20 }, { x: 60, y: 0, width: 60, height: 30 }, { x: 220, y: 160, width: 80, height: 40 },
  ]);
  assert.deepEqual(normalizePresentation(document), before);
  assert.deepEqual(result, resolveLayout(document, { composition: [rule()] }));
});

test("overlapping spans share capacity rather than producing false overflow", () => {
  assert.deepEqual(solveTracks([content(), content(), content()], 100, 0, [{ start: 0, span: 2, size: 100 }, { start: 1, span: 2, size: 100 }]), [0, 100, 0]);
});

test("bounded interval allocation agrees with independently enumerated integer feasibility", () => {
  // 3 tracks, capacities 0..2, 3 demand intervals, budgets 0..6.
  // Every integer assignment is enumerated without using the solver's algorithm.
  for (let mask = 0; mask < 27; mask++) {
    const caps = [mask % 3, Math.floor(mask / 3) % 3, Math.floor(mask / 9)];
    for (let left = 0; left <= 4; left++) for (let right = 0; right <= 4; right++) for (let middle = 0; middle <= 2; middle++) {
      const requirements = [{ start: 0, span: 2, size: left }, { start: 1, span: 2, size: right }, { start: 1, span: 1, size: middle }];
      const candidates = [];
      for (let a = 0; a <= caps[0]; a++) for (let b = 0; b <= caps[1]; b++) for (let c = 0; c <= caps[2]; c++) if (a + b >= left && b + c >= right && b >= middle) candidates.push(a + b + c);
      const minimum = Math.min(...candidates);
      for (let budget = 0; budget <= 6; budget++) {
        const actual = solveTracks(caps.map((max) => content(0, max)), budget, 0, requirements);
        assert.equal(actual !== undefined, minimum <= budget, JSON.stringify({ caps, requirements, budget }));
        if (actual) assert.equal(actual.reduce((a, b) => a + b, 0), minimum);
      }
    }
  }
});

test("fraction allocation redistributes capped growth and handles large weights", () => {
  assert.deepEqual(solveTracks([fraction(1, 20), fraction(3)], 100, 0, []), [20, 80]);
  assert.deepEqual(solveTracks([fraction(Number.MAX_VALUE), fraction(Number.MAX_VALUE)], 100, 0, []), [50, 50]);
  assert.deepEqual(solveTracks([content(10), fixed(20)], 100, 5, []), [10, 20]);
});

test("span requirements count interior gaps and ignore selection order", () => {
  const demands = [{ start: 0, span: 2, size: 60 }, { start: 1, span: 1, size: 20 }];
  assert.deepEqual(solveTracks([fixed(30), content()], 60, 10, demands), [30, 20]);
  assert.deepEqual(solveTracks([fixed(30), content()], 60, 10, demands), solveTracks([fixed(30), content()], 60, 10, [...demands].reverse()));
});

test("unsatisfiable grids are atomic and subsequent rules continue", () => {
  const document = deck();
  const grid = { ...rule(), columns: [fixed(1), fixed(1), fixed(1)] };
  const result = resolveLayout(document, { composition: [grid] });
  assert.equal(result.analysis.diagnostics[0].code, "composition-grid-unsatisfiable");
  assert.deepEqual(result.slides, resolveLayout(document).slides);
  const next = resolveLayout(document, { composition: [grid, { type: "align", container: grid.container, elementIds: ["a"], horizontal: "end" }] });
  assert.equal(next.slides[0].elements[0].box.x, 260);
});

test("invalid grid configuration rejects explicitly", () => {
  for (const patch of [
    { columns: [] }, { columns: [fixed(-1)] }, { rows: [content(20, 10)] },
    { columns: [fraction(0)] }, { rowGap: NaN }, { placements: [{ elementId: "a", row: 0, column: 3 }] },
    { placements: [{ elementId: "a", row: 0, column: 0, rowSpan: 0 }] },
    { placements: [{ elementId: "a", row: 0, column: 0 }, { elementId: "a", row: 1, column: 0 }] },
  ]) assert.throws(() => resolveLayout(deck(), { composition: [{ ...rule(), ...patch }] }));
});


test("grid uses transformed visual extents and explicit group-local coordinates", () => {
  const document = createPresentation();
  document.addSlide({ id: "s" }).addElement({ type: "group", id: "g", box: { x: 100, y: 100, width: 400, height: 200 }, coordinateSize: { width: 200, height: 100 }, transform: { rotation: 25, flipH: true }, children: [
    { type: "shape", id: "rotated", shape: "rect", box: { x: 0, y: 0, width: 40, height: 20 }, transform: { rotation: 90 } },
  ] });
  const grid = { type: "grid", container: { kind: "group", id: "g" }, bounds: "visual", padding: 10, rows: [fraction()], columns: [fraction()], placements: [{ elementId: "rotated", row: 0, column: 0, horizontal: "end", vertical: "end" }] };
  const result = resolveLayout(document, { composition: [grid] });
  assert.deepEqual(result.analysis.diagnostics, []);
  const child = result.slides[0].elements[0].children[0];
  assert.equal(child.box.x, 160);
  assert.equal(child.box.y, 60);
  assert.deepEqual(result.slides[0].elements[0].coordinateSize, { width: 200, height: 100 });
});

test("hidden grid items do not demand capacity, but explicit tracks remain", () => {
  const document = deck(true);
  const grid = rule(); grid.columns[1] = content(0, 0);
  assert.deepEqual(resolveLayout(document, { composition: [grid] }).analysis.diagnostics, []);
  assert.equal(resolveLayout(document, { composition: [{ ...grid, includeHidden: true }] }).analysis.diagnostics[0].code, "composition-grid-unsatisfiable");
});

test("grid placement order and overlapping cells never reorder drawing elements", () => {
  const document = deck();
  const grid = { ...rule(), columns: [fraction()], rows: [fraction()], placements: ["c", "a", "b"].map((elementId) => ({ elementId, row: 0, column: 0, horizontal: "center", vertical: "center" })) };
  const a = resolveLayout(document, { composition: [grid] });
  const b = resolveLayout(document, { composition: [{ ...grid, placements: [...grid.placements].reverse() }] });
  assert.deepEqual(a, b);
  assert.deepEqual(a.slides[0].elements.map((e) => e.id), ["a", "b", "c"]);
});


test("grid content tracks consume intrinsic sizes before anchored connectors resolve", () => {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const slide = document.addSlide({ id: "s" });
  slide.addElement({ type: "text", id: "label", content: "Label", frame: { margin: 0 }, box: { x: 0, y: 0, width: 10, height: 20 } });
  slide.addElement({ type: "shape", id: "box", shape: "rect", box: { x: 0, y: 0, width: 30, height: 20 } });
  slide.addElement({ type: "connector", id: "edge", start: { elementId: "label", anchor: "right" }, end: { elementId: "box", anchor: "left" } });
  const result = resolveLayout(document, {
    textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 100, height: 20 } }) },
    intrinsicSizing: [{ elementId: "label", width: { type: "content" } }],
    composition: [{ type: "grid", container: { kind: "slide", id: "s" }, columns: [content(), fraction()], rows: [fixed(20)], columnGap: 10, placements: [{ elementId: "label", row: 0, column: 0 }, { elementId: "box", row: 0, column: 1 }] }],
  });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.equal(result.slides[0].elements[0].box.width, 100);
  assert.equal(result.slides[0].elements[1].box.x, 110);
  assert.deepEqual(result.slides[0].elements[2].start, { x: 100, y: 10 });
  assert.deepEqual(result.slides[0].elements[2].end, { x: 110, y: 10 });
});
