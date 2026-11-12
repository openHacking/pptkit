import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";
import { constraintTargets } from "../dist/composition/constraints.js";

const ref = (id, metric, factor = 1, constant = 0) => ({ terms: [{ source: id === "container" ? id : { elementId: id }, metric, factor }], constant });
const rule = (elements, extra = {}) => ({ type: "constraints", container: { kind: "slide", id: "slide" }, elements, ...extra });
function fixture() {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const slide = document.addSlide({ id: "slide" });
  for (const [id, x, width] of [["a", 10, 40], ["b", 70, 60], ["c", 150, 80]]) slide.addElement({ type: "shape", id, shape: "rect", box: { x, y: 20, width, height: 30 } });
  return document;
}
const boxes = (layout) => layout.slides[0].elements.map((element) => element.box);

test("affine references resolve resized peers independent of rule item order", () => {
  const document = fixture(), before = normalizePresentation(document);
  const elements = [
    { elementId: "c", x: { value: ref("b", "right", 1, 5) }, width: { value: ref("a", "width", 0.5) } },
    { elementId: "b", x: { value: ref("a", "right", 1, 10) }, width: { value: ref("a", "width") } },
    { elementId: "a", x: { value: 0 }, width: { value: 70 } },
  ];
  const result = resolveLayout(document, { composition: [rule(elements)] });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.deepEqual(boxes(result).map(({ x, width }) => ({ x, width })), [{ x: 0, width: 70 }, { x: 80, width: 70 }, { x: 155, width: 35 }]);
  assert.deepEqual(result, resolveLayout(document, { composition: [rule([...elements].reverse())] }));
  assert.deepEqual(normalizePresentation(document), before);
});

test("affine multi-term centering and container padding use local dimensions", () => {
  const center = { terms: [{ source: "container", metric: "centerX" }, { source: { elementId: "a" }, metric: "width", factor: -0.5 }] };
  const result = resolveLayout(fixture(), { composition: [rule([{ elementId: "a", width: { value: ref("container", "width", 0.5) }, x: { value: center }, y: { value: ref("container", "bottom", 1, -30) } }], { padding: 10 })] });
  assert.deepEqual(boxes(result)[0], { x: 80, y: 160, width: 140, height: 30 });
});

test("bounds-only rules clamp incoming frames; exact values never silently clamp", () => {
  const document = fixture();
  const result = resolveLayout(document, { composition: [rule([{ elementId: "a", width: { min: 60, max: 100 }, x: { max: 5 } }])] });
  assert.equal(boxes(result)[0].width, 60); assert.equal(boxes(result)[0].x, 5);
  for (const constraint of [{ value: 120, max: 100 }, { min: 100, max: 90 }, { value: -1 }, { max: -1 }]) {
    const failed = resolveLayout(document, { composition: [rule([{ elementId: "b", x: { value: 0 } }, { elementId: "a", width: constraint }])] });
    assert.equal(failed.analysis.diagnostics[0].code, "composition-constraint-unsatisfiable");
    assert.deepEqual(boxes(failed), boxes(resolveLayout(document)));
  }
});

test("bound references are dependencies and exact values within them are retained", () => {
  const result = resolveLayout(fixture(), { composition: [rule([{ elementId: "a", width: { min: ref("b", "width"), max: 100 } }, { elementId: "b", width: { value: 90 } }])] });
  assert.equal(boxes(result)[0].width, 90);
});

test("cycles include a stable actual dependency loop and roll back all properties", () => {
  const document = fixture();
  const elements = [{ elementId: "a", x: { value: ref("b", "right") } }, { elementId: "b", width: { value: ref("a", "x") } }, { elementId: "c", y: { value: 100 } }];
  for (const items of [elements, [...elements].reverse()]) {
    const result = resolveLayout(document, { composition: [rule(items)] });
    assert.equal(result.analysis.diagnostics[0].code, "composition-constraint-cycle");
    assert.match(result.analysis.diagnostics[0].message, /a.x -> b.width -> a.x/);
    assert.deepEqual(boxes(result), boxes(resolveLayout(document)));
  }
  const self = resolveLayout(document, { composition: [rule([{ elementId: "a", x: { value: ref("a", "right") } }])] });
  assert.equal(self.analysis.diagnostics[0].code, "composition-constraint-cycle");
});

test("references to unconstrained metrics and zero factors do not create false cycles", () => {
  const result = resolveLayout(fixture(), { composition: [rule([{ elementId: "a", x: { value: ref("a", "width", 1, 5) }, y: { value: ref("a", "y", 0, 7) } }])] });
  assert.deepEqual(boxes(result)[0], { x: 45, y: 7, width: 40, height: 30 });
});

test("nonfinite products and frame extents fail atomically", () => {
  for (const elements of [[{ elementId: "a", width: { value: ref("container", "width", Number.MAX_VALUE) } }], [{ elementId: "a", x: { value: Number.MAX_VALUE }, width: { value: Number.MAX_VALUE } }]]) {
    const document = fixture();
    const result = resolveLayout(document, { composition: [rule(elements)] });
    assert.equal(result.analysis.diagnostics[0].code, "composition-constraint-invalid");
    assert.deepEqual(boxes(result), boxes(resolveLayout(document)));
  }
});

test("missing and cross-container references produce a precise rule diagnostic", () => {
  const result = resolveLayout(fixture(), { composition: [rule([{ elementId: "a", x: { value: ref("missing", "right") } }])] });
  assert.equal(result.analysis.diagnostics[0].code, "composition-constraint-reference-missing");
  assert.equal(result.analysis.diagnostics[0].path, "options.composition.0");
  assert.match(result.analysis.diagnostics[0].message, /a.x.*missing/);
});

test("hidden targets are skipped, while hidden reference frames remain explicit inputs", () => {
  const normalized = normalizePresentation(fixture()); normalized.slides[0].elements[0].hidden = true;
  const rules = [rule([{ elementId: "a", width: { value: ref("missing", "width") } }, { elementId: "b", width: { value: ref("a", "width") } }])];
  const result = resolveNormalizedLayout(normalized, { composition: rules });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.equal(boxes(result)[1].width, 40);
  assert.equal(resolveNormalizedLayout(normalized, { composition: [{ ...rules[0], includeHidden: true }] }).analysis.diagnostics[0].code, "composition-constraint-reference-missing");
});

test("constraints preserve group coordinates and intentionally rescale its children", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "group", id: "g", box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 50, height: 50 }, children: [{ type: "shape", id: "child", shape: "rect", box: { x: 5, y: 5, width: 10, height: 10 } }] });
  const result = resolveLayout(document, { composition: [rule([{ elementId: "g", width: { value: 200 } }]), rule([{ elementId: "child", x: { value: ref("container", "right", 1, -10) } }], { container: { kind: "group", id: "g" } })] });
  const group = result.slides[0].elements[3];
  assert.equal(group.box.width, 200); assert.equal(group.coordinateSize.width, 50); assert.equal(group.children[0].box.x, 40);
});

test("anchored connectors follow constrained targets but cannot supply unresolved references", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "link", start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" } });
  const result = resolveLayout(document, { composition: [rule([{ elementId: "a", width: { value: 100 } }])] });
  assert.deepEqual(result.slides[0].elements[3].start, { x: 110, y: 35 });
  const failed = resolveLayout(document, { composition: [rule([{ elementId: "a", x: { value: ref("link", "right") } }])] });
  assert.equal(failed.analysis.diagnostics[0].code, "composition-constraint-unresolved-reference");
});

test("free connectors translate every point, and resize requests fail without partial changes", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "free", start: { x: 0, y: 0 }, route: [{ x: 10, y: 20 }], end: { x: 30, y: 10 } });
  const moved = resolveLayout(document, { composition: [rule([{ elementId: "free", x: { value: 50 }, y: { value: 60 } }])] });
  assert.deepEqual(moved.slides[0].elements[3].route, [{ x: 60, y: 80 }]);
  const failed = resolveLayout(document, { composition: [rule([{ elementId: "free", width: { value: 100 } }])] });
  assert.equal(failed.analysis.diagnostics[0].code, "composition-constraint-resize-unsupported");
});

test("table rows/columns cannot silently disagree with constrained frame dimensions", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "table", id: "table", box: { x: 0, y: 0, width: 40, height: 20 }, columns: [40], rows: [{ height: 20, cells: [{ content: "A" }] }] });
  const failed = resolveLayout(document, { composition: [rule([{ elementId: "table", height: { value: 40 } }, { elementId: "a", x: { value: 200 } }])] });
  assert.equal(failed.analysis.diagnostics[0].code, "composition-constraint-resize-unsupported");
  assert.deepEqual(boxes(failed), boxes(resolveLayout(document)));
});

test("composition ordering has explicit precedence before and after a failed rule", () => {
  const document = fixture();
  const result = resolveLayout(document, { composition: [rule([{ elementId: "a", width: { value: 50 } }]), rule([{ elementId: "a", x: { value: ref("a", "right") } }]), { type: "align", container: { kind: "slide", id: "slide" }, elementIds: ["a"], horizontal: "end" }] });
  assert.deepEqual(boxes(result)[0], { x: 250, y: 20, width: 50, height: 30 });
  assert.equal(result.analysis.diagnostics[0].path, "options.composition.1");
});

test("long acyclic chains do not depend on JavaScript recursion limits", () => {
  const size = 12000;
  const ids = Array.from({ length: size }, (_, i) => `id${i}`);
  const incoming = new Map(ids.map((id) => [id, { x: 0, y: 0, width: 1, height: 1 }]));
  const elements = ids.map((elementId, i) => ({ elementId, x: { value: i === 0 ? 0 : ref(ids[i - 1], "right", 1, 1) } }));
  const result = constraintTargets(rule(elements), ids, incoming, { x: 0, y: 0, width: 100, height: 100 });
  assert.ok("targets" in result);
  assert.equal(result.targets.at(-1).x, (size - 1) * 2);
});

test("malformed constraint configuration rejects before any layout mutation", () => {
  for (const elements of [[{ elementId: "a" }], [{ elementId: "a", x: {} }], [{ elementId: "a", x: { value: Infinity } }], [{ elementId: "a", x: { value: { terms: [{ source: "container", metric: "unknown" }] } } }], [{ elementId: "a", x: { value: 1 } }, { elementId: "a", y: { value: 2 } }]]) assert.throws(() => resolveLayout(fixture(), { composition: [rule(elements)] }));
  assert.throws(() => resolveLayout(fixture(), { composition: [rule([{ elementId: "a", x: { value: 0 } }], { bounds: "visual" })] }));
});

test("grid placement consumes constrained sizes and keeps the original drawing order", () => {
  const result = resolveLayout(fixture(), { composition: [
    rule([{ elementId: "a", width: { value: 100 } }]),
    { type: "grid", container: { kind: "slide", id: "slide" }, columns: [{ type: "content" }, { type: "fraction" }], rows: [{ type: "content" }], columnGap: 10, placements: [{ elementId: "a", row: 0, column: 0 }, { elementId: "b", row: 0, column: 1 }] },
  ] });
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.equal(boxes(result)[0].width, 100); assert.equal(boxes(result)[1].x, 110);
});

test("a seeded independent affine DAG oracle covers negative factors and branching dependencies", () => {
  let seed = 12345;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed; };
  for (let run = 0; run < 100; run++) {
    const ids = Array.from({ length: 30 }, (_, i) => `n${i}`);
    const input = new Map(ids.map((id) => [id, { x: 0, y: 0, width: 1, height: 1 }]));
    const expected = [], elements = [];
    for (let i = 0; i < ids.length; i++) {
      const constant = random() % 21 - 10;
      const terms = [];
      let value = constant;
      for (let j = 0; j < Math.min(i, 3); j++) {
        const parent = random() % i, factor = (random() % 5 - 2) / 2;
        terms.push({ source: { elementId: ids[parent] }, metric: "x", factor });
        value += expected[parent] * factor;
      }
      expected.push(value); elements.push({ elementId: ids[i], x: { value: { constant, terms } } });
    }
    const result = constraintTargets(rule([...elements].reverse()), ids, input, { x: 0, y: 0, width: 100, height: 100 });
    assert.ok("targets" in result);
    assert.deepEqual(result.targets.map((box) => box.x), expected);
  }
});
