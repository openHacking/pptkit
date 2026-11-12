import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout, resolveNormalizedLayout } from "../dist/index.js";

function fixture() {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  const slide = document.addSlide({ id: "slide" });
  for (const [id, width, height] of [["a", 40, 20], ["b", 60, 40], ["c", 20, 60]]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: 7, y: 9, width, height } });
  return document;
}
const base = { container: { kind: "slide", id: "slide" }, elementIds: ["a", "b", "c"] };
const stack = { ...base, type: "stack", direction: "horizontal", gap: 10, padding: 10, align: "center" };
const boxes = (result) => result.slides[0].elements.map(({ box }) => box);
const composed = (document, ...rules) => resolveLayout(document, { composition: rules });

test('many composition findings remain ordered without exceeding the function argument limit', () => {
  const document = createPresentation(); document.addSlide({ id: 'slide' });
  const missing = { type: 'align', container: { kind: 'slide', id: 'missing' }, elementIds: [], horizontal: 'start' };
  const count = 150000;
  const result = resolveLayout(document, { composition: Array(count).fill(missing) });
  assert.equal(result.analysis.diagnostics.length, count);
  assert.equal(result.analysis.diagnostics[0].code, 'composition-container-missing');
  assert.equal(result.analysis.diagnostics[0].path, 'options.composition.0');
  assert.equal(result.analysis.diagnostics[count - 1].path, `options.composition.${count - 1}`);
});

test('composition handles a large free connector route without spreading all points as arguments', () => {
  const document = createPresentation({ size: { width: 500, height: 300 } });
  const route = Array.from({ length: 150000 }, (_, index) => ({ x: index % 101, y: index % 51 }));
  document.addSlide({ id: 'slide', elements: [{ id: 'line', type: 'connector', start: { x: 0, y: 0 }, end: { x: 100, y: 50 }, route }] });
  const result = resolveLayout(document, { composition: [{ type: 'align', container: { kind: 'slide', id: 'slide' }, elementIds: ['line'], horizontal: 'end' }] });
  const line = result.slides[0].elements[0];
  assert.deepEqual(line.box, { x: 400, y: 0, width: 100, height: 50 });
  assert.deepEqual(line.start, { x: 400, y: 0 });
  assert.deepEqual(line.end, { x: 500, y: 50 });
  assert.equal(line.route.length, 150000);
  assert.equal(document.slides[0].elements[0].start.x, 0);
});

test("horizontal stack uses fixed sizes, padding/gap and cross-axis alignment without changing drawing order", () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const result = composed(document, { ...stack, elementIds: ["c", "a", "b"] });
  assert.deepEqual(boxes(result), [
    { x: 40, y: 90, width: 40, height: 20 },
    { x: 90, y: 80, width: 60, height: 40 },
    { x: 10, y: 70, width: 20, height: 60 },
  ]);
  assert.deepEqual(result.slides[0].elements.map(({ id }) => id), ["a", "b", "c"]);
  assert.deepEqual(result.analysis.diagnostics, []);
  assert.deepEqual(normalizePresentation(document), before);
});

test("vertical stack and every main-axis packing mode use remaining space explicitly", () => {
  for (const [justify, positions] of [["start", [10, 40, 90]], ["center", [30, 60, 110]], ["end", [50, 80, 130]], ["space-between", [10, 60, 130]]]) {
    const result = composed(fixture(), { ...stack, direction: "vertical", justify, align: "end" });
    assert.deepEqual(boxes(result).map(({ y }) => y), positions);
    assert.deepEqual(boxes(result).map(({ x }) => x), [250, 230, 270]);
  }
});

test("all alignment combinations preserve size and unspecified axes", () => {
  for (const [horizontal, xs] of [["start", [10, 10, 10]], ["center", [130, 120, 140]], ["end", [250, 230, 270]]]) {
    for (const [vertical, ys] of [["start", [10, 10, 10]], ["center", [90, 80, 70]], ["end", [170, 150, 130]]]) {
      const result = composed(fixture(), { ...base, type: "align", horizontal, vertical, padding: 10 });
      assert.deepEqual(boxes(result).map(({ x }) => x), xs);
      assert.deepEqual(boxes(result).map(({ y }) => y), ys);
    }
  }
  assert.deepEqual(boxes(composed(fixture(), { ...base, type: "align", horizontal: "start" })).map(({ y }) => y), [9, 9, 9]);
});

test("distribution uses equal edge gaps in explicit ID order and preserves the other axis", () => {
  const result = composed(fixture(), { ...base, type: "distribute", direction: "horizontal", padding: { left: 10, right: 20 } });
  assert.deepEqual(boxes(result).map(({ x }) => x), [10, 125, 260]);
  assert.deepEqual(boxes(result).map(({ y }) => y), [9, 9, 9]);
  assert.deepEqual(boxes(composed(fixture(), { ...base, type: "distribute", direction: "vertical" })).map(({ y }) => y), [0, 60, 140]);
});

test("empty selections are no-ops and single-item spread starts at the padded origin", () => {
  assert.deepEqual(boxes(composed(fixture(), { ...stack, elementIds: [] })), boxes(resolveLayout(fixture())));
  for (const rule of [{ ...stack, justify: "space-between" }, { ...base, type: "distribute", direction: "horizontal", padding: 10 }]) {
    const result = composed(fixture(), { ...rule, elementIds: ["a"] });
    assert.equal(boxes(result)[0].x, 10);
    assert.ok(Object.values(boxes(result)[0]).every(Number.isFinite));
  }
});

test("unsatisfiable rules are atomic and report their rule path and owner", () => {
  const document = fixture();
  const plain = boxes(resolveLayout(document));
  for (const rule of [
    { ...stack, area: { x: 0, y: 0, width: 100, height: 200 } },
    { ...stack, area: { x: 0, y: 0, width: 300, height: 30 } },
    { ...base, type: "align", horizontal: "center", area: { x: 0, y: 0, width: 50, height: 200 } },
    { ...stack, padding: 160 },
  ]) {
    const result = composed(document, rule);
    assert.deepEqual(boxes(result), plain);
    assert.equal(result.analysis.diagnostics.length, 1);
    assert.equal(result.analysis.diagnostics[0].path, "options.composition.0");
    assert.equal(result.analysis.diagnostics[0].slideId, "slide");
    assert.match(result.analysis.diagnostics[0].code, /^composition-(overflow|space-unavailable)$/);
  }
});

test("rules compose in declared order and a failed later rule retains earlier valid placement", () => {
  const document = fixture();
  const earlier = { ...base, type: "align", horizontal: "end", padding: 10 };
  const failure = { ...stack, area: { x: 0, y: 0, width: 1, height: 1 } };
  const result = composed(document, earlier, failure, { ...base, type: "align", vertical: "end" });
  assert.deepEqual(boxes(result).map(({ x }) => x), [250, 230, 270]);
  assert.deepEqual(boxes(result).map(({ y }) => y), [180, 160, 140]);
  assert.equal(result.analysis.diagnostics[0].path, "options.composition.1");
});

test("missing container, non-child and duplicate references never partially arrange elements", () => {
  for (const [rule, code] of [
    [{ ...stack, container: { kind: "slide", id: "missing" } }, "composition-container-missing"],
    [{ ...stack, elementIds: ["a", "missing"] }, "composition-element-missing"],
  ]) {
    const result = composed(fixture(), rule);
    assert.equal(result.analysis.diagnostics[0].code, code);
    assert.deepEqual(boxes(result), boxes(resolveLayout(fixture())));
  }
  assert.throws(() => composed(fixture(), { ...stack, elementIds: ["a", "a"] }), /unique/);
});

test("hidden elements and hidden ancestor subtrees are excluded unless explicitly included", () => {
  const normalized = normalizePresentation(fixture());
  normalized.slides[0].elements[1].hidden = true;
  const result = resolveNormalizedLayout(normalized, { composition: [stack] });
  assert.equal(boxes(result)[1].x, 7);
  assert.equal(boxes(result)[2].x, 60);
  normalized.slides[0].hidden = true;
  assert.deepEqual(boxes(resolveNormalizedLayout(normalized, { composition: [stack] })), normalized.slides[0].elements.map(({ box }) => box));
  assert.equal(boxes(resolveNormalizedLayout(normalized, { composition: [{ ...stack, includeHidden: true }] }))[1].x, 60);
});

test("nested group rules operate in local coordinateSize and moving groups leaves child geometry intact", () => {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  document.addSlide({ id: "slide" }).addElement({ type: "group", id: "outer", box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 100, height: 50 }, transform: { rotation: 20 }, children: [
    { type: "group", id: "inner", box: { x: 5, y: 5, width: 50, height: 30 }, coordinateSize: { width: 100, height: 60 }, children: [
      { type: "shape", id: "a", shape: "rect", box: { x: 0, y: 0, width: 20, height: 10 } },
      { type: "shape", id: "b", shape: "rect", box: { x: 0, y: 0, width: 30, height: 20 } },
    ] },
  ] });
  const rules = [
    { ...base, elementIds: ["outer"], type: "align", horizontal: "end" },
    { ...stack, container: { kind: "group", id: "inner" }, elementIds: ["a", "b"], gap: 5, padding: 0 },
  ];
  const result = composed(document, ...rules);
  const outer = result.slides[0].elements[0];
  assert.equal(outer.box.x, 100);
  assert.deepEqual(outer.children[0].box, { x: 5, y: 5, width: 50, height: 30 });
  assert.deepEqual(outer.children[0].children.map(({ box }) => box), [
    { x: 0, y: 25, width: 20, height: 10 }, { x: 25, y: 20, width: 30, height: 20 },
  ]);
});

test("visual alignment accounts for rotation while frame alignment preserves authored frame semantics", () => {
  const normalized = normalizePresentation(fixture());
  normalized.slides[0].elements[0].transform.rotation = 90;
  const rule = { ...base, elementIds: ["a"], type: "align", horizontal: "start", vertical: "start" };
  const frame = resolveNormalizedLayout(normalized, { composition: [rule] });
  const visual = resolveNormalizedLayout(normalized, { composition: [{ ...rule, bounds: "visual" }] });
  assert.deepEqual(boxes(frame)[0], { x: 0, y: 0, width: 40, height: 20 });
  assert.ok(Math.abs(boxes(visual)[0].x + 10) < 1e-9);
  assert.ok(Math.abs(boxes(visual)[0].y - 10) < 1e-9);
});

test("anchored connectors follow arranged shapes and cannot be independently stacked", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "link", start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" } });
  const result = composed(document, stack);
  assert.deepEqual(result.slides[0].elements[3].start, { x: 50, y: 100 });
  assert.deepEqual(result.slides[0].elements[3].end, { x: 60, y: 100 });
  const invalid = composed(document, { ...stack, elementIds: ["a", "link"] });
  assert.equal(invalid.analysis.diagnostics[0].code, "composition-anchored-connector");
  assert.deepEqual(boxes(invalid), boxes(resolveLayout(document)));
});

test("free connectors translate all points even if an authored box differs from path geometry", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "line", box: { x: 999, y: 999, width: 1, height: 1 }, start: { x: 10, y: 20 }, route: [{ x: 30, y: 40 }], end: { x: 50, y: 20 } });
  const result = composed(document, { ...base, elementIds: ["line"], type: "align", horizontal: "start", vertical: "start" });
  const line = result.slides[0].elements[3];
  assert.deepEqual(line.start, { x: 0, y: 0 });
  assert.deepEqual(line.end, { x: 40, y: 0 });
  assert.deepEqual(line.route, [{ x: 20, y: 20 }]);
});

test("layout elements can be arranged without duplicating them onto slides", () => {
  const document = fixture();
  document.defineSlideLayout({ id: "template", name: "Template", elements: [{ type: "shape", id: "template-shape", shape: "rect", box: { x: 1, y: 2, width: 10, height: 20 } }] });
  const result = composed(document, { type: "align", container: { kind: "layout", id: "template" }, elementIds: ["template-shape"], horizontal: "end", vertical: "end" });
  assert.deepEqual(result.layouts.find(({ id }) => id === "template").elements[0].box, { x: 290, y: 180, width: 10, height: 20 });
  assert.equal(result.slides[0].elements.length, 3);
});

test("invalid numeric configuration is rejected and representable decimal fits do not fail from rounding", () => {
  for (const invalid of [
    { gap: -1 }, { gap: Infinity }, { padding: { left: null } }, { padding: "10" }, { bounds: "painted" },
    { direction: "diagonal" }, { justify: "stretch" }, { align: "stretch" }, { includeHidden: 1 },
    { area: { x: NaN, y: 0, width: 1, height: 1 } },
  ]) assert.throws(() => composed(fixture(), { ...stack, ...invalid }));
  const normalized = normalizePresentation(fixture());
  for (const element of normalized.slides[0].elements) element.box.width = 0.1;
  const result = resolveNormalizedLayout(normalized, { composition: [{ ...stack, gap: 0, padding: 0, area: { x: 0, y: 0, width: 0.3, height: 100 } }] });
  assert.deepEqual(result.analysis.diagnostics, []);
});

test("composition and overflow analysis run on the final detached geometry", () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const options = { composition: [{ ...stack, area: { x: 400, y: 0, width: 300, height: 200 } }], diagnostics: { slideBounds: true } };
  const first = resolveLayout(document, options);
  assert.equal(first.analysis.diagnostics.filter(({ code }) => code === "slide-overflow").length, 3);
  assert.deepEqual(first, resolveLayout(document, options));
  first.slides[0].elements[0].box.x = -999;
  assert.deepEqual(normalizePresentation(document), before);
});


test("rotated exact fits tolerate only floating-point rounding", () => {
  const normalized = normalizePresentation(fixture());
  normalized.slides[0].elements[0].transform.rotation = 90;
  const result = resolveNormalizedLayout(normalized, { composition: [{ ...stack, elementIds: ["a"], bounds: "visual", padding: 0, area: { x: 0, y: 0, width: 20, height: 40 } }] });
  assert.deepEqual(result.analysis.diagnostics, []);
});

test("a connector translation that exceeds numeric range cannot partially apply its rule", () => {
  const document = fixture();
  document.slides[0].addElement({ type: "connector", id: "line", start: { x: 0, y: 0 }, end: { x: 10, y: 0 } });
  const normalized = normalizePresentation(document);
  const connector = normalized.slides[0].elements[3];
  connector.start.x = -Number.MAX_VALUE;
  connector.end.x = -Number.MAX_VALUE;
  connector.box = { x: -Number.MAX_VALUE, y: 0, width: 0, height: 0 };
  const before = structuredClone(normalized);
  const result = resolveNormalizedLayout(normalized, { composition: [{ ...base, type: "align", horizontal: "start", elementIds: ["a", "line"], area: { x: Number.MAX_VALUE, y: 0, width: 100, height: 100 } }] });
  assert.equal(result.analysis.diagnostics[0].code, "composition-space-unavailable");
  assert.deepEqual(result.slides[0].elements[0].box, before.slides[0].elements[0].box);
  assert.deepEqual(normalized, before);
});

test("stack spacing and size invariants hold across heterogeneous dimensions and directions", () => {
  for (const direction of ["horizontal", "vertical"]) {
    for (let sample = 1; sample <= 20; sample++) {
      const normalized = normalizePresentation(fixture());
      const source = normalized.slides[0].elements;
      source.forEach((element, i) => { element.box.width = 1 + (sample * (i + 1) * 7) % 50; element.box.height = 1 + (sample * (i + 2) * 11) % 40; });
      const gap = sample / 3;
      const result = resolveNormalizedLayout(normalized, { composition: [{ ...stack, direction, gap, justify: "center" }] });
      assert.deepEqual(result.analysis.diagnostics, []);
      const positioned = boxes(result);
      const axis = direction === "horizontal" ? "x" : "y";
      const dimension = direction === "horizontal" ? "width" : "height";
      positioned.forEach((box, i) => {
        assert.equal(box.width, source[i].box.width);
        assert.equal(box.height, source[i].box.height);
        if (i > 0) assert.ok(Math.abs(box[axis] - positioned[i - 1][axis] - positioned[i - 1][dimension] - gap) < 1e-9);
      });
      const extent = direction === "horizontal" ? 300 : 200;
      assert.ok(Math.abs(positioned[0][axis] - (extent - positioned.at(-1)[axis] - positioned.at(-1)[dimension])) < 1e-9);
    }
  }
});
