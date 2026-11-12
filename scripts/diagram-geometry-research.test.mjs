import assert from "node:assert/strict";
import test from "node:test";
import { nativeLengthPoints, observeDiagramGeometry, compareNativeGeometryObservations } from "./research/diagram-geometry-observation.mjs";

const point = (elementIndex, modelId, kind = "pres", identityStatus = "unique") => ({ elementIndex, modelId, kind, identityStatus });
const model = () => ({ status: "inspected", points: [point(1, "doc", "doc"), point(2, "15"), point(3, "18"), point(4, "24")] });
const shape = (ordinal = 1, name = "15", overrides = {}) => ({ ordinal, kind: "custom-shape", name, x: "1in", y: "2in", width: "2.54cm", height: "25.4mm", ancestorTransforms: [], paragraphs: ["Node"], ...overrides });
const snapshot = shapes => ({ frame: { x: 72, y: 144, width: 300, height: 200 }, shapes });

test("native length conversion has independent unit expectations, preserves zero and rejects unresolved forms", () => {
  for (const [raw, expected] of [["1in", 72], ["2.54cm", 72], ["25.4mm", 72], ["6pc", 72], ["72pt", 72], ["-0.5in", -36], [".5pt", .5], ["0", 0], ["+1pt", 1]]) {
    assert.ok(Math.abs(nativeLengthPoints(raw) - expected) < 1e-10, raw);
  }
  for (const raw of [undefined, null, 1, "", "1", "1px", "1em", "1%", "NaNpt", "Infinitypt", "1e3pt", "999999999999cm", "1pt junk", "1 pt"]) assert.equal(nativeLengthPoints(raw), undefined);
});

test("untransformed boxes use page points and an explicit frame tolerance without mutating source", () => {
  const input = snapshot([shape(), shape(2, "24", { x: "71.95pt" }), shape(3, "18", { x: "71.8pt" })]), before = structuredClone(input);
  const actual = observeDiagramGeometry(input, model());
  assert.deepEqual(actual.shapes[0].boundsPoints, { x: 72, y: 144, width: 72, height: 72 });
  assert.deepEqual(actual.shapes.map(s => s.outsideFrame), [false, false, true]);
  assert.equal(observeDiagramGeometry(input, model(), { tolerancePoints: 0 }).shapes[1].outsideFrame, true);
  assert.deepEqual(input, before);
  actual.shapes[0].raw.paragraphs[0] = "changed"; actual.framePoints.x = -100;
  assert.deepEqual(input, before);
  const degenerate = observeDiagramGeometry(snapshot([shape(1, "15", { width: "0", height: "0pt" })]), model());
  assert.equal(degenerate.shapes[0].geometryStatus, "axis-aligned");
});

test("rotated/group-transformed drawings retain raw geometry and never invent an axis-aligned box", () => {
  const input = snapshot([shape(1, "18", { x: undefined, y: undefined, transform: "rotate (-1.047) translate (14cm 7cm)" }),
    shape(2, "15", { ancestorTransforms: ["scale (2 2)"] }), shape(3, "24", { width: "10%" }), shape(4, "24", { height: "-1cm" })]);
  const actual = observeDiagramGeometry(input, model());
  assert.deepEqual(actual.shapes.map(s => s.geometryStatus), ["transform-source-only", "transform-source-only", "incomplete", "incomplete"]);
  assert.equal(actual.summary.axisAligned, 0);
  for (const s of actual.shapes) { assert.equal(s.boundsPoints, undefined); assert.equal(s.outsideFrame, undefined); }
  assert.equal(actual.shapes[0].raw.transform, input.shapes[0].transform);
  assert.deepEqual(actual.shapes[1].raw.ancestorTransforms, ["scale (2 2)"]);
});

test("producer names retain repeated occurrences and source ambiguity rather than assigning semantic ownership", () => {
  const input = snapshot([shape(1, "18"), shape(2, "18"), shape(3, "unknown"), shape(4, "doc"), shape(5, ""), shape(6, "collision")]);
  const source = model(); source.points.push(point(5, "collision", "pres", "duplicate"), point(6, "collision", "node", "duplicate"));
  const actual = observeDiagramGeometry(input, source);
  assert.deepEqual(actual.shapes.map(s => s.association.status), ["presentation-name-match", "presentation-name-match", "unmatched-name", "non-presentation-name", "missing-name", "ambiguous-source-identity"]);
  assert.deepEqual(actual.shapes[0].association.repeatedNameOrdinals, [1, 2]);
  assert.deepEqual(actual.shapes[5].association.sourceElementIndices, [5, 6]);
  assert.equal(actual.summary.repeatedNames, 1);
  assert.ok(observeDiagramGeometry(input, { ...source, status: "partial" }).shapes.every(s => s.association.status === "source-model-incomplete"));
  const noNames = observeDiagramGeometry(snapshot([shape(1, "15", { name: undefined })]), source);
  assert.equal(noNames.shapes[0].association.status, "missing-name");
});

test("geometry work/text/shape limits include repeated-name amplification and allow an unmodified retry", () => {
  const input = snapshot([shape(), shape(2, "18")]), before = structuredClone(input);
  for (const options of [{ maxShapes: 1 }, { maxWork: 1 }, { maxTextBytes: 3 }]) assert.throws(() => observeDiagramGeometry(input, model(), options));
  const repeated = snapshot(Array.from({ length: 100 }, (_, i) => shape(i + 1, "18")));
  assert.throws(() => observeDiagramGeometry(repeated, model(), { maxWork: 1000 }), /maxWork/);
  assert.equal(observeDiagramGeometry(repeated, model()).summary.shapes, 100);
  assert.deepEqual(input, before); assert.equal(observeDiagramGeometry(input, model()).summary.axisAligned, 2);
});

test("invalid occurrence/frame/text/options reject and deterministic geometry ignores record storage order", () => {
  for (const input of [snapshot([shape(), shape()]), snapshot([shape(0)]), snapshot([shape(1, "15", { paragraphs: [42] })]),
    { frame: { x: 0, y: 0, width: 0, height: 1 }, shapes: [] }, { frame: { x: 0, width: 1, height: 1 }, shapes: [] },
    snapshot([shape(1, "15", { x: NaN })])]) assert.throws(() => observeDiagramGeometry(input, model()));
  for (const options of [{ tolerancePoints: -1 }, { maxWork: Infinity }, { maxShapes: 0 }, { future: 1 }]) assert.throws(() => observeDiagramGeometry(snapshot([]), model(), options));
  const getter = { get shapes() { throw new Error("must not execute"); } }; assert.throws(() => observeDiagramGeometry(getter, model()), /accessor/);
  const source = model(); source.points.push(point(2, "extra")); assert.throws(() => observeDiagramGeometry(snapshot([]), source), /indices/);
  const a = observeDiagramGeometry(snapshot([shape(1), shape(2, "24")]), model());
  const b = observeDiagramGeometry(snapshot([shape(2, "24"), shape(1)]), model());
  assert.deepEqual(b, a);
});

test("cache-control comparison ignores unknown producer names but detects text, bounds, raw transformed size and association changes", () => {
  const observed = shapes => observeDiagramGeometry(snapshot(shapes), model());
  const a = observed([shape(1, "random-background-a"), shape(2, "18", { transform: "rotate (1)", width: "2cm" })]);
  const b = observed([shape(100, "random-background-b"), shape(200, "18", { transform: "rotate (1)", width: "2cm" })]);
  assert.equal(compareNativeGeometryObservations(a, b).exactObservedGeometryEqual, true);
  for (const mutate of [
    c => { c.shapes[0].boundsPoints.x += .01; },
    c => { c.shapes[0].raw.paragraphs[0] = "other"; },
    c => { c.shapes[1].raw.width = "3cm"; },
    c => { c.shapes[1].raw.transform = "rotate (2)"; },
    c => { c.shapes[1].association.sourceElementIndices = [99]; },
    c => { c.shapes[1].raw.textAreas = "0 0 1 1"; },
  ]) {
    const changed = structuredClone(b); mutate(changed);
    assert.equal(compareNativeGeometryObservations(a, changed).exactObservedGeometryEqual, false);
  }
  assert.throws(() => compareNativeGeometryObservations(a, { ...b, profile: "future" }));
  assert.deepEqual(compareNativeGeometryObservations(a, a).exactObservedGeometryEqual, true);
});
