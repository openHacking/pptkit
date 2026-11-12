import assert from "node:assert/strict";
import test from "node:test";
import { observeCycleSpacing } from "./research/cycle-spacing-observation.mjs";
const snapshot = (gap = 0.5) => {
  const d = 40, r = 20 * Math.SQRT2 * (1 + gap), xy = [[100, 100 - r], [100 + r, 100], [100, 100 + r], [100 - r, 100]];
  return { frame: { x: 0, y: 0, width: 200, height: 200 }, circles: xy.map(([x, y], i) => ({ id: `n${i}`, box: { x: x - 20, y: y - 20, width: d, height: d }, proposed: { x: x - 20, y: y - 20, width: d, height: d }, strokeWidth: 1, strokeAlignment: "ctr" })) };
};
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-8, `${a} != ${b}`);

test("hand-derived circle ring infers spacing and keeps it descriptive rather than changing geometry", () => {
  const s = snapshot(0.75), before = structuredClone(s), r = observeCycleSpacing(s);
  assert.equal(r.consistentOrderedCircleRing, true); near(r.observedGapRatio, 0.75); near(r.center.x, 100); near(r.center.y, 100); near(r.radius, 35 * Math.SQRT2); near(r.diameter, 40); near(r.maxFitResidual, 0);
  near(r.minimumBoxInset, 80 - 35 * Math.SQRT2); near(r.rows[0].centeredStrokeInsetHypothesis, r.rows[0].minimumBoxInset - 0.5); assert.deepEqual(s, before);
  r.rows[0].fitted.x = 0; assert.deepEqual(s, before);
});

test("nearest EMU correction stays bounded and cannot hide larger geometric discrepancies", () => {
  const s = snapshot(); for (const c of s.circles) c.proposed.width += 0.2;
  const r = observeCycleSpacing(s); assert.equal(r.rounding.nearestQuantumMeetsComparison, false); assert.ok(r.rounding.maxRoundingAdjustment <= 1 / 25400 + 1e-12);
  assert.ok(r.rounding.roundedMaxCoordinateDelta > 0.19);
  const small = snapshot(); for (const c of small.circles) c.proposed.width += 0.00002;
  assert.equal(observeCycleSpacing(small).rounding.nearestQuantumMeetsComparison, true);
});

test("wrong order, non-ring coordinates, duplicate/invalid geometry and unknown options remain distinct", () => {
  const s = snapshot(); s.circles.reverse(); assert.equal(observeCycleSpacing(s).consistentOrderedCircleRing, false); assert.equal(observeCycleSpacing(s).observedGapRatio, undefined);
  const distorted = snapshot(); distorted.circles[0].box.x += 2; const r = observeCycleSpacing(distorted); assert.equal(r.consistentOrderedCircleRing, false); assert.ok(r.maxFitResidual > 0.1);
  for (const mutate of [s => { s.circles[0].box.width = 0; }, s => { s.circles[0].box.height += 1; }, s => { s.circles[0].id = s.circles[1].id; }, s => { s.circles[0].strokeWidth = -1; }, s => { s.frame.width = NaN; }, s => { s.circles.pop(); s.circles.pop(); }]) { const s = snapshot(); mutate(s); assert.throws(() => observeCycleSpacing(s)); }
  assert.throws(() => observeCycleSpacing(snapshot(), { unknown: true })); assert.throws(() => observeCycleSpacing(snapshot(), { maxNodes: 3 }));
});

test("exact work budgets reject without mutation and permit independent retry", () => {
  const s = snapshot(), before = structuredClone(s), result = observeCycleSpacing(s);
  assert.deepEqual(observeCycleSpacing(s, { maxWork: result.work }), result);
  assert.throws(() => observeCycleSpacing(s, { maxWork: result.work - 1 }), /maxWork/); assert.deepEqual(s, before);
  for (const options of [{ quantumPoints: 0 }, { fitTolerancePoints: NaN }, { comparisonTolerancePoints: -1 }, { maxWork: 0 }]) assert.throws(() => observeCycleSpacing(s, options));
});
