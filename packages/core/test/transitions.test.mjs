import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation } from "../dist/index.js";

test("transition snapshots, clear, copies and normalized defaults are independent", () => {
  const p = createPresentation(), input = { type: "fade", durationMs: 1234, afterMs: 0, onClick: false };
  const s = p.addSlide({ transition: input }); input.durationMs = 8;
  const before = normalizePresentation(p), copy = p.duplicateSlide(s.id);
  assert.equal(s.transition.durationMs, 1234); assert.ok(Object.isFrozen(s.transition));
  assert.notEqual(copy.transition, s.transition); assert.deepEqual(copy.transition, s.transition);
  assert.throws(() => { s.transition.durationMs = 10; }, TypeError);
  s.setTransition({ type: "fade", durationMs: 0 });
  assert.equal(normalizePresentation(p).slides[0].transition.onClick, true);
  assert.equal(copy.transition.durationMs, 1234); assert.equal(before.slides[0].transition.durationMs, 1234);
  s.setTransition(); assert.equal(s.transition, undefined); assert.equal("transition" in normalizePresentation(p).slides[0], false);
  assert.deepEqual(validatePresentation(p), []);
});

test("invalid transition values are located and snapshot failures leave state intact", () => {
  for (const transition of [null, [], { type: "morph", durationMs: 1 }, { type: "fade", durationMs: -1 }, { type: "fade", durationMs: .5 }, { type: "fade", durationMs: Infinity }, { type: "fade", durationMs: 4294967296 }, { type: "fade", durationMs: 2, afterMs: NaN }, { type: "fade", durationMs: 2, onClick: 1 }, { type: "fade", durationMs: 2, opaque: true }]) {
    const p = createPresentation(); p.addSlide({ id: "target", transition });
    const errors = validatePresentation(p); assert.ok(errors.length); assert.ok(errors.every(d => d.slideId === "target" && d.path.startsWith("slides.0.transition")));
  }
  const s = createPresentation().addSlide({ transition: { type: "fade", durationMs: 12 } });
  assert.throws(() => s.setTransition({ get type() { throw new Error("capture failed"); } }), /capture failed/);
  assert.equal(s.transition.durationMs, 12);
  for (const durationMs of [0, 1, 4294967295]) { const p = createPresentation(); p.addSlide({ transition: { type: "fade", durationMs, onClick: false, afterMs: 0 } }); assert.deepEqual(validatePresentation(p), []); }
});


test("Push/Wipe require cardinal direction and retain all eight variants independently", () => {
  for (const type of ["push", "wipe"]) for (const direction of ["left", "right", "up", "down"]) {
    const p = createPresentation(), s = p.addSlide({ transition: { type, direction, durationMs: 765 } });
    assert.deepEqual(validatePresentation(p), []);
    assert.equal(normalizePresentation(p).slides[0].transition.direction, direction);
    assert.deepEqual(p.duplicateSlide(s.id).transition, s.transition);
  }
  for (const transition of [{ type: "push", durationMs: 1 }, { type: "wipe", durationMs: 1, direction: "diagonal" }, { type: "fade", durationMs: 1, direction: "left" }]) {
    const p = createPresentation(); p.addSlide({ transition }); assert.ok(validatePresentation(p).length);
  }
});
