import assert from "node:assert/strict";
import test from "node:test";
import { compileFiniteTimeline as compile, sampleFiniteTimeline as sample, FiniteTimelineError } from "../dist/index.js";
const span = (id, durationMs, extra = {}) => ({ kind: "span", id, durationMs, ...extra });
const group = (id, kind, children, extra = {}) => ({ kind, id, children, ...extra });
const raises = (run, code, path) => assert.throws(run, error => error instanceof FiniteTimelineError && error.code === code && (path === undefined || error.path === path));

test("public finite schedules preserve hand-calculated delayed sequence/parallel and stable equal-time boundaries", () => {
  const source = group("root", "sequence", [span("a", 10), group("pair", "parallel", [span("b", 4), span("c", 3, { delayMs: 5 })], { delayMs: 2 }), span("zero", 0), span("last", 2)], { delayMs: 5 });
  const before = structuredClone(source), plan = compile(source);
  assert.equal(plan.schemaVersion, "pptkit-finite-timeline-v1"); assert.equal(plan.nodeCount, 7); assert.equal(plan.depth, 3); assert.equal(plan.durationMs, 27);
  assert.deepEqual(plan.leaves.map(x => [x.id, x.startMs, x.endMs]), [["a", 5, 15], ["b", 17, 21], ["c", 22, 25], ["zero", 25, 25], ["last", 25, 27]]);
  assert.deepEqual(plan.boundaries.filter(b => b.timeMs === 25).map(b => [b.id, b.type]), [["c", "end"], ["zero", "start"], ["zero", "end"], ["last", "start"]]);
  assert.deepEqual(source, before);
  for (const value of [plan, plan.leaves, ...plan.leaves, plan.boundaries, ...plan.boundaries]) assert.ok(Object.isFrozen(value));
  source.children[0].durationMs = 999; source.children.reverse(); assert.equal(plan.leaves[0].durationMs, 10);
  assert.throws(() => { plan.leaves[0].endMs = 99; }, TypeError);
});

test("start/repeat/final/zero boundaries and both fills are exact and independent of seek order", () => {
  for (const fill of ["remove", "hold"]) {
    const plan = compile(span("repeat", 10, { delayMs: 5, repeat: 3, fill }));
    const expected = new Map([[0, ["before", null, null]], [4, ["before", null, null]], [5, ["active", 0, 0]], [14, ["active", .9, 0]], [15, ["active", 0, 1]], [24, ["active", .9, 1]], [25, ["active", 0, 2]], [34, ["active", .9, 2]], [35, ["ended", fill === "hold" ? 1 : null, 2]], [Number.MAX_SAFE_INTEGER, ["ended", fill === "hold" ? 1 : null, 2]]]);
    for (const at of [35, 25, 4, 15, 14, 24, 5, 34, 0, 15, Number.MAX_SAFE_INTEGER]) {
      const samples = sample(plan, at), row = samples[0]; assert.deepEqual([row.phase, row.progress, row.iteration], expected.get(at));
      assert.ok(Object.isFrozen(samples) && Object.isFrozen(row));
    }
    const zero = compile(span("zero", 0, { delayMs: 7, repeat: 19, fill }));
    assert.equal(sample(zero, 6)[0].phase, "before"); assert.deepEqual(sample(zero, 7)[0], { id: "zero", phase: "ended", progress: fill === "hold" ? 1 : null, iteration: 18 });
    assert.deepEqual(zero.boundaries.map(b => b.type), ["start", "end"]);
  }
  const empty = compile(group("empty", "parallel", [], { delayMs: 6 })); assert.equal(empty.durationMs, 6); assert.deepEqual(sample(empty, 100), []);
  assert.equal(compile(group("s", "sequence", [group("e", "parallel", [], { delayMs: 9 }), span("a", 2)])).durationMs, 11);
});

test("large integer repeat is algebraic, independent of node count and exact near the safe integer ceiling", () => {
  const options = { maxNodes: 1, maxRepeat: Number.MAX_SAFE_INTEGER, maxTimeMs: Number.MAX_SAFE_INTEGER };
  const plan = compile(span("all", 1, { repeat: Number.MAX_SAFE_INTEGER, fill: "hold" }), options);
  assert.equal(plan.leaves.length, 1); assert.equal(plan.boundaries.length, 2);
  assert.deepEqual(sample(plan, Number.MAX_SAFE_INTEGER - 1, options)[0], { id: "all", phase: "active", progress: 0, iteration: Number.MAX_SAFE_INTEGER - 1 });
  assert.equal(sample(plan, Number.MAX_SAFE_INTEGER, options)[0].iteration, Number.MAX_SAFE_INTEGER - 1);
  const zero = compile(span("z", 0, { repeat: Number.MAX_SAFE_INTEGER }), options); assert.equal(sample(zero, 0, options)[0].iteration, Number.MAX_SAFE_INTEGER - 1);
  raises(() => compile(span("x", 3, { repeat: 3002399751580331 }), options), "timeline-time-overflow", "/repeat");
  raises(() => compile(span("x", Number.MAX_SAFE_INTEGER, { delayMs: 1 }), options), "timeline-time-overflow", "/endMs");
  raises(() => compile(group("r", "sequence", [span("a", Number.MAX_SAFE_INTEGER), span("b", 1)]), { ...options, maxNodes: 3 }), "timeline-time-overflow");
  raises(() => compile(span("x", 1, { repeat: 1000001 })), "timeline-limit-exceeded", "/repeat");
});

test("unsupported timing features reject explicitly instead of being approximated", () => {
  for (const extra of [{ repeat: Infinity }, { repeat: "indefinite" }, { repeat: .5 }, { repeat: 1.5 }, { autoReverse: false }, { autoReverse: true }, { easing: "linear" }, { trigger: "click" }, { target: { elementId: "x" } }, { restart: "always" }, { fill: "freeze" }, { children: [] }]) raises(() => compile(span("x", 1, extra)), "unsupported-timeline-feature");
  for (const root of [{ id: "x", kind: "future" }, group("x", "parallel", [], { repeat: 2 }), group("x", "sequence", [], { durationMs: 2 })]) raises(() => compile(root), "unsupported-timeline-feature");
  raises(() => compile(span("x", 1), { maxMilliseconds: 10 }), "unsupported-timeline-feature", "/limits");
});

test("malformed data, inherited/accessor fields, sparse/custom arrays, shared nodes and ancestor cycles reject without invoking accessors", () => {
  const cycle = group("cycle", "sequence", []); cycle.children.push(cycle);
  const leaf = span("same", 1), sparse = new Array(1), getterArray = [leaf]; let calls = 0;
  Object.defineProperty(getterArray, "0", { get() { calls++; throw Error("must not run"); }, enumerable: true });
  const iteratorArray = [leaf]; iteratorArray[Symbol.iterator] = () => { calls++; throw Error("must not run"); };
  const accessor = { kind: "span", durationMs: 1, get id() { calls++; throw Error("must not run"); } };
  const nonenumerable = span("x", 1); Object.defineProperty(nonenumerable, "durationMs", { enumerable: false });
  const bad = [null, [], new Date(), Object.create({ kind: "span", id: "x", durationMs: 1 }), span("", 1), span("x", -1), span("x", .5), span("x", NaN), span("x", Infinity), span("x", 1, { repeat: 0 }), span("x", 1, { repeat: null }), span("x", 1, { delayMs: null }), accessor, nonenumerable, group("r", "parallel", sparse), group("r", "parallel", getterArray), group("r", "parallel", iteratorArray), group("r", "parallel", [leaf, leaf]), cycle];
  for (const value of bad) assert.throws(() => compile(value), FiniteTimelineError);
  assert.equal(calls, 0);
  const plain = Object.assign(Object.create(null), span("null-record", 1)); assert.equal(compile(plain).durationMs, 1);
  assert.equal(compile(span("negative-zero", -0, { delayMs: -0 })).durationMs, 0);
  assert.ok(!Object.is(compile(span("negative-zero", -0)).leaves[0].durationMs, -0));
});

test("exact node/depth/time/identifier limits fail atomically and retry with the same unchanged input", () => {
  const source = group("root", "parallel", [span("first", 10), span("second", 20)]), before = structuredClone(source);
  const exact = { maxNodes: 3, maxDepth: 2, maxTimeMs: 20, maxIdLength: 6, maxTotalIdLength: 15 };
  assert.equal(compile(source, exact).durationMs, 20);
  for (const options of [{ ...exact, maxNodes: 2 }, { ...exact, maxDepth: 1 }, { ...exact, maxTimeMs: 19 }, { ...exact, maxIdLength: 5 }, { ...exact, maxTotalIdLength: 14 }]) raises(() => compile(source, options), "timeline-limit-exceeded");
  for (const options of [null, { maxNodes: 100001 }, { maxDepth: 257 }, { maxTimeMs: -1 }, { maxNodes: Infinity }, { maxRepeat: 0 }, { maxIdLength: 65537 }, { maxTotalIdLength: 16000001 }]) assert.throws(() => compile(source, options), FiniteTimelineError);
  assert.deepEqual(source, before); assert.equal(compile(source, exact).nodeCount, 3);
  let nested = span("leaf", 0); for (let i = 0; i < 255; i++) nested = group(`g${i}`, "parallel", [nested]);
  raises(() => compile(nested), "timeline-limit-exceeded"); assert.equal(compile(nested, { maxDepth: 256 }).depth, 256);
});

test("transported plans sample identically and malformed metadata/intervals/boundaries are never trusted", () => {
  const original = compile(group("root", "parallel", [span("a", 10), span("zero", 0), span("b", 3, { delayMs: 5, repeat: 2 })]));
  const transported = JSON.parse(JSON.stringify(original)); assert.deepEqual(sample(transported, 7), sample(original, 7));
  const mutations = [p => { p.schemaVersion = "future"; }, p => { p.nodeCount = 1; }, p => { p.depth = 10; }, p => { p.durationMs = 2; }, p => { p.leaves[0].endMs++; }, p => { p.leaves[0].repeat = 2; }, p => { p.leaves[1].id = p.leaves[0].id; }, p => { p.leaves[0].order = 2; }, p => { p.leaves[0].fill = "freeze"; }, p => { p.boundaries.pop(); }, p => { p.boundaries[0].timeMs++; }, p => { p.boundaries[0].order = 2; }, p => { p.boundaries[1] = structuredClone(p.boundaries[0]); }, p => { p.boundaries.reverse(); }, p => { p.boundaries[0].opaque = true; }];
  for (const mutate of mutations) { const p = structuredClone(original); mutate(p); raises(() => sample(p, 0), "invalid-timeline-plan"); }
  let accessed = false; const p = structuredClone(original); Object.defineProperty(p.leaves[0], "startMs", { get() { accessed = true; return 0; }, enumerable: true }); raises(() => sample(p, 0), "invalid-timeline-plan"); assert.equal(accessed, false);
  const snapshot = structuredClone(transported); sample(transported, 999); assert.deepEqual(transported, snapshot);
  for (const at of [-1, .5, NaN, Infinity]) raises(() => sample(original, at), "invalid-timeline-input", "/timeMs");
  raises(() => sample(original, 0, { maxNodes: 2 }), "timeline-limit-exceeded");
  transported.leaves[0].endMs++; assert.equal(sample(original, 0)[0].phase, "active");
});

// Independent relative intervals and per-millisecond expanded repeat samples.
function oracle(node) {
  const delay = node.delayMs ?? 0;
  if (node.kind === "span") return { duration: delay + node.durationMs * (node.repeat ?? 1), rows: [{ ...node, start: delay, end: delay + node.durationMs * (node.repeat ?? 1) }] };
  const children = node.children.map(oracle), offsets = []; let cursor = 0;
  for (const c of children) { offsets.push(node.kind === "sequence" ? cursor : 0); cursor = node.kind === "sequence" ? cursor + c.duration : Math.max(cursor, c.duration); }
  return { duration: delay + cursor, rows: children.flatMap((c, i) => c.rows.map(r => ({ ...r, start: r.start + delay + offsets[i], end: r.end + delay + offsets[i] }))) };
}
function oracleSample(row, at) {
  if (at < row.start) return { id: row.id, phase: "before", progress: null, iteration: null };
  const count = row.repeat ?? 1;
  for (let i = 0; i < count; i++) for (let t = 0; t < row.durationMs; t++) if (at === row.start + i * row.durationMs + t) return { id: row.id, phase: "active", progress: t / row.durationMs, iteration: i };
  return { id: row.id, phase: "ended", progress: row.fill === "hold" ? 1 : null, iteration: count - 1 };
}

test("400 seeded trees match a separate relative-interval and expanded per-millisecond oracle under shuffled seeks", () => {
  let state = 8451, nextId = 0; const random = n => { state = (Math.imul(state, 1664525) + 1013904223) >>> 0; return state % n; };
  const make = depth => !depth || random(3) === 0 ? span(`n${++nextId}`, random(7), { delayMs: random(4), repeat: 1 + random(3), fill: random(2) ? "hold" : "remove" }) : group(`g${++nextId}`, random(2) ? "sequence" : "parallel", Array.from({ length: random(4) }, () => make(depth - 1)), { delayMs: random(4) });
  for (let i = 0; i < 400; i++) {
    const source = make(4), expected = oracle(source), plan = compile(source);
    assert.equal(plan.durationMs, expected.duration); assert.deepEqual(plan.leaves.map(r => [r.id, r.startMs, r.endMs]), expected.rows.map(r => [r.id, r.start, r.end]));
    for (const at of [expected.duration + 1, 0, random(expected.duration + 2), random(expected.duration + 2)]) assert.deepEqual(sample(plan, at), expected.rows.map(r => oracleSample(r, at)));
  }
});

test("wide schedules meet the default node boundary without repeat expansion and remain stable under record-key permutation", () => {
  const children = Array.from({ length: 9999 }, (_, i) => span(`n${i}`, i % 3, { delayMs: i % 2 }));
  const source = group("root", "parallel", children), plan = compile(source);
  assert.equal(plan.nodeCount, 10000); assert.equal(plan.leaves.length, 9999); assert.equal(plan.boundaries.length, 19998);
  assert.equal(sample(plan, 1).length, 9999);
  const reordered = { children: children.map(n => ({ durationMs: n.durationMs, delayMs: n.delayMs, id: n.id, kind: n.kind })), id: "root", kind: "parallel" };
  assert.deepEqual(compile(reordered), plan);
  const overflow = group("root", "parallel", [...children, span("extra", 0)]);
  raises(() => compile(overflow), "timeline-limit-exceeded"); assert.equal(compile(overflow, { maxNodes: 10001 }).nodeCount, 10001);
  const badFuture = structuredClone(plan); badFuture.leaves.at(-1).repeat = Infinity;
  raises(() => sample(badFuture, 0), "invalid-timeline-plan");
  let calls = 0; const options = { get maxTimeMs() { calls++; return 100; } };
  raises(() => compile(span("x", 1), options), "invalid-timeline-input", "/limits/maxTimeMs"); assert.equal(calls, 0);
});
