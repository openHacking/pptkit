import assert from "node:assert/strict";
import test from "node:test";
import { compileEventTimeline as compile, replayEventTimeline as replay, EventTimelineError } from "../dist/index.js";
const inventory = ["shape", "button", "other"];
const span = (durationMs, extra = {}) => ({ id: "leaf", kind: "span", durationMs, ...extra });
const stage = (id, type, duration = 10, extra = {}) => ({ id, target: { pageId: "p", elementId: "shape" }, trigger: type === "object-click" ? { type, elementId: "button" } : { type }, timing: span(duration), ...extra });
const input = stages => ({ pageId: "p", elementIds: inventory, stages });
const click = (id, timeMs, extra = {}) => ({ id, timeMs, pageId: "p", visitId: "visit", type: "page-click", ...extra });
const request = (timeMs, events = [], extra = {}) => ({ timeMs, events, visitId: "visit", elementIds: inventory, ...extra });
const raises = (run, code, path) => assert.throws(run, error => error instanceof EventTimelineError && error.code === code && (path === undefined || error.path === path));

test("hand-calculated activation differs from leaf delay and with/after refer to the immediate predecessor", () => {
  const plan = compile(input([stage("a", "enter", 100, { timing: span(100, { delayMs: 5 }) }), stage("b", "with-previous", 20, { timing: span(20, { delayMs: 2 }) }), stage("c", "after-previous", 5), stage("d", "page-click")]));
  const result = replay(plan, request(30, [click("early", 26), click("ready", 27)]));
  assert.deepEqual(result.stages.map(s => [s.id, s.startMs, s.endMs, s.eventId]), [["a", 0, 105, null], ["b", 0, 22, null], ["c", 22, 27, null], ["d", 27, 37, "ready"]]);
  assert.equal(result.pendingStageId, null); assert.equal(result.events[0].reason, "before-ready"); assert.equal(result.events[1].stageId, "d");
  const zero = replay(plan, request(0)); assert.deepEqual(zero.stages.map(s => s.id), ["a", "b"]); assert.ok(zero.stages.every(s => s.samples[0].phase === "before"));
  assert.equal(zero.pendingStageId, "c"); assert.equal(plan.timingNodeCount, 4);
});

test("scope/source/early clicks and equal-time consumption have an explicit complete observation table", () => {
  const plan = compile(input([stage("a", "enter"), stage("b", "object-click", 0), stage("c", "page-click", 0), stage("d", "page-click", 0)]));
  const events = [click("early", 1), click("page", 10, { pageId: "elsewhere" }), click("visit", 10, { visitId: "old" }), click("unknown", 10, { type: "object-click", elementId: "deleted" }), click("wrong-type", 10), click("wrong-object", 10, { type: "object-click", elementId: "shape" }), click("object", 10, { type: "object-click", elementId: "button" }), click("first-page", 10), click("second-page", 10), click("unused", 10), click("future", 11)];
  const result = replay(plan, request(10, events));
  assert.deepEqual(result.stages.map(s => [s.id, s.startMs, s.eventIndex]), [["a", 0, null], ["b", 10, 6], ["c", 10, 7], ["d", 10, 8]]);
  assert.deepEqual(result.events.map(e => e.status === "ignored" ? e.reason : e.status), ["before-ready", "other-page", "other-visit", "unknown-source", "wrong-trigger", "wrong-source", "consumed", "consumed", "consumed", "no-gate", "future"]);
  assert.equal(new Set(result.stages.map(s => s.eventIndex).filter(i => i !== null)).size, 3);
  const oneClick = replay(plan, request(10, events.slice(0, 8))); assert.equal(oneClick.pendingStageId, "d"); assert.equal(oneClick.stages.length, 3);
});

test("random-access seek, frozen pause inputs, replay and fresh visits have no dispatch or sticky consumption state", () => {
  const plan = compile(input([stage("initial", "enter", 2), stage("gated", "page-click", 4), stage("follow", "after-previous", 0)])), events = [click("old", 2, { visitId: "old" }), click("current", 3)];
  const before = structuredClone(events), captured = new Map([0, 2, 3, 5, 7, 100].map(at => [at, replay(plan, request(at, events))]));
  for (const at of [100, 3, 0, 7, 2, 5, 3, 5]) assert.deepEqual(replay(plan, request(at, events)), captured.get(at));
  assert.deepEqual(replay(plan, request(5, events)), replay(plan, request(5, events)), "paused active-time/log snapshots remain equal");
  const reentered = replay(plan, request(0, events, { visitId: "new" })); assert.equal(reentered.stages.length, 1); assert.equal(reentered.pendingStageId, "gated"); assert.ok(reentered.events.every(e => e.status === "ignored" && e.reason === "other-visit"));
  const freshEvents = [...events, click("current", 3, { visitId: "new" })]; assert.equal(replay(plan, request(3, freshEvents, { visitId: "new" })).stages[1].eventId, "current");
  const scoped = compile(input([stage("gate", "page-click", 0)]));
  const collidingIfJoined = [click("event", 0, { visitId: "v|x" }), click("event", 0, { pageId: "p|v", visitId: "x" })];
  const scopedResult = replay(scoped, request(0, collidingIfJoined, { visitId: "v|x" }));
  assert.equal(scopedResult.events[0].status, "consumed"); assert.equal(scopedResult.events[1].reason, "other-page");
  assert.deepEqual(events, before);
  const result = captured.get(5);
  for (const value of [plan, plan.elementIds, plan.stages, ...plan.stages, ...plan.stages.flatMap(s => [s.target, s.trigger, s.timing]), result, result.stages, result.events, ...result.events, ...result.stages.flatMap(s => [s, s.target, s.samples, ...s.samples])]) assert.ok(Object.isFrozen(value));
  assert.throws(() => { result.events[1].status = "future"; }, TypeError); assert.equal(replay(plan, request(5, events)).events[1].status, "consumed");
});

test("current inventories reject deleted targets and trigger sources, including inactive future stages", () => {
  const stages = [stage("a", "enter"), stage("future", "object-click", 1, { target: { pageId: "p", elementId: "other" } })], plan = compile(input(stages));
  for (const elementIds of [["shape", "other"], ["shape", "button"], []]) raises(() => replay(plan, request(0, [], { elementIds })), "invalid-event-timeline-input", "/request/elementIds");
  const smaller = compile(input([stage("a", "page-click", 0)])); assert.equal(replay(smaller, request(0, [], { elementIds: ["shape"] })).pendingStageId, "a");
  assert.equal(replay(smaller, request(0, [click("stale", 0, { type: "object-click", elementId: "other" })], { elementIds: ["shape", "added"] })).events[0].reason, "unknown-source");
  for (const changed of [{ ...stages[0], target: { pageId: "wrong", elementId: "shape" } }, { ...stages[0], target: { pageId: "p", elementId: "gone" } }, { ...stages[0], trigger: { type: "object-click", elementId: "gone" } }]) raises(() => compile(input([changed])), "invalid-event-timeline-input");
});

test("programs are snapshots and timing node identity is local to each stage, not a drawing identity", () => {
  const stages = [stage("a", "enter", 10), stage("b", "after-previous", 2)], source = input(stages), original = structuredClone(source), plan = compile(source);
  stages[0].target.elementId = "missing"; stages[0].timing.durationMs = 999; source.elementIds = []; stages.reverse();
  assert.deepEqual(replay(plan, request(11)).stages.map(s => [s.id, s.samples[0].id, s.endMs]), [["a", "leaf", 10], ["b", "leaf", 12]]);
  assert.equal(plan.elementIds.length, original.elementIds.length); assert.equal(plan.stages[0].target.elementId, "shape");
});

test("unknown contracts, malformed logs and invalid future entries reject rather than being filtered or sorted", () => {
  const good = stage("s", "page-click"), plan = compile(input([good]));
  for (const trigger of [{ type: "future" }, { type: "media-end" }, { type: "independent-track" }, { type: "page-click", elementId: undefined }]) raises(() => compile(input([{ ...good, trigger }])), "unsupported-event-timeline-feature");
  for (const changed of [{ ...good, target: { ...good.target, paragraphIndex: 0 } }, { ...good, timing: span(1, { autoReverse: true }) }, { ...good, restart: true }]) raises(() => compile(input([changed])), "unsupported-event-timeline-feature");
  for (const stages of [[stage("s", "with-previous")], [stage("s", "after-previous")], [stage("s", "enter"), stage("next", "enter")], [good, good]]) raises(() => compile(input(stages)), "invalid-event-timeline-input");
  const badLogs = [[click("x", -1)], [click("x", 2), click("y", 1)], [click("x", 0), click("x", 0)], [click("x", .5)], [click("x", Infinity)], [click("x", 0, { visitId: "" })], [click("x", 0, { type: "object-click", elementId: "" })]];
  for (const events of badLogs) raises(() => replay(plan, request(0, events)), "invalid-event-timeline-input");
  raises(() => replay(plan, request(0, [click("future", 10, { type: "media-end", pageId: "other" })])), "unsupported-event-timeline-feature");
  raises(() => replay(plan, request(0, [click("future", 10, { elementId: undefined })])), "unsupported-event-timeline-feature");
  raises(() => replay(plan, { timeMs: 0, visitId: "v", events: [] }), "invalid-event-timeline-input", "/request/elementIds");
  raises(() => compile(input([good]), { maxUnknown: 2 }), "unsupported-event-timeline-feature");
});

test("accessors and sparse/custom arrays never execute while validating programs, request inventory or log", () => {
  let calls = 0; const getter = () => { calls++; throw Error("must not execute"); }, source = input([stage("a", "enter")]);
  const getStage = structuredClone(source); Object.defineProperty(getStage.stages[0].target, "elementId", { get: getter, enumerable: true });
  raises(() => compile(getStage), "invalid-event-timeline-input");
  const plan = compile(source), r = request(0, [click("x", 0)]); Object.defineProperty(r.events[0], "timeMs", { get: getter, enumerable: true }); raises(() => replay(plan, r), "invalid-event-timeline-input");
  const sparse = new Array(1); raises(() => replay(plan, request(0, sparse)), "invalid-event-timeline-input");
  const custom = []; custom[Symbol.iterator] = getter; raises(() => compile({ ...source, stages: custom }), "invalid-event-timeline-input");
  raises(() => replay(plan, request(0, [], { elementIds: ["shape", "shape"] })), "invalid-event-timeline-input");
  const opts = { get maxEvents() { return getter(); } }; raises(() => compile(source, opts), "invalid-event-timeline-input"); assert.equal(calls, 0);
});

test("aggregate nodes, elements, stages, events, identifier text and time bounds are enforced with atomic retry", () => {
  const timing = { kind: "parallel", id: "r", children: Array.from({ length: 6000 }, (_, i) => ({ kind: "parallel", id: String(i), children: [] })) };
  const source = input([stage("a", "enter", 0, { timing }), stage("b", "after-previous", 0, { timing })]);
  raises(() => compile(source), "event-timeline-limit-exceeded"); assert.equal(compile({ ...source, stages: source.stages.slice(0, 1) }).timingNodeCount, 6001);
  const small = input([stage("a", "page-click", 0)]), plan = compile(small), events = [click("one", 0), click("two", 0)];
  raises(() => compile(small, { maxElements: 2 }), "event-timeline-limit-exceeded"); raises(() => compile(input([stage("a", "enter"), stage("b", "after-previous")]), { maxStages: 1 }), "event-timeline-limit-exceeded");
  raises(() => replay(plan, request(0, events), { maxEvents: 1 }), "event-timeline-limit-exceeded"); assert.equal(replay(plan, request(0, events), { maxEvents: 2 }).stages[0].eventId, "one");
  const empty = { pageId: "p", elementIds: [], stages: [] }, minimal = compile(empty, { maxTotalIdLength: 1 });
  raises(() => replay(minimal, { visitId: "v", timeMs: 0, elementIds: [], events: [] }, { maxTotalIdLength: 1 }), "event-timeline-limit-exceeded");
  assert.equal(replay(minimal, { visitId: "v", timeMs: 0, elementIds: [], events: [] }, { maxTotalIdLength: 2 }).stages.length, 0);
  raises(() => replay(plan, request(5, [click("future", 11)]), { maxTimeMs: 10 }), "event-timeline-limit-exceeded");
  for (const options of [{ maxStages: 10001 }, { maxElements: 100001 }, { maxEvents: 100001 }, { maxEvents: 0 }, { maxStages: Infinity }]) raises(() => compile(small, options), "invalid-event-timeline-input");
  const later = compile(input([stage("later", "page-click", 2)]), { maxTimeMs: Number.MAX_SAFE_INTEGER });
  raises(() => replay(later, request(Number.MAX_SAFE_INTEGER, [click("late", Number.MAX_SAFE_INTEGER)]), { maxTimeMs: Number.MAX_SAFE_INTEGER }), "event-timeline-time-overflow", "/plan/stages/0/endMs");
  raises(() => replay(later, request(49, [click("horizon", 49)]), { maxTimeMs: 50 }), "event-timeline-limit-exceeded", "/plan/stages/0/endMs");
  assert.equal(replay(later, request(1, [click("ok", 0)])).stages[0].endMs, 2);
});

test("JSON transport validates full programs, totals and future timing before any replay output", () => {
  const plan = compile(input([stage("a", "enter"), stage("b", "page-click")])), transported = JSON.parse(JSON.stringify(plan)), events = [click("ready", 10)];
  assert.deepEqual(replay(transported, request(11, events)), replay(plan, request(11, events)));
  for (const mutate of [p => { p.schemaVersion = "future"; }, p => { p.timingNodeCount++; }, p => { p.stages[1].timing.boundaries.pop(); }, p => { p.stages[1].target.elementId = "gone"; }, p => { p.stages[1].trigger.type = "future"; }, p => { p.stages[1].timing.leaves[0].repeat = Infinity; }]) {
    const changed = structuredClone(plan); mutate(changed); raises(() => replay(changed, request(0)), "invalid-event-timeline-plan");
  }
  const before = structuredClone(transported); replay(transported, request(11, events)); assert.deepEqual(transported, before);
});

// A separate event-first state machine advances the automatic frontier before
// feeding each log row once. Production replay instead scans a cursor per gate.
function gateOracle(stages, events, at, visit = "visit") {
  const result = []; let index = 0;
  const automatic = () => {
    while (index < stages.length) {
      const s = stages[index], previous = result.at(-1), kind = s.trigger.type;
      if (!["enter", "with-previous", "after-previous"].includes(kind)) break;
      const start = kind === "enter" ? 0 : kind === "with-previous" ? previous.start : previous.end;
      if (start > at) break;
      result.push({ id: s.id, start, end: start + (s.timing.delayMs ?? 0) + s.timing.durationMs * (s.timing.repeat ?? 1), eventId: null }); index++;
    }
  };
  automatic();
  for (const e of events) {
    if (e.timeMs > at) break;
    const s = stages[index]; if (!s || ["enter", "with-previous", "after-previous"].includes(s.trigger.type) || e.pageId !== "p" || e.visitId !== visit || (e.type === "object-click" && !inventory.includes(e.elementId)) || e.timeMs < (result.at(-1)?.end ?? 0) || e.type !== s.trigger.type || (e.type === "object-click" && e.elementId !== s.trigger.elementId)) continue;
    result.push({ id: s.id, start: e.timeMs, end: e.timeMs + (s.timing.delayMs ?? 0) + s.timing.durationMs * (s.timing.repeat ?? 1), eventId: e.id }); index++; automatic();
  }
  return result;
}

test("300 generated programs/logs agree with an independent event-first state table under shuffled seeks", () => {
  let seed = 871, identity = 0; const rand = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  const kinds = ["with-previous", "after-previous", "page-click", "object-click"];
  for (let n = 0; n < 300; n++) {
    const stages = Array.from({ length: 1 + rand(12) }, (_, i) => stage(`s${i}`, i === 0 ? rand(2) ? "enter" : "page-click" : kinds[rand(kinds.length)], rand(5), { timing: span(rand(5), { delayMs: rand(3), repeat: 1 + rand(3) }) }));
    let t = 0; const events = Array.from({ length: 10 + rand(40) }, () => { t += rand(3); return click(`e${identity++}`, t, { pageId: rand(8) ? "p" : "other", visitId: rand(8) ? "visit" : "old", ...(rand(2) ? { type: "object-click", elementId: rand(2) ? "button" : "shape" } : {}) }); });
    const plan = compile(input(stages));
    for (const at of [0, t + 20, rand(t + 21), rand(t + 21)]) {
      const r = replay(plan, request(at, events)); assert.deepEqual(r.stages.map(s => ({ id: s.id, start: s.startMs, end: s.endMs, eventId: s.eventId })), gateOracle(stages, events, at));
      assert.equal(new Set(r.events.filter(e => e.status === "consumed").map(e => e.index)).size, r.events.filter(e => e.status === "consumed").length);
    }
  }
});

test("nested finite trees retain delayed leaf intervals/fills while click gates use the complete stage span", () => {
  const timing = { kind: "sequence", id: "tree", delayMs: 2, children: [
    { kind: "span", id: "a", durationMs: 3, fill: "hold" },
    { kind: "parallel", id: "p", children: [{ kind: "span", id: "b", delayMs: 1, durationMs: 2 }, { kind: "span", id: "zero", delayMs: 4, durationMs: 0 }] },
    { kind: "span", id: "d", durationMs: 1, repeat: 2 },
  ] };
  const plan = compile(input([stage("nested", "page-click", 1, { timing }), stage("gate", "object-click", 0)]));
  const events = [click("activate", 5), click("early", 15, { type: "object-click", elementId: "button" }), click("at-end", 16, { type: "object-click", elementId: "button" })];
  const result = replay(plan, request(14, events));
  assert.equal(result.stages[0].startMs, 5); assert.equal(result.stages[0].endMs, 16); assert.equal(result.pendingStageId, "gate");
  assert.deepEqual(result.stages[0].samples, [{ id: "a", phase: "ended", progress: 1, iteration: 0 }, { id: "b", phase: "ended", progress: null, iteration: 0 }, { id: "zero", phase: "ended", progress: null, iteration: 0 }, { id: "d", phase: "active", progress: 0, iteration: 0 }]);
  assert.equal(replay(plan, request(15, events)).stages[0].samples[3].iteration, 1);
  const ended = replay(plan, request(16, events)); assert.equal(ended.stages[1].eventId, "at-end"); assert.equal(ended.events[1].reason, "before-ready"); assert.ok(ended.stages[0].samples.every(s => s.phase === "ended"));
});

test("one thousand zero-duration gates consume one thousand distinct equal-time events without reuse or state leakage", () => {
  const stages = Array.from({ length: 1000 }, (_, i) => stage(`s${i}`, "page-click", 0)), source = input(stages), plan = compile(source);
  const events = Array.from({ length: 1000 }, (_, i) => click(`e${i}`, 0)), result = replay(plan, request(0, events));
  assert.equal(result.stages.length, 1000); assert.equal(result.pendingStageId, null); assert.ok(result.events.every((e, i) => e.status === "consumed" && e.stageId === `s${i}`));
  assert.deepEqual(result.stages.map(s => s.eventIndex), events.map((_, i) => i));
  assert.equal(replay(plan, request(0, events.slice(0, 999))).pendingStageId, "s999");
  raises(() => compile(input([...stages, stage("overflow", "page-click", 0)])), "event-timeline-limit-exceeded");
  assert.equal(compile(input([...stages, stage("overflow", "page-click", 0)]), { maxStages: 1001 }).stages.length, 1001);
  const tooMany = Array.from({ length: 10001 }, (_, i) => click(`e${i}`, 0)); raises(() => replay(plan, request(0, tooMany)), "event-timeline-limit-exceeded");
  assert.equal(replay(plan, request(0, events)).events.length, 1000);
});
