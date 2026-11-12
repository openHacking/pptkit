import type { FiniteTimelineBoundary, FiniteTimelineLeaf, FiniteTimelineLimits, FiniteTimelinePlan, FiniteTimelineSample } from "../types/timeline.js";
import { arithmetic, array, fail, id, integer, limits, record, repeat, time } from "./input.js";
import { boundaryOrder } from "./compile.js";

/** Validates transported schedules too; no hidden global plan registry. */
export function captureFiniteTimelineData(plan: FiniteTimelinePlan, options?: FiniteTimelineLimits): { leaves: readonly FiniteTimelineLeaf[]; nodeCount: number; depth: number; durationMs: number; idCharacters: number } {
  const budget = limits(options), code = "invalid-timeline-plan";
  const source = record(plan, ["schemaVersion", "nodeCount", "depth", "durationMs", "leaves", "boundaries"], "", code);
  if (source.schemaVersion !== "pptkit-finite-timeline-v1") return fail(code, "/schemaVersion", "Unknown finite timeline plan schema.");
  const count = integer(source.nodeCount, "/nodeCount", 1, code), depth = integer(source.depth, "/depth", 1, code);
  if (count > budget.maxNodes || depth > budget.maxDepth) return fail("timeline-limit-exceeded", "", "Timeline plan exceeds its node/depth budget.");
  if (depth > count) return fail(code, "/depth", "Plan depth cannot exceed its node count.");
  const duration = time(source.durationMs, "/durationMs", budget, code);
  const rows = array(source.leaves, budget.maxNodes, "/leaves", code), ids = new Set<string>();
  if (rows.length > count || (count > 1 && rows.length === count) || (count === 1 && depth !== 1)) return fail(code, "/nodeCount", "Plan node/depth/leaf counts disagree.");
  const leaves: FiniteTimelineLeaf[] = [];
  let characters = 0;
  for (let i = 0; i < rows.length; i++) {
    const path = `/leaves/${i}`, row = record(rows[i], ["id", "order", "startMs", "endMs", "durationMs", "repeat", "fill"], path, code);
    const identity = id(row.id, `${path}/id`, budget, code);
    characters += identity.length;
    if (characters > budget.maxTotalIdLength) return fail("timeline-limit-exceeded", `${path}/id`, "Timeline plan exceeds maxTotalIdLength.");
    if (ids.has(identity) || row.order !== i) return fail(code, path, "Plan leaf identities/order must be unique and contiguous.");
    ids.add(identity);
    const start = time(row.startMs, `${path}/startMs`, budget, code), end = time(row.endMs, `${path}/endMs`, budget, code);
    const span = time(row.durationMs, `${path}/durationMs`, budget, code), repetitions = repeat(row.repeat, `${path}/repeat`, budget, code);
    const expectedEnd = arithmetic(start, arithmetic(span, repetitions, "multiply", `${path}/repeat`, budget, code), "add", `${path}/endMs`, budget, code);
    if (end !== expectedEnd || end > duration) return fail(code, `${path}/endMs`, "Plan interval disagrees with its repeat span or total duration.");
    if (row.fill !== "remove" && row.fill !== "hold") return fail(code, `${path}/fill`, "Unknown plan fill behavior.");
    leaves.push({ id: identity, order: i, startMs: start, endMs: end, durationMs: span, repeat: repetitions, fill: row.fill });
  }
  const actual = array(source.boundaries, budget.maxNodes * 2, "/boundaries", code);
  if (actual.length !== leaves.length * 2) return fail(code, "/boundaries", "Plan boundary count disagrees with its leaves.");
  const seen = new Uint8Array(leaves.length);
  let previous: FiniteTimelineBoundary | undefined;
  for (let i = 0; i < actual.length; i++) {
    const path = `/boundaries/${i}`, row = record(actual[i], ["id", "order", "timeMs", "type"], path, code);
    const order = integer(row.order, `${path}/order`, 0, code), leaf = leaves[order];
    if (!leaf || row.id !== leaf.id || (row.type !== "start" && row.type !== "end") || row.timeMs !== (row.type === "start" ? leaf.startMs : leaf.endMs)) return fail(code, path, "Plan boundaries disagree with their leaf intervals.");
    const bit = row.type === "start" ? 1 : 2;
    if (seen[order]! & bit) return fail(code, path, "Plan contains duplicate boundary records.");
    seen[order] = seen[order]! | bit;
    const current: FiniteTimelineBoundary = { id: leaf.id, order, timeMs: row.timeMs as number, type: row.type };
    if (previous && boundaryOrder(previous, current) > 0) return fail(code, path, "Plan boundaries are not in deterministic order.");
    previous = current;
  }
  return { leaves, nodeCount: count, depth, durationMs: duration, idCharacters: characters };
}

/** Random-access abstract progress. Never dispatches events or alters Core. */
export function sampleFiniteTimeline(plan: FiniteTimelinePlan, timeMs: number, options?: FiniteTimelineLimits): readonly FiniteTimelineSample[] {
  const at = integer(timeMs, "/timeMs"), { leaves } = captureFiniteTimelineData(plan, options);
  return sampleFiniteTimelineLeaves(leaves, at);
}
/** Private validated-leaf sampler shared by event replay. */
export function sampleFiniteTimelineLeaves(leaves: readonly FiniteTimelineLeaf[], at: number): readonly FiniteTimelineSample[] {
  return Object.freeze(leaves.map((leaf): FiniteTimelineSample => {
    if (at < leaf.startMs) return Object.freeze({ id: leaf.id, phase: "before", progress: null, iteration: null });
    if (at >= leaf.endMs) return Object.freeze({ id: leaf.id, phase: "ended", progress: leaf.fill === "hold" ? 1 : null, iteration: leaf.repeat - 1 });
    const elapsed = BigInt(at) - BigInt(leaf.startMs), duration = BigInt(leaf.durationMs);
    return Object.freeze({ id: leaf.id, phase: "active", progress: Number(elapsed % duration) / leaf.durationMs, iteration: Number(elapsed / duration) });
  }));
}
