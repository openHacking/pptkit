import type { FiniteTimelineBoundary, FiniteTimelineLeaf, FiniteTimelineLimits, FiniteTimelineNode, FiniteTimelinePlan } from "../types/timeline.js";
import { arithmetic, array, fail, id, limits, record, repeat, time } from "./input.js";
export function boundaryOrder(a: FiniteTimelineBoundary, b: FiniteTimelineBoundary): number {
  return a.timeMs - b.timeMs || a.order - b.order || (a.type === b.type ? 0 : a.type === "start" ? -1 : 1);
}
/** Compiles only abstract finite timing. No clock, effects, targets or I/O. */
export function compileFiniteTimeline(root: FiniteTimelineNode, options?: FiniteTimelineLimits): FiniteTimelinePlan {
  return compileFiniteTimelineData(root, options).plan;
}
/** Private aggregate accounting for event programs; no production tree rewalk. */
export function compileFiniteTimelineData(root: FiniteTimelineNode, options?: FiniteTimelineLimits): { plan: FiniteTimelinePlan; idCharacters: number } {
  const budget = limits(options), ids = new Set<string>(), active = new Set<object>();
  const leaves: FiniteTimelineLeaf[] = [], boundaries: FiniteTimelineBoundary[] = [];
  let nodeCount = 0, maximumDepth = 0, characters = 0;
  function walk(value: unknown, base: number, depth: number, path: string): number {
    if (++nodeCount > budget.maxNodes || depth > budget.maxDepth) return fail("timeline-limit-exceeded", path, "Timeline exceeds its node/depth budget.");
    maximumDepth = Math.max(maximumDepth, depth);
    const node = record(value, ["kind", "id", "delayMs", "durationMs", "repeat", "fill", "children"], path);
    if (active.has(value as object)) return fail("invalid-timeline-input", path, "Timeline contains an ancestor cycle.");
    const identity = id(node.id, `${path}/id`, budget);
    characters += identity.length;
    if (characters > budget.maxTotalIdLength) return fail("timeline-limit-exceeded", `${path}/id`, "Timeline exceeds maxTotalIdLength.");
    if (ids.has(identity)) return fail("invalid-timeline-input", `${path}/id`, "Timeline IDs must be globally unique within the input tree.");
    if (!["span", "sequence", "parallel"].includes(node.kind as string)) return fail("unsupported-timeline-feature", `${path}/kind`, "Unknown timeline node kind.");
    const allowed = node.kind === "span" ? ["kind", "id", "delayMs", "durationMs", "repeat", "fill"] : ["kind", "id", "delayMs", "children"];
    for (const key of Object.keys(node)) if (!allowed.includes(key)) return fail("unsupported-timeline-feature", `${path}/${key}`, "This property is not supported on this node kind.");
    ids.add(identity); active.add(value as object);
    const start = arithmetic(base, time(node.delayMs === undefined ? 0 : node.delayMs, `${path}/delayMs`, budget), "add", `${path}/startMs`, budget);
    let end = start;
    if (node.kind === "span") {
      const duration = time(node.durationMs, `${path}/durationMs`, budget), repetitions = repeat(node.repeat === undefined ? 1 : node.repeat, `${path}/repeat`, budget);
      const fill = node.fill === undefined ? "remove" : node.fill;
      if (fill !== "remove" && fill !== "hold") return fail("unsupported-timeline-feature", `${path}/fill`, "Unknown fill behavior.");
      end = arithmetic(start, arithmetic(duration, repetitions, "multiply", `${path}/repeat`, budget), "add", `${path}/endMs`, budget);
      const order = leaves.length;
      leaves.push(Object.freeze({ id: identity, order, startMs: start, endMs: end, durationMs: duration, repeat: repetitions, fill }));
      boundaries.push(Object.freeze({ id: identity, order, timeMs: start, type: "start" }), Object.freeze({ id: identity, order, timeMs: end, type: "end" }));
    } else {
      const children = array(node.children, budget.maxNodes - nodeCount, `${path}/children`);
      for (let i = 0; i < children.length; i++) end = Math.max(end, walk(children[i], node.kind === "sequence" ? end : start, depth + 1, `${path}/children/${i}`));
    }
    active.delete(value as object);
    return end;
  }
  const durationMs = walk(root, 0, 1, "");
  boundaries.sort(boundaryOrder);
  const plan: FiniteTimelinePlan = Object.freeze({ schemaVersion: "pptkit-finite-timeline-v1", nodeCount, depth: maximumDepth, durationMs, leaves: Object.freeze(leaves), boundaries: Object.freeze(boundaries) });
  return { plan, idCharacters: characters };
}
