// Executable research contract for ADV-06.1.1. Not an engine/public API.
const record = value => {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError("Timing nodes must be plain records.");
  for (const d of Object.values(Object.getOwnPropertyDescriptors(value))) if (!Object.hasOwn(d, "value")) throw new TypeError("Timing records cannot contain accessors.");
};
const integer = (n, name, min, max) => { if (!Number.isSafeInteger(n) || n < min || n > max) throw new RangeError(`Invalid ${name}.`); return n; };

export function compileTiming(root, options = {}) {
  record(options);
  for (const key of Object.keys(options)) if (!["maxNodes", "maxDepth", "maxTimeMs"].includes(key)) throw new TypeError("Unknown timing limit.");
  const maxNodes = integer(options.maxNodes ?? 10000, "maxNodes", 1, 100000);
  const maxDepth = integer(options.maxDepth ?? 128, "maxDepth", 1, 256);
  const maxTimeMs = integer(options.maxTimeMs ?? 86400000, "maxTimeMs", 0, Number.MAX_SAFE_INTEGER);
  const time = (n, label) => integer(n, label, 0, maxTimeMs);
  const ids = new Set(), active = new Set(), leaves = [], events = [];
  let visited = 0;
  function walk(node, base, depth) {
    if (++visited > maxNodes || depth > maxDepth) throw new RangeError("Timing complexity limit exceeded.");
    record(node);
    if (active.has(node)) throw new TypeError("Timing cycle.");
    if (typeof node.id !== "string" || !node.id.length || ids.has(node.id)) throw new TypeError("Timing IDs must be nonempty and unique.");
    ids.add(node.id); active.add(node);
    if (!["effect", "sequence", "parallel"].includes(node.kind)) throw new TypeError("Unknown timing node; no approximation is permitted.");
    const keys = node.kind === "effect" ? ["id", "kind", "delayMs", "durationMs", "repeat", "fill"] : ["id", "kind", "delayMs", "children"];
    for (const key of Object.keys(node)) if (!keys.includes(key)) throw new TypeError("Unknown timing property.");
    const startMs = time(base + time(node.delayMs ?? 0, "delayMs"), "startMs");
    let endMs = startMs;
    if (node.kind === "effect") {
      const durationMs = time(node.durationMs, "durationMs"), repeat = integer(node.repeat ?? 1, "repeat", 1, maxNodes);
      const fill = node.fill ?? "remove";
      if (!["hold", "remove"].includes(fill)) throw new TypeError("Invalid fill.");
      endMs = time(startMs + time(durationMs * repeat, "repeated duration"), "endMs");
      const order = leaves.length;
      leaves.push(Object.freeze({ id: node.id, startMs, endMs, durationMs, repeat, fill, order }));
      events.push(Object.freeze({ id: node.id, timeMs: startMs, type: "start", order }), Object.freeze({ id: node.id, timeMs: endMs, type: "end", order }));
    } else {
      if (!Array.isArray(node.children) || node.children.length > maxNodes - visited) throw new RangeError("Invalid or over-budget children.");
      for (const child of node.children) {
        const end = walk(child, node.kind === "sequence" ? endMs : startMs, depth + 1);
        endMs = Math.max(endMs, end);
      }
    }
    active.delete(node); return endMs;
  }
  const durationMs = walk(root, 0, 1);
  events.sort((a,b) => a.timeMs - b.timeMs || a.order - b.order || (a.type === "start" ? -1 : 1));
  return Object.freeze({ nodeCount: visited, durationMs, leaves: Object.freeze(leaves), events: Object.freeze(events) });
}

/** Absolute-time sampling; no mutable playhead or clock and no Core writes. */
export function sampleTiming(plan, timeMs) {
  integer(timeMs, "sample time", 0, Number.MAX_SAFE_INTEGER);
  return plan.leaves.map(leaf => {
    if (timeMs < leaf.startMs) return { id: leaf.id, phase: "before", progress: null, iteration: null };
    if (timeMs >= leaf.endMs) return { id: leaf.id, phase: "ended", progress: leaf.fill === "hold" ? 1 : null, iteration: leaf.repeat - 1 };
    const elapsed = timeMs - leaf.startMs;
    return { id: leaf.id, phase: "active", progress: (elapsed % leaf.durationMs) / leaf.durationMs, iteration: Math.floor(elapsed / leaf.durationMs) };
  });
}
