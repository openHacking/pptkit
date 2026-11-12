import { compileEventTimeline, replayEventTimeline, type EventTimelineEvent } from "@pptkit/core";

export function replayClickExample(timeMs: number, visitId = "visit-1") {
  const plan = compileEventTimeline({
    pageId: "page", elementIds: ["shape", "button"],
    stages: [
      { id: "initial", target: { pageId: "page", elementId: "shape" },
        trigger: { type: "enter" }, timing: { kind: "span", id: "time", durationMs: 100 } },
      { id: "clicked", target: { pageId: "page", elementId: "shape" },
        trigger: { type: "object-click", elementId: "button" },
        timing: { kind: "span", id: "time", durationMs: 200, fill: "hold" } },
    ],
  });
  const events: readonly EventTimelineEvent[] = [
    { id: "early", pageId: "page", visitId: "visit-1", type: "object-click", elementId: "button", timeMs: 50 },
    { id: "ready", pageId: "page", visitId: "visit-1", type: "object-click", elementId: "button", timeMs: 100 },
  ];
  // Supply the current inventory on every evaluation. Removing button or
  // shape rejects the stale program until the caller explicitly recompiles.
  return replayEventTimeline(plan, { visitId, timeMs, elementIds: ["shape", "button"], events });
}
