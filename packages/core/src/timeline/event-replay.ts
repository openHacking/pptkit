import type { EventTimelineEvent, EventTimelineLimits, EventTimelineObservation, EventTimelinePlan, EventTimelineReplay, EventTimelineReplayRequest, EventTimelineStageSample } from "../types/event-timeline.js";
import type { FiniteTimelinePlan } from "../types/timeline.js";
import { arithmetic, array, fail, record, time } from "./input.js";
import { captureEventProgram, EventCapture, eventError } from "./event-input.js";
import { captureFiniteTimelineData, sampleFiniteTimelineLeaves } from "./sample.js";

/** Pure foreground replay: a consumption record is an observation, not dispatch. */
export function replayEventTimeline(plan: EventTimelinePlan, request: EventTimelineReplayRequest, options?: EventTimelineLimits): EventTimelineReplay {
  try {
    const context = new EventCapture(options), program = captureEventProgram(plan, context, true);
    const stages = program.stages.map((stage, i) => {
      const path = `/plan/stages/${i}/timing`, finiteOptions = context.finiteOptions(path);
      let timing;
      try { timing = captureFiniteTimelineData(stage.timing as FiniteTimelinePlan, finiteOptions); } catch (error) { return eventError(error, path); }
      context.nodes += timing.nodeCount; context.charge(timing.idCharacters, path);
      return { ...stage, timing };
    });
    if (program.timingNodeCount !== context.nodes) fail("invalid-timeline-plan", "/plan/timingNodeCount", "Event plan aggregate timing node count disagrees with its stages.");
    const input = record(request, ["visitId", "timeMs", "elementIds", "events"], "/request");
    const visitId = context.identity(input.visitId, "/request/visitId"), at = time(input.timeMs, "/request/timeMs", context.budget);
    const current = new Set(context.inventory(input.elementIds, "/request/elementIds"));
    for (const stage of stages) if (!current.has(stage.target.elementId) || (stage.trigger.type === "object-click" && !current.has(stage.trigger.elementId))) fail("invalid-timeline-input", "/request/elementIds", "Current inventory no longer contains an event target or trigger source; recompile an explicit valid program.");
    const rows = array(input.events, context.budget.maxEvents, "/request/events"), events: EventTimelineEvent[] = [], observations: EventTimelineObservation[] = [];
    const seen = new Set<string>(); let previousTime = -1;
    for (let i = 0; i < rows.length; i++) {
      const path = `/request/events/${i}`, row = record(rows[i], ["id", "pageId", "visitId", "timeMs", "type", "elementId"], path);
      const identity = context.identity(row.id, `${path}/id`), pageId = context.identity(row.pageId, `${path}/pageId`), eventVisit = context.identity(row.visitId, `${path}/visitId`), atMs = time(row.timeMs, `${path}/timeMs`, context.budget);
      const key = JSON.stringify([pageId, eventVisit, identity]);
      if (seen.has(key)) fail("invalid-timeline-input", `${path}/id`, "Duplicate event identity within one page visit."); seen.add(key);
      if (atMs < previousTime) fail("invalid-timeline-input", `${path}/timeMs`, "Event log must retain nondecreasing time and explicit array order at equal times."); previousTime = atMs;
      if (row.type !== "page-click" && row.type !== "object-click") fail("unsupported-timeline-feature", `${path}/type`, "Unknown event kind.");
      if (row.type === "page-click" && Object.hasOwn(row, "elementId")) fail("unsupported-timeline-feature", `${path}/elementId`, "Page clicks do not identify an object source.");
      const elementId = row.type === "object-click" ? context.identity(row.elementId, `${path}/elementId`) : undefined;
      const base = { id: identity, pageId, visitId: eventVisit, timeMs: atMs };
      const event: EventTimelineEvent = elementId === undefined ? { ...base, type: "page-click" } : { ...base, type: "object-click", elementId };
      events.push(event);
      const observation = { index: i, eventId: identity, pageId, visitId: eventVisit, timeMs: atMs };
      const reason = pageId !== program.pageId ? "other-page" : eventVisit !== visitId ? "other-visit" : elementId !== undefined && !current.has(elementId) ? "unknown-source" : undefined;
      observations.push(reason ? { ...observation, status: "ignored", reason } : atMs > at ? { ...observation, status: "future" } : { ...observation, status: "ignored", reason: "no-gate" });
    }
    const started: EventTimelineStageSample[] = []; let cursor = 0;
    for (let i = 0; i < stages.length; i++) {
      const stage = stages[i]!, previous = started.at(-1); let start: number | undefined, consumed: number | null = null;
      if (stage.trigger.type === "enter") start = 0;
      else if (stage.trigger.type === "with-previous") start = previous!.startMs;
      else if (stage.trigger.type === "after-previous") start = previous!.endMs;
      else {
        const ready = previous?.endMs ?? 0;
        for (; cursor < events.length; cursor++) {
          const event = events[cursor]!, observation = observations[cursor]!;
          if (event.timeMs > at) break;
          if (observation.status === "ignored" && observation.reason !== "no-gate") continue;
          const reason = event.timeMs < ready ? "before-ready" : event.type !== stage.trigger.type ? "wrong-trigger" : stage.trigger.type === "object-click" && event.type === "object-click" && event.elementId !== stage.trigger.elementId ? "wrong-source" : undefined;
          const identity = { index: cursor, eventId: event.id, pageId: event.pageId, visitId: event.visitId, timeMs: event.timeMs };
          if (reason) { observations[cursor] = { ...identity, status: "ignored", reason }; continue; }
          start = event.timeMs; consumed = cursor;
          observations[cursor] = { ...identity, status: "consumed", stageId: stage.id };
          cursor++; break;
        }
      }
      if (start === undefined || start > at) break;
      const end = arithmetic(start, stage.timing.durationMs, "add", `/plan/stages/${i}/endMs`, context.budget);
      started.push(Object.freeze({ id: stage.id, target: stage.target, startMs: start, endMs: end, eventId: consumed === null ? null : events[consumed]!.id, eventIndex: consumed,
        samples: sampleFiniteTimelineLeaves(stage.timing.leaves, at - start) }));
    }
    return Object.freeze({ schemaVersion: "pptkit-event-timeline-replay-v1", pageId: program.pageId, visitId, timeMs: at, pendingStageId: stages[started.length]?.id ?? null,
      stages: Object.freeze(started), events: Object.freeze(observations.map(row => Object.freeze(row))) });
  } catch (error) { return eventError(error); }
}
