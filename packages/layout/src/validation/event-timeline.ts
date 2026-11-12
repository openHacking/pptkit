import { compileEventTimeline, EventTimelineError, replayEventTimeline, type NormalizedElement, type NormalizedSlide, type EventTimelinePlan, type EventTimelineStageInput } from "@pptkit/core";
function data(value: unknown, keys: readonly string[], path: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new EventTimelineError("invalid-event-timeline-plan", "Expected a bound timing data record.", path);
  const result: Record<string, unknown> = Object.create(null);
  for (const key of Reflect.ownKeys(value)) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
    if (typeof key !== "string" || !keys.includes(key) || !("value" in descriptor) || !descriptor.enumerable) throw new EventTimelineError("invalid-event-timeline-plan", "Unknown/accessor bound timing field.", path);
    result[key] = descriptor.value;
  }
  return result;
}
/** Compares already validated plain data without invoking JSON hooks. */
function same(a: unknown, b: unknown): boolean {
  const pending: [unknown, unknown][] = [[a, b]];
  while (pending.length) {
    const [x, y] = pending.pop()!;
    if (Object.is(x, y)) continue;
    if (!x || !y || typeof x !== "object" || typeof y !== "object" || Array.isArray(x) !== Array.isArray(y)) return false;
    const left = x as Record<string, unknown>, right = y as Record<string, unknown>, keys = Object.keys(left);
    if (keys.length !== Object.keys(right).length) return false;
    for (const key of keys) { if (!Object.hasOwn(right, key)) return false; pending.push([left[key], right[key]]); }
  }
  return true;
}
/** Targets are source-owned whole objects; layout elements remain shared. */
export function validateSlideEventTimeline(slide: NormalizedSlide): void {
  if (slide.eventTimeline === undefined) return;
  const bound = data(slide.eventTimeline, ["schemaVersion", "source", "plan"], "/eventTimeline");
  if (bound.schemaVersion !== "pptkit-slide-event-timeline-v1") throw new EventTimelineError("invalid-event-timeline-plan", "Unknown bound slide timing schema.", "/eventTimeline/schemaVersion");
  const source = data(bound.source, ["stages"], "/eventTimeline/source");
  if (slide.elements.length > 10000) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timeline exceeds its owned object budget.", "/eventTimeline");
  const pending: NormalizedElement[] = [...slide.elements], elementIds: string[] = [];
  while (pending.length) {
    const element = pending.pop()!;
    if (elementIds.length >= 10000) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timeline exceeds its owned object budget.", "/eventTimeline");
    elementIds.push(element.id);
    if (element.type === "group") {
      if (element.children.length > 10000 - elementIds.length - pending.length) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timeline exceeds its owned object budget.", "/eventTimeline");
      for (const child of element.children) pending.push(child);
    }
  }
  // Validate the stored plan completely before comparing it, including future
  // records. Recompile source too so group identity/topology is never discarded.
  const plan = bound.plan as EventTimelinePlan;
  replayEventTimeline(plan, { visitId: "layout-validation", timeMs: 0, elementIds, events: [] });
  let expected: EventTimelinePlan;
  try { expected = compileEventTimeline({ pageId: slide.id, elementIds: elementIds.sort(), stages: source.stages as readonly EventTimelineStageInput[] }); }
  catch (error) {
    if (!(error instanceof EventTimelineError)) throw error;
    throw new EventTimelineError(error.code === "event-timeline-limit-exceeded" ? error.code : "invalid-event-timeline-plan", error.message, `/eventTimeline/source${error.path}`, { cause: error });
  }
  if (!same(expected, plan)) throw new EventTimelineError("invalid-event-timeline-plan", "Bound timing source, owner/inventory and compiled intervals disagree.", "/eventTimeline/plan");
}
