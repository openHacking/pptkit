import type { EventTimelinePlan, EventTimelineStageInput, SlideEventTimelineInput } from "../types/event-timeline.js";
import { EventTimelineError } from "../types/event-timeline.js";
import type { PresentationElement } from "../types/element.js";
import { compileEventTimeline } from "../timeline/event-compile.js";
import { record } from "../timeline/input.js";
import { eventError } from "../timeline/event-input.js";
import { deepClone, deepFreeze } from "../utils/clone.js";

export function slideElementIds(elements: readonly PresentationElement[]): string[] {
  if (elements.length > 10000) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timelines admit at most 10000 owned objects.", "/elementIds");
  const ids: string[] = [], pending = [...elements]; let characters = 0;
  while (pending.length) {
    const element = pending.pop()!;
    if (ids.length >= 10000) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timelines admit at most 10000 owned objects.", "/elementIds");
    if (typeof element.id !== "string" || !element.id.length) throw new EventTimelineError("invalid-event-timeline-input", "Owned object identities must be nonempty strings.", "/elementIds");
    characters += element.id.length;
    if (element.id.length > 256 || characters > 1000000) throw new EventTimelineError("event-timeline-limit-exceeded", "Owned object identities exceed the slide timing identifier budget.", "/elementIds");
    ids.push(element.id);
    if (element.type === "group") {
      if (element.children.length > 10000 - ids.length - pending.length) throw new EventTimelineError("event-timeline-limit-exceeded", "Slide event timelines admit at most 10000 owned objects.", "/elementIds");
      for (const child of element.children) pending.push(child);
    }
  }
  return ids.sort();
}
export function compileSlideEventTimeline(pageId: string, elements: readonly PresentationElement[], input: SlideEventTimelineInput): EventTimelinePlan {
  try {
    const data = record(input, ["stages"], "");
    return compileEventTimeline({ pageId, elementIds: slideElementIds(elements), stages: data.stages as readonly EventTimelineStageInput[] });
  } catch (error) { return eventError(error); }
}
export function captureSlideEventTimeline(pageId: string, elements: readonly PresentationElement[], input: SlideEventTimelineInput): SlideEventTimelineInput {
  compileSlideEventTimeline(pageId, elements, input);
  return deepFreeze(deepClone(input));
}
export function remapSlideEventTimeline(input: SlideEventTimelineInput, pageId: string, elementIds: ReadonlyMap<string, string>): SlideEventTimelineInput {
  return { stages: input.stages.map(stage => ({ ...deepClone(stage), target: { pageId, elementId: elementIds.get(stage.target.elementId)! },
    trigger: stage.trigger.type === "object-click" ? { type: "object-click", elementId: elementIds.get(stage.trigger.elementId)! } : { ...stage.trigger } })) };
}
