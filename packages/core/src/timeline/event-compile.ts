import type { EventTimelineInput, EventTimelineLimits, EventTimelinePlan, EventTimelineStage } from "../types/event-timeline.js";
import type { FiniteTimelineNode } from "../types/timeline.js";
import { compileFiniteTimelineData } from "./compile.js";
import { captureEventProgram, EventCapture, eventError } from "./event-input.js";

/** Detached foreground event gates, not native effects or a document mutation. */
export function compileEventTimeline(input: EventTimelineInput, options?: EventTimelineLimits): EventTimelinePlan {
  try {
    const context = new EventCapture(options), program = captureEventProgram(input, context, false), stages: EventTimelineStage[] = [];
    for (let i = 0; i < program.stages.length; i++) {
      const stage = program.stages[i]!, path = `/stages/${i}/timing`, finiteOptions = context.finiteOptions(path);
      let compiled;
      try { compiled = compileFiniteTimelineData(stage.timing as FiniteTimelineNode, finiteOptions); } catch (error) { eventError(error, path); }
      context.nodes += compiled.plan.nodeCount; context.charge(compiled.idCharacters, path);
      stages.push(Object.freeze({ id: stage.id, target: stage.target, trigger: stage.trigger, timing: compiled.plan }));
    }
    return Object.freeze({ schemaVersion: "pptkit-event-timeline-v1", pageId: program.pageId, elementIds: program.elementIds, timingNodeCount: context.nodes, stages: Object.freeze(stages) });
  } catch (error) { return eventError(error); }
}
