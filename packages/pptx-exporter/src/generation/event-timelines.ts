import type { NormalizedPresentation } from "@pptkit/core";
import { PptxExportError, type ExportWarning } from "../types/export.js";

/** Mutates only the freshly normalized private output candidate on explicit omission. */
export function prepareEventTimelineOutput(document: NormalizedPresentation, mode: "reject" | "omit"): ExportWarning[] {
  const diagnostics: ExportWarning[] = [];
  for (let i = 0; i < document.slides.length; i++) {
    const slide = document.slides[i]!, timeline = slide.eventTimeline;
    if (timeline === undefined) continue;
    const base = { code: mode === "omit" ? "event-timeline-omitted" : "event-timeline-export-unsupported", slideId: slide.id, path: `slides.${i}.eventTimeline`, message: mode === "omit" ? "Explicit static output omits authored foreground targets, triggers and finite timing semantics; native timing output is unavailable." : "Native foreground timing output is unavailable; generation was blocked before output. Explicit omit would discard its targets, triggers and finite timing semantics." };
    if (!timeline.plan.stages.length) diagnostics.push(base);
    else for (let j = 0; j < timeline.plan.stages.length; j++) {
      const stage = timeline.plan.stages[j]!;
      diagnostics.push({ ...base, elementId: stage.target.elementId, stageId: stage.id, path: `${base.path}.stages.${j}` });
    }
  }
  if (diagnostics.length && mode !== "omit") throw new PptxExportError(diagnostics);
  if (mode === "omit") for (const slide of document.slides) delete slide.eventTimeline;
  return diagnostics;
}
