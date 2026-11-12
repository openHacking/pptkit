import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function boundSlideTimingExample() {
  const document = createPresentation();
  const slide = document.addSlide({ id: "page" });
  slide.addElement({ id: "text", type: "text", box: { x: 30, y: 30, width: 300, height: 50 }, content: "Before" });
  slide.addElement({ id: "button", type: "shape", shape: "rect", box: { x: 30, y: 100, width: 100, height: 40 } });
  slide.setEventTimeline({ stages: [{
    id: "stage", target: { pageId: slide.id, elementId: "text" },
    trigger: { type: "object-click", elementId: "button" },
    timing: { kind: "span", id: "time", durationMs: 200 },
  }] });
  // The logical text ID stays the same; no temporary dangling target exists.
  slide.replaceElement("text", { type: "text", box: { x: 30, y: 30, width: 300, height: 50 }, content: "After" });
  const copy = document.duplicateSlide(slide.id);
  const normalized = normalizePresentation(document);
  const layout = resolveLayout(document);
  // Default PPTX generation rejects this unimplemented native mapping.
  // Explicit omission authorizes a static deck with located loss warnings.
  const staticPptx = await generatePptx(document, { eventTimelineMode: "omit" });
  return { document, copy, normalized, layout, staticPptx };
}
