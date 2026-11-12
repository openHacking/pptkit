import { createPresentation, prepareLayoutRebind, normalizePresentation, type LayoutRebindPlan, type PresentationData } from "@pptkit/core";
const document = createPresentation();
for (const [id, x] of [["source", 20], ["target", 40]] as const) document.defineSlideLayout({ id, name: id, placeholders: [{ key: id, kind: "body", index: 0, box: { x, y: 20, width: 300, height: 120 } }] });
const slide = document.addSlide({ layoutId: "source", elements: [{ id: "body", type: "text", placeholderKey: "source", content: "Retained" }] });
const values: PresentationData = document;
const plan: LayoutRebindPlan = prepareLayoutRebind(values, [{ slideId: slide.id, layoutId: "target" }], { localAppearance: "adopt", reset: ["geometry", "text"] });
if (plan.status === "ready") document.applyLayoutRebind(plan);
export const normalized = normalizePresentation(values);
export const findings = plan.findings;
