import { createPresentation } from "@pptkit/core";
import { resolveLayout, type ConstraintsComposition } from "@pptkit/layout";

export function createConstrainedCards() {
  const document = createPresentation({ size: { width: 600, height: 300 } });
  const slide = document.addSlide({ id: "cards" });
  for (const id of ["left", "right"]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: 0, y: 30, width: 100, height: 80 } });
  const rule: ConstraintsComposition = {
    type: "constraints", container: { kind: "slide", id: slide.id }, padding: 20,
    elements: [
      { elementId: "left", x: { value: 20 }, width: { value: 180, min: 100, max: 200 } },
      { elementId: "right", x: { value: { terms: [{ source: { elementId: "left" }, metric: "right" }], constant: 20 } }, width: { value: { terms: [{ source: { elementId: "left" }, metric: "width" }] } } },
    ],
  };
  return resolveLayout(document, { composition: [rule] });
}
