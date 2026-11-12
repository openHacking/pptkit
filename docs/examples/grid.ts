import { createPresentation } from "@pptkit/core";
import { resolveLayout, type GridComposition } from "@pptkit/layout";

export function createGrid() {
  const document = createPresentation({ size: { width: 600, height: 300 } });
  const slide = document.addSlide({ id: "slide" });
  for (const id of ["label", "content"]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: 0, y: 0, width: 100, height: 60 } });
  const grid: GridComposition = {
    type: "grid", container: { kind: "slide", id: slide.id }, padding: 20,
    columns: [{ type: "content", max: 140 }, { type: "fraction" }],
    rows: [{ type: "fraction" }], columnGap: 20,
    placements: [
      { elementId: "label", row: 0, column: 0, vertical: "center" },
      { elementId: "content", row: 0, column: 1, horizontal: "center", vertical: "center" },
    ],
  };
  return resolveLayout(document, { composition: [grid] });
}
