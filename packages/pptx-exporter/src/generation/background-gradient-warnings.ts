import type { ShapeFillPaintInput, PresentationDocument, PresentationElement } from "@pptkit/core";
import type { ExportWarning } from "../types/export.js";

/** Report precision loss that collapses a nonzero gradient interval into a hard edge. */
export function backgroundGradientWarnings(document: PresentationDocument): ExportWarning[] {
  const warnings: ExportWarning[] = [];
  const inspect = (paint: ShapeFillPaintInput | undefined, path: string, owner: Pick<ExportWarning, "masterId" | "layoutId" | "slideId" | "elementId">, shape = false) => {
    if (paint?.type !== "linearGradient") return;
    for (let index = 1; index < paint.stops.length; index += 1) {
      const previous = paint.stops[index - 1]!, current = paint.stops[index]!;
      if (previous.position !== current.position && Math.round(previous.position * 100000) === Math.round(current.position * 100000)) {
        warnings.push({ code: shape ? "shape-gradient-stops-collapsed" : "background-gradient-stops-collapsed", message: "Distinct gradient stops round to the same PPTX position, collapsing their interval.", path: `${path}.stops.${index}.position`, ...owner });
      }
    }
  };
  const elements = (items: readonly PresentationElement[] | undefined, path: string, owner: Pick<ExportWarning, "masterId" | "layoutId" | "slideId">) => {
    items?.forEach((element, index) => {
      const current = `${path}.${index}`;
      if (element.type === "shape") inspect(element.style?.fill, `${current}.style.fill`, { ...owner, elementId: element.id }, true);
      else if (element.type === "group") elements(element.children, `${current}.children`, owner);
    });
  };
  document.masters?.forEach((master, index) => (inspect(master.background, `masters.${index}.background`, { masterId: master.id }), elements(master.elements, `masters.${index}.elements`, { masterId: master.id })));
  document.layouts.forEach((layout, index) => (inspect(layout.background, `layouts.${index}.background`, { layoutId: layout.id }), elements(layout.elements, `layouts.${index}.elements`, { layoutId: layout.id })));
  document.slides.forEach((slide, index) => (inspect(slide.background, `slides.${index}.background`, { slideId: slide.id }), elements(slide.elements, `slides.${index}.elements`, { slideId: slide.id })));
  return warnings;
}
