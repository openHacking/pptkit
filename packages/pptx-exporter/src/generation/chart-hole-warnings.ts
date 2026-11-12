import type { PresentationDocument, PresentationElement } from "@pptkit/core";
import type { ExportWarning } from "../types/export.js";

/** Office accepts 1–9, while older Transitional specifications state a minimum of 10. */
export function chartHoleWarnings(document: PresentationDocument): ExportWarning[] {
  const warnings: ExportWarning[] = [];
  const elements = (items: readonly PresentationElement[] | undefined, path: string, owner: Pick<ExportWarning, "masterId" | "layoutId" | "slideId">) => {
    items?.forEach((element, index) => {
      const current = `${path}.${index}`;
      if (element.type === "group") elements(element.children, `${current}.children`, owner);
      else if (element.type === "chart" && element.chartType === "doughnut" && element.holeSize !== undefined && element.holeSize < 10) {
        warnings.push({ code: "doughnut-small-hole-office-extension", message: "A 1–9% doughnut hole is an Office-compatible extension to the older Transitional minimum of 10%; other applications may not accept it.", path: `${current}.holeSize`, ...owner, elementId: element.id });
      }
    });
  };
  document.masters?.forEach((master, index) => elements(master.elements, `masters.${index}.elements`, { masterId: master.id }));
  document.layouts.forEach((layout, index) => elements(layout.elements, `layouts.${index}.elements`, { layoutId: layout.id }));
  document.slides.forEach((slide, index) => elements(slide.elements, `slides.${index}.elements`, { slideId: slide.id }));
  return warnings;
}
