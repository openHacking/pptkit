import type { PresentationDocument, PresentationElement, TextContentInput } from "@pptkit/core";
import type { ExportWarning } from "../types/export.js";

/** A concrete native start preserves appearance but not the authoring dependency. */
export function listContinuationWarnings(document: PresentationDocument): ExportWarning[] {
  const warnings: ExportWarning[] = [];
  const inspect = (content: TextContentInput | undefined, path: string, owner: Pick<ExportWarning, "masterId" | "layoutId" | "slideId" | "elementId">) => {
    if (!Array.isArray(content)) return;
    content.forEach((paragraph, index) => {
      const bullet = paragraph.style?.bullet;
      if (bullet?.type === "number" && bullet.startAt === "continue") warnings.push({ code: "numbered-list-continuation-materialized", message: "The continuation is exported as an explicit numbering start; future native edits do not preserve the authoring dependency.", path: `${path}.${index}.style.bullet.startAt`, ...owner });
    });
  };
  const elements = (items: readonly PresentationElement[] | undefined, path: string, owner: Pick<ExportWarning, "masterId" | "layoutId" | "slideId">) => {
    items?.forEach((element, index) => {
      const current = `${path}.${index}`, identity = { ...owner, elementId: element.id };
      if (element.type === "text") inspect(element.content, `${current}.content`, identity);
      else if (element.type === "shape") inspect(element.text?.content, `${current}.text.content`, identity);
      else if (element.type === "table") element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => inspect(cell.content, `${current}.rows.${rowIndex}.cells.${cellIndex}.content`, identity)));
      else if (element.type === "group") elements(element.children, `${current}.children`, owner);
    });
  };
  document.masters?.forEach((master, index) => elements(master.elements, `masters.${index}.elements`, { masterId: master.id }));
  document.layouts.forEach((layout, index) => elements(layout.elements, `layouts.${index}.elements`, { layoutId: layout.id }));
  document.slides.forEach((slide, index) => {
    const owner = { slideId: slide.id };
    elements(slide.elements, `slides.${index}.elements`, owner); inspect(slide.notes, `slides.${index}.notes`, owner);
  });
  return warnings;
}
