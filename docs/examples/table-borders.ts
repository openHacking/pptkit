import { createPresentation, normalizePresentation, type TableCellEdge } from "@pptkit/core";
import { resolveLayout, resolveTableBorders, type TableBorderLimits } from "@pptkit/layout";

export function directTableBorders() {
  const document = createPresentation();
  const edge: TableCellEdge = "right";
  document.addSlide().addElement({
    type: "table", box: { x: 30, y: 30, width: 200, height: 40 }, columns: [100, 100],
    rows: [{ height: 40, cells: [
      { content: "Explicit edge", style: { stroke: { width: 1 }, borders: { [edge]: { width: 3, paint: { type: "solid", color: "FF0000" } } } } },
      { content: "Uniform fallback", style: { stroke: { width: 2, paint: { type: "solid", color: "0000FF" } } } },
    ] }],
  });
  const limits: TableBorderLimits = { maxGridSlots: 1000, maxEdges: 3000, maxWorkUnits: 20000 };
  const normalized = normalizePresentation(document), table = normalized.slides[0]!.elements[0]!;
  if (table.type !== "table") throw new Error("Expected table.");
  return { document, standalone: resolveTableBorders(table, limits), layout: resolveLayout(document, { tableBorderLimits: limits }) };
}
