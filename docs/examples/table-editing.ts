import { createPresentation, editTable, type TableElementInput } from "@pptkit/core";

export function editMergedTable() {
  const document = createPresentation();
  const slide = document.addSlide();
  const input: TableElementInput = {
    id: "table", type: "table", box: { x: 30, y: 30, width: 240, height: 60 }, columns: [120, 120],
    rows: [{ height: 30, cells: [{ content: "Anchor", rowSpan: 2 }, { content: "First" }] }, { height: 30, cells: [{ content: "Second" }] }],
  };
  const table = slide.addElement(input);
  if (table.type !== "table") throw new Error("Expected a table.");
  const result = editTable(table, [
    { type: "insert-row", index: 1, height: 24 },
    { type: "insert-column", index: 2, width: 60 },
    { type: "split", rowIndex: 1, columnIndex: 0 },
  ], { frame: "fit-grid" });
  slide.replaceElement(table.id, result.table);
  return { document, mapping: result.cellMapping };
}
