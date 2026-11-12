const content = text => [{ runs: [{ text, style: { fontFamily: "Arial", fontSize: 9 } }] }];
const cell = (text, properties = {}) => ({ content: content(text), style: { margin: 0 }, ...properties });
export const tableEditBatches = [
  [{ type: "insert-row", index: 1, height: 25, cell: cell("Inserted row") }, { type: "insert-column", index: 1, width: 30, cell: cell("Inserted column") }],
  [{ type: "delete-row", index: 0 }, { type: "delete-column", index: 0 }],
  [{ type: "split", rowIndex: 1, columnIndex: 1 }, { type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 2 }, content: "append-paragraphs" }, { type: "insert-row", index: 3, height: 35, cell: cell("Tail") }],
];
export function tableEditSource() {
  return { type: "table", id: "table", name: "Editable table", box: { x: 40, y: 40, width: 300, height: 120 }, columns: [100, 100, 100], rows: [
    { height: 40, cells: [cell("Anchor A", { rowSpan: 2, colSpan: 2, style: { margin: 0, fill: { type: "solid", color: "DDEEFF" } } }), cell("Delete B")] },
    { height: 40, cells: [cell("Control C")] }, { height: 40, cells: [cell("Delete D"), cell("Control E"), cell("Control F")] },
  ] };
}
export function tableEditDeck(createPresentation) {
  const document = createPresentation({ size: { width: 480, height: 320 } }), slide = document.addSlide({ id: "edited" });
  slide.addElement(tableEditSource());
  document.addSlide({ id: "control" }).addElement({ type: "text", box: { x: 40, y: 40, width: 300, height: 60 }, content: "Unchanged table control" });
  return { document, slide };
}
