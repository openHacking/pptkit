import { createPresentation, normalizePresentation, type TableStyleInput } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createRegionalTable() {
  const style: TableStyleInput = {
    wholeTable: { margin: 3, text: { run: { fontFamily: { theme: "body" }, fontSize: 12 } } },
    bandedRows: {
      first: { fill: { type: "solid", color: { theme: "accent1" }, opacity: 0.15 } },
      second: { fill: { type: "none" } },
    },
    firstRow: { fill: { type: "solid", color: { theme: "accent1" } }, text: { run: { bold: true, color: "FFFFFF" } } },
    borders: { outer: { width: 2, paint: { type: "solid", color: "444444" } }, innerHorizontal: { width: 0.5, paint: { type: "solid", color: "AAAAAA" } } },
  };
  const document = createPresentation();
  document.addSlide().addElement({
    type: "table", style, box: { x: 30, y: 30, width: 240, height: 90 }, columns: [120, 120],
    rows: [
      { height: 30, cells: [{ content: "Item" }, { content: "Value" }] },
      { height: 30, cells: [{ content: "First" }, { content: "10" }] },
      { height: 30, cells: [{ content: "Second" }, { content: "20", style: { fill: { type: "solid", color: "FFCC66" } } }] },
    ],
  });
  return { document, effective: normalizePresentation(document), output: await generatePptx(document) };
}
