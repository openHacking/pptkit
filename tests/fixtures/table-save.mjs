import { createTextOwnerFixture } from "./text-owners.mjs";

/** Mixed source fixture with compact merged anchors and unrelated owners. */
export function createTableSaveFixture(createPresentation, imageBytes) {
  const document = createTextOwnerFixture(createPresentation, imageBytes);
  document.slides[0].addElement({
    type: "table", name: "Merged source table", columns: [70, 70, 90],
    box: { x: 370, y: 270, width: 230, height: 90 },
    rows: [
      { height: 30, cells: [{ content: [{ runs: [{ text: "Merged anchor control", style: { bold: true } }] }], rowSpan: 2, colSpan: 2 }, { content: "Top right control" }] },
      { height: 30, cells: [{ content: "Bottom right source" }] },
      { height: 30, cells: [{ content: "Last row source" }, { content: "" }, { content: "Last cell control" }] },
    ],
  });
  return document;
}
