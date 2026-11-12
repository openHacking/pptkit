import { createPresentation, resolvePresentationInheritance } from "@pptkit/core";

export function tableFormatOrigins() {
  const document = createPresentation();
  document.addSlide({ elements: [{ id: "table", type: "table", box: { x: 10, y: 10, width: 200, height: 60 }, columns: [100, 100], rows: [{ cells: [{ content: "Header" }] }, { cells: [] }], style: {
    wholeTable: { text: { run: { bold: true, scriptFontFamilies: { Hans: "Arial" } } } },
    lastRow: { fill: { type: "solid", color: "FFFFFF", opacity: 0 }, text: { run: { scriptFontFamilies: null } } },
  } }] });
  const resolution = resolvePresentationInheritance(document);
  return resolution.origins.filter(origin => origin.target.kind === "element" && origin.target.id === "table" && origin.property.startsWith("rows."));
}
