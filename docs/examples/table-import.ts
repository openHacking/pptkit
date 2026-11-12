import { importPptx } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function exportProjectedTableEdit(source: Uint8Array) {
  const session = await importPptx(source, { projection: "content" });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "table"));
    if (!slide) throw new Error("No table in the supported content projection.");
    const index = slide.elements.findIndex(e => e.type === "table");
    const table = slide.elements[index]!;
    if (table.type !== "table") throw new Error("Expected a table.");
    const rows = table.rows.map((row, ri) => ({
      ...row,
      cells: row.cells.map((cell, ci) => ri === 0 && ci === 0
        ? { ...cell, content: "Edited anchor text" }
        : cell),
    }));
    slide.removeElement(table.id);
    slide.insertElement(index, { ...table, rows });
    // This exports the supported Core projection. Unsupported source content
    // is not included, so callers must review the import report first.
    const fresh = await generatePptx(session.document);
    const preservation = await session.prepareSave({ mode: "edit" });
    return { fresh, importReport: session.report, preservation };
  } finally {
    session.dispose();
  }
}
