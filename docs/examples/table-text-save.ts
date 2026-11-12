import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editImportedTableRun(
  source: Uint8Array,
  rowIndex: number,
  cellIndex: number,
  text: string,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v3",
  });
  try {
    const slide = session.document.slides.find(item =>
      item.elements.some(element => element.type === "table"));
    const element = slide?.elements.find(item => item.type === "table");
    if (!slide || !element || element.type !== "table") {
      throw new Error("No mapped source table.");
    }
    const updated = structuredClone(element);
    const cell = updated.rows[rowIndex]?.cells[cellIndex];
    if (!cell || !Array.isArray(cell.content) || !cell.content[0]?.runs[0]) {
      throw new Error("No existing first run at that compact anchor.");
    }
    cell.content[0].runs[0].text = text;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
