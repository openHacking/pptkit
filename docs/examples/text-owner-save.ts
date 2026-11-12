import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editShapeAndNotes(
  source: Uint8Array,
  shapeText: string,
  notesText: string,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v2",
  });
  try {
    const slide = session.document.slides[0];
    const element = slide?.elements.find(item => item.type === "shape" && item.text);
    if (!slide || !element || element.type !== "shape" || !element.text) {
      throw new Error("No imported top-level shape text.");
    }
    const shape = structuredClone(element);
    const notes = structuredClone(slide.notes);
    if (!shape.text || typeof shape.text.content === "string" || !Array.isArray(notes) ||
        !shape.text.content[0]?.runs[0] || !notes[0]?.runs[0]) {
      throw new Error("Both existing first runs are required.");
    }
    shape.text.content[0].runs[0].text = shapeText;
    notes[0].runs[0].text = notesText;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, shape);
    slide.setNotes(notes);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
