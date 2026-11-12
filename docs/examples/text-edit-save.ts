import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editFirstImportedRun(
  source: Uint8Array,
  text: string,
  signal?: AbortSignal,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v1",
    ...(signal ? { signal } : {}),
  });
  try {
    const slide = session.document.slides[0];
    const element = slide?.elements.find(item => item.type === "text");
    if (!slide || !element || element.type !== "text") {
      throw new Error("No imported top-level text box.");
    }
    const updated = structuredClone(element);
    if (typeof updated.content === "string" || !updated.content[0]?.runs[0]) {
      throw new Error("No imported first run.");
    }
    updated.content[0].runs[0].text = text;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({
      mode: "edit",
      ...(signal ? { signal } : {}),
    });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
