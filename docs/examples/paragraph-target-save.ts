import { importPptx } from "@pptkit/pptx-parser";

export async function editStableParagraphTarget(bytes: Uint8Array, replacement: string) {
  const session = await importPptx(bytes, {
    projection: "content",
    editSaveProfile: "text-runs-v2",
    textTargetSaveProfile: "paragraph-stable-v1",
  });
  try {
    const slide = session.document.slides[0];
    const element = slide?.elements.find(e => e.type === "text");
    if (!slide || !element || element.type !== "text") throw new Error("No imported text box.");
    const updated = structuredClone(element);
    if (typeof updated.content === "string" || !updated.content[0]?.runs[0]) throw new Error("No existing first run.");
    updated.content[0].runs[0].text = replacement;
    slide.replaceElement(element.id, updated);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error(prepared.report.diagnostics.map(d => d.message).join("\n"));
    return await session.save(prepared.plan);
  } finally { session.dispose(); }
}
