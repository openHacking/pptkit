import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function replaceFirstImportedPicture(
  source: Uint8Array,
  image: Uint8Array,
  mimeType: "image/png" | "image/jpeg" | "image/gif",
  width: number,
  height: number,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, { projection: "content", pictureEditSaveProfile: "replace-v1" });
  try {
    const slide = session.document.slides[0];
    const picture = slide?.elements.find(item => item.type === "image");
    if (!slide || !picture || picture.type !== "image") throw new Error("No imported top-level picture.");
    session.document.registerAsset({ id: "replacement", kind: "image", source: { type: "bytes", value: image }, mimeType, width, height });
    slide.replaceElement(picture.id, { ...picture, assetId: "replacement" });
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
