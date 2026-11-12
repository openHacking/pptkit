import { importPptx } from "@pptkit/pptx-parser";

export async function removeSourceTitle(source: Uint8Array): Promise<Uint8Array> {
  const session = await importPptx(source, { projection: "content", metadataSaveProfile: "metadata-default-remove-v17" });
  try {
    const title = session.report.coreMetadataSources?.find(item => item.field === "title");
    if (!title?.removable || !session.document.setMetadata) throw new Error("Title source is not eligible for removal.");
    const metadata = { ...session.document.metadata };
    delete metadata.title;
    session.document.setMetadata(metadata);
    const prepared = await session.prepareSave({ mode: "edit", coreMetadataRemovals: ["title"] });
    if (prepared.status !== "ready") throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    return (await session.save(prepared.plan)).bytes;
  } finally {
    session.dispose();
  }
}
