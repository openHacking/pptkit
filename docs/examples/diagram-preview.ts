import { importPptx } from "@pptkit/pptx-parser";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function previewCachedDiagrams(bytes: Uint8Array) {
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
  try {
    return { rendering: await renderPresentationToSvg(session.document), diagrams: session.report.diagramPreviews };
  } finally {
    session.dispose();
  }
}
