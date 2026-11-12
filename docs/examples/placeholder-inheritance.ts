import { normalizePresentation } from "@pptkit/core";
import { importPptx, type PptxPlaceholderInfo } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function projectPlaceholderDimensions(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    sessionLimits: { maxPlaceholderEntries: 100_000 },
  });
  try {
    const placeholders: readonly PptxPlaceholderInfo[] =
      session.report.placeholderInheritance?.placeholders ?? [];
    const widths = placeholders.map(ph => ({
      drawing: ph.source,
      parent: ph.parent,
      width: ph.dimensions.width,
      status: ph.geometryStatus,
    }));
    // Fresh output is a new local-value document, not a source template save.
    const fresh = await generatePptx(session.document);
    const original = await session.prepareSave({ mode: "no-op" });
    if (original.status !== "ready") throw new Error("Original-byte save is unavailable.");
    const saved = await session.save(original.plan);
    return {
      document: normalizePresentation(session.document),
      widths,
      freshBytes: fresh.bytes,
      originalBytes: saved.bytes,
    };
  } finally {
    session.dispose();
  }
}
