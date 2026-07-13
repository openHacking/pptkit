import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import { createZip } from "../archive/create-zip.js";
import { buildPackage } from "../packaging/build-package.js";
import type { ExportWarning, GeneratePptxResult } from "../types/export.js";
import type { AssetLoader } from "../types/internal.js";

export async function generatePptxWith(
  document: PresentationDocument,
  loadAsset: AssetLoader,
): Promise<GeneratePptxResult> {
  const normalized = normalizePresentation(document);
  const layout = resolveNormalizedLayout(normalized);
  const warnings: ExportWarning[] = [];
  const bytes = createZip(await buildPackage(normalized.title, layout, normalized.assets, warnings, loadAsset));
  return {
    bytes,
    slideCount: layout.slideCount,
    byteLength: bytes.byteLength,
    warnings,
    status: warnings.length === 0 ? "generated" : "generated-with-warnings",
  };
}
