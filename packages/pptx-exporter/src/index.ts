import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import { createZip } from "./archive/create-zip.js";
import { writeOutput } from "./output/write-output.js";
import { buildPackage } from "./packaging/build-package.js";
import type { ExportPptxOptions, ExportResult, ExportWarning } from "./types/export.js";

export type { ExportPptxOptions, ExportResult, ExportWarning } from "./types/export.js";

export async function exportPptx(document: PresentationDocument, options: ExportPptxOptions): Promise<ExportResult> {
  const normalized = normalizePresentation(document);
  const layout = resolveNormalizedLayout(normalized);
  const warnings: ExportWarning[] = [];
  const bytes = createZip(await buildPackage(normalized.title, layout, normalized.assets, warnings));
  const written = await writeOutput(options.output, bytes);
  return {
    output: written.output,
    slideCount: layout.slideCount,
    byteLength: written.byteLength,
    warnings,
    status: warnings.length === 0 ? "written" : "written-with-warnings",
  };
}
