import type { PresentationDocument } from "@pptkit/core";
import { loadNodeAsset } from "./assets/load-asset.node.js";
import { generatePptxWith } from "./generation/generate-pptx.js";
import { writeOutput } from "./output/write-output.js";
import type { GeneratePptxResult, WritePptxOptions, WritePptxResult } from "./types/export.js";

export type { ExportWarning, GeneratePptxResult, WritePptxOptions, WritePptxResult } from "./types/export.js";

export async function generatePptx(document: PresentationDocument): Promise<GeneratePptxResult> {
  return generatePptxWith(document, loadNodeAsset);
}

export async function writePptx(document: PresentationDocument, options: WritePptxOptions): Promise<WritePptxResult> {
  const generated = await generatePptx(document);
  const written = await writeOutput(options.output, generated.bytes);
  return {
    output: written.output,
    slideCount: generated.slideCount,
    byteLength: written.byteLength,
    warnings: generated.warnings,
    status: generated.warnings.length === 0 ? "written" : "written-with-warnings",
  };
}
