import type { PresentationDocument } from "@pptkit/core";
import { loadAsset } from "./assets/load-asset.js";
import { generatePptxWith } from "./generation/generate-pptx.js";
import type { GeneratePptxResult } from "./types/export.js";

export type { ExportWarning, GeneratePptxResult } from "./types/export.js";

export async function generatePptx(document: PresentationDocument): Promise<GeneratePptxResult> {
  return generatePptxWith(document, loadAsset);
}
