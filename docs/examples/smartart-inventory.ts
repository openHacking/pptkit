import { inspectPptx, type PptxDiagramInfo } from "@pptkit/pptx-parser";

export async function listDiagrams(bytes: Uint8Array): Promise<readonly { part: string; diagram: PptxDiagramInfo }[]> {
  const inspection = await inspectPptx(bytes);
  return inspection.slides.flatMap(slide => (slide.diagrams ?? []).map(diagram => ({ part: slide.part, diagram })));
}
