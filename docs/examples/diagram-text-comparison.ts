import { inspectPptx, type PptxDiagramTextComparison } from "@pptkit/pptx-parser";

export async function compareDirectDiagramText(bytes: Uint8Array): Promise<readonly PptxDiagramTextComparison[]> {
  const report = await inspectPptx(bytes, { diagramTextComparison: "direct-paragraphs" });
  return report.diagramInventory?.textComparisons ?? [];
}
