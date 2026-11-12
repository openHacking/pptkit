import { inspectPptx, type PptxDiagramDataModelInfo } from "@pptkit/pptx-parser";

export async function readDiagramModels(bytes: Uint8Array): Promise<readonly PptxDiagramDataModelInfo[]> {
  const inspection = await inspectPptx(bytes, {
    diagramLimits: { maxWorkUnits: 100_000, maxFindings: 1_000 },
  });
  return inspection.diagramInventory?.dataModels ?? [];
}
