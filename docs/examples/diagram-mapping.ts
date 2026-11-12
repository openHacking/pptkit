import { inspectPptx, type PptxDiagramPresentationMapping } from "@pptkit/pptx-parser";

export async function traceFirstCachedDiagram(bytes: Uint8Array): Promise<readonly {
  cacheElementIndex: number;
  dataElementIndex: number | undefined;
  presentation: PptxDiagramPresentationMapping | undefined;
}[]> {
  const inventory = (await inspectPptx(bytes)).diagramInventory;
  const cache = inventory?.cacheMappings?.[0];
  if (!cache) return [];
  const model = inventory?.dataModels?.find(item => item.part === cache.dataPart);
  const presentations = new Map(model?.presentationMappings?.map(item => [item.elementIndex, item]));
  return cache.shapes.map(shape => ({
    cacheElementIndex: shape.elementIndex,
    dataElementIndex: shape.model.targetElementIndex,
    presentation: shape.model.targetElementIndex === undefined ? undefined : presentations.get(shape.model.targetElementIndex),
  }));
}
