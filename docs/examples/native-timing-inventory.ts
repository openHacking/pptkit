import { inspectPptxTiming, type PptxTimingInspectionOptions } from "@pptkit/pptx-parser";

export async function inventoryNativeTiming(bytes: Uint8Array, signal?: AbortSignal) {
  const options: PptxTimingInspectionOptions = {
    timingLimits: { maxRecords: 100_000, maxTextCharacters: 1_000_000 },
    ...(signal ? { signal } : {}),
  };
  const report = await inspectPptxTiming(bytes, options);
  return {
    report,
    targets: report.slides.flatMap(slide => slide.nodes
      .filter(node => node.drawingReference !== undefined)
      .map(node => ({ slideId: slide.slideId, source: node.source, reference: node.drawingReference }))),
  };
}
