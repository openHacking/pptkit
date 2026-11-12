import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type LayoutOptions, type TextMeasurementOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function fitAndExport(document: PresentationDocument, textMeasurement: TextMeasurementOptions) {
  const options: LayoutOptions = {
    textMeasurement,
    intrinsicSizing: [
      { elementId: "title", width: { type: "content", max: 400 }, height: { type: "content" } },
      { elementId: "photo", height: { type: "content" } },
    ],
    diagnostics: { textOverflow: true },
  };
  const resolved = resolveLayout(document, options);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return { generated: await generatePptx(document, { layout: options }) };
}
