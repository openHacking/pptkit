import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type LayoutOptions, type TextMeasurer } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderLayoutToSvg } from "@pptkit/svg-renderer";

export async function fitTable(document: PresentationDocument, measure: TextMeasurer) {
  const layout: LayoutOptions = {
    textMeasurement: { measure },
    tableSizing: [{
      elementId: "report-table",
      autoRows: [{ row: 1, minHeight: 24, maxHeight: 120 }, { row: 2, minHeight: 24 }],
      overflow: "preserve",
    }],
  };
  const resolved = resolveLayout(document, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return {
    preview: await renderLayoutToSvg(resolved),
    pptx: await generatePptx(document, { layout }),
  };
}
