import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type TextMeasurer } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

// Supply a synchronous provider backed by fonts loaded by your application.
export async function inspectAndExport(document: PresentationDocument, measure: TextMeasurer) {
  const layout = {
    textMeasurement: { measure },
    diagnostics: { textOverflow: true, slideBounds: true },
  };
  const resolved = resolveLayout(document, layout);
  const diagnostics = resolved.analysis?.diagnostics ?? [];
  // The application decides whether warnings should block its workflow.
  if (diagnostics.length > 0) return { diagnostics };
  return { diagnostics, generated: await generatePptx(document, { layout }) };
}
