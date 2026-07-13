import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";

export interface ExportPptxOptions {
  output: string;
}

export interface ExportAttempt {
  output: string;
  slideCount: number;
  status: "not-implemented";
}

export async function exportPptx(
  document: PresentationDocument,
  options: ExportPptxOptions,
): Promise<ExportAttempt> {
  const normalized = normalizePresentation(document);
  const layout = resolveLayout(document);

  return {
    output: options.output,
    slideCount: Math.min(layout.slideCount, normalized.slides.length),
    status: "not-implemented",
  };
}
