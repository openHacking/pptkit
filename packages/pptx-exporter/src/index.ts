import type { PresentationDocument } from "@pptkit/core";
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
  const layout = resolveLayout(document);

  return {
    output: options.output,
    slideCount: layout.slideCount,
    status: "not-implemented",
  };
}

